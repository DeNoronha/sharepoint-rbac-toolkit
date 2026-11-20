<#
.SYNOPSIS
    Brings the permissions in SharePoint in line with the access matrix.
    Dry run by default; nothing changes without -Apply.

.DESCRIPTION
    Uses the same comparison as 02-Audit-Permissions.ps1 and then:

      ADD                 grants the missing group its permission level
      FIX_LEVEL           replaces a wrong permission level with the right one
      BREAK_INHERITANCE   gives the library its own permissions (existing ones are copied
                          first, then cleaned up)
      REMOVE_*            removes groups, people, sharing links and SharePoint groups that
                          the matrix does not allow

    Groups are added before anything is removed. REVIEW and KEEP rows are never touched:
    site owners, site visitors and permissions granted on the site itself are reported only.

    Limited Access entries are left alone. They belong to files and folders that were
    shared individually; use 06-Audit-ItemShares.ps1 for those.

.PARAMETER ConfigPath
    Path to config.json or to the folder that contains it.

.PARAMETER Target
    Optional. One or more target names. Start with one target when you apply for the first time.

.PARAMETER KeepDirectUsers
    Leave people with direct access in place.

.PARAMETER KeepSharePointGroups
    Leave SharePoint groups on libraries in place.

.PARAMETER AddOnly
    Only add and fix; remove nothing.

.PARAMETER Apply
    Make the changes. Without this switch the script only shows what it would do.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/03-Sync-Permissions.ps1 -ConfigPath ./config/contoso
    pwsh -NoProfile -File ./scripts/03-Sync-Permissions.ps1 -ConfigPath ./config/contoso -Target Finance -Apply

.NOTES
    Needs: PnP.PowerShell, a fresh snapshot from 01-Export-EntraSnapshot.ps1, and
    Full Control on the sites (site collection admin or SharePoint Administrator).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [string[]]$Target,
    [switch]$KeepDirectUsers,
    [switch]$KeepSharePointGroups,
    [switch]$AddOnly,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
$cfg = Get-SpRbacConfig -Path $ConfigPath
Write-SpRbacHeader -Title 'SYNC PERMISSIONS WITH THE MATRIX' -Config $cfg -Apply:$Apply
Assert-SpRbacModule -Name 'PnP.PowerShell'
$index = Get-SpRbacGroupIndex -Config $cfg -MaxAgeHours 4
Write-Host "  Entra snapshot: $(Split-Path $index.Path -Leaf) ($($index.AgeHours) h old)`n" -ForegroundColor Gray

# pwsh -File passes "A,B" as one string
$Target = @($Target | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$targets = @($cfg.Targets | Where-Object { $_.Scope -ne 'None' })
if ($Target) {
    $unknown = @($Target | Where-Object { $_ -notin $cfg.Targets.Name })
    if ($unknown) { throw "Unknown target(s): $($unknown -join ', '). Valid names: $($cfg.Targets.Name -join ', ')" }
    $targets = @($targets | Where-Object { $_.Name -in $Target })
}

$order = @{ BREAK_INHERITANCE = 0; ADD = 1; FIX_LEVEL = 2 }   # everything else (REMOVE_*) comes last
$log = [System.Collections.Generic.List[object]]::new()
$planned = 0; $done = 0; $errors = 0
$diffArgs = @{ Config = $cfg; GroupIndex = $index; KeepDirectUsers = $KeepDirectUsers; KeepSharePointGroups = $KeepSharePointGroups }

foreach ($group in ($targets | Group-Object Site)) {
    Write-Host "Site $($group.Name)" -ForegroundColor Cyan
    try { Connect-SpRbacSite -Config $cfg -SiteUrl $group.Name }
    catch { Write-Host "  ERROR: cannot connect: $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }

    foreach ($t in $group.Group) {
        try { $rows = @(Get-SpRbacTargetDiff @diffArgs -Target $t) }
        catch { Write-Host "  [$($t.Name)] ERROR: $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }

        $blocking = @($rows | Where-Object Action -in 'MISSING_GROUP', 'MISSING_TARGET')
        if ($blocking) {
            foreach ($b in $blocking) { Write-Host "  [$($t.Name)] $($b.Action): $($b.Principal) $($b.Note)" -ForegroundColor Red; $log.Add(($b | Select-Object * -ExcludeProperty Handle | Add-Member Status 'blocked' -PassThru)) }
            Write-Host "  [$($t.Name)] skipped: solve the missing item first, so no access is removed on incomplete information." -ForegroundColor Red
            $errors++; continue
        }

        # A library that still inherits gets its own permissions first; then compare again
        $break = $rows | Where-Object Action -eq 'BREAK_INHERITANCE' | Select-Object -First 1
        if ($break) {
            $planned++
            Write-Host "  [$($t.Name)] BREAK_INHERITANCE  (existing permissions are copied, then cleaned up)" -ForegroundColor Yellow
            if ($Apply) {
                try { Invoke-SpRbacChange -Row $break | Out-Null; $done++; $rows = @(Get-SpRbacTargetDiff @diffArgs -Target $t) }
                catch { Write-Host "      ERROR: $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }
            }
            $log.Add(($break | Select-Object * -ExcludeProperty Handle | Add-Member Status $(if ($Apply) { 'applied' } else { 'dry run' }) -PassThru))
        }

        $changes = @($rows | Where-Object { $_.Action -eq 'ADD' -or $_.Action -eq 'FIX_LEVEL' -or $_.Action -like 'REMOVE_*' })
        if ($AddOnly) { $changes = @($changes | Where-Object { $_.Action -notlike 'REMOVE_*' }) }
        $changes = @($changes | Sort-Object { if ($order.ContainsKey($_.Action)) { $order[$_.Action] } else { 9 } }, Principal)
        $review = @($rows | Where-Object Action -eq 'REVIEW')

        if (-not $changes.Count -and -not $break) { Write-Host "  [$($t.Name)] in line with the matrix$(if ($review.Count) { " ($($review.Count) to review, see 02-Audit-Permissions.ps1)" })" -ForegroundColor Green; continue }
        if ($changes.Count) { Write-Host "  [$($t.Name)] $($changes.Count) change(s)" -ForegroundColor Yellow }

        foreach ($c in $changes) {
            $planned++
            $lvl = if ($c.Action -in 'ADD', 'FIX_LEVEL') { " -> $($c.ExpectedLevel)" } else { '' }
            Write-Host ("      {0,-18} {1}  [{2}{3}]  {4}" -f $c.Action, $c.Principal, $c.CurrentLevel, $lvl, $c.Container) -ForegroundColor Yellow
            $status = 'dry run'
            if ($Apply) {
                try { Invoke-SpRbacChange -Row $c | Out-Null; $done++; $status = 'applied'; Write-Host '         done' -ForegroundColor Green }
                catch { $errors++; $status = "ERROR: $($_.Exception.Message)"; Write-Host "         $status" -ForegroundColor Red }
            }
            $log.Add(($c | Select-Object * -ExcludeProperty Handle | Add-Member Status $status -PassThru))
        }
    }
}

$file = New-SpRbacOutputPath -Config $cfg -Name 'permissions-sync'
$log | Export-Csv -LiteralPath $file -Encoding utf8
Write-SpRbacFooter -Apply:$Apply -Planned $planned -Done $done -Errors $errors
Write-Host "Log: $file" -ForegroundColor Gray
if ($Apply) { Write-Host 'Next: run 02-Audit-Permissions.ps1 to confirm the result.' -ForegroundColor Gray }
if ($errors) { exit 1 }
