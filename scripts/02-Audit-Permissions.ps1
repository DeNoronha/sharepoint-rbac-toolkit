<#
.SYNOPSIS
    Compares the permissions in SharePoint with the access matrix. Read-only.

.DESCRIPTION
    For every target in the config (a library or a site) this script reads who has access
    and compares that with what the matrix says.

    Library target  The role groups from the matrix must sit directly on the library, at
                    the right permission level, and nothing else.
    Site target     The access group(s) from the matrix must be in the site's Members
                    group and the admin groups in the Owners group.

    The result is a table on screen and a CSV in the output folder. Each row has an action:

      OK                  as the matrix says
      ADD                 expected group is missing
      FIX_LEVEL           expected group has the wrong permission level
      BREAK_INHERITANCE   library still inherits its permissions from the site
      REMOVE_GROUP / REMOVE_USER / REMOVE_SHARINGLINK / REMOVE_SPGROUP / REMOVE_OTHER
                          access that the matrix does not allow
      KEEP                left alone on request
      REVIEW              outside what the toolkit manages; a person has to judge it
      MISSING_GROUP       expected group does not exist in Entra (or the snapshot is old)
      MISSING_TARGET      library or site group not found

    03-Sync-Permissions.ps1 applies the ADD, FIX_LEVEL, BREAK_INHERITANCE and REMOVE rows.

.PARAMETER ConfigPath
    Path to config.json or to the folder that contains it.

.PARAMETER Target
    Optional. One or more target names (matrix columns). Default: all targets.

.PARAMETER KeepDirectUsers
    Report people with direct access as KEEP instead of REMOVE_USER.

.PARAMETER KeepSharePointGroups
    Report SharePoint groups on a library as KEEP instead of REMOVE_SPGROUP.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/02-Audit-Permissions.ps1 -ConfigPath ./config/contoso

.NOTES
    Needs: PnP.PowerShell, and a fresh snapshot from 01-Export-EntraSnapshot.ps1.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [string[]]$Target,
    [switch]$KeepDirectUsers,
    [switch]$KeepSharePointGroups
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
$cfg = Get-SpRbacConfig -Path $ConfigPath
Write-SpRbacHeader -Title 'AUDIT PERMISSIONS AGAINST THE MATRIX' -Config $cfg -ReadOnly
Assert-SpRbacModule -Name 'PnP.PowerShell'
$index = Get-SpRbacGroupIndex -Config $cfg
Write-Host "  Entra snapshot: $(Split-Path $index.Path -Leaf) ($($index.AgeHours) h old)`n" -ForegroundColor Gray

# pwsh -File passes "A,B" as one string
$Target = @($Target | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$targets = @($cfg.Targets | Where-Object { $_.Scope -ne 'None' })
if ($Target) {
    $unknown = @($Target | Where-Object { $_ -notin $cfg.Targets.Name })
    if ($unknown) { throw "Unknown target(s): $($unknown -join ', '). Valid names: $($cfg.Targets.Name -join ', ')" }
    $targets = @($targets | Where-Object { $_.Name -in $Target })
}
$skipped = @($cfg.Targets | Where-Object { $_.Scope -eq 'None' -and (-not $Target -or $_.Name -in $Target) })

$all = [System.Collections.Generic.List[object]]::new()
$errors = 0
foreach ($group in ($targets | Group-Object Site)) {
    Write-Host "Site $($group.Name)" -ForegroundColor Cyan
    try { Connect-SpRbacSite -Config $cfg -SiteUrl $group.Name }
    catch { Write-Host "  ERROR: cannot connect: $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }
    foreach ($t in $group.Group) {
        try {
            $rows = @(Get-SpRbacTargetDiff -Config $cfg -Target $t -GroupIndex $index -KeepDirectUsers:$KeepDirectUsers -KeepSharePointGroups:$KeepSharePointGroups)
        }
        catch { Write-Host "  [$($t.Name)] ERROR: $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }
        $rows | ForEach-Object { $all.Add($_) }
        $ok = @($rows | Where-Object Action -eq 'OK').Count
        $drift = @($rows | Where-Object { $_.Action -notin 'OK', 'KEEP', 'REVIEW' })
        $review = @($rows | Where-Object Action -eq 'REVIEW')
        $color = if ($drift.Count) { 'Yellow' } else { 'Green' }
        Write-Host ("  [{0}] ok: {1}   to change: {2}   to review: {3}" -f $t.Name, $ok, $drift.Count, $review.Count) -ForegroundColor $color
        foreach ($r in $drift + $review) {
            $lvl = if ($r.Action -in 'ADD', 'FIX_LEVEL') { " -> $($r.ExpectedLevel)" } else { '' }
            Write-Host ("      {0,-18} {1}  [{2}{3}]  {4}" -f $r.Action, $r.Principal, $r.CurrentLevel, $lvl, $r.Container) -ForegroundColor $(if ($r.Action -eq 'REVIEW') { 'DarkYellow' } else { 'Yellow' })
        }
    }
}
foreach ($t in $skipped) { Write-Host "  [$($t.Name)] has no site or library in the config: nothing to audit in SharePoint." -ForegroundColor DarkGray }

$file = New-SpRbacOutputPath -Config $cfg -Name 'permissions-audit'
$all | Select-Object * -ExcludeProperty Handle | Export-Csv -LiteralPath $file -Encoding utf8

$summary = $all | Group-Object Action | Sort-Object Name | ForEach-Object { "{0}: {1}" -f $_.Name, $_.Count }
$toChange = @($all | Where-Object { $_.Action -notin 'OK', 'KEEP', 'REVIEW' }).Count
Write-Host "`n$($summary -join '   ')"
if ($errors) { Write-Host "$errors target(s) or site(s) could not be read. See the errors above." -ForegroundColor Red }
if ($errors) { Write-Host "RESULT: INCOMPLETE. $toChange difference(s) found so far; fix the errors and run again." -ForegroundColor Red }
elseif ($toChange) { Write-Host "RESULT: $toChange difference(s) with the matrix." -ForegroundColor Yellow }
else { Write-Host 'RESULT: SharePoint matches the matrix.' -ForegroundColor Green }
Write-Host "Saved: $file" -ForegroundColor Gray
if ($errors) { exit 1 }
