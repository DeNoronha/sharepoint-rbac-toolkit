<#
.SYNOPSIS
    Fixes group names that SharePoint still shows under an old name after the group was
    renamed in Entra ID. Dry run by default; nothing changes without -Apply.

.DESCRIPTION
    SharePoint keeps its own copy of the display name of every group it has seen. When you
    rename a group in Entra, sites keep showing the old name, sometimes for months. Access
    is not affected (SharePoint works with the object id), but it confuses everyone who
    looks at the permissions page.

    For every site in the config this script compares the name SharePoint shows with the
    current name in the Entra snapshot, matched on object id, and corrects the difference.
    Only the display name changes; no permissions.

.PARAMETER ConfigPath
    Path to config.json or to the folder that contains it.

.PARAMETER Apply
    Make the changes. Without this switch the script only shows what it would do.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/10-Update-GroupTitles.ps1 -ConfigPath ./config/contoso

.NOTES
    Needs: PnP.PowerShell, a fresh snapshot from 01-Export-EntraSnapshot.ps1, Full Control
    on the sites.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
$cfg = Get-SpRbacConfig -Path $ConfigPath
Write-SpRbacHeader -Title 'UPDATE STALE GROUP NAMES IN SHAREPOINT' -Config $cfg -Apply:$Apply
Assert-SpRbacModule -Name 'PnP.PowerShell'
$index = Get-SpRbacGroupIndex -Config $cfg -MaxAgeHours 4

$log = [System.Collections.Generic.List[object]]::new()
$planned = 0; $done = 0; $errors = 0
foreach ($site in (Get-SpRbacSites -Config $cfg)) {
    try { Connect-SpRbacSite -Config $cfg -SiteUrl $site }
    catch { Write-Host "[$site] ERROR: cannot connect: $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }
    $web = Get-PnPWeb -Includes SiteUserInfoList
    $stale = @(foreach ($u in (Get-PnPUser)) {
            $info = Get-SpRbacPrincipalInfo -Member $u
            if ($info.Kind -ne 'EntraGroup' -or -not $index.ById.ContainsKey($info.Id)) { continue }
            $current = $index.ById[$info.Id]
            if ($u.Title -cne $current) { [pscustomobject]@{ Id = $u.Id; Old = $u.Title; New = $current; EntraId = $info.Id } }
        })
    if (-not $stale.Count) { Write-Host "[$site] all group names are current" -ForegroundColor Green; continue }
    Write-Host "[$site] $($stale.Count) stale name(s)" -ForegroundColor Yellow
    foreach ($s in $stale) {
        $planned++; Write-Host "    '$($s.Old)'  ->  '$($s.New)'" -ForegroundColor Yellow
        $status = 'dry run'
        if ($Apply) {
            $status = ''
            # The user information list refuses a plain update on some sites; try the alternatives in turn
            foreach ($type in 'Update', 'SystemUpdate', 'UpdateOverwriteVersion') {
                try { Set-PnPListItem -List $web.SiteUserInfoList.Id -Identity $s.Id -Values @{ Title = $s.New } -UpdateType $type -ErrorAction Stop | Out-Null; $status = "applied ($type)"; break }
                catch { $last = $_.Exception.Message }
            }
            if ($status) { $done++ } else { $errors++; $status = "ERROR: $last"; Write-Host "       $status" -ForegroundColor Red }
        }
        $log.Add([pscustomobject]@{ Site = $site; OldName = $s.Old; NewName = $s.New; EntraId = $s.EntraId; Status = $status })
    }
}

$file = New-SpRbacOutputPath -Config $cfg -Name 'group-titles'
$log | Export-Csv -LiteralPath $file -Encoding utf8
Write-SpRbacFooter -Apply:$Apply -Planned $planned -Done $done -Errors $errors
Write-Host "Log: $file" -ForegroundColor Gray
if ($errors) { exit 1 }
