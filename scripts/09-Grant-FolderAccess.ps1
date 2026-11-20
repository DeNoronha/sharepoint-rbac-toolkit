<#
.SYNOPSIS
    Gives one Entra group access to one folder: an exception outside the matrix.
    Dry run by default; nothing changes without -Apply.

.DESCRIPTION
    Sometimes people need one folder and nothing else of the site. The rule stays the
    same: access goes to a group, never to a person. Create a dedicated group for the
    exception, put the people in it, and let this script grant the group access.

      1. If the folder still inherits its permissions, inheritance is broken. Everything
         that was inherited is copied first, so nobody loses access.
      2. The group gets the requested permission level on the folder.

    Keep exceptions rare and write each one down: after this, changes to the permissions
    of the library no longer reach this folder.

    A group that was created minutes ago may not be known to SharePoint yet. The script
    retries for a few minutes.

.PARAMETER ConfigPath
    Path to config.json or to the folder that contains it.

.PARAMETER Site
    Site URL, or the part after the tenant: /sites/Operations

.PARAMETER Folder
    Server-relative path of the folder: /sites/Operations/Shared Documents/Board meeting

.PARAMETER Group
    Full name of the Entra group. It must be in the snapshot: run
    01-Export-EntraSnapshot.ps1 after creating it, and list it under 'exceptionGroups'
    in the config so the snapshot and the report include it.

.PARAMETER Permission
    Read, Contribute (default), Edit, or the name of a custom permission level.

.PARAMETER Apply
    Make the changes. Without this switch the script only shows what it would do.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/09-Grant-FolderAccess.ps1 -ConfigPath ./config/contoso -Site /sites/Operations -Folder "/sites/Operations/Shared Documents/Board meeting" -Group "SP - Board Meeting Guests"

.NOTES
    Needs: PnP.PowerShell, a snapshot from 01-Export-EntraSnapshot.ps1, Full Control on the site.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [Parameter(Mandatory)][string]$Site,
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$Group,
    [string]$Permission = 'Contribute',
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
$cfg = Get-SpRbacConfig -Path $ConfigPath
Write-SpRbacHeader -Title 'GRANT A GROUP ACCESS TO ONE FOLDER' -Config $cfg -Apply:$Apply
Assert-SpRbacModule -Name 'PnP.PowerShell'
$index = Get-SpRbacGroupIndex -Config $cfg -MaxAgeHours 4

$groupId = $index.ByName[$Group]
if (-not $groupId) { throw "Group '$Group' is not in the Entra snapshot. Add it to 'exceptionGroups' in the config and run 01-Export-EntraSnapshot.ps1." }
$siteUrl = if ($Site -match '^https://') { $Site.TrimEnd('/') } else { $cfg.TenantUrl + '/' + $Site.Trim('/') }
$login = "c:0t.c|tenant|$groupId"
Write-Host "  Site   : $siteUrl`n  Folder : $Folder`n  Group  : $Group ($groupId)`n  Level  : $Permission`n" -ForegroundColor Gray

Connect-SpRbacSite -Config $cfg -SiteUrl $siteUrl
$ctx = Get-PnPContext
$web = Get-PnPWeb
$planned = 0; $done = 0; $errors = 0

function Show-FolderAccess($item) {
    foreach ($ra in (Get-PnPProperty -ClientObject $item -Property RoleAssignments)) {
        Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings | Out-Null
        $levels = @(Get-SpRbacLevels -Bindings $ra.RoleDefinitionBindings)
        if ($levels.Count) { $info = Get-SpRbacPrincipalInfo -Member $ra.Member; $name = if ($info.Kind -eq 'EntraGroup' -and $index.ById.ContainsKey($info.Id)) { $index.ById[$info.Id] } else { $info.Title }; Write-Host ("    {0,-50} {1,-14} {2}" -f $name, ($levels -join ', '), $info.Kind) }
    }
}

try { $f = Get-PnPFolder -Url $Folder -Includes ListItemAllFields.HasUniqueRoleAssignments, ListItemAllFields.Id }
catch { throw "Folder not found: $Folder. Use the server-relative path, starting with /sites/..." }
$item = $f.ListItemAllFields
Write-Host "Folder found (item $($item.Id)). Own permissions: $($item.HasUniqueRoleAssignments)" -ForegroundColor Cyan
Write-Host '  Access now:'; Show-FolderAccess $item

if (-not $item.HasUniqueRoleAssignments) {
    $planned++; Write-Host "`n  BREAK_INHERITANCE  (existing permissions are copied)" -ForegroundColor Yellow
    if ($Apply) { $item.BreakRoleInheritance($true, $false); $ctx.ExecuteQuery(); $done++ }
}

$has = $false
foreach ($ra in (Get-PnPProperty -ClientObject $item -Property RoleAssignments)) {
    Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings | Out-Null
    if ($ra.Member.LoginName -eq $login -and (@(Get-SpRbacLevels -Bindings $ra.RoleDefinitionBindings) -contains $Permission)) { $has = $true }
}
if ($has) { Write-Host "`n  OK  $Group already has $Permission on this folder." -ForegroundColor Green }
else {
    $planned++; Write-Host "`n  GRANT  $Group -> $Permission" -ForegroundColor Yellow
    if ($Apply) {
        try {
            $principal = $null
            for ($i = 1; $i -le 10 -and -not $principal; $i++) {
                try { $p = $web.EnsureUser($login); $ctx.Load($p); $ctx.ExecuteQuery(); $principal = $p }
                catch { Write-Host "      group not resolvable yet (attempt $i/10), waiting 30 s ..." -ForegroundColor DarkYellow; Start-Sleep -Seconds 30 }
            }
            if (-not $principal) { throw "SharePoint cannot resolve '$Group' yet. Try again in a few minutes." }
            $rd = Get-SpRbacRoleDefinition -Web $web -Level $Permission
            $bindings = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
            $bindings.Add($rd)
            $item.RoleAssignments.Add($principal, $bindings) | Out-Null
            $ctx.ExecuteQuery(); $done++
        }
        catch { $errors++; Write-Host "      ERROR: $($_.Exception.Message)" -ForegroundColor Red }
    }
}

if ($Apply -and $done) {
    $item = (Get-PnPFolder -Url $Folder -Includes ListItemAllFields.HasUniqueRoleAssignments).ListItemAllFields
    Write-Host "`n  Access after the change (own permissions: $($item.HasUniqueRoleAssignments)):"; Show-FolderAccess $item
}
Write-SpRbacFooter -Apply:$Apply -Planned $planned -Done $done -Errors $errors
if ($Apply -and $done) { Write-Host 'Write this exception down (who asked, why, which group) and mention it in the report notes.' -ForegroundColor Gray }
if ($errors) { exit 1 }
