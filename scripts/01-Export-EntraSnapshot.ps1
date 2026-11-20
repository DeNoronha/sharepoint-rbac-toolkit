<#
.SYNOPSIS
    Reads the authorization groups, their members and the users from Entra ID. Read-only.

.DESCRIPTION
    Writes two CSV files to the output folder of the config:

      entra-groups_<timestamp>.csv   one row per group member (group id, member, type)
      entra-users_<timestamp>.csv    one row per user (job title, department, groups)

    Every other script reads these files. The SharePoint scripts need them to translate
    group names into object ids, because they cannot query Microsoft Graph themselves.
    Run this script first, and again after every change in Entra.

    Groups included: every group whose name starts with one of 'snapshotPrefixes'
    (default: 'groupPrefix'), plus the admin and exception groups from the config.

.PARAMETER ConfigPath
    Path to config.json or to the folder that contains it.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/01-Export-EntraSnapshot.ps1 -ConfigPath ./config/contoso

.NOTES
    Needs: Microsoft.Graph.Authentication. Delegated scopes: Group.Read.All, User.Read.All.
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$ConfigPath)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
$cfg = Get-SpRbacConfig -Path $ConfigPath
Write-SpRbacHeader -Title 'EXPORT ENTRA SNAPSHOT' -Config $cfg -ReadOnly
Connect-SpRbacGraph -Config $cfg -Scopes 'Group.Read.All', 'GroupMember.Read.All', 'User.Read.All'

# ── Groups and members ───────────────────────────────────────────────
$extra = @($cfg.AdminGroups | ForEach-Object Group) + @($cfg.ExceptionGroups) + @($cfg.EveryoneGroup)
$groups = @(Get-SpRbacGraphGroups -Prefixes $cfg.SnapshotPrefixes -Names $extra)
if (-not $groups.Count) {
    throw "No groups found whose name starts with: $($cfg.SnapshotPrefixes -join ', '). Check 'groupPrefix' / 'snapshotPrefixes' in the config, and check that you signed in to the right tenant ($((Get-MgContext).Account))."
}
Write-Host "Found $($groups.Count) groups. Reading members ..." -ForegroundColor Cyan

$dupes = @($groups | Group-Object displayName | Where-Object Count -gt 1)
foreach ($d in $dupes) { Write-Host "  WARNING: $($d.Count) groups share the name '$($d.Name)'. Rename one; names must be unique." -ForegroundColor Yellow }

$groupRows = [System.Collections.Generic.List[object]]::new()
$i = 0
foreach ($g in $groups) {
    $i++
    Write-Progress -Activity 'Reading group members' -Status "$i / $($groups.Count): $($g.displayName)" -PercentComplete (100 * $i / $groups.Count)
    $kind = if ($g.groupTypes -contains 'Unified') { 'Microsoft 365' } elseif ($g.securityEnabled) { 'Security' } else { 'Other' }
    $members = @(Invoke-SpRbacGraphPaged "https://graph.microsoft.com/v1.0/groups/$($g.id)/members?`$select=id,displayName,userPrincipalName,accountEnabled")
    if (-not $members.Count) {
        $groupRows.Add([pscustomobject]@{ GroupId = $g.id; GroupName = $g.displayName; GroupDescription = $g.description; GroupKind = $kind; MemberId = ''; MemberName = '(empty)'; MemberUPN = ''; MemberType = ''; AccountEnabled = '' })
        continue
    }
    foreach ($m in $members) {
        $type = switch ($m.'@odata.type') { '#microsoft.graph.user' { 'User' } '#microsoft.graph.group' { 'Nested Group' } default { 'Other' } }
        $groupRows.Add([pscustomobject]@{
                GroupId = $g.id; GroupName = $g.displayName; GroupDescription = $g.description; GroupKind = $kind
                MemberId = $m.id; MemberName = $m.displayName; MemberUPN = $m.userPrincipalName; MemberType = $type
                AccountEnabled = if ($type -eq 'User') { "$($m.accountEnabled)" } else { '' }
            })
    }
}
Write-Progress -Activity 'Reading group members' -Completed

# ── Users ────────────────────────────────────────────────────────────
Write-Host 'Reading users ...' -ForegroundColor Cyan
$allUsers = @(Invoke-SpRbacGraphPaged "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,jobTitle,department,companyName,accountEnabled,userType&`$count=true")
$groupsByUser = @{}
foreach ($grp in ($groupRows | Where-Object MemberType -eq 'User' | Group-Object MemberId)) { $groupsByUser[$grp.Name] = @($grp.Group | ForEach-Object GroupName) }

$domains = @($cfg.UserFilter.domains | Where-Object { $_ })
$companies = @($cfg.UserFilter.companies | Where-Object { $_ })
$hasFilter = ($domains.Count + $companies.Count) -gt 0
$userRows = foreach ($u in $allUsers) {
    $inGroups = $groupsByUser.ContainsKey($u.id)
    $match = $false
    foreach ($d in $domains) { if ($u.userPrincipalName -like "*@$d") { $match = $true } }
    foreach ($c in $companies) { if ($u.companyName -and $u.companyName -like "*$c*") { $match = $true } }
    # Without a userFilter only people who are in a group are listed
    if (-not $inGroups -and (-not $hasFilter -or -not $match)) { continue }
    $gs = if ($inGroups) { @($groupsByUser[$u.id] | Sort-Object -Unique) } else { @() }
    [pscustomobject]@{
        DisplayName = $u.displayName; UserPrincipalName = $u.userPrincipalName; JobTitle = $u.jobTitle
        Department = $u.department; CompanyName = $u.companyName; UserType = $u.userType
        AccountEnabled = "$($u.accountEnabled)"; InScope = "$($match -or -not $hasFilter)"
        GroupCount = $gs.Count; Groups = $gs -join '; '
    }
}
$userRows = @($userRows | Sort-Object DisplayName)

# ── Write ────────────────────────────────────────────────────────────
$groupFile = New-SpRbacOutputPath -Config $cfg -Name 'entra-groups'
$userFile = New-SpRbacOutputPath -Config $cfg -Name 'entra-users'
$groupRows | Sort-Object GroupName, MemberName | Export-Csv -LiteralPath $groupFile -Encoding utf8
$userRows | Export-Csv -LiteralPath $userFile -Encoding utf8
Disconnect-MgGraph | Out-Null

$empty = @($groupRows | Where-Object MemberName -eq '(empty)').Count
$disabled = @($groupRows | Where-Object { $_.MemberType -eq 'User' -and $_.AccountEnabled -eq 'False' })
$noGroup = @($userRows | Where-Object { $_.GroupCount -eq 0 -and $_.AccountEnabled -eq 'True' }).Count
Write-Host ''
Write-Host "Groups                         : $($groups.Count)  ($empty empty)"
Write-Host "Users listed                   : $($userRows.Count)  ($noGroup active without any group)"
Write-Host "Disabled accounts still in groups: $($disabled.Count)" -ForegroundColor $(if ($disabled.Count) { 'Yellow' } else { 'Gray' })
foreach ($d in $disabled) { Write-Host "    $($d.MemberName) <$($d.MemberUPN)> in $($d.GroupName)" -ForegroundColor Yellow }
Write-Host "`nSaved:`n  $groupFile`n  $userFile" -ForegroundColor Green
