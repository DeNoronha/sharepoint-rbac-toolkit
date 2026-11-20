<#
.SYNOPSIS
    Adds people to, or removes people from, a role group in Entra ID.
    Dry run by default; nothing changes without -Apply.

.DESCRIPTION
    This is how a person gets or loses access: through membership of a role group.
    Nothing is changed in SharePoint.

    The script checks, before it changes anything:
      - that the role is a row in the matrix (unless you pass -Group)
      - that the group exists and its name is unique
      - that every account exists; a disabled account is not added

.PARAMETER ConfigPath
    Path to config.json or to the folder that contains it.

.PARAMETER Role
    Role name as in the first column of the matrix, without the group prefix.

.PARAMETER Group
    Full group name, for groups outside the matrix (for example a folder exception group).
    Use either -Role or -Group.

.PARAMETER Add
    One or more user principal names (sign-in addresses) to add.

.PARAMETER Remove
    One or more user principal names to remove.

.PARAMETER Apply
    Make the changes. Without this switch the script only shows what it would do.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/05-Set-GroupMembership.ps1 -ConfigPath ./config/contoso -Role "Finance Team" -Add jane@contoso.com
    pwsh -NoProfile -File ./scripts/05-Set-GroupMembership.ps1 -ConfigPath ./config/contoso -Role "Finance Team" -Add jane@contoso.com -Apply

.EXAMPLE
    # Someone changes role: add to the new group and remove from the old one
    pwsh -NoProfile -File ./scripts/05-Set-GroupMembership.ps1 -ConfigPath ./config/contoso -Role "Sales Team" -Add jane@contoso.com -Apply
    pwsh -NoProfile -File ./scripts/05-Set-GroupMembership.ps1 -ConfigPath ./config/contoso -Role "Finance Team" -Remove jane@contoso.com -Apply

.NOTES
    Needs: Microsoft.Graph.Authentication. Delegated scopes: GroupMember.ReadWrite.All,
    User.Read.All. Role: Groups Administrator or owner of the group.
    When calling with -File, separate several addresses with commas and no spaces.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [string]$Role,
    [string]$Group,
    [string[]]$Add = @(),
    [string[]]$Remove = @(),
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
$cfg = Get-SpRbacConfig -Path $ConfigPath
Write-SpRbacHeader -Title 'SET GROUP MEMBERSHIP' -Config $cfg -Apply:$Apply

# pwsh -File passes "a@x.com,b@x.com" as one string
$Add = @($Add | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$Remove = @($Remove | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

if ([bool]$Role -eq [bool]$Group) { throw 'Pass either -Role (a row of the matrix) or -Group (a full group name).' }
if (-not $Add.Count -and -not $Remove.Count) { throw 'Nothing to do: pass -Add and/or -Remove with one or more user principal names.' }
$both = @($Add | Where-Object { $_ -in $Remove })
if ($both) { throw "These accounts are in both -Add and -Remove: $($both -join ', ')" }
if ($Role) {
    $match = $cfg.Roles | Where-Object { $_ -eq $Role } | Select-Object -First 1
    if (-not $match) { throw "'$Role' is not a role in the matrix. Roles: $($cfg.Roles -join ', ')" }
    $Group = $cfg.GroupPrefix + $match
}

Connect-SpRbacGraph -Config $cfg -Scopes 'GroupMember.ReadWrite.All', 'Group.Read.All', 'User.Read.All'
$g = Get-SpRbacGraphGroupByName -Name $Group
if (-not $g) { Disconnect-MgGraph | Out-Null; throw "Group '$Group' does not exist in Entra. Create it with 04-Sync-GroupNesting.ps1 -CreateMissingGroups." }
Write-Host "Group: $($g.displayName)  ($($g.id))" -ForegroundColor Cyan
$current = @{}
foreach ($m in @(Invoke-SpRbacGraphPaged "https://graph.microsoft.com/v1.0/groups/$($g.id)/members?`$select=id,displayName,userPrincipalName")) { $current[$m.id] = $m }
Write-Host "  Current members: $($current.Count)`n" -ForegroundColor Gray

$log = [System.Collections.Generic.List[object]]::new()
$planned = 0; $done = 0; $errors = 0
foreach ($entry in @($Add | ForEach-Object { @{ Upn = $_; Action = 'ADD' } }) + @($Remove | ForEach-Object { @{ Upn = $_; Action = 'REMOVE' } })) {
    $upn = $entry.Upn; $action = $entry.Action; $status = ''
    $u = Get-SpRbacGraphUser -Upn $upn
    if (-not $u) { Write-Host "  ERROR   ${upn}: no such account in Entra" -ForegroundColor Red; $errors++; $log.Add([pscustomobject]@{ Group = $Group; Action = $action; User = $upn; Status = 'ERROR: account not found' }); continue }
    $isMember = $current.ContainsKey($u.id)
    if ($action -eq 'ADD') {
        if ($isMember) { Write-Host "  OK      $($u.displayName) <$upn> is already a member" -ForegroundColor Green; $status = 'already a member' }
        elseif (-not $u.accountEnabled) { Write-Host "  ERROR   $($u.displayName) <$upn> - account is disabled, not added" -ForegroundColor Red; $errors++; $status = 'ERROR: account disabled' }
        else {
            $planned++; Write-Host "  ADD     $($u.displayName) <$upn>" -ForegroundColor Yellow; $status = 'dry run'
            if ($Apply) {
                try { Add-SpRbacGraphGroupMember -GroupId $g.id -MemberId $u.id; $done++; $status = 'applied' }
                catch { $errors++; $status = "ERROR: $($_.Exception.Message)"; Write-Host "          $status" -ForegroundColor Red }
            }
        }
    }
    else {
        if (-not $isMember) { Write-Host "  OK      $($u.displayName) <$upn> is not a member" -ForegroundColor Green; $status = 'not a member' }
        else {
            $planned++; Write-Host "  REMOVE  $($u.displayName) <$upn>" -ForegroundColor Yellow; $status = 'dry run'
            if ($Apply) {
                try { Remove-SpRbacGraphGroupMember -GroupId $g.id -MemberId $u.id; $done++; $status = 'applied' }
                catch { $errors++; $status = "ERROR: $($_.Exception.Message)"; Write-Host "          $status" -ForegroundColor Red }
            }
        }
    }
    $log.Add([pscustomobject]@{ Group = $Group; Action = $action; User = "$($u.displayName) <$upn>"; Status = $status })
}

if ($Apply -and $done) {
    $after = @(Invoke-SpRbacGraphPaged "https://graph.microsoft.com/v1.0/groups/$($g.id)/members?`$select=id,displayName,userPrincipalName")
    Write-Host "`nMembers of $Group after the change ($($after.Count)):" -ForegroundColor Cyan
    $after | Sort-Object { $_.displayName } | ForEach-Object { Write-Host "  $($_.displayName) <$($_.userPrincipalName)>" }
}
Disconnect-MgGraph | Out-Null
$file = New-SpRbacOutputPath -Config $cfg -Name 'membership'
$log | Export-Csv -LiteralPath $file -Encoding utf8
Write-SpRbacFooter -Apply:$Apply -Planned $planned -Done $done -Errors $errors
Write-Host "Log: $file" -ForegroundColor Gray
if ($Apply -and $done) { Write-Host 'Access follows within about an hour. Next: 01-Export-EntraSnapshot.ps1, then rebuild the report.' -ForegroundColor Gray }
if ($errors) { exit 1 }
