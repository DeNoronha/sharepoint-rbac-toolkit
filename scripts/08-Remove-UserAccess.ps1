<#
.SYNOPSIS
    Removes all SharePoint access of someone who has left. Two steps, each in its own run.
    Dry run by default; nothing changes without -Apply.

.DESCRIPTION
    -Step Groups   (Entra ID) Removes the account from every group the toolkit manages:
                   groups whose name starts with one of 'snapshotPrefixes', plus the admin,
                   everyone and exception groups from the config.

    -Step Sites    (SharePoint) Looks the account up on every site in the config (targets
                   plus 'extraSites') and removes it from the site. That drops everything
                   the person held on that site at once: memberships of SharePoint groups,
                   direct permissions on libraries, files and folders, and sharing links
                   they had used.

    Run Groups first, then Sites. They are separate runs because the Graph and PnP modules
    cannot be loaded into the same PowerShell process.

    This is for leavers. For someone who changes role, use 05-Set-GroupMembership.ps1.

.PARAMETER ConfigPath
    Path to config.json or to the folder that contains it.

.PARAMETER User
    User principal name (sign-in address) of the person.

.PARAMETER Step
    Groups or Sites.

.PARAMETER Apply
    Make the changes. Without this switch the script only shows what it would do.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/08-Remove-UserAccess.ps1 -ConfigPath ./config/contoso -User jane@contoso.com -Step Groups
    pwsh -NoProfile -File ./scripts/08-Remove-UserAccess.ps1 -ConfigPath ./config/contoso -User jane@contoso.com -Step Groups -Apply
    pwsh -NoProfile -File ./scripts/08-Remove-UserAccess.ps1 -ConfigPath ./config/contoso -User jane@contoso.com -Step Sites -Apply

.NOTES
    Groups: Microsoft.Graph.Authentication, scopes GroupMember.ReadWrite.All + User.Read.All.
    Sites:  PnP.PowerShell, Full Control on the sites.
    An account that was already deleted from Entra cannot be found by -Step Groups (its
    memberships are gone with it); -Step Sites still works, because SharePoint keeps its own
    copy of the user.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [Parameter(Mandatory)][string]$User,
    [Parameter(Mandatory)][ValidateSet('Groups', 'Sites')][string]$Step,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
$cfg = Get-SpRbacConfig -Path $ConfigPath
Write-SpRbacHeader -Title "REMOVE USER ACCESS - step $Step - $User" -Config $cfg -Apply:$Apply
$log = [System.Collections.Generic.List[object]]::new()
$planned = 0; $done = 0; $errors = 0

if ($Step -eq 'Groups') {
    Connect-SpRbacGraph -Config $cfg -Scopes 'GroupMember.ReadWrite.All', 'Group.Read.All', 'User.Read.All'
    $u = Get-SpRbacGraphUser -Upn $User
    if (-not $u) { Disconnect-MgGraph | Out-Null; throw "No account '$User' in Entra. If it was deleted, its group memberships are gone already; continue with -Step Sites." }
    Write-Host "Account: $($u.displayName) <$($u.userPrincipalName)>  enabled: $($u.accountEnabled)" -ForegroundColor Cyan

    $named = @($cfg.AdminGroups | ForEach-Object Group) + @($cfg.ExceptionGroups) + @($cfg.EveryoneGroup) | Where-Object { $_ }
    $memberOf = @(Invoke-SpRbacGraphPaged "https://graph.microsoft.com/v1.0/users/$($u.id)/memberOf?`$select=id,displayName")
    $mine = @($memberOf | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' } | Where-Object {
            $n = $_.displayName; ($n -in $named) -or (@($cfg.SnapshotPrefixes | Where-Object { $n.StartsWith($_) }).Count -gt 0) } | Sort-Object { $_.displayName })
    $other = @($memberOf | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' -and $_.id -notin $mine.id })
    Write-Host "  Member of $($mine.Count) authorization group(s) and $($other.Count) other group(s) (the others are not touched).`n" -ForegroundColor Gray

    foreach ($g in $mine) {
        $planned++; Write-Host "  REMOVE  from $($g.displayName)" -ForegroundColor Yellow
        $status = 'dry run'
        if ($Apply) {
            try { Remove-SpRbacGraphGroupMember -GroupId $g.id -MemberId $u.id; $done++; $status = 'applied' }
            catch { $errors++; $status = "ERROR: $($_.Exception.Message)"; Write-Host "          $status" -ForegroundColor Red }
        }
        $log.Add([pscustomobject]@{ Step = 'Groups'; Where = $g.displayName; Detail = 'group membership'; Status = $status })
    }
    if (-not $mine.Count) { Write-Host '  Not a member of any authorization group.' -ForegroundColor Green }
    Disconnect-MgGraph | Out-Null
}

if ($Step -eq 'Sites') {
    Assert-SpRbacModule -Name 'PnP.PowerShell'
    $login = "i:0#.f|membership|$($User.ToLower())"
    foreach ($site in (Get-SpRbacSites -Config $cfg)) {
        try { Connect-SpRbacSite -Config $cfg -SiteUrl $site }
        catch { Write-Host "  [$site] ERROR: cannot connect: $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }
        $ctx = Get-PnPContext
        $su = Get-PnPUser | Where-Object { "$($_.LoginName)".ToLower() -eq $login } | Select-Object -First 1
        if (-not $su) { Write-Host "  [$site] not known on this site" -ForegroundColor Green; continue }

        $ctx.Load($su.Groups); $ctx.ExecuteQuery()
        $groups = @($su.Groups | Where-Object { $_.Title -notlike 'Limited Access System Group*' })
        $links = @($groups | Where-Object { $_.Title -like 'SharingLinks.*' })
        $spGroups = @($groups | Where-Object { $_.Title -notlike 'SharingLinks.*' })
        $planned++
        Write-Host "  [$site] REMOVE from site: $($spGroups.Count) SharePoint group(s), $($links.Count) sharing link(s), plus any direct permissions" -ForegroundColor Yellow
        foreach ($g in $spGroups) { Write-Host "            group: $($g.Title)" -ForegroundColor DarkYellow }
        $status = 'dry run'
        if ($Apply) {
            try { Remove-PnPUser -Identity $su.Id -Force; $done++; $status = 'applied' }
            catch { $errors++; $status = "ERROR: $($_.Exception.Message)"; Write-Host "          $status" -ForegroundColor Red }
        }
        $log.Add([pscustomobject]@{ Step = 'Sites'; Where = $site; Detail = "SharePoint groups: $(($spGroups.Title) -join ' | '); sharing links: $($links.Count)"; Status = $status })
    }
}

$file = New-SpRbacOutputPath -Config $cfg -Name 'offboard'
$log | Export-Csv -LiteralPath $file -Encoding utf8
Write-SpRbacFooter -Apply:$Apply -Planned $planned -Done $done -Errors $errors
Write-Host "Log: $file" -ForegroundColor Gray
if ($Step -eq 'Groups') { Write-Host 'Next: the same command with -Step Sites.' -ForegroundColor Gray }
if ($errors) { exit 1 }
