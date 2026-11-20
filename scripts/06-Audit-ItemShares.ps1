<#
.SYNOPSIS
    Lists the files and folders that were shared with individual people. Read-only.

.DESCRIPTION
    Library permissions can be clean while single files and folders inside the library are
    still shared with people directly or through a sharing link. This script finds every
    file and folder that has its own permissions and lists who got access that way.

    For each row it also says whether that person already reaches the library through a
    role group. If so, removing the share costs them nothing.

    Output: itemshares_<timestamp>.csv in the output folder, with an empty Decision column.
    Fill in Keep or Remove per row (in Excel: report/itemshares.py builds a workbook with a
    dropdown) and feed the file to 07-Remove-ItemShares.ps1.

    Library targets are scanned as configured. For a site target, every document library
    on the site is scanned.

.PARAMETER ConfigPath
    Path to config.json or to the folder that contains it.

.PARAMETER Target
    Optional. One or more target names. Default: all targets that have a site.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/06-Audit-ItemShares.ps1 -ConfigPath ./config/contoso -Target Finance,Board

.NOTES
    Needs: PnP.PowerShell and a snapshot from 01-Export-EntraSnapshot.ps1.
    Large libraries take time: every item with its own permissions costs one request.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [string[]]$Target
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
$cfg = Get-SpRbacConfig -Path $ConfigPath
Write-SpRbacHeader -Title 'AUDIT INDIVIDUALLY SHARED FILES AND FOLDERS' -Config $cfg -ReadOnly
Assert-SpRbacModule -Name 'PnP.PowerShell'
$index = Get-SpRbacGroupIndex -Config $cfg

$Target = @($Target | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$targets = @($cfg.Targets | Where-Object { $_.Scope -ne 'None' })
if ($Target) {
    $unknown = @($Target | Where-Object { $_ -notin $cfg.Targets.Name })
    if ($unknown) { throw "Unknown target(s): $($unknown -join ', '). Valid names: $($cfg.Targets.Name -join ', ')" }
    $targets = @($targets | Where-Object { $_.Name -in $Target })
}

# Who reaches a target through a group: expand the expected groups to people (nested groups included)
$membersOf = @{}
foreach ($r in $index.Rows) { if (-not $membersOf.ContainsKey($r.GroupName)) { $membersOf[$r.GroupName] = @() }; $membersOf[$r.GroupName] += $r }
function Get-PeopleInGroup([string]$Name, [hashtable]$Seen) {
    if ($Seen.ContainsKey($Name)) { return }; $Seen[$Name] = $true
    foreach ($m in $membersOf[$Name]) {
        if ($m.MemberType -eq 'User') { [pscustomobject]@{ Upn = "$($m.MemberUPN)".ToLower(); Via = $Name } }
        elseif ($m.MemberType -eq 'Nested Group') { Get-PeopleInGroup $m.MemberName $Seen | ForEach-Object { [pscustomobject]@{ Upn = $_.Upn; Via = $_.Via } } }
    }
}
function Get-AccessMap($t) {
    $map = @{}
    foreach ($e in (Get-SpRbacExpectedGroups -Config $cfg -Target $t | Where-Object Kind -eq 'Access')) {
        foreach ($p in (Get-PeopleInGroup $e.Group @{})) { if (-not $map.ContainsKey($p.Upn)) { $map[$p.Upn] = @() }; if ($p.Via -notin $map[$p.Upn]) { $map[$p.Upn] += $p.Via } }
    }
    $map
}
$knownUsers = @{}
foreach ($r in $index.Rows) { if ($r.MemberType -eq 'User') { $knownUsers["$($r.MemberUPN)".ToLower()] = $true } }

$rows = [System.Collections.Generic.List[object]]::new()
$errors = 0
foreach ($group in ($targets | Group-Object Site)) {
    Write-Host "Site $($group.Name)" -ForegroundColor Cyan
    try { Connect-SpRbacSite -Config $cfg -SiteUrl $group.Name }
    catch { Write-Host "  ERROR: cannot connect: $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }
    $ctx = Get-PnPContext

    foreach ($t in $group.Group) {
        $access = Get-AccessMap $t
        $lists = @()
        if ($t.Scope -eq 'Library') { $l = Resolve-SpRbacLibrary -Target $t; if ($l) { $lists = @($l) } else { Write-Host "  [$($t.Name)] library not found" -ForegroundColor Red; $errors++; continue } }
        else { $lists = @(Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden }) }

        foreach ($list in $lists) {
            $root = (Get-PnPProperty -ClientObject $list -Property RootFolder).ServerRelativeUrl
            $items = @(Get-PnPListItem -List $list -PageSize 1000 -Includes HasUniqueRoleAssignments, FileSystemObjectType)
            $unique = @($items | Where-Object { $_.HasUniqueRoleAssignments })
            Write-Host ("  [{0}] {1}: {2} items, {3} with their own permissions" -f $t.Name, $list.Title, $items.Count, $unique.Count)

            foreach ($item in $unique) {
                $path = "$($item.FieldValues['FileRef'])"
                $rel = if ($path.StartsWith($root)) { $path.Substring($root.Length).TrimStart('/') } else { $path }
                $kind = if ("$($item.FileSystemObjectType)" -eq 'Folder') { 'Folder' } else { 'File' }
                $ras = Get-PnPProperty -ClientObject $item -Property RoleAssignments
                foreach ($ra in $ras) { $ctx.Load($ra.Member); $ctx.Load($ra.RoleDefinitionBindings) }
                $ctx.ExecuteQuery()
                foreach ($ra in $ras) {
                    $levels = @(Get-SpRbacLevels -Bindings $ra.RoleDefinitionBindings)
                    if (-not $levels.Count) { continue }
                    $info = Get-SpRbacPrincipalInfo -Member $ra.Member
                    $people = @()
                    if ($info.Kind -in 'User', 'Guest') { $people = @([pscustomobject]@{ Title = $info.Title; Login = $info.Login; Upn = $info.Id; Via = 'Direct'; LinkGroup = '' }) }
                    elseif ($info.Kind -eq 'SharingLink') {
                        $linkType = ($info.Login -split '\.')[2]
                        $linkMembers = @(Get-PnPGroupMember -Group $ra.Member.Id)
                        if (-not $linkMembers.Count) { $people = @([pscustomobject]@{ Title = '(link not used by anyone yet)'; Login = ''; Upn = ''; Via = "Sharing link ($linkType)"; LinkGroup = $info.Login }) }
                        foreach ($u in $linkMembers) { $ui = Get-SpRbacPrincipalInfo -Member $u; $people += [pscustomobject]@{ Title = $ui.Title; Login = $ui.Login; Upn = $ui.Id; Via = "Sharing link ($linkType)"; LinkGroup = $info.Login } }
                    }
                    foreach ($p in $people) {
                        $isGuest = $p.Upn -like '*#ext#*'
                        $account = if ($isGuest) { (($p.Upn -split '#ext#')[0] -replace '_(?=[^_]*$)', '@') + ' (guest)' } else { $p.Upn }
                        $already = if (-not $p.Upn) { '' }
                        elseif ($isGuest) { 'No - external guest' }
                        elseif ($access.ContainsKey($p.Upn)) { 'Yes - ' + (($access[$p.Upn] | Sort-Object) -join ', ') }
                        elseif (-not $knownUsers.ContainsKey($p.Upn)) { 'No - account is not in any role group' }
                        else { 'No' }
                        $rows.Add([pscustomobject]@{
                                Target = $t.Name; Library = $list.Title; Type = $kind; Path = $rel
                                SharedWith = ($p.Title -split ' \| ')[0]; Account = $account; How = $p.Via; Access = $levels -join ', '
                                AlreadyHasAccessViaGroup = $already; Decision = ''
                                SiteUrl = $t.Site; ListId = "$($list.Id)"; ItemId = $item.Id; LoginName = $p.Login; LinkGroup = $p.LinkGroup
                            })
                    }
                }
            }
        }
    }
}

$sorted = @($rows | Sort-Object Target, Library, Path, SharedWith)
$file = New-SpRbacOutputPath -Config $cfg -Name 'itemshares'
$sorted | Export-Csv -LiteralPath $file -Encoding utf8
Write-Host ''
$sorted | Format-Table Library, Type, @{ n = 'Path'; e = { if ($_.Path.Length -gt 60) { '...' + $_.Path.Substring($_.Path.Length - 57) } else { $_.Path } } }, SharedWith, How, Access, AlreadyHasAccessViaGroup -AutoSize | Out-String -Width 260 | Write-Host
$items = @($sorted | Group-Object SiteUrl, ListId, ItemId).Count
$loses = @($sorted | Where-Object { $_.AlreadyHasAccessViaGroup -like 'No*' }).Count
Write-Host "$($sorted.Count) share(s) on $items file(s) or folder(s). $loses of them give access the person does not have through a group." -ForegroundColor $(if ($sorted.Count) { 'Yellow' } else { 'Green' })
Write-Host "Saved: $file" -ForegroundColor Gray
if ($errors) { Write-Host "$errors site(s) or librar(ies) could not be read." -ForegroundColor Red; exit 1 }
