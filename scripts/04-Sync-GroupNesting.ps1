<#
.SYNOPSIS
    Makes the groups in Entra ID match the access matrix: which groups exist, and (in the
    nested model) which role group is a member of which access group.
    Dry run by default; nothing changes without -Apply.

.DESCRIPTION
    Both models:
      - checks that every role in the matrix has a group named <groupPrefix><role>
      - with -CreateMissingGroups, creates the missing security groups

    Nested model only (role group -> access group -> site):
      - every access group contains exactly the role groups that have a mark in its
        matrix column: missing ones are added, others are removed
      - the 'everyoneGroup', when configured, contains every access group
      - people who are a direct member of an access group are reported, never removed:
        people belong in role groups

    People are never added to or removed from role groups by this script. Use
    05-Set-GroupMembership.ps1 for that.

.PARAMETER ConfigPath
    Path to config.json or to the folder that contains it.

.PARAMETER CreateMissingGroups
    Create the role, access and everyone groups that do not exist yet.

.PARAMETER Apply
    Make the changes. Without this switch the script only shows what it would do.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/04-Sync-GroupNesting.ps1 -ConfigPath ./config/contoso
    pwsh -NoProfile -File ./scripts/04-Sync-GroupNesting.ps1 -ConfigPath ./config/contoso -CreateMissingGroups -Apply

.NOTES
    Needs: Microsoft.Graph.Authentication. Delegated scopes: Group.ReadWrite.All
    (Group.Read.All is enough for a dry run). Role: Groups Administrator or owner of the groups.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [switch]$CreateMissingGroups,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
$cfg = Get-SpRbacConfig -Path $ConfigPath
Write-SpRbacHeader -Title 'SYNC GROUPS AND NESTING WITH THE MATRIX' -Config $cfg -Apply:$Apply
$scopes = if ($Apply) { 'Group.ReadWrite.All' } else { 'Group.Read.All', 'GroupMember.Read.All' }
Connect-SpRbacGraph -Config $cfg -Scopes $scopes

$log = [System.Collections.Generic.List[object]]::new()
$planned = 0; $done = 0; $errors = 0
function Add-Log($Group, $Action, $Member, $Status) { $log.Add([pscustomobject]@{ Group = $Group; Action = $Action; Member = $Member; Status = $Status }) }

# ── Which groups should exist ────────────────────────────────────────
$wanted = [ordered]@{}
foreach ($role in $cfg.Roles) { $wanted[$cfg.GroupPrefix + $role] = 'role group' }
if ($cfg.Model -eq 'nested') {
    foreach ($t in ($cfg.Targets | Where-Object InMatrix)) { $wanted[$t.AccessGroup] = 'access group' }
    if ($cfg.EveryoneGroup) { $wanted[$cfg.EveryoneGroup] = 'everyone group' }
}
foreach ($t in $cfg.Targets) { foreach ($f in $t.FixedGroups) { if (-not $wanted.Contains($f.Group)) { $wanted[$f.Group] = 'fixed group' } } }
foreach ($a in $cfg.AdminGroups) { if (-not $wanted.Contains($a.Group)) { $wanted[$a.Group] = 'admin group' } }

Write-Host "Checking $($wanted.Count) groups ..." -ForegroundColor Cyan
$ids = @{}
foreach ($name in $wanted.Keys) {
    $g = Get-SpRbacGraphGroupByName -Name $name
    if ($g) { $ids[$name] = $g.id; continue }
    $planned++
    if (-not $CreateMissingGroups) {
        Write-Host "  MISSING  $name  ($($wanted[$name])) - add -CreateMissingGroups to create it" -ForegroundColor Red
        Add-Log $name 'MISSING' '' 'not created (no -CreateMissingGroups)'; $errors++; continue
    }
    Write-Host "  CREATE   $name  ($($wanted[$name]))" -ForegroundColor Yellow
    if ($Apply) {
        try {
            $short = if ($name.StartsWith($cfg.GroupPrefix)) { $name.Substring($cfg.GroupPrefix.Length) } else { $name }
            $new = New-SpRbacGraphSecurityGroup -Name $name -Description ($cfg.GroupDescription.Replace('{name}', $short))
            $ids[$name] = $new.id; $done++; Add-Log $name 'CREATE' '' 'applied'
        }
        catch { $errors++; Add-Log $name 'CREATE' '' "ERROR: $($_.Exception.Message)"; Write-Host "      ERROR: $($_.Exception.Message)" -ForegroundColor Red }
    }
    else { Add-Log $name 'CREATE' '' 'dry run' }
}
$missing = @($wanted.Keys | Where-Object { -not $ids.ContainsKey($_) })
if (-not $missing.Count) { Write-Host '  All groups exist.' -ForegroundColor Green }

# ── Nesting (nested model) ───────────────────────────────────────────
if ($cfg.Model -ne 'nested') {
    Write-Host "`nModel is 'direct': role groups sit on the libraries themselves, so there is no nesting to sync." -ForegroundColor Gray
}
else {
    $containers = [ordered]@{}
    foreach ($t in ($cfg.Targets | Where-Object InMatrix)) {
        $containers[$t.AccessGroup] = @(Get-SpRbacRolesForTarget -Config $cfg -Target $t | ForEach-Object Group)
    }
    if ($cfg.EveryoneGroup) { $containers[$cfg.EveryoneGroup] = @($cfg.Targets | Where-Object InMatrix | ForEach-Object AccessGroup) }

    Write-Host "`nChecking nesting of $($containers.Count) groups ..." -ForegroundColor Cyan
    foreach ($container in $containers.Keys) {
        if (-not $ids.ContainsKey($container)) { Write-Host "  [$container] skipped: group does not exist yet" -ForegroundColor Red; continue }
        $members = @(Invoke-SpRbacGraphPaged "https://graph.microsoft.com/v1.0/groups/$($ids[$container])/members?`$select=id,displayName,userPrincipalName")
        $nested = @{}; $people = @()
        foreach ($m in $members) { if ($m.'@odata.type' -eq '#microsoft.graph.group') { $nested[$m.id] = $m.displayName } else { $people += $m } }

        $desired = @($containers[$container])
        $unresolved = @($desired | Where-Object { -not $ids.ContainsKey($_) })
        $desiredIds = @($desired | Where-Object { $ids.ContainsKey($_) } | ForEach-Object { $ids[$_] })
        $toAdd = @($desired | Where-Object { $ids.ContainsKey($_) -and -not $nested.ContainsKey($ids[$_]) })
        $toRemove = @($nested.Keys | Where-Object { $_ -notin $desiredIds })

        $changes = $toAdd.Count + $toRemove.Count
        $color = if ($changes -or $people.Count -or $unresolved.Count) { 'Yellow' } else { 'Green' }
        Write-Host ("  [{0}] nested: {1}   to add: {2}   to remove: {3}" -f $container, $nested.Count, $toAdd.Count, $toRemove.Count) -ForegroundColor $color

        foreach ($name in $toAdd) {
            $planned++; Write-Host "      ADD     $name" -ForegroundColor Yellow
            $status = 'dry run'
            if ($Apply) {
                try { Add-SpRbacGraphGroupMember -GroupId $ids[$container] -MemberId $ids[$name]; $done++; $status = 'applied' }
                catch { $errors++; $status = "ERROR: $($_.Exception.Message)"; Write-Host "         $status" -ForegroundColor Red }
            }
            Add-Log $container 'ADD' $name $status
        }
        # Never remove while a desired group cannot be resolved: the picture is incomplete
        if ($unresolved.Count) {
            foreach ($u in $unresolved) { Write-Host "      WAITING $u does not exist yet; nothing is removed from this group until it does" -ForegroundColor Red }
        }
        else {
            foreach ($id in $toRemove) {
                $planned++; Write-Host "      REMOVE  $($nested[$id])  (no mark in the matrix)" -ForegroundColor Yellow
                $status = 'dry run'
                if ($Apply) {
                    try { Remove-SpRbacGraphGroupMember -GroupId $ids[$container] -MemberId $id; $done++; $status = 'applied' }
                    catch { $errors++; $status = "ERROR: $($_.Exception.Message)"; Write-Host "         $status" -ForegroundColor Red }
                }
                Add-Log $container 'REMOVE' $nested[$id] $status
            }
        }
        foreach ($p in $people) {
            Write-Host "      REVIEW  $($p.displayName) <$($p.userPrincipalName)> is a direct member; people belong in role groups" -ForegroundColor DarkYellow
            Add-Log $container 'REVIEW' "$($p.displayName) <$($p.userPrincipalName)>" 'direct member, not changed'
        }
    }
}

Disconnect-MgGraph | Out-Null
$file = New-SpRbacOutputPath -Config $cfg -Name 'nesting-sync'
$log | Export-Csv -LiteralPath $file -Encoding utf8
Write-SpRbacFooter -Apply:$Apply -Planned $planned -Done $done -Errors $errors
Write-Host "Log: $file" -ForegroundColor Gray
if ($Apply -and $done) { Write-Host 'Next: run 01-Export-EntraSnapshot.ps1 so the other scripts see the new state.' -ForegroundColor Gray }
if ($errors) { exit 1 }
