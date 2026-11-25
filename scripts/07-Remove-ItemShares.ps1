<#
.SYNOPSIS
    Removes the individual shares that were marked 'Remove' in the list from
    06-Audit-ItemShares.ps1. Dry run by default; nothing changes without -Apply.

.DESCRIPTION
    Reads a decisions file: the CSV from 06-Audit-ItemShares.ps1 with the Decision column
    filled in (or the CSV written by 'report/itemshares.py decisions' from the workbook).

    Only rows with Decision = Remove are processed. Rows with Keep or an empty decision
    are left alone.

      Direct share    the person's own permission on the file or folder is removed
      Sharing link    the person is removed from the link; an unused link is deleted

    The file or folder keeps its own permissions. It does not go back to inheriting from
    the library, because other shares on the same item may have been kept.

.PARAMETER ConfigPath
    Path to config.json or to the folder that contains it.

.PARAMETER DecisionFile
    CSV with the decisions. Default: the newest itemshares-decisions_*.csv, or else the
    newest itemshares_*.csv, in the output folder.

.PARAMETER Apply
    Make the changes. Without this switch the script only shows what it would do.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/07-Remove-ItemShares.ps1 -ConfigPath ./config/contoso
    pwsh -NoProfile -File ./scripts/07-Remove-ItemShares.ps1 -ConfigPath ./config/contoso -Apply

.NOTES
    Needs: PnP.PowerShell and Full Control on the sites.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [string]$DecisionFile,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
$cfg = Get-SpRbacConfig -Path $ConfigPath
Write-SpRbacHeader -Title 'REMOVE INDIVIDUAL SHARES' -Config $cfg -Apply:$Apply
Assert-SpRbacModule -Name 'PnP.PowerShell'

if (-not $DecisionFile) {
    $f = Get-SpRbacLatestOutput -Config $cfg -Name 'itemshares-decisions'
    if (-not $f) { $f = Get-SpRbacLatestOutput -Config $cfg -Name 'itemshares' }
    if (-not $f) { throw "No itemshares CSV found in $($cfg.OutputDir). Run 06-Audit-ItemShares.ps1 first." }
    $DecisionFile = $f.FullName
}
$all = @(Import-Csv -LiteralPath $DecisionFile -Encoding utf8)
foreach ($col in 'Decision', 'SiteUrl', 'ListId', 'ItemId', 'LoginName', 'LinkGroup', 'How') {
    if ($all.Count -and $all[0].PSObject.Properties.Name -notcontains $col) { throw "Decision file has no column '$col'. Use the CSV from 06-Audit-ItemShares.ps1 (or from 'itemshares.py decisions')." }
}
$todo = @($all | Where-Object { "$($_.Decision)".Trim() -eq 'Remove' })
$odd = @($all | Where-Object { "$($_.Decision)".Trim() -notin '', 'Keep', 'Remove' })
Write-Host "Decisions: $(Split-Path $DecisionFile -Leaf)" -ForegroundColor Gray
Write-Host "  $($all.Count) rows: $($todo.Count) Remove, $(@($all | Where-Object { "$($_.Decision)".Trim() -eq 'Keep' }).Count) Keep, $(@($all | Where-Object { "$($_.Decision)".Trim() -eq '' }).Count) undecided`n" -ForegroundColor Gray
if ($odd) { throw "Unknown value(s) in the Decision column: $(($odd.Decision | Select-Object -Unique) -join ', '). Use Keep or Remove." }

$log = [System.Collections.Generic.List[object]]::new()
$planned = 0; $done = 0; $errors = 0
foreach ($group in ($todo | Group-Object SiteUrl)) {
    Write-Host "Site $($group.Name)" -ForegroundColor Cyan
    try { Connect-SpRbacSite -Config $cfg -SiteUrl $group.Name }
    catch { Write-Host "  ERROR: cannot connect: $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }
    $ctx = Get-PnPContext

    foreach ($r in $group.Group) {
        $planned++
        Write-Host ("  REMOVE  {0}  <-  {1}  [{2}]" -f $r.Path, $r.SharedWith, $r.How) -ForegroundColor Yellow
        $status = 'dry run'
        try {
            $item = Get-PnPListItem -List $r.ListId -Id $r.ItemId
            $ras = Get-PnPProperty -ClientObject $item -Property RoleAssignments
            foreach ($ra in $ras) { $ctx.Load($ra.Member) }
            $ctx.ExecuteQuery()
            if ($r.How -eq 'Direct') {
                $ra = $ras | Where-Object { $_.Member.LoginName -eq $r.LoginName } | Select-Object -First 1
                if (-not $ra) { $status = 'already gone'; $planned--; Write-Host '          already gone' -ForegroundColor Green }
                elseif ($Apply) { $ra.DeleteObject(); $ctx.ExecuteQuery(); $done++; $status = 'applied' }
            }
            else {
                $ra = $ras | Where-Object { $_.Member.LoginName -eq $r.LinkGroup } | Select-Object -First 1
                if (-not $ra) { $status = 'already gone'; $planned--; Write-Host '          link already gone' -ForegroundColor Green }
                elseif (-not $r.LoginName) {
                    # Unused link: remove the link's permission from the item
                    if ($Apply) { $ra.DeleteObject(); $ctx.ExecuteQuery(); $done++; $status = 'applied (unused link removed)' }
                }
                else {
                    $spg = $ctx.Web.SiteGroups.GetById($ra.Member.Id); $ctx.Load($spg.Users); $ctx.ExecuteQuery()
                    if (-not ($spg.Users | Where-Object { $_.LoginName -eq $r.LoginName })) { $status = 'already gone'; $planned--; Write-Host '          already gone' -ForegroundColor Green }
                    elseif ($Apply) {
                        $spg.Users.RemoveByLoginName($r.LoginName); $ctx.ExecuteQuery(); $done++; $status = 'applied'
                        # A link nobody is left on still works for whoever has the URL, so remove the link itself
                        $ctx.Load($spg.Users); $ctx.ExecuteQuery()
                        if ($spg.Users.Count -eq 0) { $ra.DeleteObject(); $ctx.ExecuteQuery(); $status = 'applied (last person; link removed)'; Write-Host '          link is now empty: link removed' -ForegroundColor DarkGray }
                    }
                }
            }
        }
        catch { $errors++; $status = "ERROR: $($_.Exception.Message)"; Write-Host "          $status" -ForegroundColor Red }
        $log.Add(($r | Select-Object Target, Library, Type, Path, SharedWith, Account, How, Access | Add-Member Status $status -PassThru))
    }
}

$file = New-SpRbacOutputPath -Config $cfg -Name 'itemshares-removed'
$log | Export-Csv -LiteralPath $file -Encoding utf8
Write-SpRbacFooter -Apply:$Apply -Planned $planned -Done $done -Errors $errors
Write-Host "Log: $file" -ForegroundColor Gray
if ($Apply) { Write-Host 'Next: run 06-Audit-ItemShares.ps1 again to confirm what is left.' -ForegroundColor Gray }
if ($errors) { exit 1 }
