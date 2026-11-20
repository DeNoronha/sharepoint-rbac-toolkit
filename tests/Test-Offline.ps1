<#
.SYNOPSIS
    Self-test without any network or sign-in: checks the config and matrix logic against
    the example configurations and builds an example report from a made-up snapshot.

.EXAMPLE
    pwsh -NoProfile -File ./tests/Test-Offline.ps1
    pwsh -NoProfile -File ./tests/Test-Offline.ps1 -Python /usr/bin/python3
#>
[CmdletBinding()]
param([string]$Python)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $root 'scripts/lib/SpRbac.psm1') -Force
$script:failed = 0; $script:passed = 0
function Check([string]$Name, [scriptblock]$Test) {
    try { if (& $Test) { $script:passed++; Write-Host "  [ok]   $Name" -ForegroundColor Green } else { $script:failed++; Write-Host "  [FAIL] $Name" -ForegroundColor Red } }
    catch { $script:failed++; Write-Host "  [FAIL] $Name : $($_.Exception.Message)" -ForegroundColor Red }
}
function Throws([scriptblock]$Code, [string]$Pattern) { try { & $Code | Out-Null; $false } catch { "$($_.Exception.Message)" -match $Pattern } }

Write-Host "`nScripts" -ForegroundColor White
foreach ($f in (Get-ChildItem (Join-Path $root 'scripts') -Recurse -Include *.ps1, *.psm1 | Sort-Object Name)) {
    $tokens = $null; $errs = $null
    [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errs) | Out-Null
    Check "$($f.Name) parses without syntax errors" { -not $errs.Count }
    if ($f.Extension -eq '.ps1') { Check "$($f.Name) has a help block with a synopsis" { (Get-Content $f.FullName -Raw) -match '(?s)^<#.*?\.SYNOPSIS.*?#>' } }
}
$changing = '03-Sync-Permissions', '04-Sync-GroupNesting', '05-Set-GroupMembership', '07-Remove-ItemShares', '08-Remove-UserAccess', '09-Grant-FolderAccess', '10-Update-GroupTitles'
foreach ($sn in $changing) { $cmd = Get-Command (Join-Path $root "scripts/$sn.ps1"); Check "$sn is a dry run unless -Apply is given" { $cmd.Parameters.ContainsKey('Apply') } }
foreach ($sn in '01-Export-EntraSnapshot', '02-Audit-Permissions', '06-Audit-ItemShares') { $cmd = Get-Command (Join-Path $root "scripts/$sn.ps1"); Check "$sn is read-only (no -Apply switch)" { -not $cmd.Parameters.ContainsKey('Apply') } }

Write-Host "`nConfig and matrix (direct model)" -ForegroundColor White
$d = Get-SpRbacConfig -Path (Join-Path $root 'examples/contoso-direct')
$t = @{}; foreach ($x in $d.Targets) { $t[$x.Name] = $x }
Check 'loads 7 roles and 7 targets' { $d.Roles.Count -eq 7 -and $d.Targets.Count -eq 7 }
Check 'tenant for sign-in is derived from the URL' { $d.TenantId -eq 'contoso.onmicrosoft.com' }
Check 'relative site becomes a full URL' { $t['Finance'].Site -eq 'https://contoso.sharepoint.com/sites/Warehouse' }
$fin = @(Get-SpRbacExpectedGroups -Config $d -Target $t['Finance'])
Check 'Finance: 3 role groups + admin group' { $fin.Count -eq 4 -and @($fin | Where-Object Kind -eq 'Admin').Count -eq 1 }
Check 'role override: General Manager gets Edit' { ($fin | Where-Object Group -eq 'SP - UK General Manager').Permission -eq 'Edit' }
Check 'default level: Finance Team gets Contribute' { ($fin | Where-Object Group -eq 'SP - UK Finance Team').Permission -eq 'Contribute' }
Check 'level in a cell: Auditor gets Read' { ($fin | Where-Object Group -eq 'SP - UK Auditor').Permission -eq 'Read' }
Check 'admin group gets Full Control' { ($fin | Where-Object Group -eq 'SP - IT Admins').Permission -eq 'Full Control' }
$ev = @(Get-SpRbacExpectedGroups -Config $d -Target $t['Everyone'])
Check 'fixedGroups: Everyone library gets the everyone group' { @($ev | Where-Object Group -eq 'SP - UK Everyone').Count -eq 1 -and $ev.Count -eq 2 }
Check 'empty cell means no access' { -not (@(Get-SpRbacExpectedGroups -Config $d -Target $t['Board']) | Where-Object Group -eq 'SP - UK Auditor') }

Write-Host "`nConfig and matrix (nested model)" -ForegroundColor White
$n = Get-SpRbacConfig -Path (Join-Path $root 'examples/contoso-nested')
$nt = @{}; foreach ($x in $n.Targets) { $nt[$x.Name] = $x }
$sales = @(Get-SpRbacExpectedGroups -Config $n -Target $nt['Sales'])
Check 'site gets the access group, not the role groups' { @($sales | Where-Object Kind -eq 'Access').Count -eq 1 -and ($sales | Where-Object Kind -eq 'Access').Group -eq 'SP - NL Sales' }
Check 'roles nested in the Sales access group' { (@(Get-SpRbacRolesForTarget -Config $n -Target $nt['Sales']).Group -join '|') -eq 'SP - NL General Manager|SP - NL Head of Sales|SP - NL Account Manager|SP - NL Sales Support' }
Check 'target without a site has scope None' { $nt['Archive'].Scope -eq 'None' }
Check 'sites are listed once' { (Get-SpRbacSites -Config $n).Count -eq 5 }

Write-Host "`nValidation" -ForegroundColor White
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("sprbac-test-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmp | Out-Null
try {
    Copy-Item (Join-Path $root 'examples/contoso-direct/*') $tmp
    $cfgFile = Join-Path $tmp 'config.json'
    $good = Get-Content $cfgFile -Raw
    $good.Replace('"model": "direct"', '"model": "flat"') | Set-Content $cfgFile
    Check 'rejects an unknown model' { Throws { Get-SpRbacConfig -Path $tmp } "must be 'direct' or 'nested'" }
    $good.Replace('"Customers": { "library": "Customers" },', '') | Set-Content $cfgFile
    Check 'rejects a matrix column without a target' { Throws { Get-SpRbacConfig -Path $tmp } "Matrix column 'Customers'" }
    $good.Replace('"General Manager": "Edit"', '"Managing Director": "Edit"') | Set-Content $cfgFile
    Check 'rejects rolePermissions for an unknown role' { Throws { Get-SpRbacConfig -Path $tmp } "Managing Director" }
    $good.Replace('"clientId": "00000000-0000-0000-0000-000000000000"', '"clientId": "my-app"') | Set-Content $cfgFile
    Check 'rejects a clientId that is not a GUID' { Throws { Get-SpRbacConfig -Path $tmp } 'clientId' }
    $good | Set-Content $cfgFile
    "Role,Procedures,Customers,Operations,Finance,Board,Human Resources,Everyone`nGeneral Manager,yes,,,Edit,,,`n" | Set-Content (Join-Path $tmp 'matrix.csv')
    $c = Get-SpRbacConfig -Path $tmp
    Check 'reads a comma-separated matrix' { $c.Roles.Count -eq 1 -and $c.Matrix['General Manager']['Procedures'] -eq '*' -and $c.Matrix['General Manager']['Finance'] -eq 'Edit' }

    Write-Host "`nSharePoint principals" -ForegroundColor White
    function P($login, $title = 'x', $type = 'User') { Get-SpRbacPrincipalInfo -Member ([pscustomobject]@{ LoginName = $login; Title = $title; PrincipalType = $type }) }
    Check 'Entra security group' { $p = P 'c:0t.c|tenant|0F8FAD5B-D9CB-469F-A165-70867728950E' 'g' 'SecurityGroup'; $p.Kind -eq 'EntraGroup' -and $p.Id -eq '0f8fad5b-d9cb-469f-a165-70867728950e' }
    Check 'Microsoft 365 group' { (P 'c:0o.c|federateddirectoryclaimprovider|0f8fad5b-d9cb-469f-a165-70867728950e' 'g' 'SecurityGroup').Kind -eq 'M365Group' }
    Check 'internal user' { $p = P 'i:0#.f|membership|Jane@Contoso.com'; $p.Kind -eq 'User' -and $p.Id -eq 'jane@contoso.com' }
    Check 'guest user' { (P 'i:0#.f|membership|joe_partner.com#ext#@contoso.onmicrosoft.com').Kind -eq 'Guest' }
    Check 'everyone except external users' { (P 'c:0-.f|rolemanager|spo-grid-all-users/0f8fad5b-d9cb-469f-a165-70867728950e' 'Everyone except external users' 'SecurityGroup').Kind -eq 'Everyone' }
    Check 'sharing link' { (P 'SharingLinks.0f8fad5b.Flexible.abc' 'SharingLinks.0f8fad5b.Flexible.abc' 'SharePointGroup').Kind -eq 'SharingLink' }
    Check 'SharePoint group' { (P 'Warehouse Members' 'Warehouse Members' 'SharePointGroup').Kind -eq 'SharePointGroup' }
    Check 'limited access system group is ignored' { (P 'Limited Access System Group' 'Limited Access System Group For List 1' 'SharePointGroup').Kind -eq 'System' }
    $b = @([pscustomobject]@{ Hidden = $false; RoleTypeKind = 'Contributor'; Name = 'Bijdragen' }, [pscustomobject]@{ Hidden = $true; RoleTypeKind = 'Guest'; Name = 'Limited Access' })
    Check 'permission levels are language-independent' { (@(Get-SpRbacLevels -Bindings $b) -join ',') -eq 'Contribute' }
    Check 'custom permission level keeps its name' { (@(Get-SpRbacLevels -Bindings @([pscustomobject]@{ Hidden = $false; RoleTypeKind = 'None'; Name = 'Approve only' })) -join ',') -eq 'Approve only' }

    Write-Host "`nSnapshot and report" -ForegroundColor White
    $c = Get-SpRbacConfig -Path $tmp
    Copy-Item (Join-Path $root 'examples/contoso-direct/matrix.csv') (Join-Path $tmp 'matrix.csv') -Force
    $c = Get-SpRbacConfig -Path $tmp
    Check 'no snapshot gives a clear error' { Throws { Get-SpRbacGroupIndex -Config $c } 'Run 01-Export-EntraSnapshot' }
    $g = @(
        [pscustomobject]@{ GroupId = '11111111-1111-1111-1111-111111111111'; GroupName = 'SP - UK General Manager'; GroupDescription = 'd'; GroupKind = 'Security'; MemberId = 'u1'; MemberName = 'Ann Lee | Contoso'; MemberUPN = 'ann@contoso.co.uk'; MemberType = 'User'; AccountEnabled = 'True' }
        [pscustomobject]@{ GroupId = '22222222-2222-2222-2222-222222222222'; GroupName = 'SP - UK Finance Team'; GroupDescription = 'd'; GroupKind = 'Security'; MemberId = 'u2'; MemberName = 'Bo Chen'; MemberUPN = 'bo@contoso.co.uk'; MemberType = 'User'; AccountEnabled = 'True' }
        [pscustomobject]@{ GroupId = '22222222-2222-2222-2222-222222222222'; GroupName = 'SP - UK Finance Team'; GroupDescription = 'd'; GroupKind = 'Security'; MemberId = 'u3'; MemberName = 'Cy Old'; MemberUPN = 'cy@contoso.co.uk'; MemberType = 'User'; AccountEnabled = 'False' }
        [pscustomobject]@{ GroupId = '33333333-3333-3333-3333-333333333333'; GroupName = 'SP - UK Auditor'; GroupDescription = 'd'; GroupKind = 'Security'; MemberId = ''; MemberName = '(empty)'; MemberUPN = ''; MemberType = ''; AccountEnabled = '' }
    )
    $g | Export-Csv -LiteralPath (New-SpRbacOutputPath -Config $c -Name 'entra-groups') -Encoding utf8
    @([pscustomobject]@{ DisplayName = 'Ann Lee | Contoso'; UserPrincipalName = 'ann@contoso.co.uk'; JobTitle = 'General Manager'; Department = ''; CompanyName = ''; UserType = 'Member'; AccountEnabled = 'True'; InScope = 'True'; GroupCount = 1; Groups = 'SP - UK General Manager' }) |
        Export-Csv -LiteralPath (New-SpRbacOutputPath -Config $c -Name 'entra-users') -Encoding utf8
    $idx = Get-SpRbacGroupIndex -Config $c
    Check 'group index maps names to ids and back' { $idx.ByName['SP - UK Finance Team'] -eq '22222222-2222-2222-2222-222222222222' -and $idx.ById['11111111-1111-1111-1111-111111111111'] -eq 'SP - UK General Manager' }

    $py = if ($Python) { $Python } else { @('python3', 'python') | Where-Object { Get-Command $_ -ErrorAction SilentlyContinue } | Select-Object -First 1 }
    $hasPkgs = $false
    if ($py) { & $py -c 'import xlsxwriter, openpyxl' 2>$null; $hasPkgs = ($LASTEXITCODE -eq 0) }
    if (-not $hasPkgs) { Write-Host '  [skip] report tests: Python with xlsxwriter and openpyxl not found (pass -Python <path>)' -ForegroundColor Yellow }
    else {
        $out = & $py (Join-Path $root 'report/build_report.py') --config $tmp 2>&1
        $xlsx = Get-ChildItem (Join-Path $tmp 'reports') -Filter '*.xlsx' -ErrorAction SilentlyContinue | Select-Object -First 1
        Check 'report builder writes a workbook' { $LASTEXITCODE -eq 0 -and $xlsx }
        Check 'report lists the empty role group' { "$out" -match 'Empty role groups:.*Auditor' }
        $probe = "import openpyxl,sys; wb=openpyxl.load_workbook(sys.argv[1]); ws=wb['User → Role Groups']; rows=[[c for c in r] for r in ws.iter_rows(min_row=3,values_only=True)]; print(len(rows), rows[0][0], rows[0][1], '|'.join(str(r[0]) for r in rows)); print(','.join(wb.sheetnames))"
        $res = & $py -c $probe $xlsx.FullName
        Check 'report: names are shortened, job title filled, disabled account left out' { "$($res[0])" -eq '2 Ann Lee General Manager Ann Lee|Bo Chen' }
        Check 'report: all seven tabs' { "$($res[1])" -eq 'Access Matrix,User → Role Groups,Groups → Members,Users → Groups,Summary,No Access,Notes' }
        $imp = & $py (Join-Path $root 'report/import_matrix.py') --config $tmp $xlsx.FullName 2>&1
        Check 'matrix round trip: report tab equals the matrix file' { "$imp" -match 'No differences' }
    }
}
finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:failed) { Write-Host "RESULT: $($script:failed) failed, $($script:passed) passed." -ForegroundColor Red; exit 1 }
Write-Host "RESULT: all $($script:passed) checks passed." -ForegroundColor Green
