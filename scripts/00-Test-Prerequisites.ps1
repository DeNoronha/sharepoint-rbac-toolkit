<#
.SYNOPSIS
    Checks that this computer has everything the toolkit needs. Changes nothing and does
    not sign in anywhere.

.DESCRIPTION
    Checks:
      - PowerShell version (7.4 or newer)
      - PowerShell modules PnP.PowerShell and Microsoft.Graph.Authentication
      - Python 3 with the packages xlsxwriter and openpyxl (for the Excel report)
      - with -ConfigPath: that the config and the matrix are valid, and what they contain

    Sign-in and permissions are not tested here. 01-Export-EntraSnapshot.ps1 tests the
    Entra side and 02-Audit-Permissions.ps1 the SharePoint side; both are read-only.

.PARAMETER ConfigPath
    Optional. Path to config.json or to the folder that contains it.

.PARAMETER Python
    Optional. Python executable to check. Default: python3, then python.

.EXAMPLE
    pwsh -NoProfile -File ./scripts/00-Test-Prerequisites.ps1 -ConfigPath ./config/contoso
#>
[CmdletBinding()]
param([string]$ConfigPath, [string]$Python)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'lib/SpRbac.psm1') -Force
Write-Host "`n=== CHECK PREREQUISITES ===`n" -ForegroundColor Cyan
$fail = 0; $warn = 0
function Pass($t) { Write-Host "  [ok]    $t" -ForegroundColor Green }
function Fail($t, $fix) { Write-Host "  [FAIL]  $t" -ForegroundColor Red; if ($fix) { Write-Host "          fix: $fix" -ForegroundColor Gray }; $script:fail++ }
function Warn($t, $fix) { Write-Host "  [warn]  $t" -ForegroundColor Yellow; if ($fix) { Write-Host "          $fix" -ForegroundColor Gray }; $script:warn++ }

Write-Host 'PowerShell' -ForegroundColor White
$v = $PSVersionTable.PSVersion
if ($v -ge [version]'7.4') { Pass "PowerShell $v" } else { Fail "PowerShell $v is too old" 'install PowerShell 7.4 or newer: https://aka.ms/powershell' }
foreach ($m in @(@{ Name = 'PnP.PowerShell'; Min = '3.0.0'; For = 'SharePoint' }, @{ Name = 'Microsoft.Graph.Authentication'; Min = '2.0.0'; For = 'Entra ID' })) {
    $found = Get-Module -ListAvailable -Name $m.Name | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $found) { Fail "$($m.Name) is not installed (needed for $($m.For))" "Install-Module $($m.Name) -Scope CurrentUser" }
    elseif ($found.Version -lt [version]$m.Min) { Fail "$($m.Name) $($found.Version) is older than $($m.Min)" "Update-Module $($m.Name)" }
    else { Pass "$($m.Name) $($found.Version)" }
}

Write-Host "`nPython (only for the Excel report)" -ForegroundColor White
$candidates = if ($Python) { @($Python) } else { @('python3', 'python') }
$py = $null
foreach ($c in $candidates) { if (Get-Command $c -ErrorAction SilentlyContinue) { $py = $c; break } }
if (-not $py) { Warn 'Python 3 not found' 'install Python 3 if you want the Excel report; the PowerShell scripts work without it' }
else {
    $pv = (& $py --version 2>&1) -join ' '
    if ($pv -match 'Python 3\.(\d+)' -and [int]$Matches[1] -ge 8) { Pass "$pv ($py)" } else { Warn "$pv ($py) - Python 3.8 or newer is expected" '' }
    foreach ($pkg in 'xlsxwriter', 'openpyxl') {
        & $py -c "import $pkg" 2>$null
        if ($LASTEXITCODE -eq 0) { Pass "Python package $pkg" } else { Warn "Python package $pkg is missing for $py" "$py -m pip install $pkg   (or pass another interpreter with -Python)" }
    }
}

if ($ConfigPath) {
    Write-Host "`nConfiguration" -ForegroundColor White
    try {
        $cfg = Get-SpRbacConfig -Path $ConfigPath
        Pass "Config '$($cfg.Name)' is valid ($($cfg.ConfigPath))"
        Pass "Matrix: $($cfg.Roles.Count) roles x $(@($cfg.Targets | Where-Object InMatrix).Count) targets ($(Split-Path $cfg.MatrixPath -Leaf))"
        Write-Host "          model        : $($cfg.Model)" -ForegroundColor Gray
        Write-Host "          tenant       : $($cfg.TenantUrl)  (sign-in tenant: $(if ($cfg.TenantId) { $cfg.TenantId } else { 'not set' }))" -ForegroundColor Gray
        Write-Host "          group prefix : '$($cfg.GroupPrefix)'" -ForegroundColor Gray
        Write-Host "          sites        : $((Get-SpRbacSites -Config $cfg).Count)" -ForegroundColor Gray
        $noSite = @($cfg.Targets | Where-Object Scope -eq 'None')
        if ($noSite) { Warn "Targets without a site or library: $($noSite.Name -join ', ')" 'their groups are managed, but nothing is checked in SharePoint for them' }
        $emptyRoles = @($cfg.Roles | Where-Object { -not $cfg.Matrix[$_].Count })
        if ($emptyRoles) { Warn "Roles without any mark in the matrix: $($emptyRoles -join ', ')" '' }
        $emptyTargets = @($cfg.Targets | Where-Object { $_.InMatrix -and -not @(Get-SpRbacRolesForTarget -Config $cfg -Target $_).Count -and -not $_.FixedGroups.Count })
        if ($emptyTargets) { Warn "Targets that no role has access to: $($emptyTargets.Name -join ', ')" '' }
        if (-not $cfg.AdminGroups.Count) { Warn "No 'adminGroups' in the config" 'without an admin group nobody but site collection admins can manage a library after a sync' }
        $snap = Get-SpRbacLatestOutput -Config $cfg -Name 'entra-groups'
        if ($snap) { Pass "Entra snapshot present: $($snap.Name)" } else { Warn 'No Entra snapshot yet' 'run 01-Export-EntraSnapshot.ps1 first' }
    }
    catch { Fail "$($_.Exception.Message)" '' }
}

Write-Host ''
if ($fail) { Write-Host "RESULT: $fail problem(s) to fix, $warn warning(s)." -ForegroundColor Red; exit 1 }
elseif ($warn) { Write-Host "RESULT: ready, with $warn warning(s)." -ForegroundColor Yellow }
else { Write-Host 'RESULT: ready.' -ForegroundColor Green }
