# SpRbac.psm1 - shared functions for the SharePoint RBAC toolkit.
#
# This module deliberately imports neither PnP.PowerShell nor Microsoft.Graph. The two
# modules ship conflicting assemblies and cannot be loaded into the same PowerShell
# process, so every script imports exactly one of them and runs in its own process.

$script:LevelToRoleType = [ordered]@{
    'Read'         = 'Reader'
    'Contribute'   = 'Contributor'
    'Edit'         = 'Editor'
    'Design'       = 'WebDesigner'
    'Full Control' = 'Administrator'
}

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

function Get-SpRbacConfig {
    <#
    .SYNOPSIS
        Loads and validates a config.json plus its access matrix.
    .PARAMETER Path
        Path to config.json, or to the folder that contains it.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (Test-Path -LiteralPath $Path -PathType Container) { $Path = Join-Path $Path 'config.json' }
    if (-not (Test-Path -LiteralPath $Path)) { throw "Config file not found: $Path" }
    $full = (Resolve-Path -LiteralPath $Path).Path
    $dir  = Split-Path $full -Parent
    try { $raw = Get-Content -LiteralPath $full -Raw | ConvertFrom-Json -AsHashtable }
    catch { throw "Config file is not valid JSON: $full`n$($_.Exception.Message)" }

    $problems = [System.Collections.Generic.List[string]]::new()
    foreach ($key in 'name', 'tenantUrl', 'clientId', 'groupPrefix', 'model', 'targets') {
        if (-not $raw.ContainsKey($key) -or $null -eq $raw[$key] -or "$($raw[$key])" -eq '') { $problems.Add("Missing required setting '$key'.") }
    }
    if ($problems.Count) { throw "Config $full is incomplete:`n - " + ($problems -join "`n - ") }

    $model = "$($raw.model)".ToLower()
    if ($model -notin 'direct', 'nested') { $problems.Add("'model' must be 'direct' or 'nested' (found '$($raw.model)').") }
    if ("$($raw.clientId)" -notmatch '^[0-9a-fA-F-]{36}$') { $problems.Add("'clientId' must be the GUID of your PnP app registration.") }
    $tenantUrl = "$($raw.tenantUrl)".TrimEnd('/')
    if ($tenantUrl -notmatch '^https://[^/]+$') { $problems.Add("'tenantUrl' must look like https://contoso.sharepoint.com (no path).") }

    $prefix         = "$($raw.groupPrefix)"
    $defaultLevel   = if ($raw.defaultPermission) { "$($raw.defaultPermission)" } else { 'Contribute' }
    $rolePerms      = if ($raw.rolePermissions) { $raw.rolePermissions } else { @{} }
    $matrixFile     = if ($raw.matrixFile) { "$($raw.matrixFile)" } else { 'matrix.csv' }
    $matrixPath     = if ([IO.Path]::IsPathRooted($matrixFile)) { $matrixFile } else { Join-Path $dir $matrixFile }
    $outputDir      = if ($raw.outputDir) { "$($raw.outputDir)" } else { 'output' }
    if (-not [IO.Path]::IsPathRooted($outputDir)) { $outputDir = Join-Path $dir $outputDir }
    $snapPrefixes   = if ($raw.snapshotPrefixes) { @($raw.snapshotPrefixes | ForEach-Object { "$_" }) } else { @($prefix.TrimEnd()) }

    $adminGroups = @(foreach ($a in @($raw.adminGroups)) {
            if ($null -eq $a) { continue }
            if ($a -is [string]) { [pscustomobject]@{ Group = $a; Permission = 'Full Control' } }
            else { [pscustomobject]@{ Group = "$($a.group)"; Permission = $(if ($a.permission) { "$($a.permission)" } else { 'Full Control' }) } }
        })

    # Matrix
    $roles = @(); $matrix = @{}; $matrixTargets = @()
    if (-not (Test-Path -LiteralPath $matrixPath)) { $problems.Add("Matrix file not found: $matrixPath") }
    else {
        $m = Import-SpRbacMatrix -Path $matrixPath
        $roles = $m.Roles; $matrix = $m.Cells; $matrixTargets = $m.Targets
    }

    # Targets
    $targets = [System.Collections.Generic.List[object]]::new()
    foreach ($tName in $raw.targets.Keys) {
        $t = $raw.targets[$tName]; if ($null -eq $t) { $t = @{} }
        $site = if ($t.site) { "$($t.site)" } elseif ($raw.defaultSite) { "$($raw.defaultSite)" } else { '' }
        if ($site -and $site -notmatch '^https://') { $site = $tenantUrl + '/' + $site.TrimStart('/') }
        $site = $site.TrimEnd('/')
        $scope = if ($t.library) { 'Library' } elseif ($site) { 'Site' } else { 'None' }
        if ($model -eq 'direct' -and $scope -eq 'None') { $problems.Add("Target '$tName' needs a 'site' (or a 'defaultSite' in the config).") }
        $fixed = @(foreach ($f in @($t.fixedGroups)) {
                if ($null -eq $f) { continue }
                if ($f -is [string]) { [pscustomobject]@{ Group = $f; Permission = $defaultLevel } }
                else { [pscustomobject]@{ Group = "$($f.group)"; Permission = $(if ($f.permission) { "$($f.permission)" } else { $defaultLevel }) } }
            })
        $targets.Add([pscustomobject]@{
                Name         = "$tName"
                Scope        = $scope
                Site         = $site
                Library      = if ($t.library) { "$($t.library)" } else { '' }
                Aliases      = @($t.aliases | Where-Object { $_ } | ForEach-Object { "$_" })
                AccessGroup  = if ($t.accessGroup) { "$($t.accessGroup)" } else { $prefix + $tName }
                Permission   = if ($t.permission) { "$($t.permission)" } else { '' }
                FixedGroups  = $fixed
                InMatrix     = ($matrixTargets -contains "$tName")
            })
    }
    $targetNames = @($targets | ForEach-Object Name)
    foreach ($mt in $matrixTargets) { if ($targetNames -notcontains $mt) { $problems.Add("Matrix column '$mt' has no entry under 'targets' in the config.") } }
    foreach ($t in $targets) { if (-not $t.InMatrix -and -not $t.FixedGroups.Count) { $problems.Add("Target '$($t.Name)' is not a column in the matrix and has no 'fixedGroups'.") } }
    foreach ($r in $rolePerms.Keys) { if ($roles -notcontains $r) { $problems.Add("rolePermissions mentions '$r', which is not a row in the matrix.") } }

    if ($problems.Count) { throw "Config $full has problems:`n - " + ($problems -join "`n - ") }

    [pscustomobject]@{
        Name              = "$($raw.name)"
        ConfigPath        = $full
        ConfigDir         = $dir
        TenantUrl         = $tenantUrl
        TenantId          = if ($raw.tenantId) { "$($raw.tenantId)" } elseif ($tenantUrl -match '^https://([^.]+)\.sharepoint\.com$') { "$($Matches[1]).onmicrosoft.com" } else { '' }
        ClientId          = "$($raw.clientId)"
        PnPAuth           = if ($raw.pnpAuth) { "$($raw.pnpAuth)" } else { 'Interactive' }
        PersistLogin      = if ($raw.ContainsKey('persistLogin')) { [bool]$raw.persistLogin } else { $true }
        GroupPrefix       = $prefix
        SnapshotPrefixes  = $snapPrefixes
        Model             = $model
        DefaultPermission = $defaultLevel
        RolePermissions   = $rolePerms
        AdminGroups       = $adminGroups
        EveryoneGroup     = if ($raw.everyoneGroup) { "$($raw.everyoneGroup)" } else { '' }
        ExceptionGroups   = @($raw.exceptionGroups | Where-Object { $_ } | ForEach-Object { "$_" })
        KeepPrincipals    = @($raw.keepPrincipals | Where-Object { $_ } | ForEach-Object { "$_" })
        ExtraSites        = @($raw.extraSites | Where-Object { $_ } | ForEach-Object { $s = "$_"; if ($s -notmatch '^https://') { $s = $tenantUrl + '/' + $s.TrimStart('/') }; $s.TrimEnd('/') })
        UserFilter        = if ($raw.userFilter) { $raw.userFilter } else { @{} }
        GroupDescription  = if ($raw.groupDescription) { "$($raw.groupDescription)" } else { 'Security group for SharePoint authorization: {name}' }
        Report            = if ($raw.report) { $raw.report } else { @{} }
        MatrixPath        = $matrixPath
        OutputDir         = $outputDir
        Roles             = $roles
        Matrix            = $matrix
        Targets           = $targets.ToArray()
    }
}

function Import-SpRbacMatrix {
    <#
    .SYNOPSIS
        Reads the access matrix: first column = role, other columns = targets.
        An empty cell means no access. Any mark (x, yes, 1, a tick) means access at the
        default level; a permission level name (Read, Contribute, Edit, ...) sets that level.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $lines = @(Get-Content -LiteralPath $Path -Encoding utf8 | Where-Object { $_.Trim() -ne '' })
    if ($lines.Count -lt 1) { throw "Matrix file is empty: $Path" }
    $header = $lines[0]
    $delim = if (($header.Split(';').Count) -gt ($header.Split(',').Count)) { ';' } else { ',' }
    $rows = @($lines | ConvertFrom-Csv -Delimiter $delim)
    $cols = @(($lines[0] | ConvertFrom-Csv -Delimiter $delim -Header (1..200 | ForEach-Object { "c$_" })).PSObject.Properties |
        Where-Object { $null -ne $_.Value -and "$($_.Value)".Trim() -ne '' } | ForEach-Object { "$($_.Value)".Trim() })
    $roleCol = $cols[0]
    $targets = @($cols | Select-Object -Skip 1)
    $dupes = @($targets | Group-Object | Where-Object Count -gt 1 | ForEach-Object Name)
    if ($dupes.Count) { throw "Matrix has duplicate columns: $($dupes -join ', ')" }

    $roles = [System.Collections.Generic.List[string]]::new()
    $cells = @{}
    foreach ($row in $rows) {
        $role = "$($row.$roleCol)".Trim()
        if (-not $role) { continue }
        if ($roles.Contains($role)) { throw "Matrix has duplicate role row: $role" }
        $roles.Add($role)
        $cells[$role] = @{}
        foreach ($t in $targets) {
            $v = "$($row.$t)".Trim()
            if ($v -eq '' -or $v -in '-', '0', 'no', 'n', 'false') { continue }
            $level = $script:LevelToRoleType.Keys | Where-Object { $_ -eq $v } | Select-Object -First 1
            $cells[$role][$t] = if ($level) { $level } else { '*' }   # '*' = default level
        }
    }
    [pscustomobject]@{ Roles = $roles.ToArray(); Targets = $targets; Cells = $cells }
}

function Get-SpRbacRolesForTarget {
    # Roles that have a mark for this target, with the effective permission level.
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)]$Target)
    foreach ($role in $Config.Roles) {
        $cell = $Config.Matrix[$role][$Target.Name]
        if (-not $cell) { continue }
        $level = if ($cell -ne '*') { $cell }
        elseif ($Config.RolePermissions.ContainsKey($role)) { "$($Config.RolePermissions[$role])" }
        elseif ($Target.Permission) { $Target.Permission }
        else { $Config.DefaultPermission }
        [pscustomobject]@{ Role = $role; Group = $Config.GroupPrefix + $role; Permission = $level }
    }
}

function Get-SpRbacExpectedGroups {
    <#
    .SYNOPSIS
        The Entra groups that should hold access on a target, with their permission level.
        Kind = Access (from the matrix or fixedGroups) or Admin (adminGroups).
    #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)]$Target)
    $out = [ordered]@{}
    if ($Config.Model -eq 'nested' -and $Target.InMatrix) {
        $level = if ($Target.Permission) { $Target.Permission } else { $Config.DefaultPermission }
        $out[$Target.AccessGroup] = [pscustomobject]@{ Group = $Target.AccessGroup; Permission = $level; Kind = 'Access' }
    }
    elseif ($Target.InMatrix) {
        foreach ($r in (Get-SpRbacRolesForTarget -Config $Config -Target $Target)) {
            $out[$r.Group] = [pscustomobject]@{ Group = $r.Group; Permission = $r.Permission; Kind = 'Access' }
        }
    }
    foreach ($f in $Target.FixedGroups) { $out[$f.Group] = [pscustomobject]@{ Group = $f.Group; Permission = $f.Permission; Kind = 'Access' } }
    foreach ($a in $Config.AdminGroups) { $out[$a.Group] = [pscustomobject]@{ Group = $a.Group; Permission = $a.Permission; Kind = 'Admin' } }
    @($out.Values)
}

function Get-SpRbacSites {
    # Unique site URLs touched by this config (targets plus extraSites).
    param([Parameter(Mandatory)]$Config)
    @(@($Config.Targets | Where-Object Site | ForEach-Object Site) + @($Config.ExtraSites) | Where-Object { $_ } | Select-Object -Unique)
}

# ─────────────────────────────────────────────────────────────────────────────
# Output and console helpers
# ─────────────────────────────────────────────────────────────────────────────

function New-SpRbacOutputPath {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$Name, [string]$Extension = 'csv')
    if (-not (Test-Path -LiteralPath $Config.OutputDir)) { New-Item -ItemType Directory -Path $Config.OutputDir -Force | Out-Null }
    Join-Path $Config.OutputDir ("{0}_{1}.{2}" -f $Name, (Get-Date -Format 'yyyyMMdd_HHmmss'), $Extension)
}

function Get-SpRbacLatestOutput {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$Name)
    if (-not (Test-Path -LiteralPath $Config.OutputDir)) { return $null }
    Get-ChildItem -LiteralPath $Config.OutputDir -Filter "$($Name)_*.csv" | Sort-Object Name | Select-Object -Last 1
}

function Write-SpRbacHeader {
    param([Parameter(Mandatory)][string]$Title, $Config, [switch]$ReadOnly, [bool]$Apply)
    Write-Host "`n=== $Title ===" -ForegroundColor Cyan
    if ($Config) { Write-Host "  Config : $($Config.Name)  ($($Config.ConfigPath))" -ForegroundColor Gray }
    if ($ReadOnly) { Write-Host "  Mode   : READ-ONLY - this script never changes anything" -ForegroundColor Gray }
    elseif ($Apply) { Write-Host "  Mode   : APPLY - changes will be made" -ForegroundColor Red }
    else { Write-Host "  Mode   : DRY RUN - nothing will be changed (add -Apply to apply)" -ForegroundColor Yellow }
    Write-Host ''
}

function Write-SpRbacFooter {
    param([bool]$Apply, [int]$Planned = 0, [int]$Done = 0, [int]$Errors = 0)
    Write-Host ''
    if ($Apply) {
        $color = if ($Errors) { 'Red' } else { 'Green' }
        Write-Host "APPLIED: $Done change(s) made, $Errors error(s)." -ForegroundColor $color
    }
    elseif ($Planned) { Write-Host "DRY RUN: $Planned change(s) would be made. Nothing was changed. Re-run with -Apply to apply." -ForegroundColor Yellow }
    else { Write-Host 'DRY RUN: nothing to change.' -ForegroundColor Green }
}

function Assert-SpRbacModule {
    # Verifies that the wanted module is installed and the conflicting one is not loaded.
    param([Parameter(Mandatory)][ValidateSet('PnP.PowerShell', 'Microsoft.Graph.Authentication')][string]$Name)
    $other = if ($Name -eq 'PnP.PowerShell') { 'Microsoft.Graph.Authentication' } else { 'PnP.PowerShell' }
    if (Get-Module -Name $other) {
        throw "$other is already loaded in this PowerShell session. It conflicts with $Name. Start the script in a fresh process: pwsh -NoProfile -File <script>"
    }
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "PowerShell module $Name is not installed. Run: Install-Module $Name -Scope CurrentUser"
    }
    Import-Module $Name -ErrorAction Stop
}

# ─────────────────────────────────────────────────────────────────────────────
# Entra snapshot (written by 01-Export-EntraSnapshot.ps1, read by the SharePoint scripts)
# ─────────────────────────────────────────────────────────────────────────────

function Get-SpRbacGroupIndex {
    <#
    .SYNOPSIS
        Name <-> object id lookup for Entra groups, read from the latest snapshot.
        The SharePoint scripts use it because they cannot query Microsoft Graph themselves.
    #>
    param([Parameter(Mandatory)]$Config, [int]$MaxAgeHours = 24)
    $file = Get-SpRbacLatestOutput -Config $Config -Name 'entra-groups'
    if (-not $file) { throw "No Entra snapshot found in $($Config.OutputDir). Run 01-Export-EntraSnapshot.ps1 first." }
    $age = [math]::Round(((Get-Date) - $file.LastWriteTime).TotalHours, 1)
    if ($age -gt $MaxAgeHours) { Write-Host "  WARNING: the Entra snapshot is $age hours old ($($file.Name)). Run 01-Export-EntraSnapshot.ps1 for a fresh one." -ForegroundColor Yellow }
    $rows = @(Import-Csv -LiteralPath $file.FullName -Encoding utf8)
    $byName = @{}; $byId = @{}
    foreach ($r in $rows) { $byName[$r.GroupName] = $r.GroupId.ToLower(); $byId[$r.GroupId.ToLower()] = $r.GroupName }
    [pscustomobject]@{ Path = $file.FullName; AgeHours = $age; ByName = $byName; ById = $byId; Rows = $rows }
}

# ─────────────────────────────────────────────────────────────────────────────
# Microsoft Graph helpers (REST only: the typed Graph cmdlets are not needed)
# ─────────────────────────────────────────────────────────────────────────────

function Connect-SpRbacGraph {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string[]]$Scopes)
    Assert-SpRbacModule -Name 'Microsoft.Graph.Authentication'
    $p = @{ Scopes = $Scopes; NoWelcome = $true }
    if ($Config.TenantId) { $p.TenantId = $Config.TenantId }
    Write-Host 'Signing in to Microsoft Graph (a browser window may open) ...' -ForegroundColor Cyan
    Connect-MgGraph @p
    $ctx = Get-MgContext
    if (-not $ctx -or -not $ctx.Account) {
        throw 'Sign-in to Microsoft Graph did not complete. Finish the sign-in in the browser window (it may be behind other windows) and run the script again.'
    }
    Write-Host "  Signed in as $($ctx.Account)`n" -ForegroundColor Green
}

function Invoke-SpRbacGraphPaged {
    param([Parameter(Mandatory)][string]$Uri)
    $all = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -Headers @{ ConsistencyLevel = 'eventual' }
        foreach ($v in $resp.value) { $all.Add($v) }
        $next = $resp.'@odata.nextLink'
    }
    $all.ToArray()   # enumerated on purpose: callers wrap the call in @()
}

function Get-SpRbacGraphGroups {
    # All groups whose display name starts with one of the prefixes, plus named extras.
    param([string[]]$Prefixes, [string[]]$Names)
    $select = 'id,displayName,description,securityEnabled,mailEnabled,groupTypes'
    $found = @{}
    foreach ($p in $Prefixes) {
        $esc = $p.Replace("'", "''")
        foreach ($g in (Invoke-SpRbacGraphPaged "https://graph.microsoft.com/v1.0/groups?`$select=$select&`$count=true&`$filter=startsWith(displayName,'$esc')")) { $found[$g.id] = $g }
    }
    foreach ($n in ($Names | Where-Object { $_ } | Select-Object -Unique)) {
        if ($found.Values | Where-Object { $_.displayName -eq $n }) { continue }
        $esc = $n.Replace("'", "''")
        foreach ($g in (Invoke-SpRbacGraphPaged "https://graph.microsoft.com/v1.0/groups?`$select=$select&`$count=true&`$filter=displayName eq '$esc'")) { $found[$g.id] = $g }
    }
    @($found.Values | Sort-Object { $_.displayName })
}

function Get-SpRbacGraphGroupByName {
    # Exactly one group with this display name, or $null. Throws when the name is ambiguous.
    param([Parameter(Mandatory)][string]$Name)
    $esc = $Name.Replace("'", "''")
    $hits = @(Invoke-SpRbacGraphPaged "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,description&`$count=true&`$filter=displayName eq '$esc'")
    if ($hits.Count -gt 1) { throw "More than one Entra group is named '$Name'. Rename one of them first." }
    if ($hits.Count -eq 1) { return $hits[0] }
    $null
}

function Get-SpRbacGraphUser {
    # $null when the account does not exist.
    param([Parameter(Mandatory)][string]$Upn)
    try { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Upn))?`$select=id,displayName,userPrincipalName,accountEnabled,jobTitle" }
    catch { if ("$($_.Exception.Message)" -match 'NotFound|404|does not exist') { $null } else { throw } }
}

function Add-SpRbacGraphGroupMember {
    param([Parameter(Mandatory)][string]$GroupId, [Parameter(Mandatory)][string]$MemberId)
    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$MemberId" } | ConvertTo-Json
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref" -Body $body -ContentType 'application/json' | Out-Null
}

function Remove-SpRbacGraphGroupMember {
    param([Parameter(Mandatory)][string]$GroupId, [Parameter(Mandatory)][string]$MemberId)
    Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/$MemberId/`$ref" | Out-Null
}

function New-SpRbacGraphSecurityGroup {
    param([Parameter(Mandatory)][string]$Name, [string]$Description)
    $nick = ($Name -replace '[^A-Za-z0-9]', '-' -replace '-+', '-').Trim('-')
    if ($nick.Length -gt 60) { $nick = $nick.Substring(0, 60) }
    $body = @{ displayName = $Name; description = $Description; mailEnabled = $false; mailNickname = $nick; securityEnabled = $true } | ConvertTo-Json
    Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/groups' -Body $body -ContentType 'application/json'
}

# ─────────────────────────────────────────────────────────────────────────────
# SharePoint helpers (PnP.PowerShell + CSOM)
# ─────────────────────────────────────────────────────────────────────────────

function Connect-SpRbacSite {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$SiteUrl)
    $p = @{ Url = $SiteUrl; ClientId = $Config.ClientId }
    if ($Config.TenantId) { $p.Tenant = $Config.TenantId }
    if ($Config.PnPAuth -eq 'DeviceLogin') { $p.DeviceLogin = $true } else { $p.Interactive = $true }
    # PersistLogin keeps the token in the operating system's credential store, so only the
    # first run opens a browser. Clear it with: Disconnect-PnPOnline -ClearPersistedLogin
    if ($Config.PersistLogin) { $p.PersistLogin = $true }
    if (-not $script:SignInHintShown) { Write-Host '  Signing in to SharePoint (a browser window may open; if nothing happens, look for it) ...' -ForegroundColor DarkGray; $script:SignInHintShown = $true }
    Connect-PnPOnline @p
}

function Get-SpRbacPrincipalInfo {
    <#
    .SYNOPSIS
        Classifies a SharePoint principal by its login name.
        Kind: EntraGroup, M365Group, User, Guest, Everyone, SharingLink, SharePointGroup, System, Other.
    #>
    param([Parameter(Mandatory)]$Member)
    $login = "$($Member.LoginName)"; $title = "$($Member.Title)"; $ptype = "$($Member.PrincipalType)"
    $kind = 'Other'; $id = ''
    if ($login -match 'c:0t\.c\|tenant\|([0-9a-fA-F-]{36})') { $kind = 'EntraGroup'; $id = $Matches[1].ToLower() }
    elseif ($login -match 'c:0o\.c\|federateddirectoryclaimprovider\|([0-9a-fA-F-]{36})') { $kind = 'M365Group'; $id = $Matches[1].ToLower() }
    elseif ($login -match '^i:0#\.f\|membership\|(.+)$') { $id = $Matches[1].ToLower(); $kind = if ($id.Contains('#ext#')) { 'Guest' } else { 'User' } }
    elseif ($login -eq 'c:0(.s|true' -or $login -match 'spo-grid-all-users') { $kind = 'Everyone' }
    elseif ($login -match '^SHAREPOINT\\system$' -or $login -match 'app@sharepoint' -or $login -match '^i:0i\.t\|') { $kind = 'System' }
    elseif ($ptype -eq 'SharePointGroup') {
        if ($login -like 'SharingLinks.*' -or $title -like 'SharingLinks.*') { $kind = 'SharingLink' }
        elseif ($title -like 'Limited Access System Group*') { $kind = 'System' }
        else { $kind = 'SharePointGroup' }
    }
    [pscustomobject]@{ Kind = $kind; Id = $id; Login = $login; Title = $title }
}

function Get-SpRbacLevels {
    # Visible permission levels of a role assignment, as canonical English names where possible.
    param($Bindings)
    @(foreach ($b in $Bindings) {
            if ($b.Hidden) { continue }
            $kind = "$($b.RoleTypeKind)"
            if ($kind -eq 'Guest') { continue }
            $name = $script:LevelToRoleType.Keys | Where-Object { $script:LevelToRoleType[$_] -eq $kind } | Select-Object -First 1
            if ($name) { $name } else { "$($b.Name)" }
        })
}

function Get-SpRbacRoleDefinition {
    param([Parameter(Mandatory)]$Web, [Parameter(Mandatory)][string]$Level)
    if ($script:LevelToRoleType.Contains($Level)) {
        return $Web.RoleDefinitions.GetByType([Microsoft.SharePoint.Client.RoleType]::($script:LevelToRoleType[$Level]))
    }
    $Web.RoleDefinitions.GetByName($Level)   # custom permission level, looked up by its name
}

function Resolve-SpRbacLibrary {
    # Finds the library of a target by its 'library' setting or one of its aliases.
    param([Parameter(Mandatory)]$Target)
    foreach ($name in @($Target.Library) + @($Target.Aliases) | Where-Object { $_ }) {
        try { $l = Get-PnPList -Identity $name -ErrorAction Stop; if ($l) { return $l } } catch { }
    }
    $null
}

function Test-SpRbacKeep {
    param($Config, [string]$Title, [string]$Login)
    foreach ($pattern in $Config.KeepPrincipals) { if ($Title -like $pattern -or $Login -like $pattern) { return $true } }
    $false
}

function New-SpRbacDiffRow {
    param($Target, [string]$Container, [string]$Action, [string]$Principal, [string]$PrincipalType, [string]$EntraId,
        [string]$Login, [string]$CurrentLevel, [string]$ExpectedLevel, [string]$Note, $Handle)
    $row = [pscustomobject]@{
        Target        = $Target.Name
        Scope         = $Target.Scope
        Site          = $Target.Site
        Container     = $Container
        Action        = $Action
        Principal     = $Principal
        PrincipalType = $PrincipalType
        EntraId       = $EntraId
        CurrentLevel  = $CurrentLevel
        ExpectedLevel = $ExpectedLevel
        Note          = $Note
        Login         = $Login
    }
    $row | Add-Member -NotePropertyName Handle -NotePropertyValue $Handle
    $row
}

function Get-SpRbacTargetDiff {
    <#
    .SYNOPSIS
        Compares one target (library or site) with what the matrix says. Read-only.
        Requires an active PnP connection to the target's site.

        Actions returned:
          OK                  expected group, right level
          ADD                 expected group is missing
          FIX_LEVEL           expected group, wrong permission level (library targets)
          BREAK_INHERITANCE   library still inherits from the site
          REMOVE_GROUP        Entra group that the matrix does not list
          REMOVE_USER         a person with direct access
          REMOVE_SHARINGLINK  sharing link with access on the library itself
          REMOVE_SPGROUP      SharePoint group on a library
          REMOVE_OTHER        Everyone claims, Microsoft 365 groups, anything else
          KEEP                left alone on request (switch or keepPrincipals)
          REVIEW              not managed by the toolkit; a person has to judge it
          MISSING_GROUP       the expected group is not in the Entra snapshot
          MISSING_TARGET      library or site group not found
    #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)]$Target, [Parameter(Mandatory)]$GroupIndex,
        [switch]$KeepDirectUsers, [switch]$KeepSharePointGroups)

    $expected = @{}   # entra id -> expected object
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($e in (Get-SpRbacExpectedGroups -Config $Config -Target $Target)) {
        $id = $GroupIndex.ByName[$e.Group]
        if (-not $id) { $rows.Add((New-SpRbacDiffRow -Target $Target -Action 'MISSING_GROUP' -Principal $e.Group -PrincipalType 'EntraGroup' -ExpectedLevel $e.Permission -Note 'Not in the Entra snapshot. Create the group (04-Sync-GroupNesting.ps1 -CreateMissingGroups) or refresh the snapshot.')); continue }
        $expected[$id] = $e
    }
    $entraName = { param($info) if ($GroupIndex.ById.ContainsKey($info.Id)) { $GroupIndex.ById[$info.Id] } else { $info.Title } }

    if ($Target.Scope -eq 'Library') {
        $found = Resolve-SpRbacLibrary -Target $Target
        if (-not $found) { $rows.Add((New-SpRbacDiffRow -Target $Target -Container $Target.Library -Action 'MISSING_TARGET' -Note 'Library not found on the site.')); return $rows.ToArray() }
        $list = Get-PnPList -Identity $found.Id -Includes HasUniqueRoleAssignments, RoleAssignments
        $container = $list.Title
        if (-not $list.HasUniqueRoleAssignments) {
            $rows.Add((New-SpRbacDiffRow -Target $Target -Container $container -Action 'BREAK_INHERITANCE' -Note 'Library inherits its permissions from the site. The rows below show the inherited state.' -Handle $list))
        }
        $seen = @{}
        foreach ($ra in $list.RoleAssignments) {
            Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings | Out-Null
            $info = Get-SpRbacPrincipalInfo -Member $ra.Member
            $levels = @(Get-SpRbacLevels -Bindings $ra.RoleDefinitionBindings)
            if ($info.Kind -eq 'System' -or -not $levels.Count) { continue }   # Limited Access only: a side effect of item-level sharing
            $cur = $levels -join ', '
            $handle = [pscustomobject]@{ List = $list; RoleAssignment = $ra }
            $common = @{ Target = $Target; Container = $container; Login = $info.Login; CurrentLevel = $cur; Handle = $handle }
            if ($info.Kind -eq 'EntraGroup' -and $expected.ContainsKey($info.Id)) {
                $e = $expected[$info.Id]; $seen[$info.Id] = $true
                if ($levels.Count -eq 1 -and $levels[0] -eq $e.Permission) { $rows.Add((New-SpRbacDiffRow @common -Action 'OK' -Principal $e.Group -PrincipalType $info.Kind -EntraId $info.Id -ExpectedLevel $e.Permission)) }
                else { $rows.Add((New-SpRbacDiffRow @common -Action 'FIX_LEVEL' -Principal $e.Group -PrincipalType $info.Kind -EntraId $info.Id -ExpectedLevel $e.Permission)) }
                continue
            }
            if (Test-SpRbacKeep -Config $Config -Title $info.Title -Login $info.Login) { $rows.Add((New-SpRbacDiffRow @common -Action 'KEEP' -Principal $info.Title -PrincipalType $info.Kind -EntraId $info.Id -Note 'Matches keepPrincipals in the config.')); continue }
            switch ($info.Kind) {
                'EntraGroup' { $rows.Add((New-SpRbacDiffRow @common -Action 'REMOVE_GROUP' -Principal (& $entraName $info) -PrincipalType $info.Kind -EntraId $info.Id -Note 'Group is not in the matrix for this target.')) }
                { $_ -in 'User', 'Guest' } {
                    if ($KeepDirectUsers) { $rows.Add((New-SpRbacDiffRow @common -Action 'KEEP' -Principal $info.Title -PrincipalType $info.Kind -EntraId $info.Id -Note 'Direct user, kept because of -KeepDirectUsers.')) }
                    else { $rows.Add((New-SpRbacDiffRow @common -Action 'REMOVE_USER' -Principal $info.Title -PrincipalType $info.Kind -EntraId $info.Id -Note 'Person with direct access.')) }
                }
                'SharingLink' { $rows.Add((New-SpRbacDiffRow @common -Action 'REMOVE_SHARINGLINK' -Principal $info.Title -PrincipalType $info.Kind -Note 'Sharing link on the library itself.')) }
                'SharePointGroup' {
                    if ($KeepSharePointGroups) { $rows.Add((New-SpRbacDiffRow @common -Action 'KEEP' -Principal $info.Title -PrincipalType $info.Kind -Note 'SharePoint group, kept because of -KeepSharePointGroups.')) }
                    else { $rows.Add((New-SpRbacDiffRow @common -Action 'REMOVE_SPGROUP' -Principal $info.Title -PrincipalType $info.Kind -Note 'SharePoint group on a library that is governed by the matrix.')) }
                }
                default { $rows.Add((New-SpRbacDiffRow @common -Action 'REMOVE_OTHER' -Principal $info.Title -PrincipalType $info.Kind -EntraId $info.Id -Note 'Not an Entra security group from the matrix.')) }
            }
        }
        foreach ($id in $expected.Keys) {
            if ($seen.ContainsKey($id)) { continue }
            $e = $expected[$id]
            $rows.Add((New-SpRbacDiffRow -Target $Target -Container $container -Action 'ADD' -Principal $e.Group -PrincipalType 'EntraGroup' -EntraId $id -Login "c:0t.c|tenant|$id" -ExpectedLevel $e.Permission -Handle ([pscustomobject]@{ List = $list; RoleAssignment = $null })))
        }
        return $rows.ToArray()
    }

    if ($Target.Scope -eq 'Site') {
        $web = Get-PnPWeb -Includes AssociatedMemberGroup, AssociatedOwnerGroup, AssociatedVisitorGroup, RoleAssignments
        $assoc = [ordered]@{ Members = $web.AssociatedMemberGroup; Owners = $web.AssociatedOwnerGroup; Visitors = $web.AssociatedVisitorGroup }
        $assocIds = @{}
        foreach ($role in $assoc.Keys) {
            $spg = $assoc[$role]
            $missing = $false
            try { if ($null -eq $spg -or $null -eq $spg.Id) { $missing = $true } } catch { $missing = $true }
            if ($missing) {
                if ($role -ne 'Visitors') { $rows.Add((New-SpRbacDiffRow -Target $Target -Container "($role group)" -Action 'MISSING_TARGET' -Note "The site has no associated $role group.")) }
                continue
            }
            $assocIds[[int]$spg.Id] = $true
            $wantKind = if ($role -eq 'Members') { 'Access' } elseif ($role -eq 'Owners') { 'Admin' } else { 'None' }
            $seen = @{}
            foreach ($m in @(Get-PnPGroupMember -Group $spg.Id)) {
                $info = Get-SpRbacPrincipalInfo -Member $m
                if ($info.Kind -eq 'System') { continue }
                $handle = [pscustomobject]@{ SPGroup = $spg }
                $common = @{ Target = $Target; Container = $spg.Title; Login = $info.Login; CurrentLevel = "$role group"; Handle = $handle }
                if ($info.Kind -eq 'EntraGroup' -and $expected.ContainsKey($info.Id) -and $expected[$info.Id].Kind -eq $wantKind) {
                    $seen[$info.Id] = $true
                    $rows.Add((New-SpRbacDiffRow @common -Action 'OK' -Principal $expected[$info.Id].Group -PrincipalType $info.Kind -EntraId $info.Id -ExpectedLevel "$role group"))
                    continue
                }
                $name = if ($info.Kind -eq 'EntraGroup') { & $entraName $info } else { $info.Title }
                if (Test-SpRbacKeep -Config $Config -Title $info.Title -Login $info.Login) { $rows.Add((New-SpRbacDiffRow @common -Action 'KEEP' -Principal $name -PrincipalType $info.Kind -EntraId $info.Id -Note 'Matches keepPrincipals in the config.')); continue }
                if ($role -ne 'Members') { $rows.Add((New-SpRbacDiffRow @common -Action 'REVIEW' -Principal $name -PrincipalType $info.Kind -EntraId $info.Id -Note "In the $role group and not an admin group from the config. The toolkit never removes owners or visitors.")); continue }
                switch ($info.Kind) {
                    'EntraGroup' { $rows.Add((New-SpRbacDiffRow @common -Action 'REMOVE_GROUP' -Principal $name -PrincipalType $info.Kind -EntraId $info.Id -Note 'Group should not be a direct member of the site Members group.')) }
                    { $_ -in 'User', 'Guest' } {
                        if ($KeepDirectUsers) { $rows.Add((New-SpRbacDiffRow @common -Action 'KEEP' -Principal $name -PrincipalType $info.Kind -EntraId $info.Id -Note 'Direct user, kept because of -KeepDirectUsers.')) }
                        else { $rows.Add((New-SpRbacDiffRow @common -Action 'REMOVE_USER' -Principal $name -PrincipalType $info.Kind -EntraId $info.Id -Note 'Person added directly to the site Members group.')) }
                    }
                    default { $rows.Add((New-SpRbacDiffRow @common -Action 'REMOVE_OTHER' -Principal $name -PrincipalType $info.Kind -EntraId $info.Id -Note 'Not an Entra security group from the matrix.')) }
                }
            }
            if ($wantKind -ne 'None') {
                foreach ($id in $expected.Keys) {
                    $e = $expected[$id]
                    if ($e.Kind -ne $wantKind -or $seen.ContainsKey($id)) { continue }
                    $rows.Add((New-SpRbacDiffRow -Target $Target -Container $spg.Title -Action 'ADD' -Principal $e.Group -PrincipalType 'EntraGroup' -EntraId $id -Login "c:0t.c|tenant|$id" -ExpectedLevel "$role group" -Handle ([pscustomobject]@{ SPGroup = $spg })))
                }
            }
        }
        # Anything granted on the site itself, outside the three associated groups
        foreach ($ra in $web.RoleAssignments) {
            Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings | Out-Null
            if ($assocIds.ContainsKey([int]$ra.Member.Id)) { continue }
            $info = Get-SpRbacPrincipalInfo -Member $ra.Member
            $levels = @(Get-SpRbacLevels -Bindings $ra.RoleDefinitionBindings)
            if ($info.Kind -in 'System', 'SharingLink' -or -not $levels.Count) { continue }
            $name = if ($info.Kind -eq 'EntraGroup') { & $entraName $info } else { $info.Title }
            $rows.Add((New-SpRbacDiffRow -Target $Target -Container '(site permissions)' -Action 'REVIEW' -Principal $name -PrincipalType $info.Kind -EntraId $info.Id -Login $info.Login -CurrentLevel ($levels -join ', ') -Note 'Has access on the site outside the Owners, Members and Visitors groups.'))
        }
        return $rows.ToArray()
    }
    @()
}

function Invoke-SpRbacChange {
    <#
    .SYNOPSIS
        Applies one row from Get-SpRbacTargetDiff. Returns $true when something was changed.
    #>
    param([Parameter(Mandatory)]$Row)
    $ctx = Get-PnPContext
    $web = $ctx.Web
    $h = $Row.Handle
    $ensure = {
        param($login)
        # A group created minutes ago can take a while before SharePoint resolves it
        for ($i = 1; $i -le 6; $i++) {
            try { $p = $web.EnsureUser($login); $ctx.Load($p); $ctx.ExecuteQuery(); return $p }
            catch { if ($i -eq 6) { throw }; Write-Host "       not resolvable yet (attempt $i/6), waiting 30 s ..." -ForegroundColor DarkYellow; Start-Sleep -Seconds 30 }
        }
    }
    $grant = {
        param($list, $login, $level)
        $principal = & $ensure $login
        $rd = Get-SpRbacRoleDefinition -Web $web -Level $level
        $bindings = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
        $bindings.Add($rd)
        $list.RoleAssignments.Add($principal, $bindings) | Out-Null
        $ctx.ExecuteQuery()
    }
    if ($Row.Scope -eq 'Library') {
        switch -Regex ($Row.Action) {
            '^BREAK_INHERITANCE$' { $h.BreakRoleInheritance($true, $false); $ctx.ExecuteQuery(); return $true }
            '^ADD$' { & $grant $h.List $Row.Login $Row.ExpectedLevel; return $true }
            '^FIX_LEVEL$' { $h.RoleAssignment.DeleteObject(); $ctx.ExecuteQuery(); & $grant $h.List $Row.Login $Row.ExpectedLevel; return $true }
            '^REMOVE_' { $h.RoleAssignment.DeleteObject(); $ctx.ExecuteQuery(); return $true }
        }
    }
    elseif ($Row.Scope -eq 'Site') {
        switch -Regex ($Row.Action) {
            '^ADD$' { $principal = & $ensure $Row.Login; $h.SPGroup.Users.AddUser($principal) | Out-Null; $ctx.ExecuteQuery(); return $true }
            '^REMOVE_' { $h.SPGroup.Users.RemoveByLoginName($Row.Login); $ctx.ExecuteQuery(); return $true }
        }
    }
    $false
}

Export-ModuleMember -Function *-SpRbac*
