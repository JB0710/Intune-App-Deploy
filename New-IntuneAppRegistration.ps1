<#
.SYNOPSIS
    One-time setup PER TENANT: creates (or updates) the Entra ID app registration that New-IntuneApp.ps1
    signs in with, grants admin consent, and saves the details to IntuneTenant-<TenantName>.json.

.DESCRIPTION
    Built for managing several tenants (MSP): run it once in each customer tenant. Each run writes its own
    IntuneTenant-<TenantName>.json next to this script, e.g. IntuneTenant-contoso.json, and
    New-IntuneApp.ps1 asks which tenant to upload to (or takes -Tenant contoso).

    In the tenant you sign in to, it creates the app registration "Intune App Upload" (or -DisplayName) with:
      - Single tenant, public client, redirect URI http://localhost (interactive browser sign-in)
      - Microsoft Graph delegated permissions:
          DeviceManagementApps.ReadWrite.All, DeviceManagementConfiguration.ReadWrite.All,
          DeviceManagementRBAC.Read.All, Group.Read.All, offline_access
      - Tenant-wide admin consent for those permissions
    No client secret or certificate is created - you still sign in as yourself.

    <TenantName> defaults to the tenant's initial domain without .onmicrosoft.com (contoso.onmicrosoft.com
    -> contoso). Override it with -TenantName (e.g. a short customer code).

    Before changing anything it shows which tenant you're signed in to and asks you to confirm (-Force skips this),
    so you don't create the registration in the wrong customer.

    Safe to run again: an existing registration with the same name is reused and its permissions,
    redirect URI and consent are brought up to date.

    Requires the Microsoft.Graph.Authentication module (installed for the current user if missing) and an
    account in that tenant that can create app registrations and grant admin consent (Global Administrator,
    Privileged Role Administrator, or Cloud Application Administrator / Application Administrator).

.PARAMETER TenantId
    Tenant ID or domain to sign in to, e.g. contoso.onmicrosoft.com. Recommended for MSP use, so you
    always land in the intended customer tenant. Default: the home tenant of the account you sign in with.
.PARAMETER TenantName
    Short name used in the file name IntuneTenant-<TenantName>.json and in New-IntuneApp.ps1's tenant list.
    Default: the initial domain prefix (contoso).
.PARAMETER DisplayName
    Name of the app registration. Default: Intune App Upload.
.PARAMETER SkipConsent
    Don't grant admin consent (an admin can do it later in the portal: API permissions > Grant admin consent).
.PARAMETER SetAsDefault
    Single-tenant shortcut: also write this tenant's IDs into $DefaultTenantId / $DefaultClientId in every
    New-IntuneApp.ps1, so it never asks which tenant. Don't use this when you manage several tenants.
.PARAMETER Force
    Don't ask for confirmation of the tenant.

.EXAMPLE
    .\New-IntuneAppRegistration.ps1 -TenantId contoso.onmicrosoft.com
    # -> IntuneTenant-contoso.json
.EXAMPLE
    .\New-IntuneAppRegistration.ps1 -TenantId fabrikam.onmicrosoft.com -TenantName FAB
    # -> IntuneTenant-FAB.json
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$TenantName,
    [string]$DisplayName = 'Intune App Upload',
    [switch]$SkipConsent,
    [switch]$SetAsDefault,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$Root        = $PSScriptRoot
$GraphAppId  = '00000003-0000-0000-c000-000000000000'   # Microsoft Graph
$RedirectUri = 'http://localhost'
$Scopes      = 'DeviceManagementApps.ReadWrite.All', 'DeviceManagementConfiguration.ReadWrite.All',
               'DeviceManagementRBAC.Read.All', 'Group.Read.All', 'offline_access'
$Graph       = 'https://graph.microsoft.com/v1.0'

function Get-GraphItems([string]$Uri) {
    # GET with paging; returns the items of .value
    $items = @()
    while ($Uri) {
        $r = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        if ($r.value) { $items += $r.value }
        $Uri = $r.'@odata.nextLink'
    }
    return $items
}
function ConvertTo-Filter([string]$Filter) { return [uri]::EscapeDataString($Filter) }

#region ---------------- Module + sign-in ----------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host 'Installing the Microsoft.Graph.Authentication module for the current user...' -ForegroundColor Yellow
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
    }
    Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Authentication

$ConnectScopes = @('Application.ReadWrite.All', 'User.Read')   # User.Read: read the tenant's name and domains
if (-not $SkipConsent) { $ConnectScopes += 'DelegatedPermissionGrant.ReadWrite.All' }
$ConnectParams = @{ Scopes = $ConnectScopes; NoWelcome = $true }
if ($TenantId) { $ConnectParams['TenantId'] = $TenantId }

Write-Host 'Signing in to Microsoft Graph (a browser window will open)...' -ForegroundColor Cyan
Connect-MgGraph @ConnectParams
$Context = Get-MgContext
if (-not $Context) { throw 'Sign-in failed.' }
$TenantGuid = $Context.TenantId

# Tenant name + initial domain (contoso.onmicrosoft.com)
$Org = (Invoke-MgGraphRequest -Method GET -Uri "$Graph/organization?`$select=id,displayName,verifiedDomains" -OutputType PSObject).value | Select-Object -First 1
$TenantDisplayName = [string]$Org.displayName
$InitialDomain     = [string]($Org.verifiedDomains | Where-Object { $_.isInitial } | Select-Object -First 1).name
if (-not $TenantName) {
    $TenantName = if ($InitialDomain) { $InitialDomain -replace '\.onmicrosoft\.com$', '' } else { $TenantDisplayName }
}
$TenantName = ($TenantName -replace '[^A-Za-z0-9_.-]', '').Trim('.')
if (-not $TenantName) { throw 'Could not work out a tenant name - pass -TenantName.' }

Write-Host ''
Write-Host "Signed in as : $($Context.Account)" -ForegroundColor Green
Write-Host "Tenant       : $TenantDisplayName ($InitialDomain)" -ForegroundColor Green
Write-Host "Tenant ID    : $TenantGuid" -ForegroundColor Green
Write-Host "Saved as     : IntuneTenant-$TenantName.json" -ForegroundColor Green
if (-not $Force) {
    $answer = Read-Host "Create/update the '$DisplayName' app registration in this tenant? (y/n) [n]"
    if ($answer -notmatch '^(y|yes)$') { Disconnect-MgGraph | Out-Null; Write-Host 'Cancelled - nothing changed.'; return }
}
#endregion

#region ---------------- Microsoft Graph permission IDs ----------------
$GraphSp = Get-GraphItems "$Graph/servicePrincipals?`$filter=$(ConvertTo-Filter "appId eq '$GraphAppId'")&`$select=id,oauth2PermissionScopes" | Select-Object -First 1
if (-not $GraphSp) { throw 'Microsoft Graph service principal not found in the tenant.' }

$ResourceAccess = foreach ($scope in $Scopes) {
    $perm = $GraphSp.oauth2PermissionScopes | Where-Object { $_.value -eq $scope } | Select-Object -First 1
    if (-not $perm) { throw "Microsoft Graph delegated permission '$scope' not found." }
    @{ id = $perm.id; type = 'Scope' }
}
$AppBody = @{
    displayName            = $DisplayName
    signInAudience         = 'AzureADMyOrg'
    isFallbackPublicClient = $true            # allows device code sign-in too
    publicClient           = @{ redirectUris = @($RedirectUri) }
    requiredResourceAccess = @(@{ resourceAppId = $GraphAppId; resourceAccess = @($ResourceAccess) })
    notes                  = 'Used by New-IntuneApp.ps1 (IntuneWin32App module) to upload Win32 apps to Intune with interactive sign-in. Created by New-IntuneAppRegistration.ps1.'
}
#endregion

#region ---------------- Create or update the app registration ----------------
$Existing = @(Get-GraphItems "$Graph/applications?`$filter=$(ConvertTo-Filter "displayName eq '$($DisplayName.Replace("'", "''"))'")&`$select=id,appId,displayName,publicClient")
if ($Existing.Count -gt 1) { throw "More than one app registration is named '$DisplayName' - delete the extras or use -DisplayName with a unique name." }

if ($Existing.Count -eq 1) {
    $AppReg = $Existing[0]
    Write-Host "Found existing app registration '$DisplayName' (client ID $($AppReg.appId)) - updating it." -ForegroundColor Yellow
    # Keep any redirect URIs already configured, add http://localhost
    $uris = @($AppReg.publicClient.redirectUris) + $RedirectUri | Where-Object { $_ } | Select-Object -Unique
    $AppBody.publicClient = @{ redirectUris = @($uris) }
    Invoke-MgGraphRequest -Method PATCH -Uri "$Graph/applications/$($AppReg.id)" -Body ($AppBody | ConvertTo-Json -Depth 6) -ContentType 'application/json' | Out-Null
}
else {
    $AppReg = Invoke-MgGraphRequest -Method POST -Uri "$Graph/applications" -Body ($AppBody | ConvertTo-Json -Depth 6) -ContentType 'application/json' -OutputType PSObject
    Write-Host "Created app registration '$DisplayName' (client ID $($AppReg.appId))." -ForegroundColor Green
}
$ClientId = $AppReg.appId
#endregion

#region ---------------- Service principal (enterprise app) ----------------
$Sp = $null
for ($i = 0; $i -lt 10 -and -not $Sp; $i++) {
    $Sp = Get-GraphItems "$Graph/servicePrincipals?`$filter=$(ConvertTo-Filter "appId eq '$ClientId'")&`$select=id,appId" | Select-Object -First 1
    if (-not $Sp) {
        try {
            $Sp = Invoke-MgGraphRequest -Method POST -Uri "$Graph/servicePrincipals" -Body (@{ appId = $ClientId } | ConvertTo-Json) -ContentType 'application/json' -OutputType PSObject
            Write-Host 'Created the service principal (enterprise app).' -ForegroundColor Green
        }
        catch { Start-Sleep -Seconds 5 }   # new app registrations can take a few seconds to replicate
    }
}
if (-not $Sp) { throw 'Could not create the service principal for the app registration.' }
#endregion

#region ---------------- Admin consent ----------------
$ConsentDone = $false
if (-not $SkipConsent) {
    $ScopeString = $Scopes -join ' '
    try {
        $grantFilter = ConvertTo-Filter "clientId eq '$($Sp.id)' and resourceId eq '$($GraphSp.id)' and consentType eq 'AllPrincipals'"
        $Grant = Get-GraphItems "$Graph/oauth2PermissionGrants?`$filter=$grantFilter" | Select-Object -First 1
        if ($Grant) {
            $merged = (@($Grant.scope -split ' ') + $Scopes | Where-Object { $_ } | Select-Object -Unique) -join ' '
            Invoke-MgGraphRequest -Method PATCH -Uri "$Graph/oauth2PermissionGrants/$($Grant.id)" -Body (@{ scope = $merged } | ConvertTo-Json) -ContentType 'application/json' | Out-Null
        }
        else {
            $body = @{ clientId = $Sp.id; consentType = 'AllPrincipals'; resourceId = $GraphSp.id; scope = $ScopeString }
            Invoke-MgGraphRequest -Method POST -Uri "$Graph/oauth2PermissionGrants" -Body ($body | ConvertTo-Json) -ContentType 'application/json' | Out-Null
        }
        $ConsentDone = $true
        Write-Host "Granted admin consent: $ScopeString" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not grant admin consent ($($_.Exception.Message))."
        Write-Warning "Ask a Global Administrator to open Entra > App registrations > '$DisplayName' > API permissions > Grant admin consent."
    }
}
#endregion

#region ---------------- Save IntuneTenant-<TenantName>.json ----------------
$TenantFile = Join-Path $Root "IntuneTenant-$TenantName.json"
[ordered]@{
    TenantName        = $TenantName
    TenantDisplayName = $TenantDisplayName
    TenantDomain      = $InitialDomain
    TenantId          = $TenantGuid
    ClientId          = $ClientId
    AppRegistration   = $DisplayName
    Updated           = (Get-Date -Format 'yyyy-MM-dd HH:mm')
} | ConvertTo-Json | Set-Content -Path $TenantFile -Encoding UTF8

# Remove older files for the same tenant (the unnamed IntuneTenant.json, or the same tenant saved under another name)
$Replaced = @()
Get-ChildItem -Path $Root -Filter 'IntuneTenant*.json' -File | Where-Object { $_.FullName -ne $TenantFile } | ForEach-Object {
    try { $other = Get-Content -Path $_.FullName -Raw | ConvertFrom-Json } catch { return }
    if ($other.TenantId -eq $TenantGuid) { Remove-Item -Path $_.FullName -Force; $Replaced += $_.Name }
}
#endregion

#region ---------------- Optional: write IDs into every New-IntuneApp.ps1 (single-tenant use) ----------------
$Updated = @()
if ($SetAsDefault) {
    $Targets = Get-ChildItem -Path $Root -Filter 'New-IntuneApp.ps1' -File -Recurse -Depth 2 -ErrorAction SilentlyContinue
    foreach ($f in $Targets) {
        $bytes  = [System.IO.File]::ReadAllBytes($f.FullName)
        $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        $text   = [System.IO.File]::ReadAllText($f.FullName)
        if ($text -notmatch '(?m)^\$DefaultTenantId\s*=') {
            Write-Warning "$($f.FullName) has no `$DefaultTenantId line (older version) - copy New-IntuneApp.ps1 from IntuneAppTemplate into that folder."
            continue
        }
        $new = [regex]::Replace($text, "(?m)^\`$DefaultTenantId\s*=\s*'[^']*'", "`$`$DefaultTenantId = '$TenantGuid'")
        $new = [regex]::Replace($new,  "(?m)^\`$DefaultClientId\s*=\s*'[^']*'", "`$`$DefaultClientId = '$ClientId'")
        if ($new -ne $text) { [System.IO.File]::WriteAllText($f.FullName, $new, (New-Object System.Text.UTF8Encoding $hasBom)) }
        $Updated += $f.FullName.Substring($Root.Length).TrimStart('\', '/')
    }
}
#endregion

Disconnect-MgGraph | Out-Null

Write-Host ''
Write-Host '================ App registration ================' -ForegroundColor Cyan
Write-Host "Tenant        : $TenantDisplayName ($InitialDomain)"
Write-Host "Tenant ID     : $TenantGuid"
Write-Host "App reg name  : $DisplayName"
Write-Host "Client ID     : $ClientId"
Write-Host "Redirect URI  : $RedirectUri (public client)"
Write-Host "Permissions   : $($Scopes -join ', ') (delegated)"
Write-Host "Admin consent : $(if ($ConsentDone) { 'granted' } elseif ($SkipConsent) { 'skipped (-SkipConsent)' } else { 'NOT granted - see warning above' })"
Write-Host "Saved to      : IntuneTenant-$TenantName.json"
if ($Replaced) { Write-Host "Replaced      : $($Replaced -join ', ') (same tenant)" }
if ($SetAsDefault) {
    Write-Host 'Set as default in:'
    if ($Updated) { $Updated | ForEach-Object { Write-Host "  $_" } } else { Write-Host '  (no New-IntuneApp.ps1 found)' }
}
Write-Host '==================================================' -ForegroundColor Cyan
Write-Host "Next: in an app folder run .\New-IntuneApp.ps1 -Tenant $TenantName (or pick it from the list)." -ForegroundColor Yellow
