<#
.SYNOPSIS
    One-time setup PER TENANT for unattended uploads (GitHub Actions): creates the app registration
    "Intune App Upload (Automation)" with Microsoft Graph APPLICATION permissions, a certificate (or secret),
    stores the credential as GitHub environment secrets, and records the client ID in IntuneTenant-<TenantName>.json.

.DESCRIPTION
    This is separate from New-IntuneAppRegistration.ps1, which sets up interactive sign-in (no stored credential).
    Keeping them separate means the powerful app-only credential exists only where automation needs it.

    In the tenant you sign in to, it:
      1. Creates or updates the app registration (single tenant, confidential client, no redirect URI) with
         Microsoft Graph application permissions:
           DeviceManagementApps.ReadWrite.All   create/update Win32 apps, assignments, supersedence
           DeviceManagementRBAC.Read.All        scope tags
           Group.Read.All                       look up assignment groups by name
      2. Grants admin consent for those permissions.
      3. Creates the credential:
           Certificate (default): self-signed, RSA 2048, valid -ValidityMonths. The public key is added to the app
                                  (replacing earlier certificates - this is how you rotate it). The private key is
                                  exported as a password-protected .pfx.
           Secret:                a client secret valid -ValidityMonths.
      4. With -GitHubRepo (needs the GitHub CLI 'gh', signed in): creates the environments
         '<prefix><TenantName>' and '<prefix><TenantName>-production' and stores in both:
           INTUNE_CERT_PFX_BASE64 + INTUNE_CERT_PASSWORD   (certificate)   or   INTUNE_CLIENT_SECRET (secret)
         The local .pfx copy and the certificate in your store are then deleted (use -KeepLocalCopy to keep them).
         Without -GitHubRepo the credential is written to -OutputPath for you to add to GitHub by hand.
      5. Adds AutomationClientId (+ credential type and expiry) to IntuneTenant-<TenantName>.json next to this script,
         keeping the interactive ClientId that New-IntuneAppRegistration.ps1 wrote.

    Shows the tenant and asks for confirmation before changing anything (-Force skips this).

    Requires: Windows PowerShell 5.1 / PowerShell 7 on Windows (for New-SelfSignedCertificate), the
    Microsoft.Graph.Authentication module (installed for the current user if missing), and an account in that tenant
    that can grant admin consent for application permissions: Global Administrator or Privileged Role Administrator.

.PARAMETER TenantId
    Customer tenant ID or domain, e.g. contoso.onmicrosoft.com (recommended).
.PARAMETER TenantName
    Short name for IntuneTenant-<TenantName>.json and the GitHub environment names. Default: the .onmicrosoft.com prefix.
.PARAMETER DisplayName
    App registration name. Default: Intune App Upload (Automation).
.PARAMETER CredentialType
    Certificate (default, recommended) or Secret.
.PARAMETER ValidityMonths
    Credential lifetime. Default 12. Re-run the script before it expires to rotate it.
.PARAMETER GitHubRepo
    owner/repo to store the credential in, as GitHub environment secrets (uses the gh CLI).
.PARAMETER EnvironmentPrefix
    GitHub environment name prefix. Default 'intune-' -> intune-contoso and intune-contoso-production.
.PARAMETER OutputPath
    Where to write the credential when it isn't stored in GitHub (or with -KeepLocalCopy).
    Default: $HOME\IntuneAutomationCredentials\<TenantName>. Keep it out of the repository.
.PARAMETER KeepLocalCopy
    Keep the .pfx/password files and the certificate in CurrentUser\My even after storing them in GitHub.
.PARAMETER Force
    Don't ask to confirm the tenant.

.EXAMPLE
    .\New-IntuneAutomationRegistration.ps1 -TenantId contoso.onmicrosoft.com -GitHubRepo petabloc/intune-apps
.EXAMPLE
    .\New-IntuneAutomationRegistration.ps1 -TenantId fabrikam.onmicrosoft.com -TenantName FAB -GitHubRepo petabloc/intune-apps -ValidityMonths 6
.EXAMPLE
    .\New-IntuneAutomationRegistration.ps1 -TenantId contoso.onmicrosoft.com     # writes the .pfx + password to $HOME\IntuneAutomationCredentials\contoso
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$TenantName,
    [string]$DisplayName = 'Intune App Upload (Automation)',
    [ValidateSet('Certificate', 'Secret')][string]$CredentialType = 'Certificate',
    [ValidateRange(1, 24)][int]$ValidityMonths = 12,
    [string]$GitHubRepo,
    [string]$EnvironmentPrefix = 'intune-',
    [string]$OutputPath,
    [switch]$KeepLocalCopy,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$Root       = $PSScriptRoot
$GraphAppId = '00000003-0000-0000-c000-000000000000'   # Microsoft Graph
$Graph      = 'https://graph.microsoft.com/v1.0'
$AppRoles   = 'DeviceManagementApps.ReadWrite.All', 'DeviceManagementRBAC.Read.All', 'Group.Read.All'

function Get-GraphItems([string]$Uri) {
    $items = @()
    while ($Uri) {
        $r = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        if ($r.value) { $items += $r.value }
        $Uri = $r.'@odata.nextLink'
    }
    return $items
}
function ConvertTo-Filter([string]$Filter) { return [uri]::EscapeDataString($Filter) }
function Invoke-Graph([string]$Method, [string]$Uri, $Body) {
    $p = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject' }
    if ($null -ne $Body) { $p['Body'] = ($Body | ConvertTo-Json -Depth 8); $p['ContentType'] = 'application/json' }
    Invoke-MgGraphRequest @p
}
function New-RandomPassword([int]$Length = 32) {
    $chars = [char[]]'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789'
    $bytes = New-Object byte[] $Length
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
    return -join ($bytes | ForEach-Object { $chars[$_ % $chars.Length] })
}

#region ---------------- Checks ----------------
if ($CredentialType -eq 'Certificate' -and -not ($PSVersionTable.PSEdition -eq 'Desktop' -or $IsWindows)) { throw 'Run this on Windows (it creates the certificate with New-SelfSignedCertificate), or use -CredentialType Secret.' }
$UseGitHub = [bool]$GitHubRepo
if ($UseGitHub) {
    if ($GitHubRepo -notmatch '^[\w.-]+/[\w.-]+$') { throw "-GitHubRepo must be owner/repo (got '$GitHubRepo')." }
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw "The GitHub CLI (gh) isn't installed - install it (winget install GitHub.cli) or leave out -GitHubRepo." }
    & gh auth status *> $null
    if ($LASTEXITCODE -ne 0) { throw "The GitHub CLI isn't signed in - run 'gh auth login' first." }
}
#endregion

#region ---------------- Module + sign-in ----------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host 'Installing the Microsoft.Graph.Authentication module for the current user...' -ForegroundColor Yellow
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
    }
    Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Authentication

$ConnectParams = @{ Scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'User.Read'); NoWelcome = $true }
if ($TenantId) { $ConnectParams['TenantId'] = $TenantId }
Write-Host 'Signing in to Microsoft Graph (a browser window will open)...' -ForegroundColor Cyan
Connect-MgGraph @ConnectParams
$Context = Get-MgContext
if (-not $Context) { throw 'Sign-in failed.' }
$TenantGuid = $Context.TenantId

$Org = (Invoke-Graph GET "$Graph/organization?`$select=id,displayName,verifiedDomains").value | Select-Object -First 1
$TenantDisplayName = [string]$Org.displayName
$InitialDomain     = [string]($Org.verifiedDomains | Where-Object { $_.isInitial } | Select-Object -First 1).name
if (-not $TenantName) { $TenantName = if ($InitialDomain) { $InitialDomain -replace '\.onmicrosoft\.com$', '' } else { $TenantDisplayName } }
$TenantName = ($TenantName -replace '[^A-Za-z0-9_.-]', '').Trim('.')
if (-not $TenantName) { throw 'Could not work out a tenant name - pass -TenantName.' }
if (-not $OutputPath) { $OutputPath = Join-Path $HOME "IntuneAutomationCredentials\$TenantName" }
$EnvPilot = "$EnvironmentPrefix$TenantName"
$EnvProd  = "$EnvironmentPrefix$TenantName-production"

Write-Host ''
Write-Host "Signed in as : $($Context.Account)" -ForegroundColor Green
Write-Host "Tenant       : $TenantDisplayName ($InitialDomain)" -ForegroundColor Green
Write-Host "Tenant ID    : $TenantGuid" -ForegroundColor Green
Write-Host "App          : $DisplayName - application permissions: $($AppRoles -join ', ')" -ForegroundColor Green
Write-Host "Credential   : $CredentialType, valid $ValidityMonths months" -ForegroundColor Green
Write-Host "Stored in    : $(if ($UseGitHub) { "GitHub $GitHubRepo, environments $EnvPilot and $EnvProd" } else { $OutputPath })" -ForegroundColor Green
if (-not $Force) {
    $answer = Read-Host 'Create/update the automation app registration in this tenant? (y/n) [n]'
    if ($answer -notmatch '^(y|yes)$') { Disconnect-MgGraph | Out-Null; Write-Host 'Cancelled - nothing changed.'; return }
}
#endregion

#region ---------------- App registration ----------------
$GraphSp = Get-GraphItems "$Graph/servicePrincipals?`$filter=$(ConvertTo-Filter "appId eq '$GraphAppId'")&`$select=id,appRoles" | Select-Object -First 1
if (-not $GraphSp) { throw 'Microsoft Graph service principal not found in the tenant.' }
$RoleIds = foreach ($r in $AppRoles) {
    $role = $GraphSp.appRoles | Where-Object { $_.value -eq $r -and $_.allowedMemberTypes -contains 'Application' } | Select-Object -First 1
    if (-not $role) { throw "Microsoft Graph application permission '$r' not found." }
    $role.id
}
$AppBody = @{
    displayName            = $DisplayName
    signInAudience         = 'AzureADMyOrg'
    requiredResourceAccess = @(@{ resourceAppId = $GraphAppId; resourceAccess = @($RoleIds | ForEach-Object { @{ id = $_; type = 'Role' } }) })
    notes                  = 'App-only credential for unattended Intune Win32 app uploads (New-IntuneApp.ps1 in GitHub Actions). Created by New-IntuneAutomationRegistration.ps1.'
}
$Existing = @(Get-GraphItems "$Graph/applications?`$filter=$(ConvertTo-Filter "displayName eq '$($DisplayName.Replace("'", "''"))'")&`$select=id,appId,displayName,keyCredentials,passwordCredentials")
if ($Existing.Count -gt 1) { throw "More than one app registration is named '$DisplayName' - remove the extras or use -DisplayName." }
if ($Existing.Count -eq 1) {
    $AppReg = $Existing[0]
    Invoke-Graph PATCH "$Graph/applications/$($AppReg.id)" $AppBody | Out-Null
    Write-Host "Updated app registration '$DisplayName' (client ID $($AppReg.appId))." -ForegroundColor Yellow
}
else {
    $AppReg = Invoke-Graph POST "$Graph/applications" $AppBody
    Write-Host "Created app registration '$DisplayName' (client ID $($AppReg.appId))." -ForegroundColor Green
}
$ClientId = $AppReg.appId

$Sp = $null
for ($i = 0; $i -lt 10 -and -not $Sp; $i++) {
    $Sp = Get-GraphItems "$Graph/servicePrincipals?`$filter=$(ConvertTo-Filter "appId eq '$ClientId'")&`$select=id" | Select-Object -First 1
    if (-not $Sp) {
        try { $Sp = Invoke-Graph POST "$Graph/servicePrincipals" @{ appId = $ClientId } }
        catch { Start-Sleep -Seconds 5 }      # new registrations take a few seconds to replicate
    }
}
if (-not $Sp) { throw 'Could not create the service principal.' }
#endregion

#region ---------------- Admin consent (app role assignments) ----------------
$Assigned = @(Get-GraphItems "$Graph/servicePrincipals/$($Sp.id)/appRoleAssignments" | Where-Object { $_.resourceId -eq $GraphSp.id } | ForEach-Object { $_.appRoleId })
foreach ($rid in $RoleIds) {
    if ($Assigned -contains $rid) { continue }
    try { Invoke-Graph POST "$Graph/servicePrincipals/$($Sp.id)/appRoleAssignments" @{ principalId = $Sp.id; resourceId = $GraphSp.id; appRoleId = $rid } | Out-Null }
    catch { throw "Could not grant admin consent ($($_.Exception.Message)). This needs Global Administrator or Privileged Role Administrator." }
}
Write-Host "Admin consent granted: $($AppRoles -join ', ')" -ForegroundColor Green
#endregion

#region ---------------- Credential ----------------
$Expires = (Get-Date).AddMonths($ValidityMonths)
$Secrets = [ordered]@{}
$LocalFiles = @()
$Thumbprint = $null
if ($CredentialType -eq 'Certificate') {
    if ($AppReg.keyCredentials -and @($AppReg.keyCredentials).Count -gt 0) {
        Write-Warning "Replacing $(@($AppReg.keyCredentials).Count) existing certificate(s) on '$DisplayName' - anything still using them stops working now."
    }
    $Cert = New-SelfSignedCertificate -Subject "CN=$DisplayName - $TenantName" -CertStoreLocation 'Cert:\CurrentUser\My' `
        -KeyExportPolicy Exportable -KeySpec Signature -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
        -Provider 'Microsoft Enhanced RSA and AES Cryptographic Provider' -NotAfter $Expires
    $Thumbprint = $Cert.Thumbprint
    Invoke-Graph PATCH "$Graph/applications/$($AppReg.id)" @{ keyCredentials = @(@{
        type = 'AsymmetricX509Cert'; usage = 'Verify'; key = [Convert]::ToBase64String($Cert.RawData); displayName = "CN=$DisplayName - $TenantName" }) } | Out-Null
    Write-Host "Certificate $Thumbprint added to the app (expires $($Expires.ToString('yyyy-MM-dd')))." -ForegroundColor Green

    $PfxPassword = New-RandomPassword
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    $PfxFile = Join-Path $OutputPath "$TenantName-automation.pfx"
    Export-PfxCertificate -Cert $Cert -FilePath $PfxFile -Password (ConvertTo-SecureString $PfxPassword -AsPlainText -Force) | Out-Null
    $Secrets['INTUNE_CERT_PFX_BASE64'] = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($PfxFile))
    $Secrets['INTUNE_CERT_PASSWORD']   = $PfxPassword
    $LocalFiles += $PfxFile
}
else {
    $pw = Invoke-Graph POST "$Graph/applications/$($AppReg.id)/addPassword" @{ passwordCredential = @{ displayName = 'GitHub Actions'; endDateTime = $Expires.ToUniversalTime().ToString('o') } }
    $Secrets['INTUNE_CLIENT_SECRET'] = $pw.secretText
    Write-Host "Client secret created (expires $($Expires.ToString('yyyy-MM-dd')))." -ForegroundColor Green
}
#endregion

#region ---------------- Store the credential ----------------
$StoredInGitHub = $false
if ($UseGitHub) {
    try {
        foreach ($envName in $EnvPilot, $EnvProd) {
            & gh api --method PUT "repos/$GitHubRepo/environments/$envName" --silent
            if ($LASTEXITCODE -ne 0) { throw "Could not create environment '$envName' in $GitHubRepo." }
            foreach ($k in $Secrets.Keys) {
                & gh secret set $k --env $envName --repo $GitHubRepo --body $Secrets[$k]     # --body: no trailing newline added
                if ($LASTEXITCODE -ne 0) { throw "Could not set secret $k in environment '$envName'." }
            }
            Write-Host "GitHub environment '$envName': secrets $($Secrets.Keys -join ', ') set." -ForegroundColor Green
        }
        $StoredInGitHub = $true
    }
    catch { Write-Warning "$($_.Exception.Message) - the credential is kept in $OutputPath instead." }
}

if (-not $StoredInGitHub -or $KeepLocalCopy) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    foreach ($k in $Secrets.Keys) {
        if ($k -eq 'INTUNE_CERT_PFX_BASE64') { continue }       # the .pfx itself is already there
        $f = Join-Path $OutputPath "$k.txt"; Set-Content -Path $f -Value $Secrets[$k] -NoNewline; $LocalFiles += $f
    }
    if ($Secrets.Contains('INTUNE_CERT_PFX_BASE64')) {
        $f = Join-Path $OutputPath 'INTUNE_CERT_PFX_BASE64.txt'; Set-Content -Path $f -Value $Secrets['INTUNE_CERT_PFX_BASE64'] -NoNewline; $LocalFiles += $f
    }
}
elseif ($Thumbprint) {
    # Stored in GitHub: remove the local private key copies
    foreach ($f in $LocalFiles) { Remove-Item -Path $f -Force -ErrorAction SilentlyContinue }
    Remove-Item -Path "Cert:\CurrentUser\My\$Thumbprint" -Force -ErrorAction SilentlyContinue
    if ((Test-Path $OutputPath) -and -not (Get-ChildItem $OutputPath)) { Remove-Item $OutputPath -Force }
    $LocalFiles = @()
}
$Secrets.Clear()
#endregion

#region ---------------- IntuneTenant-<TenantName>.json ----------------
$TenantFile = Join-Path $Root "IntuneTenant-$TenantName.json"
$Info = [ordered]@{}
# Merge with the existing file for this tenant (same name, or the same tenant ID under another name)
$Prev = Get-ChildItem -Path $Root -Filter 'IntuneTenant*.json' -File | ForEach-Object {
    try { $j = Get-Content $_.FullName -Raw | ConvertFrom-Json; if ($j.TenantId -eq $TenantGuid) { [pscustomobject]@{ File = $_; Json = $j } } } catch { }
} | Select-Object -First 1
if ($Prev) { $Prev.Json.PSObject.Properties | ForEach-Object { $Info[$_.Name] = $_.Value } }
$Info['TenantName']        = $TenantName
$Info['TenantDisplayName'] = $TenantDisplayName
$Info['TenantDomain']      = $InitialDomain
$Info['TenantId']          = $TenantGuid
if (-not $Info.Contains('ClientId')) { $Info['ClientId'] = '' }     # interactive app (New-IntuneAppRegistration.ps1), if any
$Info['AutomationClientId']          = $ClientId
$Info['AutomationAppRegistration']   = $DisplayName
$Info['AutomationCredential']        = $CredentialType
$Info['AutomationCredentialExpires'] = $Expires.ToString('yyyy-MM-dd')
if ($Thumbprint) { $Info['AutomationCertThumbprint'] = $Thumbprint }
if ($UseGitHub)  { $Info['GitHubEnvironments'] = @($EnvPilot, $EnvProd) }
$Info['Updated'] = (Get-Date -Format 'yyyy-MM-dd HH:mm')
$Info | ConvertTo-Json | Set-Content -Path $TenantFile -Encoding UTF8
if ($Prev -and $Prev.File.FullName -ne $TenantFile) { Remove-Item $Prev.File.FullName -Force }
#endregion

Disconnect-MgGraph | Out-Null

Write-Host ''
Write-Host '================ Automation app registration ================' -ForegroundColor Cyan
Write-Host "Tenant        : $TenantDisplayName ($InitialDomain)"
Write-Host "App           : $DisplayName"
Write-Host "Client ID     : $ClientId   (saved as AutomationClientId in IntuneTenant-$TenantName.json - commit that file)"
Write-Host "Permissions   : $($AppRoles -join ', ') (application, admin consent granted)"
Write-Host "Credential    : $CredentialType$(if ($Thumbprint) { " $Thumbprint" }), expires $($Expires.ToString('yyyy-MM-dd'))"
if ($StoredInGitHub) {
    Write-Host "GitHub        : $GitHubRepo - environments $EnvPilot, $EnvProd"
    Write-Host '                Add required reviewers to the -production environment (Settings > Environments) to gate production assignments.' -ForegroundColor Yellow
}
if ($LocalFiles) {
    Write-Host "Local files   : $OutputPath" -ForegroundColor Yellow
    $LocalFiles | ForEach-Object { Write-Host "                $(Split-Path $_ -Leaf)" }
    Write-Host '                These contain the private key/secret. Add them as environment secrets, then delete them.' -ForegroundColor Yellow
}
Write-Host '=============================================================' -ForegroundColor Cyan
