<#
.SYNOPSIS
    The one script for an Intune Win32 app: generates the Install / Uninstall / Update / Detect scripts and
    the documentation, builds the .intunewin package, and uploads it to Intune (pick tenant, supersede the
    previous version, assign a group).

.DESCRIPTION
    New app:
      1. Copy the IntuneAppTemplate folder and rename the copy to the app name (e.g. "7-Zip").
      2. Put the installer (.exe or .msi) in .\Source. For an MSI you can also add .mst transform(s).
         Optional: add the app logo (.png) to .\Source - it's uploaded to Intune as the app icon (not packaged).
      3. Run .\New-IntuneApp.ps1 - it reads the installer (version, product name, publisher,
         installer type - or for an MSI: ProductCode/UpgradeCode), suggests silent switches, asks you to confirm each setting, then writes:
            Scripts\Install-<App>.ps1      Scripts\Uninstall-<App>.ps1
            Scripts\Update-<App>.ps1       Scripts\Detect-<App>.ps1
            Documentation\<App>-Intune.md
         builds Output\<installer>.intunewin (Source + scripts, Detect-*.ps1 and logo excluded),
         then asks "Upload to Intune now?" - pick the tenant, supersede the old version, assign a group.

    Test first, upload later (or to more tenants):
      .\New-IntuneApp.ps1 -SkipUpload            build only
      .\New-IntuneApp.ps1 -UploadOnly            upload the package already in Output - nothing is regenerated

    New version of an existing app:
      Replace the installer in .\Source and run .\New-IntuneApp.ps1 again. The settings already in this
      folder's scripts are reused as the defaults, so normally only the version changes.

    At each prompt press Enter to accept the [default], type a new value, or type - for "none".

.PARAMETER AppName
    Short name used in script and log file names. Default: folder name (+ becomes P, other symbols removed).
.PARAMETER DisplayName
    Name as shown in Add/Remove Programs (-like pattern, wildcards allowed).
    Default: the MSI's exact ProductName, or *<ProductName>* for an EXE.
.PARAMETER InstallArgs
    Silent install switches. Default: based on the detected installer type.
.PARAMETER InstallLogArg
    Installer log switch, {0} = log path. Use '' for none.
.PARAMETER UninstallArgs
    Silent switches for the app's own uninstaller (ignored when the app uninstalls via MSI).
.PARAMETER ProcessesToClose
    Process names (without .exe) closed before install/update/uninstall.
.PARAMETER InstallIfMissing
    Update script installs the app even if it isn't present.
.PARAMETER NoPrompt
    Don't ask - use parameters, existing settings and detected defaults.
.PARAMETER SkipPackage
    Only generate the scripts and documentation - don't build the .intunewin package (no upload).
.PARAMETER Upload
    Upload to Intune after building, without asking first (the summary is still confirmed unless -NoPrompt).
.PARAMETER SkipUpload
    Build only - don't upload and don't ask.
.PARAMETER UploadOnly
    Skip generation and packaging; upload the existing package in Output (e.g. after testing, or to another tenant).
.PARAMETER Tenant
    Tenant to upload to: the TenantName from ..\IntuneTenant-<TenantName>.json (e.g. contoso), or its domain
    or tenant ID. Asked for when there's more than one tenant file.
.PARAMETER GroupName
    Entra ID group to assign the app to (display name). Asked for if not given (Enter = no assignment).
.PARAMETER GroupId
    Object ID of the group instead of the name.
.PARAMETER Intent
    required (default) or available.
.PARAMETER Notification
    hideAll (default), showReboot or showAll.
.PARAMETER Supersede
    Supersede (update) the newest previous version found in Intune without asking.
.PARAMETER NoSupersede
    Don't supersede, and don't ask.
.PARAMETER MinimumOS
    Minimum Windows release requirement. Default W10_1809.
.PARAMETER IconPath
    PNG/JPG app logo. Default: the .png/.jpg in Source, then Documentation, otherwise the icon in the .exe.
.PARAMETER TenantId
    Override: tenant ID or domain. Use together with -ClientId to skip the tenant files.
.PARAMETER ClientId
    Override: application (client) ID of the app registration in that tenant.
.PARAMETER CertificateThumbprint
    App-only sign-in with a certificate from CurrentUser\My or LocalMachine\My (instead of the browser).
.PARAMETER InstallerUrl
    Download the installer from this URL into Source (replacing older .exe/.msi files there) and save the URL,
    file name and SHA256 to Source\download.json - use it for a new version.
.PARAMETER InstallerSha256
    Expected SHA256 of the download (recommended - from the vendor). The download is rejected if it doesn't match.
.PARAMETER InstallerFileName
    File name to save the download as, when the URL doesn't reveal it.
.PARAMETER AssignOnly
    Don't upload: assign the already uploaded "<Product> <Version>" app to -GroupName/-GroupId
    (used for the approval-gated production assignment in GitHub Actions).

.NOTES
    Unattended / CI (GitHub Actions): the script never prompts when GITHUB_ACTIONS, TF_BUILD or CI=true is set.
    App-only sign-in is used when one of these is set (parameters or environment variables):
      INTUNE_CERT_PFX_BASE64 + INTUNE_CERT_PASSWORD   certificate (.pfx as base64) - recommended
      INTUNE_CERT_THUMBPRINT / -CertificateThumbprint certificate in the Windows certificate store
      INTUNE_CLIENT_SECRET                            client secret
    The client ID is AutomationClientId from ..\IntuneTenant-<TenantName>.json (written by
    ..\New-IntuneAutomationRegistration.ps1) or INTUNE_CLIENT_ID.
    If Source has no installer but Source\download.json exists, the installer is downloaded and its SHA256 checked.

.EXAMPLE
    .\New-IntuneApp.ps1                                   # everything, with prompts
.EXAMPLE
    .\New-IntuneApp.ps1 -SkipUpload                       # build only, test, then:
    .\New-IntuneApp.ps1 -UploadOnly -Tenant contoso
.EXAMPLE
    .\New-IntuneApp.ps1 -NoPrompt -Upload -Tenant contoso -Supersede -GroupName 'SG-App-7Zip' -Intent required
    # new version, fully unattended: reuse settings, rebuild, upload, supersede, assign
.EXAMPLE
    .\New-IntuneApp.ps1 -SkipPackage                      # scripts + docs only
.EXAMPLE
    .\New-IntuneApp.ps1 -InstallerUrl 'https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.9.9/npp.8.9.9.Installer.x64.exe' -InstallerSha256 <sha256> -NoPrompt -SkipUpload
    # new version from a download link
.EXAMPLE
    .\New-IntuneApp.ps1 -AssignOnly -Tenant contoso -GroupName 'All Staff' -Intent required
.EXAMPLE
    .\New-IntuneApp.ps1 -DisplayName 'Notepad++*' -ProcessesToClose notepad++
#>
[CmdletBinding()]
param(
    [string]$AppName,
    [string]$DisplayName,
    [string]$InstallArgs,
    [AllowEmptyString()][string]$InstallLogArg,
    [string]$UninstallArgs,
    [AllowEmptyCollection()][string[]]$ProcessesToClose,
    [switch]$InstallIfMissing,
    [switch]$NoPrompt,
    [switch]$SkipPackage,

    # ---- Upload to Intune ----
    [switch]$Upload,
    [switch]$SkipUpload,
    [switch]$UploadOnly,
    [string]$Tenant,
    [string]$GroupName,
    [string]$GroupId,
    [ValidateSet('required', 'available')][string]$Intent = 'required',
    [ValidateSet('hideAll', 'showReboot', 'showAll')][string]$Notification = 'hideAll',
    [switch]$Supersede,
    [switch]$NoSupersede,
    [ValidateSet('W10_1607', 'W10_1703', 'W10_1709', 'W10_1803', 'W10_1809', 'W10_1903', 'W10_1909', 'W10_2004',
                 'W10_20H2', 'W10_21H1', 'W10_21H2', 'W10_22H2', 'W11_21H2', 'W11_22H2')]
    [string]$MinimumOS = 'W10_1809',
    [string]$IconPath,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,

    # ---- Installer download / CI ----
    [string]$InstallerUrl,
    [string]$InstallerSha256,
    [string]$InstallerFileName,
    [switch]$AssignOnly
)

#region ---------------- Optional fixed tenant (single-tenant use: New-IntuneAppRegistration.ps1 -SetAsDefault) ----------------
# Leave empty when you manage several tenants - the tenant is then picked from ..\IntuneTenant-*.json.
$DefaultTenantId = ''
$DefaultClientId = ''
#endregion ------------------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
$Bound = $PSBoundParameters

# Unattended (GitHub Actions / Azure DevOps / any CI): never prompt
$IsCI = ($env:GITHUB_ACTIONS -eq 'true') -or [bool]$env:TF_BUILD -or ($env:CI -eq 'true')
if ($IsCI) { $NoPrompt = [switch]$true }

$Root          = $PSScriptRoot
$FolderName    = Split-Path $Root -Leaf
$SourcePath    = Join-Path $Root 'Source'
$ScriptsPath   = Join-Path $Root 'Scripts'
$DocsPath      = Join-Path $Root 'Documentation'
$TemplatesPath = Join-Path $Root 'Templates'
$OutputPath    = Join-Path $Root 'Output'
$ToolPath      = Join-Path $Root 'Tool\IntuneWinAppUtil.exe'
$TemplateNames = 'Install', 'Uninstall', 'Update', 'Detect'

#region ---------------- Silent switch defaults per installer type ----------------
$InstallerProfiles = @{
    'Inno Setup'         = @{ Install = '/VERYSILENT /NORESTART /SUPPRESSMSGBOXES /SP-'; Log = '/LOG="{0}"'; Uninstall = '/VERYSILENT /NORESTART /SUPPRESSMSGBOXES' }
    'NSIS'               = @{ Install = '/S';                                             Log = '';           Uninstall = '/S' }
    'WiX Burn'           = @{ Install = '/quiet /norestart';                              Log = '/log "{0}"'; Uninstall = '/uninstall /quiet /norestart' }
    'InstallShield'      = @{ Install = '/s /v"/qn REBOOT=ReallySuppress"';               Log = '';           Uninstall = '/s' }
    'Advanced Installer' = @{ Install = '/exenoui /qn /norestart';                        Log = '';           Uninstall = '/exenoui /qn /norestart' }
    'MSI'                = @{ Install = '/qn /norestart';                                 Log = '/l*v "{0}"'; Uninstall = '' }
    'Unknown'            = @{ Install = '/S';                                             Log = '';           Uninstall = '/S' }
}
#endregion

#region ---------------- Helpers ----------------
function ConvertTo-Version {
    param([string]$Text)
    if ($Text -match '(\d+(\.\d+){1,3})') {
        $parts = [System.Collections.Generic.List[string]]$Matches[1].Split('.')
        while ($parts.Count -lt 4) { $parts.Add('0') }
        return [version]($parts -join '.')
    }
    return $null
}

function Get-VersionText([string]$Text) {
    # The dotted version exactly as the installer writes it (e.g. 24.08 stays 24.08)
    if ($Text -match '(\d+(\.\d+){1,3})') { return $Matches[1] }
    return $null
}

function Format-Version([version]$v) {
    if ($v.Revision -gt 0) { return $v.ToString(4) }
    return $v.ToString(3)
}

function Get-InstallerType([string]$Path) {
    # Look for the installer engine's signature in the first 8 MB of the file
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $buffer = New-Object byte[] ([Math]::Min($fs.Length, 8MB))
        $read = 0
        while ($read -lt $buffer.Length) {
            $n = $fs.Read($buffer, $read, $buffer.Length - $read)
            if ($n -le 0) { break }
            $read += $n
        }
    }
    finally { $fs.Dispose() }
    $text = [System.Text.Encoding]::GetEncoding(28591).GetString($buffer, 0, $read)
    if ($text.Contains('.wixburn'))           { return 'WiX Burn' }
    if ($text.Contains('Inno Setup'))         { return 'Inno Setup' }
    if ($text.Contains('Nullsoft'))           { return 'NSIS' }
    if ($text.Contains('InstallShield'))      { return 'InstallShield' }
    if ($text.Contains('Advanced Installer')) { return 'Advanced Installer' }
    return 'Unknown'
}

function Get-MsiProperty {
    # Reads a value from an MSI's Property table
    param([string]$Path, [string]$Property)
    $installer = $db = $view = $record = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db     = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 0))
        $view   = $db.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $db, @("SELECT Value FROM Property WHERE Property='$Property'"))
        [void]$view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
        $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        $value  = if ($record) { $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1) } else { $null }
        [void]$view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)
        return $value
    }
    catch { return $null }
    finally {
        foreach ($o in $record, $view, $db, $installer) { if ($o) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Get-ScriptSettings([string]$File) {
    # Reads the literal $Variable = value assignments from a previously generated script
    $result = @{}
    if (-not $File -or -not (Test-Path $File)) { return $result }
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($File, [ref]$null, [ref]$null)
    $assignments = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $false)
    foreach ($a in $assignments) {
        if ($a.Left -isnot [System.Management.Automation.Language.VariableExpressionAst]) { continue }
        if ($a.Right -isnot [System.Management.Automation.Language.CommandExpressionAst]) { continue }
        try { $result[$a.Left.VariablePath.UserPath] = $a.Right.Expression.SafeGetValue() } catch { }
    }
    return $result
}

function Read-Setting {
    # Precedence: command-line parameter > setting already in this folder's scripts > detected default.
    # Then (unless -NoPrompt) ask, showing that value as the default.
    param([string]$Name, [string]$Prompt, [string]$Detected)
    if ($Bound.ContainsKey($Name)) { return [string]$Bound[$Name] }
    $default = $Detected
    if ($Existing.ContainsKey($Name) -and $null -ne $Existing[$Name]) { $default = [string]$Existing[$Name] }
    if ($NoPrompt) { return $default }
    $shown  = if ($default) { $default } else { '(none)' }
    $answer = Read-Host "$Prompt [$shown]"
    if ([string]::IsNullOrWhiteSpace($answer)) { return $default }
    if ($answer.Trim() -eq '-') { return '' }
    return $answer.Trim()
}

function ConvertTo-PSLiteral([string]$Value) { return "'" + $Value.Replace("'", "''") + "'" }

function Read-Answer([string]$Prompt, [string]$Default) {
    if ($NoPrompt) { return $Default }
    $shown  = if ($Default) { $Default } else { '(none)' }
    $answer = Read-Host "$Prompt [$shown]"
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
    if ($answer.Trim() -eq '-') { return '' }
    return $answer.Trim()
}
#endregion

function Get-Installer {
    # Downloads an installer into Source, checks its SHA256, returns @{ FileName; Sha256 }
    param([string]$Url, [string]$FileName, [string]$Sha256)
    $ProgressPreference = 'SilentlyContinue'      # the progress bar makes Invoke-WebRequest very slow
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $tmp  = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName())
    $resp = $null
    Write-Host "Downloading $Url ..." -ForegroundColor Cyan
    for ($try = 1; $try -le 3; $try++) {
        try { $resp = Invoke-WebRequest -Uri $Url -OutFile $tmp -PassThru -UseBasicParsing -MaximumRedirection 10; break }
        catch {
            if ($try -eq 3) { throw "Download failed after 3 attempts: $($_.Exception.Message)" }
            Write-Warning "Download attempt $try failed - retrying..."; Start-Sleep -Seconds (5 * $try)
        }
    }
    if (-not $FileName -and $resp) {
        $cd = [string]$resp.Headers['Content-Disposition']
        if ($cd -match 'filename\*?=(?:UTF-8'''')?"?([^";]+)"?') { $FileName = [uri]::UnescapeDataString($Matches[1]) }
    }
    if (-not $FileName) {
        $final = [uri]$Url
        if ($resp -and $resp.BaseResponse) {
            if ($resp.BaseResponse.ResponseUri) { $final = $resp.BaseResponse.ResponseUri }                                    # Windows PowerShell 5.1
            elseif ($resp.BaseResponse.RequestMessage) { $final = $resp.BaseResponse.RequestMessage.RequestUri }              # PowerShell 7
        }
        $FileName = [System.IO.Path]::GetFileName($final.AbsolutePath)
    }
    if ($FileName -notmatch '\.(exe|msi)$') {
        Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
        throw "Couldn't tell the installer's file name from the download ('$FileName') - pass -InstallerFileName."
    }
    $hash = (Get-FileHash -Path $tmp -Algorithm SHA256).Hash
    if ($Sha256 -and $hash -ne $Sha256.Trim().ToUpper()) {
        Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
        throw "Checksum mismatch for $FileName - expected $Sha256, got $hash. The download was rejected."
    }
    if (-not $Sha256) { Write-Warning "No SHA256 given - recorded $hash. Compare it with the checksum the vendor publishes." }
    $dest = Join-Path $SourcePath $FileName
    Move-Item -Path $tmp -Destination $dest -Force
    Write-Host "Downloaded Source\$FileName ($([math]::Round((Get-Item $dest).Length / 1MB, 1)) MB, SHA256 $hash)" -ForegroundColor Green
    return [pscustomobject]@{ FileName = $FileName; Sha256 = $hash }
}

function Write-CiResult {
    # GitHub Actions: step outputs + job summary (no-op elsewhere)
    param([hashtable]$Outputs, [string]$Markdown)
    if ($env:GITHUB_OUTPUT -and $Outputs) { foreach ($k in $Outputs.Keys) { Add-Content -Path $env:GITHUB_OUTPUT -Value "$k=$($Outputs[$k])" } }
    if ($env:GITHUB_STEP_SUMMARY -and $Markdown) { Add-Content -Path $env:GITHUB_STEP_SUMMARY -Value $Markdown }
}

#region ================ Upload to Intune (used at the end, or on its own with -UploadOnly) ================
function Invoke-IntuneUpload {
    # Uploads Output\*.intunewin using Documentation\<App>-Intune.json. Runs in its own scope so its variables
    # never mix with the generation steps above.
    $TenantRoot = Split-Path $Root -Parent      # where New-IntuneAppRegistration.ps1 saves IntuneTenant-<TenantName>.json
    $MinModule  = [version]'1.5.0'

    #region ---------------- Read the app settings ----------------
    if ((Split-Path $Root -Leaf) -eq 'IntuneAppTemplate') { throw 'This is the template folder - run this script in an app folder.' }

    $JsonFile = Get-ChildItem -Path $DocsPath -Filter '*-Intune.json' -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $JsonFile) { throw "No Documentation\*-Intune.json found - run .\New-IntuneApp.ps1 (without -UploadOnly) first." }
    $App = Get-Content -Path $JsonFile.FullName -Raw | ConvertFrom-Json

    $PackageFile = Join-Path $Root $App.PackageFile
    $DetectFile  = Join-Path $Root $App.DetectScript
    if (-not $AssignOnly) {
        if (-not (Test-Path $PackageFile)) { throw "Package not found: $($App.PackageFile) - run .\New-IntuneApp.ps1 (without -SkipPackage)." }
        if (-not (Test-Path $DetectFile))  { throw "Detection script not found: $($App.DetectScript) - run .\New-IntuneApp.ps1." }
    }

    $InstallerInSource = Join-Path $Root "Source\$($App.InstallerFile)"
    if ($AssignOnly -or $IsCI) { }      # no installer needed to upload/assign; in CI only the package is downloaded
    elseif (-not (Test-Path $InstallerInSource)) {
        Write-Warning "Source\$($App.InstallerFile) is no longer in Source - the settings may be out of date. Re-run .\New-IntuneApp.ps1 if you replaced the installer."
    }
    elseif ((Get-Item $PackageFile).LastWriteTime -lt (Get-Item $JsonFile.FullName).LastWriteTime.AddMinutes(-5)) {
        Write-Warning "The package is older than the app settings - re-run .\New-IntuneApp.ps1 to rebuild it if anything changed."
    }

    # Logo: -IconPath > .png/.jpg in Source > .png/.jpg in Documentation > extracted from the .exe
    if ($IconPath -and -not (Test-Path $IconPath)) { throw "Icon not found: $IconPath" }
    if (-not $IconPath) {
        $img = foreach ($dir in (Join-Path $Root 'Source'), $DocsPath) {
            Get-ChildItem -Path $dir -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -in '.png', '.jpg', '.jpeg' } |
                Sort-Object @{ Expression = { $_.Extension -ne '.png' } }, Name     # prefer .png
        }
        $img = @($img) | Where-Object { $_ } | Select-Object -First 1
        if ($img) { $IconPath = $img.FullName }
        elseif ($App.InstallerFile -like '*.exe' -and (Test-Path $InstallerInSource) -and ($PSVersionTable.PSEdition -eq 'Desktop' -or $IsWindows)) {
            $extracted = Join-Path $DocsPath "$($App.AppName)-Icon.png"
            try {
                # Run in a child scope so a System.Drawing load failure is caught here and never stops the upload
                & {
                    param($exe, $png)
                    Add-Type -AssemblyName System.Drawing
                    $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exe)
                    $bmp  = $icon.ToBitmap()
                    $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
                    $bmp.Dispose(); $icon.Dispose()
                } $InstallerInSource $extracted
                $IconPath = $extracted
                Write-Host "Extracted icon from the installer: Documentation\$(Split-Path $extracted -Leaf)" -ForegroundColor DarkGray
            }
            catch { Write-Warning "Could not extract an icon from the installer: $($_.Exception.Message)" }
        }
    }
    #endregion

    #region ---------------- Tenant / app registration ----------------
    # Known tenants: ..\IntuneTenant-<TenantName>.json (plus a legacy, unnamed ..\IntuneTenant.json)
    $Tenants = @(Get-ChildItem -Path $TenantRoot -Filter 'IntuneTenant*.json' -File -ErrorAction SilentlyContinue | ForEach-Object {
        try { $j = Get-Content -Path $_.FullName -Raw | ConvertFrom-Json } catch { Write-Warning "Skipping unreadable $($_.Name)"; return }
        if (-not $j.TenantId -or -not ($j.ClientId -or $j.AutomationClientId)) { return }
        $name = if ($j.TenantName) { $j.TenantName } elseif ($_.BaseName -match '^IntuneTenant-(.+)$') { $Matches[1] } else { '(unnamed)' }
        [pscustomobject]@{ TenantName = $name; DisplayName = [string]$j.TenantDisplayName; Domain = [string]$j.TenantDomain
                           TenantId = $j.TenantId; ClientId = [string]$j.ClientId; AutomationClientId = [string]$j.AutomationClientId; File = $_.Name }
    } | Sort-Object TenantName)

    $Selected = $null
    if ($TenantId -and $ClientId) {
        $Selected = [pscustomobject]@{ TenantName = $TenantId; DisplayName = ''; Domain = ''; TenantId = $TenantId; ClientId = $ClientId; File = '(parameters)' }
    }
    elseif ($Tenant) {
        $hits = @($Tenants | Where-Object { $_.TenantName -eq $Tenant -or $_.Domain -eq $Tenant -or $_.TenantId -eq $Tenant -or $_.DisplayName -eq $Tenant })
        if ($hits.Count -eq 0) { throw "Tenant '$Tenant' not found. Known tenants: $(($Tenants.TenantName) -join ', '). Run ..\New-IntuneAppRegistration.ps1 -TenantId <domain> to add it." }
        if ($hits.Count -gt 1) { throw "'$Tenant' matches more than one tenant file: $(($hits.File) -join ', ')" }
        $Selected = $hits[0]
    }
    elseif ($DefaultTenantId -and $DefaultClientId) {
        $Selected = [pscustomobject]@{ TenantName = '(script default)'; DisplayName = ''; Domain = ''; TenantId = $DefaultTenantId; ClientId = $DefaultClientId; File = 'New-IntuneApp.ps1' }
    }
    elseif ($Tenants.Count -eq 1) {
        $Selected = $Tenants[0]
    }
    elseif ($Tenants.Count -gt 1) {
        if ($NoPrompt) { throw "Several tenants are configured - pass -Tenant <name>: $(($Tenants.TenantName) -join ', ')" }
        Write-Host 'Upload to which tenant?' -ForegroundColor Cyan
        for ($i = 0; $i -lt $Tenants.Count; $i++) {
            $tn   = $Tenants[$i]
            $desc = (@($tn.DisplayName, $tn.Domain) | Where-Object { $_ }) -join ', '
            Write-Host ("  {0,2}) {1,-20} {2}" -f ($i + 1), $tn.TenantName, $desc)
        }
        $pick = Read-Host 'Number or name'
        if ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $Tenants.Count) { $Selected = $Tenants[[int]$pick - 1] }
        else { $Selected = $Tenants | Where-Object { $_.TenantName -eq $pick } | Select-Object -First 1 }
        if (-not $Selected) { throw "No tenant selected." }
    }
    else {
        throw "No tenant configured. Run ..\New-IntuneAppRegistration.ps1 -TenantId <customer domain> once for each tenant (it creates ..\IntuneTenant-<TenantName>.json)."
    }

    $TenantId = $Selected.TenantId
    $ClientId = $Selected.ClientId
    $TenantLabel = (@($Selected.DisplayName, $Selected.Domain) | Where-Object { $_ }) -join ', '
    $TenantLabel = if ($TenantLabel) { "$($Selected.TenantName) ($TenantLabel)" } else { "$($Selected.TenantName) ($TenantId)" }
    Write-Host "Tenant: $TenantLabel" -ForegroundColor Yellow
    #endregion

    #region ---------------- Module + sign-in ----------------
    $Module = Get-Module -ListAvailable -Name IntuneWin32App | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $Module -or $Module.Version -lt $MinModule) {
        Write-Host "Installing the IntuneWin32App module (>= $MinModule) for the current user..." -ForegroundColor Yellow
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
        }
        Install-Module -Name IntuneWin32App -MinimumVersion $MinModule -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name IntuneWin32App -MinimumVersion $MinModule -Force

    # App-only sign-in (CI / unattended) when a certificate or secret is supplied, otherwise interactive in a browser
    $CertB64 = $env:INTUNE_CERT_PFX_BASE64
    $Secret  = $env:INTUNE_CLIENT_SECRET
    $Thumb   = if ($CertificateThumbprint) { $CertificateThumbprint } else { $env:INTUNE_CERT_THUMBPRINT }
    if ($CertB64 -or $Secret -or $Thumb) {
        $AppClientId = if ($env:INTUNE_CLIENT_ID) { $env:INTUNE_CLIENT_ID } else { $Selected.AutomationClientId }
        if (-not $AppClientId) { throw "No automation app for $TenantLabel - run ..\New-IntuneAutomationRegistration.ps1 -TenantId <domain> (adds AutomationClientId to its IntuneTenant json), or set INTUNE_CLIENT_ID." }
        if ($CertB64) {
            $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]'UserKeySet, Exportable'
            $Cert  = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String($CertB64), [string]$env:INTUNE_CERT_PASSWORD, $flags)
        }
        elseif ($Thumb) {
            $Cert = Get-ChildItem -Path Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $Thumb } | Select-Object -First 1
            if (-not $Cert) { throw "Certificate $Thumb not found in CurrentUser\My or LocalMachine\My." }
        }
        if ($Cert) {
            if ($Cert.NotAfter -lt (Get-Date).AddDays(30)) { Write-Warning "The automation certificate expires $($Cert.NotAfter.ToString('yyyy-MM-dd')) - renew it with New-IntuneAutomationRegistration.ps1." }
            Write-Host "Signing in to $TenantLabel as app $AppClientId (certificate $($Cert.Thumbprint))..." -ForegroundColor Cyan
            $AuthHeader = Connect-MSIntuneGraph -TenantID $TenantId -ClientID $AppClientId -ClientCert $Cert
        }
        else {
            Write-Host "Signing in to $TenantLabel as app $AppClientId (client secret)..." -ForegroundColor Cyan
            $AuthHeader = Connect-MSIntuneGraph -TenantID $TenantId -ClientID $AppClientId -ClientSecret $Secret
        }
    }
    else {
        if ($IsCI) { throw 'No app-only credentials in this CI run - add INTUNE_CERT_PFX_BASE64 + INTUNE_CERT_PASSWORD (or INTUNE_CLIENT_SECRET) as secrets of the GitHub environment.' }
        if (-not $ClientId) { throw "$TenantLabel has no interactive app registration (ClientId) - run ..\New-IntuneAppRegistration.ps1 -TenantId <domain>." }
        Write-Host "Signing in to $TenantLabel (a browser window will open - use an account in that tenant)..." -ForegroundColor Cyan
        $AuthHeader = Connect-MSIntuneGraph -TenantID $TenantId -ClientID $ClientId
    }
    if (-not $AuthHeader) { $AuthHeader = $Global:AuthenticationHeader }
    if (-not $Global:AccessToken) { throw 'Sign-in failed.' }
    #endregion

    #region ---------------- Existing versions / supersedence ----------------
    $IntuneName = $App.IntuneDisplayName
    $SupersedeApp = $null
    if (-not $AssignOnly) {
    $Existing = @(Get-IntuneWin32App -DisplayName $App.ProductName -WarningAction SilentlyContinue |
                  Where-Object { $_.displayName -eq $App.ProductName -or $_.displayName -match ('^' + [regex]::Escape($App.ProductName) + ' v?\d') })

    $Duplicate = $Existing | Where-Object { $_.displayName -eq $IntuneName }
    if ($Duplicate) {
        Write-Warning "'$IntuneName' already exists in Intune (id $($Duplicate[0].id))."
        if ($NoPrompt -or (Read-Answer 'Upload another copy anyway? (y/n)' 'n') -notmatch '^(y|yes)$') {
            Write-Host 'Nothing uploaded.'
            Write-CiResult -Outputs @{ app_id = $Duplicate[0].id; app_name = $IntuneName } -Markdown "| $($Selected.TenantName) | $IntuneName | already in Intune - not uploaded again | [open](https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/0/appId/$($Duplicate[0].id)) |"
            return
        }
    }

    # Newest previous version first (by version number, then creation date)
    $Previous = @($Existing | Where-Object { $_.displayName -ne $IntuneName } |
                  Sort-Object -Property @{ Expression = {
                      $v = if ($_.displayVersion) { $_.displayVersion } else { $_.displayName }
                      if ($v -match '(\d+(\.\d+){1,3})') { $parts = [System.Collections.Generic.List[string]]$Matches[1].Split('.'); while ($parts.Count -lt 4) { $parts.Add('0') }; [version]($parts -join '.') } else { [version]'0.0' }
                  }; Descending = $true }, @{ Expression = 'createdDateTime'; Descending = $true })
    if ($Previous -and -not $NoSupersede) {
        Write-Host 'Existing versions in Intune:' -ForegroundColor Cyan
        $Previous | ForEach-Object { Write-Host "  $($_.displayName)  (id $($_.id))" }
        if ($Supersede) { $SupersedeApp = $Previous[0] }
        elseif (-not $NoPrompt) {
            if ((Read-Answer "Supersede '$($Previous[0].displayName)' with the new version? (y/n)" 'y') -match '^(y|yes)$') { $SupersedeApp = $Previous[0] }
        }
    }
    }   # end if (-not $AssignOnly)
    #endregion

    #region ---------------- Assignment choice ----------------
    if (-not $GroupName -and -not $GroupId -and -not $NoPrompt) {
        $GroupName = Read-Answer 'Assign to Entra ID group (name, Enter = no assignment)' ''
        if ($GroupName) {
            $Intent       = Read-Answer 'Intent (required/available)' $Intent
            if ($Intent -notin 'required', 'available') { throw "Intent must be 'required' or 'available'." }
            $Notification = Read-Answer 'Notifications (hideAll/showReboot/showAll)' $Notification
            if ($Notification -notin 'hideAll', 'showReboot', 'showAll') { throw "Notification must be hideAll, showReboot or showAll." }
        }
    }

    $Group = $null
    if ($GroupId) {
        $Group = Invoke-RestMethod -Method Get -Headers $AuthHeader -Uri "https://graph.microsoft.com/v1.0/groups/$($GroupId)?`$select=id,displayName"
    }
    elseif ($GroupName) {
        $filter = [uri]::EscapeDataString("displayName eq '$($GroupName.Replace("'", "''"))'")
        $found  = @((Invoke-RestMethod -Method Get -Headers $AuthHeader -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=$filter&`$select=id,displayName").value)
        if ($found.Count -eq 0) { throw "Group '$GroupName' not found." }
        if ($found.Count -gt 1) { throw "More than one group is named '$GroupName' - use -GroupId instead: $($found.id -join ', ')" }
        $Group = $found[0]
    }
    #endregion

    #region ---------------- Assign-only (approval step) ----------------
    if ($AssignOnly) {
        if (-not $Group) { throw '-AssignOnly needs a group: -GroupName or -GroupId.' }
        $Target = @(Get-IntuneWin32App -DisplayName $App.ProductName -WarningAction SilentlyContinue | Where-Object { $_.displayName -eq $IntuneName } |
                    Sort-Object createdDateTime -Descending)
        if (-not $Target) { throw "'$IntuneName' not found in $TenantLabel - upload it first." }
        $Target = $Target[0]
        Write-Host ''
        Write-Host "Assign '$IntuneName' (id $($Target.id)) in $TenantLabel to '$($Group.displayName)' - $Intent, notifications $Notification" -ForegroundColor Cyan
        if (-not $NoPrompt -and (Read-Answer 'Assign now? (y/n)' 'y') -notmatch '^(y|yes)$') { Write-Host 'Cancelled.'; return }
        Add-IntuneWin32AppAssignmentGroup -Include -ID $Target.id -GroupID $Group.id -Intent $Intent -Notification $Notification | Out-Null
        Write-Host "Assigned to '$($Group.displayName)' as $Intent." -ForegroundColor Green
        $LogFile = Join-Path $DocsPath "$($App.AppName)-IntuneUploads.log"
        Add-Content -Path $LogFile -Value ('{0}  tenant={1}  {2}  id={3}  ASSIGNED group={4} ({5})' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $Selected.TenantName, $IntuneName, $Target.id, $Group.displayName, $Intent)
        Write-CiResult -Outputs @{ app_id = $Target.id } -Markdown ("| $($Selected.TenantName) | $IntuneName | assigned to **$($Group.displayName)** ($Intent) | [open](https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/0/appId/$($Target.id)) |")
        return
    }
    #endregion

    #region ---------------- Summary + confirm ----------------
    $PackageMB = [math]::Round((Get-Item $PackageFile).Length / 1MB, 1)
    Write-Host ''
    Write-Host '================ Upload to Intune ================' -ForegroundColor Cyan
    Write-Host "Tenant       : $TenantLabel" -ForegroundColor Yellow
    Write-Host "Name         : $IntuneName"
    Write-Host "Publisher    : $($App.Publisher)"
    Write-Host "Version      : $($App.Version)"
    Write-Host "Package      : $($App.PackageFile) ($PackageMB MB)"
    Write-Host "Install      : $($App.InstallCommand)"
    Write-Host "Uninstall    : $($App.UninstallCommand)"
    Write-Host "Detection    : $($App.DetectScript)"
    Write-Host "Requirements : $($App.Architecture), minimum $MinimumOS"
    Write-Host "Logo         : $(if ($IconPath) { $IconPath.Replace("$Root\", '').Replace("$Root/", '') } else { '(none)' })"
    Write-Host "Supersedes   : $(if ($SupersedeApp) { $SupersedeApp.displayName } else { '(none)' })"
    Write-Host "Assignment   : $(if ($Group) { "$($Group.displayName) - $Intent, notifications $Notification" } else { '(none)' })"
    Write-Host '==================================================' -ForegroundColor Cyan
    if (-not $NoPrompt -and (Read-Answer 'Upload now? (y/n)' 'y') -notmatch '^(y|yes)$') { Write-Host 'Cancelled.'; return }
    #endregion

    #region ---------------- Create the Win32 app ----------------
    $DetectionRule   = New-IntuneWin32AppDetectionRuleScript -ScriptFile $DetectFile -EnforceSignatureCheck $false -RunAs32Bit $false
    $RequirementRule = New-IntuneWin32AppRequirementRule -Architecture $App.Architecture -MinimumSupportedWindowsRelease $MinimumOS

    $AddParams = @{
        FilePath                         = $PackageFile
        DisplayName                      = $IntuneName
        Description                      = $App.Description
        Publisher                        = $App.Publisher
        AppVersion                       = $App.Version
        InstallExperience                = 'system'
        RestartBehavior                  = 'basedOnReturnCode'
        DetectionRule                    = @($DetectionRule)
        RequirementRule                  = $RequirementRule
        InstallCommandLine               = $App.InstallCommand
        UninstallCommandLine             = $App.UninstallCommand
        MaximumInstallationTimeInMinutes = 60
        Notes                            = "Created by New-IntuneApp.ps1 from folder '$(Split-Path $Root -Leaf)' on $(Get-Date -Format 'yyyy-MM-dd HH:mm')."
    }
    if ($IconPath) { $AddParams['Icon'] = New-IntuneWin32AppIcon -FilePath $IconPath }

    Write-Host "Uploading $($App.PackageFile) - this can take a few minutes for large packages..." -ForegroundColor Cyan
    $Win32App = Add-IntuneWin32App @AddParams
    if (-not $Win32App -or -not $Win32App.id) { throw 'Add-IntuneWin32App did not return an app - see the warnings above.' }
    Write-Host "Created '$($Win32App.displayName)' (id $($Win32App.id))." -ForegroundColor Green
    #endregion

    #region ---------------- Supersedence + assignment ----------------
    if ($SupersedeApp) {
        try {
            $rel = New-IntuneWin32AppSupersedence -ID $SupersedeApp.id -SupersedenceType 'Update'
            Add-IntuneWin32AppSupersedence -ID $Win32App.id -Supersedence @($rel) | Out-Null
            Write-Host "Supersedes '$($SupersedeApp.displayName)' (update in place)." -ForegroundColor Green
        }
        catch { Write-Warning "App created, but supersedence failed: $($_.Exception.Message) - set it in the portal (Supersedence tab)." }
    }

    if ($Group) {
        try {
            Add-IntuneWin32AppAssignmentGroup -Include -ID $Win32App.id -GroupID $Group.id -Intent $Intent -Notification $Notification | Out-Null
            Write-Host "Assigned to '$($Group.displayName)' as $Intent (notifications: $Notification)." -ForegroundColor Green
        }
        catch { Write-Warning "App created, but the assignment failed: $($_.Exception.Message) - assign it in the portal." }
    }
    #endregion

    #region ---------------- Record the result ----------------
    $LogFile = Join-Path $DocsPath "$($App.AppName)-IntuneUploads.log"
    $line = '{0}  tenant={6}  {1}  id={2}  supersedes={3}  group={4} ({5})' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $IntuneName, $Win32App.id,
            $(if ($SupersedeApp) { $SupersedeApp.displayName } else { '-' }), $(if ($Group) { $Group.displayName } else { '-' }), $(if ($Group) { $Intent } else { '-' }), $Selected.TenantName
    Add-Content -Path $LogFile -Value $line

    Write-Host ''
    Write-Host "Done - uploaded to $TenantLabel. Open it in the Intune admin center:" -ForegroundColor Green
    Write-Host "  https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/0/appId/$($Win32App.id)"
    Write-Host "Logged to Documentation\$(Split-Path $LogFile -Leaf)" -ForegroundColor DarkGray
    Write-CiResult -Outputs @{ app_id = $Win32App.id; app_name = $IntuneName } `
        -Markdown ("| $($Selected.TenantName) | $IntuneName | uploaded$(if ($SupersedeApp) { ", supersedes $($SupersedeApp.displayName)" })$(if ($Group) { ", assigned to **$($Group.displayName)** ($Intent)" }) | [open](https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/0/appId/$($Win32App.id)) |")
    #endregion
}
#endregion =============================================================================================

if ($UploadOnly -or $AssignOnly) {
    Invoke-IntuneUpload
    return
}

#region ---------------- Checks ----------------
if ($FolderName -eq 'IntuneAppTemplate') {
    throw 'This is the template folder. Copy it, rename the copy to the app name, put the installer in .\Source, then run this script in the copy.'
}
foreach ($t in $TemplateNames) {
    if (-not (Test-Path (Join-Path $TemplatesPath "$t.ps1.template"))) { throw "Missing template: Templates\$t.ps1.template" }
}
if (-not $SkipPackage -and -not (Test-Path $ToolPath)) {
    throw "IntuneWinAppUtil.exe not found at: $ToolPath (or run with -SkipPackage)"
}
foreach ($p in $ScriptsPath, $DocsPath) {
    if (-not (Test-Path $p)) { New-Item -Path $p -ItemType Directory | Out-Null }
}

# Get the installer: -InstallerUrl (new version) > Source\download.json (when Source has no installer) > file already in Source
if (-not (Test-Path $SourcePath)) { New-Item -Path $SourcePath -ItemType Directory | Out-Null }
$DownloadFile = Join-Path $SourcePath 'download.json'
if ($InstallerUrl) {
    $dl = Get-Installer -Url $InstallerUrl -FileName $InstallerFileName -Sha256 $InstallerSha256
    Get-ChildItem -Path $SourcePath -File | Where-Object { $_.Extension -in '.exe', '.msi' -and $_.Name -ne $dl.FileName } |
        ForEach-Object { Remove-Item -Path $_.FullName -Force; Write-Host "Removed the previous installer: Source\$($_.Name)" -ForegroundColor DarkGray }
    [ordered]@{ Url = $InstallerUrl; FileName = $dl.FileName; Sha256 = $dl.Sha256 } | ConvertTo-Json | Set-Content -Path $DownloadFile -Encoding UTF8
    Write-Host 'Updated Source\download.json' -ForegroundColor Green
}
elseif (Test-Path $DownloadFile) {
    $d = Get-Content -Path $DownloadFile -Raw | ConvertFrom-Json
    $present = @(Get-ChildItem -Path $SourcePath -File | Where-Object { $_.Extension -in '.exe', '.msi' })
    if ($present.Name -contains $d.FileName) {
        $h = (Get-FileHash -Path (Join-Path $SourcePath $d.FileName) -Algorithm SHA256).Hash
        if ($d.Sha256 -and $h -ne ([string]$d.Sha256).ToUpper()) { throw "Source\$($d.FileName) doesn't match the SHA256 in download.json - delete it (it'll be downloaded again) or update download.json." }
    }
    elseif ($present.Count -eq 0) {
        if (-not $d.Sha256) { Write-Warning 'download.json has no Sha256 - the download is not verified.' }
        [void](Get-Installer -Url $d.Url -FileName $d.FileName -Sha256 $d.Sha256)
    }
    else {
        Write-Warning "download.json points to $($d.FileName), but Source contains $($present.Name -join ', ') - using that. Run with -InstallerUrl to update download.json."
    }
}

$Installers = @(Get-ChildItem -Path $SourcePath -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.exe', '.msi' })
if ($Installers.Count -eq 0) { throw "No .exe or .msi found in $SourcePath - copy the installer there, add Source\download.json, or use -InstallerUrl." }
if ($Installers.Count -gt 1) { throw "More than one installer found in $SourcePath - keep only one .exe or .msi: $($Installers.Name -join ', ')" }
$Installer = $Installers[0]
$IsMsi     = $Installer.Extension -eq '.msi'
$InstallerFilterValue = '*' + $Installer.Extension.ToLower()
#endregion

#region ---------------- Read the installer ----------------
$ProductCode = ''
$UpgradeCode = ''
$Transforms  = @()
if ($IsMsi) {
    # MSI: everything comes from the MSI's Property table
    $InstallerType = 'MSI'
    $ProductName   = [string](Get-MsiProperty $Installer.FullName 'ProductName')
    $Publisher     = [string](Get-MsiProperty $Installer.FullName 'Manufacturer')
    $ProductCode   = [string](Get-MsiProperty $Installer.FullName 'ProductCode')
    $UpgradeCode   = [string](Get-MsiProperty $Installer.FullName 'UpgradeCode')
    $VersionSource = [string](Get-MsiProperty $Installer.FullName 'ProductVersion')
    if (-not $ProductName) { Write-Warning 'Could not read the MSI properties (Windows Installer COM unavailable?) - enter the values at the prompts.' }
    $Transforms    = @(Get-ChildItem -Path $SourcePath -Filter *.mst -File -ErrorAction SilentlyContinue)
}
else {
    $vi            = $Installer.VersionInfo
    $ProductName   = if ($vi.ProductName) { $vi.ProductName.Trim() } else { '' }
    $Publisher     = if ($vi.CompanyName) { $vi.CompanyName.Trim() } else { '' }
    $InstallerType = Get-InstallerType $Installer.FullName
    # FileVersion first (numeric/standard)
    $VersionSource = $vi.FileVersion
}
$ProductName = $ProductName.Trim()
$Publisher   = $Publisher.Trim()
$TypeProfile = $InstallerProfiles[$InstallerType]

# Fallbacks: the file name, then (EXE) ProductVersion
if (-not (ConvertTo-Version $VersionSource)) { $VersionSource = $Installer.BaseName }
if (-not (ConvertTo-Version $VersionSource) -and -not $IsMsi) { $VersionSource = $vi.ProductVersion }
$Version = ConvertTo-Version $VersionSource
if (-not $Version) {
    if ($NoPrompt) { throw "Could not read a version from $($Installer.Name)." }
    $VersionSource = Read-Host "Could not read the installer version - enter it (e.g. 1.2.3)"
    $Version = ConvertTo-Version $VersionSource
    if (-not $Version) { throw 'No valid version entered.' }
}
$VersionText = Get-VersionText $VersionSource
if (-not $VersionText) { $VersionText = Format-Version $Version }

Write-Host ''
Write-Host "Installer     : $($Installer.Name)" -ForegroundColor Cyan
Write-Host "Version       : $VersionText"       -ForegroundColor Cyan
Write-Host "Product name  : $ProductName"       -ForegroundColor Cyan
Write-Host "Publisher     : $Publisher"         -ForegroundColor Cyan
Write-Host "Installer type: $InstallerType"     -ForegroundColor Cyan
if ($IsMsi) {
    Write-Host "ProductCode   : $ProductCode"   -ForegroundColor Cyan
    Write-Host "UpgradeCode   : $UpgradeCode"   -ForegroundColor Cyan
    if ($Transforms) { Write-Host "Transforms    : $($Transforms.Name -join ', ')" -ForegroundColor Cyan }
}
if ($InstallerType -eq 'Unknown') {
    Write-Warning 'Installer type not recognised - check the vendor docs for the silent install/uninstall switches.'
}
Write-Host ''
#endregion

#region ---------------- Resolve settings ----------------
$DefaultAppName = ($FolderName -replace '\+', 'P') -replace '[^A-Za-z0-9_.-]', ''
if (-not $DefaultAppName) { $DefaultAppName = 'App' }

# Reuse settings from a previous run in this folder (same app only)
$Existing = @{}
$PrevInstall = Get-ChildItem -Path $ScriptsPath -Filter 'Install-*.ps1' -File -ErrorAction SilentlyContinue | Select-Object -First 1
if ($PrevInstall) {
    $prev = Get-ScriptSettings $PrevInstall.FullName
    $prevAppName = [string]$prev['AppName']
    $prevFilter  = if ($prev['InstallerFilter']) { [string]$prev['InstallerFilter'] } else { '*.exe' }
    if ($prevAppName -and $prevAppName -notlike '__*' -and $prevAppName -eq $DefaultAppName -and $prevFilter -eq $InstallerFilterValue) {
        foreach ($k in 'AppName', 'DisplayName', 'InstallArgs', 'InstallLogArg') { if ($prev.ContainsKey($k)) { $Existing[$k] = $prev[$k] } }
        if ($prev.ContainsKey('ProcessesToClose')) { $Existing['ProcessesToClose'] = (@($prev['ProcessesToClose']) -join ', ') }
        $u = Get-ScriptSettings (Get-ChildItem -Path $ScriptsPath -Filter 'Uninstall-*.ps1' -File | Select-Object -First 1).FullName
        if ($u.ContainsKey('UninstallArgs')) { $Existing['UninstallArgs'] = $u['UninstallArgs'] }
        $up = Get-ScriptSettings (Get-ChildItem -Path $ScriptsPath -Filter 'Update-*.ps1' -File | Select-Object -First 1).FullName
        if ($up.ContainsKey('InstallIfMissing')) { $Existing['InstallIfMissing'] = [string][bool]$up['InstallIfMissing'] }
        Write-Host "Using the settings from the existing $($PrevInstall.Name) as defaults." -ForegroundColor Yellow
    }
}

# MSI: exact ProductName (that's what Add/Remove Programs shows). EXE: a *ProductName* guess.
if ($ProductName -and $IsMsi) { $DefaultDisplayName = [System.Management.Automation.WildcardPattern]::Escape($ProductName) }
elseif ($ProductName)          { $DefaultDisplayName = '*' + [System.Management.Automation.WildcardPattern]::Escape($ProductName) + '*' }
else                           { $DefaultDisplayName = "*$FolderName*" }
# If the product name contains the version (e.g. "7-Zip 24.08 (x64 edition)"), wildcard it so detection keeps matching future versions
if ($VersionText -and $DefaultDisplayName.Contains($VersionText)) { $DefaultDisplayName = $DefaultDisplayName.Replace($VersionText, '*') -replace '\*{2,}', '*' }

# MSI: apply any .mst transforms found in Source
$DefaultInstallArgs = $TypeProfile.Install
if ($IsMsi -and $Transforms) { $DefaultInstallArgs = "$DefaultInstallArgs TRANSFORMS=`"$($Transforms.Name -join ';')`"" }

if (-not $NoPrompt) { Write-Host 'Press Enter to accept [default], type a new value, or - for none.' -ForegroundColor Yellow }

$AppNameValue = Read-Setting 'AppName' 'App name (script/log file names)' $DefaultAppName
$AppNameValue = $AppNameValue -replace '[^A-Za-z0-9_.+-]', ''
$AppNameValue = $AppNameValue -replace '\+', 'P'
if (-not $AppNameValue) { throw 'App name cannot be empty.' }

$DisplayNameValue = Read-Setting 'DisplayName'   'Add/Remove Programs name (-like pattern)' $DefaultDisplayName
if (-not $DisplayNameValue) { throw 'Display name cannot be empty.' }
$InstallArgsValue   = Read-Setting 'InstallArgs'   $(if ($IsMsi) { 'msiexec switches/properties (after /i <msi>)' } else { 'Silent install switches' }) $DefaultInstallArgs
$InstallLogValue    = Read-Setting 'InstallLogArg' 'Installer log switch ({0} = log path)'     $TypeProfile.Log
if ($IsMsi) { $UninstallArgsValue = '' }   # MSI uninstall is always msiexec /x {ProductCode} /qn /norestart
else { $UninstallArgsValue = Read-Setting 'UninstallArgs' 'Silent uninstall switches' $TypeProfile.Uninstall }

if ($Bound.ContainsKey('ProcessesToClose')) { $ProcessText = ($ProcessesToClose -join ', ') }
else { $ProcessText = Read-Setting 'ProcessesToClose' 'Processes to close, comma separated (no .exe)' '' }
$Processes = @($ProcessText -split ',' | ForEach-Object { ($_.Trim()) -replace '\.exe$', '' } | Where-Object { $_ })

if ($Bound.ContainsKey('InstallIfMissing')) { $InstallIfMissingValue = [bool]$InstallIfMissing }
else {
    $ans = Read-Setting 'InstallIfMissing' 'Update script installs the app if missing? (True/False)' 'False'
    $InstallIfMissingValue = $ans -match '^(y|yes|true|1|\$true)$'
}
#endregion

#region ---------------- Generate scripts ----------------
$ProcessLiteral = if ($Processes.Count) { '@(' + (($Processes | ForEach-Object { ConvertTo-PSLiteral $_ }) -join ', ') + ')' } else { '@()' }

$Tokens = [ordered]@{
    "'__DISPLAYNAME__'"     = ConvertTo-PSLiteral $DisplayNameValue
    "'__INSTALLARGS__'"     = ConvertTo-PSLiteral $InstallArgsValue
    "'__INSTALLLOGARG__'"   = ConvertTo-PSLiteral $InstallLogValue
    "'__UNINSTALLARGS__'"   = ConvertTo-PSLiteral $UninstallArgsValue
    "'__MINVERSION__'"      = ConvertTo-PSLiteral $VersionText
    "'__INSTALLERFILTER__'" = ConvertTo-PSLiteral $InstallerFilterValue
    "'__PRODUCTCODE__'"     = ConvertTo-PSLiteral $ProductCode
    '__PROCESSES__'         = $ProcessLiteral
    '__INSTALLIFMISSING__'  = $(if ($InstallIfMissingValue) { '$true' } else { '$false' })
    '__APPNAME__'           = $AppNameValue
}

# Remove previously generated scripts (any app name) so only the current set remains
Get-ChildItem -Path $ScriptsPath -File |
    Where-Object { $_.Name -match '^(Install|Uninstall|Update|Detect)-.+\.ps1$' } |
    ForEach-Object { Remove-Item -Path $_.FullName -Force; Write-Host "Removed old script: $($_.Name)" -ForegroundColor DarkGray }

$Utf8Bom = New-Object System.Text.UTF8Encoding $true
$Generated = @()
foreach ($t in $TemplateNames) {
    $content = [System.IO.File]::ReadAllText((Join-Path $TemplatesPath "$t.ps1.template"))
    foreach ($key in $Tokens.Keys) { $content = $content.Replace($key, $Tokens[$key]) }
    $outFile = Join-Path $ScriptsPath "$t-$AppNameValue.ps1"
    [System.IO.File]::WriteAllText($outFile, $content, $Utf8Bom)

    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($outFile, [ref]$null, [ref]$errors)
    if ($errors) { Write-Warning "$(Split-Path $outFile -Leaf) has syntax errors: $($errors[0].Message)" }
    if ($content -match '__[A-Z]+__') { Write-Warning "$(Split-Path $outFile -Leaf) still contains an unreplaced token: $($Matches[0])" }
    $Generated += $outFile
    Write-Host "Created: Scripts\$(Split-Path $outFile -Leaf)" -ForegroundColor Green
}
#endregion

#region ---------------- Documentation ----------------
$PackageName    = [System.IO.Path]::GetFileNameWithoutExtension($Installer.Name) + '.intunewin'
$InstallCmd     = "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File .\Install-$AppNameValue.ps1"
$UninstallCmd   = "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File .\Uninstall-$AppNameValue.ps1"
$UpdateCmd      = "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File .\Update-$AppNameValue.ps1"
$ProcessDoc     = if ($Processes.Count) { $Processes -join ', ' } else { '(none)' }
$LogArgDoc      = if ($InstallLogValue) { $InstallLogValue } else { '(none)' }
$PublisherDoc   = if ($Publisher) { $Publisher } else { '(not set in installer)' }
$ProductDoc     = if ($ProductName) { $ProductName } else { $FolderName }
$Arch           = if ($Installer.Name -match 'x64|amd64|win64|64-bit') { '64-bit' } elseif ($Installer.Name -match 'arm64') { 'ARM64' } else { '64-bit (check - not in the file name)' }
$UninstallDoc   = if ($IsMsi) { 'msiexec /x {ProductCode} /qn /norestart' } else { $UninstallArgsValue }
$MsiDoc = ''
if ($IsMsi) {
    $TransformDoc = if ($Transforms) { $Transforms.Name -join ', ' } else { '(none)' }
    $MsiDoc = @"

## MSI details

| | |
|---|---|
| ProductCode | ``$ProductCode`` |
| UpgradeCode | ``$UpgradeCode`` |
| Transforms | $TransformDoc |

- The install script runs ``msiexec /i <msi> $InstallArgsValue`` and logs to ``$AppNameValue-Install-Setup.log``.
- Updates install the new MSI over the old one - this only works if the vendor's MSI supports upgrades (MajorUpgrade, same UpgradeCode). If an update fails with 1638, the old version must be uninstalled first.
- Alternative detection: Intune's built-in **MSI** rule with ProductCode ``$ProductCode`` (and *Value is greater than or equal to* $VersionText). The custom script is still recommended because it keeps working after the ProductCode changes in later versions.
"@
}
$MissingDoc     = if ($InstallIfMissingValue) { 'installs it' } else { 'does nothing (exit 0)' }

$Md = @"
# $ProductDoc - Intune Win32 app

Generated by ``New-IntuneApp.ps1`` on $(Get-Date -Format 'yyyy-MM-dd HH:mm') from folder ``$FolderName``.

## Installer

| | |
|---|---|
| File | ``$($Installer.Name)`` |
| Version | $VersionText |
| Product name | $ProductDoc |
| Publisher | $PublisherDoc |
| Installer type | $InstallerType |
| Installer filter | ``$InstallerFilterValue`` |
| Package | ``Output\$PackageName`` |

## Intune - App information

| Setting | Value |
|---|---|
| Name | $ProductDoc |
| Version | $VersionText |
| Publisher | $PublisherDoc |

## Intune - Program

| Setting | Value |
|---|---|
| Install command | ``$InstallCmd`` |
| Uninstall command | ``$UninstallCmd`` |
| Install behavior | System |
| Device restart behavior | Determine behavior based on return codes |

| Return code | Type |
|---|---|
| 0 | Success |
| 3010 | Soft reboot |
| 1618 | Retry (another installation was in progress) |
| 1 | Failed |

## Intune - Requirements

| Setting | Value |
|---|---|
| Operating system architecture | $Arch |
| Minimum OS | Windows 10 / 11 (as required by the vendor) |

## Intune - Detection rules

- Rules format: **Use a custom detection script**
- Script file: ``Scripts\Detect-$AppNameValue.ps1`` (not packaged - upload it here)
- Run script as 32-bit process on 64-bit clients: **No**
- Enforce script signature check: **No** (unless your scripts are signed)
- Detects: Add/Remove Programs entry like ``$DisplayNameValue`` with version **$VersionText** or newer
$MsiDoc

## Updating to a new version

Option A - **Supersedence** (recommended): create a new app from the new package with the install command above,
then in the new app's *Supersedence* tab add the old app (leave *Uninstall previous version* = No; the installer upgrades in place).

Option B - **Separate update app**: same package, install command:

``$UpdateCmd``

Add a requirement rule so it only targets devices that already have the app. If the app is missing, the update script $MissingDoc.

Steps for a new version:
1. Replace the installer in ``Source`` (keep only one .exe or .msi).
2. Run ``.\New-IntuneApp.ps1`` (reuses the current settings; only the version changes) - it regenerates the scripts and builds the package.
   The script then offers to upload it to Intune as ``$ProductDoc <version>`` and to supersede the previous version
   (or test first: ``-SkipUpload``, then ``.\New-IntuneApp.ps1 -UploadOnly``).

## Script settings

| Setting | Value |
|---|---|
| AppName | ``$AppNameValue`` |
| DisplayName (-like) | ``$DisplayNameValue`` |
| Install switches | ``$InstallArgsValue`` |
| Installer log switch | ``$LogArgDoc`` |
| Uninstall | ``$UninstallDoc`` |
| Processes closed first | $ProcessDoc |
| Update installs if missing | $InstallIfMissingValue |

The scripts relaunch in 64-bit PowerShell, skip if the same or newer version is already installed, verify the result
in Add/Remove Programs, and (install/update) delete the installer from the IMECache after a successful install.

## Logs (on the device)

``C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\``

- ``$AppNameValue-Install.log``, ``$AppNameValue-Update.log``, ``$AppNameValue-Uninstall.log``
- ``$AppNameValue-Install-Setup.log`` / ``$AppNameValue-Update-Setup.log`` (installer's own log, if it supports one)

## Test locally (elevated PowerShell)

``````powershell
# Check the Add/Remove Programs name matches the DisplayName pattern after a test install
Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' |
    Where-Object DisplayName -like '$($DisplayNameValue.Replace("'", "''"))' |
    Select-Object DisplayName, DisplayVersion, UninstallString

# Run as SYSTEM like Intune does (PsExec from Sysinternals), from a copy of Source + Scripts
psexec.exe -s -i powershell.exe -ExecutionPolicy Bypass -File .\Install-$AppNameValue.ps1
``````
"@

$DocFile = Join-Path $DocsPath "$AppNameValue-Intune.md"
[System.IO.File]::WriteAllText($DocFile, $Md, (New-Object System.Text.UTF8Encoding $false))
Write-Host "Created: Documentation\$(Split-Path $DocFile -Leaf)" -ForegroundColor Green

# Machine-readable settings for the upload step (-UploadOnly reads them back)
$IntuneArch = if ($Installer.Name -match 'arm64') { 'arm64' } elseif ($Installer.Name -match 'x64|amd64|win64|64-bit') { 'x64' } else { 'x64x86' }
$Description = $null
if (-not $IsMsi -and $vi.FileDescription -and $vi.FileDescription.Trim() -notmatch '(?i)setup|installer') { $Description = $vi.FileDescription.Trim() }
if (-not $Description) { $Description = "$ProductDoc $VersionText - installed silently by Install-$AppNameValue.ps1." }
$AppInfo = [ordered]@{
    AppName           = $AppNameValue
    IntuneDisplayName = "$ProductDoc $VersionText"
    ProductName       = $ProductDoc
    Publisher         = $(if ($Publisher) { $Publisher } else { $ProductDoc })
    Version           = $VersionText
    Description       = $Description
    InstallerFile     = $Installer.Name
    InstallerType     = $InstallerType
    ProductCode       = $ProductCode
    Architecture      = $IntuneArch
    PackageFile       = "Output\$PackageName"
    DetectScript      = "Scripts\Detect-$AppNameValue.ps1"
    InstallCommand    = $InstallCmd
    UninstallCommand  = $UninstallCmd
    DisplayNameRule   = $DisplayNameValue
    Generated         = (Get-Date -Format 'yyyy-MM-dd HH:mm')
}
$JsonFile = Join-Path $DocsPath "$AppNameValue-Intune.json"
[System.IO.File]::WriteAllText($JsonFile, ($AppInfo | ConvertTo-Json), (New-Object System.Text.UTF8Encoding $false))
Write-Host "Created: Documentation\$(Split-Path $JsonFile -Leaf)" -ForegroundColor Green
#endregion

Write-Host ''
Write-Host "Done. $ProductDoc $VersionText scripts are ready." -ForegroundColor Green
if ($DisplayNameValue -eq $DefaultDisplayName -and -not $Existing.ContainsKey('DisplayName')) {
    Write-Host "Note: DisplayName '$DisplayNameValue' was guessed from the product name - confirm it after a test install (see the .md)." -ForegroundColor Yellow
}

#region ---------------- Build the .intunewin package ----------------
if ($SkipPackage) {
    Write-Host 'Skipped packaging and upload (-SkipPackage). Run .\New-IntuneApp.ps1 to regenerate, package and upload.' -ForegroundColor Yellow
    return
}

if (-not (Test-Path $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory | Out-Null }
$StagingPath = Join-Path ([System.IO.Path]::GetTempPath()) ("IntuneWinBuild_" + $AppNameValue)
$PackageFiles = @($Generated | Where-Object { (Split-Path $_ -Leaf) -notlike 'Detect-*' })   # detection script is uploaded separately

try {
    # Staging folder = everything in Source + the generated Install/Uninstall/Update scripts
    if (Test-Path $StagingPath) { Remove-Item -Path $StagingPath -Recurse -Force }
    New-Item -Path $StagingPath -ItemType Directory | Out-Null
    # (the app logo .png/.jpg in Source is only for Intune - the upload step sends it - so it isn't packaged)
    # (download.json only tells the build where to get the installer - not packaged either)
    Get-ChildItem -Path $SourcePath | Where-Object { $_.Extension -notin '.png', '.jpg', '.jpeg' -and $_.Name -ne 'download.json' } |
        Copy-Item -Destination $StagingPath -Recurse -Force
    $PackageFiles | Copy-Item -Destination $StagingPath -Force

    Write-Host ''
    Write-Host 'Creating IntuneWin package...' -ForegroundColor Green
    Write-Host "Setup file : $($Installer.Name)" -ForegroundColor Yellow
    Write-Host "Scripts    : $(($PackageFiles | ForEach-Object { Split-Path $_ -Leaf }) -join ', ')" -ForegroundColor Yellow
    Write-Host "Output Path: $OutputPath" -ForegroundColor Yellow

    # & = call operator; -q = quiet, overwrites an existing package without prompting
    & $ToolPath -c $StagingPath -s $Installer.Name -o $OutputPath -q
    if ($LASTEXITCODE -ne 0) { throw "IntuneWinAppUtil.exe failed with exit code $LASTEXITCODE" }
}
finally {
    if (Test-Path $StagingPath) { Remove-Item -Path $StagingPath -Recurse -Force -ErrorAction SilentlyContinue }
}

$PackageFile = Join-Path $OutputPath $PackageName
if (-not (Test-Path $PackageFile)) { throw "Packaging finished but $PackageName was not found in $OutputPath" }
$Pkg = Get-Item $PackageFile
Write-Host ''
Write-Host "Package created: Output\$($Pkg.Name) ($([math]::Round($Pkg.Length / 1MB, 1)) MB)" -ForegroundColor Green
#endregion

#region ---------------- Upload? ----------------
if ($SkipUpload) { $DoUpload = $false }
elseif ($Upload) { $DoUpload = $true }
elseif ($NoPrompt) { $DoUpload = $false }
else {
    Write-Host ''
    $DoUpload = (Read-Answer 'Upload to Intune now? (y/n - you can test first and run .\New-IntuneApp.ps1 -UploadOnly later)' 'y') -match '^(y|yes)$'
}
if ($DoUpload) {
    Write-Host ''
    Invoke-IntuneUpload
}
else {
    Write-Host 'Not uploaded. Test the package, then run .\New-IntuneApp.ps1 -UploadOnly to upload it.' -ForegroundColor Yellow
}
#endregion
