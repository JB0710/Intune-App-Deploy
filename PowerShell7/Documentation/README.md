# Intune App Deploy

Package EXE and MSI installers as **Intune Win32 apps** and upload them to Intune with two PowerShell scripts: one per tenant, one per app.
Each app gets its own folder, copied from a template. One set of app folders serves **any number of tenants** (MSP use): each tenant is registered once, and you choose the tenant when uploading. The scripts generate silent install, uninstall, update and detection scripts, build the `.intunewin` package, and upload it to Intune.

```
 Once per tenant              Every new app / new version
 ───────────────────────      ─────────────────────────────────────────────────────────────────
 New-IntuneAppRegistration    Copy template ─► Installer + logo into Source ─► New-IntuneApp.ps1
   (app registration)           (or replace installer)                          ├─ scripts + docs
                                                                                ├─ .intunewin package
                                                                                └─ upload: pick tenant, supersede, assign
                                                                                   (or -SkipUpload ─► test ─► -UploadOnly)
```

---

## Contents

1. [Folder layout](#1-folder-layout)
2. [Prerequisites](#2-prerequisites)
3. [One-time setup](#3-one-time-setup)
4. [Add a new app](#4-add-a-new-app)
5. [Release a new version of an app](#5-release-a-new-version-of-an-app)
6. [Script reference](#6-script-reference)
7. [What the deployed scripts do on devices](#7-what-the-deployed-scripts-do-on-devices)
8. [Installer types and silent switches](#8-installer-types-and-silent-switches)
9. [MSI notes](#9-msi-notes)
10. [Intune settings (manual reference)](#10-intune-settings-manual-reference)
11. [Logs and troubleshooting](#11-logs-and-troubleshooting)
12. [Maintaining the template](#12-maintaining-the-template)
13. [Known apps](#13-known-apps)

---

## 1. Folder layout

```
D:\Scripts\Intune\App Deploy\
├── README.md                       ← this file
├── New-IntuneAppRegistration.ps1   ← once per tenant: creates the Entra app registration
├── IntuneTenant-contoso.json       ← one file per tenant: Tenant ID + Client ID (written by the script above)
├── IntuneTenant-fabrikam.json
├── IntuneAppTemplate\              ← copy this for every new app (never run scripts in it)
├── VSCode\                         ← one folder per app
├── Notepad++\
└── ...
```

Each app folder (and the template) contains:

| Path | Purpose |
|---|---|
| `New-IntuneApp.ps1` | **The one script per app:** reads the installer, generates the scripts and documentation, builds the `.intunewin` package, and uploads it to Intune (tenant, supersedence, group assignment) |
| `Source\` | **You add:** exactly one installer (`.exe` or `.msi`), any `.mst` transforms (MSI), and optionally the app logo (`.png`) |
| `Scripts\` | Generated `Install-<App>.ps1`, `Uninstall-<App>.ps1`, `Update-<App>.ps1`, `Detect-<App>.ps1` |
| `Documentation\` | This README, generated `<App>-Intune.md` (human-readable settings), `<App>-Intune.json` (read by the upload step), `<App>-IntuneUploads.log` |
| `Output\` | The built `<installer>.intunewin` |
| `Templates\` | Master copies of the four device scripts with placeholders |
| `Tool\` | Microsoft Win32 Content Prep Tool (`IntuneWinAppUtil.exe`) |

`<App>` is the folder name with `+` changed to `P` and other symbols removed (for example `Notepad++` → `NotepadPP`).

---

## 2. Prerequisites

**On your admin PC**

- Windows PowerShell 5.1 or PowerShell 7 on Windows.
- Internet access to the PowerShell Gallery the first time (the scripts install the modules they need for your user only):
  - `IntuneWin32App` 1.5.0 or later (used by the upload step of `New-IntuneApp.ps1`)
  - `Microsoft.Graph.Authentication` (used by `New-IntuneAppRegistration.ps1`)
- If scripts are blocked, allow local scripts once: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.
  Files downloaded from the internet may also need `Get-ChildItem -Recurse | Unblock-File`.

**Roles**

| Task | Role needed |
|---|---|
| `New-IntuneAppRegistration.ps1` (once per tenant) | Global Administrator, Privileged Role Administrator, or Cloud Application Administrator / Application Administrator |
| Uploading with `New-IntuneApp.ps1` | An Intune role that can create and assign apps, such as Intune Administrator or Application Manager |

**A test device** (or VM) enrolled in Intune, for testing packages before broad assignment.

---

## 3. One-time setup

### 3.1 Register each tenant

Since version 1.5.0 of the `IntuneWin32App` module, sign-in needs an Entra ID app registration **in each tenant** you upload to. You still sign in **interactively** in a browser, with an account in that tenant (a customer admin account, or your partner account through GDAP). The app registration only identifies the script, and it has no secret or certificate.

From `D:\Scripts\Intune\App Deploy`, run it once for each customer tenant:

```powershell
.\New-IntuneAppRegistration.ps1 -TenantId contoso.onmicrosoft.com            # -> IntuneTenant-contoso.json
.\New-IntuneAppRegistration.ps1 -TenantId fabrikam.onmicrosoft.com -TenantName FAB   # -> IntuneTenant-FAB.json
```

Always pass **`-TenantId`** so you sign in to the intended customer and not your own tenant.
After sign-in it shows the tenant's name, domain and ID, plus the file it will save, and asks you to confirm before changing anything (`-Force` skips this). Then it:

1. Creates the app registration **Intune App Upload** in that tenant, or reuses it if it already exists. It's single tenant, a public client, with redirect URI `http://localhost`.
2. Adds Microsoft Graph **delegated** permissions and grants **admin consent**:
   `DeviceManagementApps.ReadWrite.All`, `DeviceManagementConfiguration.ReadWrite.All`, `DeviceManagementRBAC.Read.All`, `Group.Read.All`, `offline_access`.
3. Saves **`IntuneTenant-<TenantName>.json`** next to the script, with the tenant name, display name, domain, Tenant ID and Client ID.
   `<TenantName>` is the tenant's `.onmicrosoft.com` prefix (`contoso`) unless you pass `-TenantName` (for example a customer code).
   An older file for the **same tenant**, such as the unnamed `IntuneTenant.json` or the same tenant under another name, is removed.

| Parameter | Description |
|---|---|
| `-TenantId` | Customer tenant ID or domain to sign in to (recommended) |
| `-TenantName` | Short name for the file and the tenant list (default: `.onmicrosoft.com` prefix) |
| `-DisplayName` | App registration name (default `Intune App Upload`) |
| `-SkipConsent` | Don't grant admin consent; an admin grants it later in the portal |
| `-SetAsDefault` | **Single-tenant only:** also write the IDs into every `New-IntuneApp.ps1`, so it never asks for a tenant |
| `-Force` | Don't ask to confirm the tenant |

Running it again for a tenant is safe. It keeps existing settings and adds anything missing.

<details>
<summary>Manual alternative (Entra admin center of the customer tenant)</summary>

1. **App registrations** > **New registration** > name `Intune App Upload`, *Accounts in this organizational directory only*.
2. **Authentication** > **Add a platform** > **Mobile and desktop applications** > custom redirect URI `http://localhost`.
3. **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated** > add the five permissions above > **Grant admin consent**.
4. Create `App Deploy\IntuneTenant-<name>.json`:
   ```json
   { "TenantName": "contoso", "TenantDisplayName": "Contoso Ltd", "TenantDomain": "contoso.onmicrosoft.com",
     "TenantId": "<Directory (tenant) ID>", "ClientId": "<Application (client) ID>" }
   ```
</details>

### 3.2 Check the template's packaging tool

`IntuneAppTemplate\Tool\IntuneWinAppUtil.exe` must be present. To update it, download the latest [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool) and replace the file in the template. New app folders copied after that get the new version.

---

## 4. Add a new app

### Step 1: Copy the template

Copy `IntuneAppTemplate` and rename the copy to the app's name, for example `7-Zip`.
The folder name becomes the app name in script and log file names.

### Step 2: Add the installer and logo to `Source\`

- **Exactly one** installer: `.exe` or `.msi`.
- **MSI only:** any `.mst` transform files. They're applied automatically with `TRANSFORMS=`.
- **Optional logo:** a `.png` (square, about 256×256). It's uploaded to Intune as the app icon and is **not** put in the package.
  Without a logo, the icon inside the `.exe` is used.

### Step 3: Run New-IntuneApp.ps1

```powershell
cd 'D:\Scripts\Intune\App Deploy\7-Zip'
.\New-IntuneApp.ps1               # build, then asks "Upload to Intune now?"
.\New-IntuneApp.ps1 -SkipUpload   # build only (recommended for a new app: test first, then step 5)
```

It reads the installer and shows what it found:

```
Installer     : 7z2408-x64.msi
Version       : 24.08.00.0
Product name  : 7-Zip 24.08 (x64 edition)
Publisher     : Igor Pavlov
Installer type: MSI
ProductCode   : {23170F69-40C1-2702-2408-000001000000}
```

Then it asks you to confirm each setting. Press **Enter** to accept the `[default]`, type a new value, or type **`-`** for "none":

| Prompt | What it means | Tip |
|---|---|---|
| App name | Used in script and log file names | Default is the folder name |
| Add/Remove Programs name | `-like` pattern used to find the app on devices | For an EXE it's a guess (`*Product*`). Check it after a test install (see §11). The version is replaced with `*` automatically. |
| Silent install switches | Arguments for the installer. For an MSI, they're added after `msiexec /i <msi>` | Pre-filled for the detected installer type (§8). Check the vendor's docs if the type is *Unknown*. |
| Installer log switch | Makes the installer write its own log (`{0}` = log path) | `-` if the installer has none |
| Silent uninstall switches | For the app's own uninstaller (not asked for MSI) | |
| Processes to close | Process names (without `.exe`) closed before install, update or uninstall | For example `Code`, `notepad++`. Leave empty if none. |
| Update installs if missing | Should the update script install the app on devices that don't have it? | Usually `False` |

It then writes:

- `Scripts\Install-<App>.ps1`, `Uninstall-<App>.ps1`, `Update-<App>.ps1`, `Detect-<App>.ps1`
- `Documentation\<App>-Intune.md` (all Intune settings) and `Documentation\<App>-Intune.json`
- `Output\<installer>.intunewin`, containing the installer, `.mst` files, and the install, uninstall and update scripts. The detect script and logo are left out.

Finally it asks **Upload to Intune now?** For a **new** app, answer **n** (or use `-SkipUpload`), test it (step 4), then upload (step 5).
Answering **y** goes straight to the upload described in step 5.

Use `-SkipPackage` to only generate scripts and docs, or `-NoPrompt` to accept all defaults (with `-NoPrompt` it only uploads if you also pass `-Upload`).

### Step 4: Test on a test device

Copy `Source\<installer>` and the three scripts from `Scripts\` (not `Detect-`) into one folder on a test device. Then, in an **elevated** PowerShell:

```powershell
# Run as SYSTEM like Intune does (PsExec from Sysinternals)
psexec.exe -s -i powershell.exe -ExecutionPolicy Bypass -File C:\Test\Install-7-Zip.ps1

# Check the detection script finds it (should print "Detected: ...")
powershell.exe -ExecutionPolicy Bypass -File C:\Test\Detect-7-Zip.ps1

# Uninstall
psexec.exe -s -i powershell.exe -ExecutionPolicy Bypass -File C:\Test\Uninstall-7-Zip.ps1
```

Check that nothing appears on screen, and look at the logs in `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\` (§11).
The installer isn't deleted during the test, because cleanup only runs from Intune's cache.

### Step 5: Upload to Intune

```powershell
.\New-IntuneApp.ps1 -UploadOnly                   # uploads the package already in Output - nothing is rebuilt
.\New-IntuneApp.ps1 -UploadOnly -Tenant contoso
```

(The same upload runs if you answer **y** to *Upload to Intune now?* in step 3.)

1. Picks the **tenant**: `-Tenant <name>` (the TenantName, domain or ID), the only registered tenant, or a numbered list to choose from:
   ```
   Upload to which tenant?
      1) contoso              Contoso Ltd, contoso.onmicrosoft.com
      2) FAB                  Fabrikam Inc, fabrikam.onmicrosoft.com
   Number or name:
   ```
   Then it signs you in to that tenant through the browser. Use an account in that tenant.
2. Lists any **older versions** of the app already in Intune and asks whether to **supersede** the newest one, so devices update in place.
3. Asks for a **group to assign** (Enter = no assignment), the **intent** (`required` or `available`) and **notifications** (`hideAll` by default).
4. Shows a summary and asks **Upload now?**
5. Creates the app, sets supersedence and the assignment, prints the link to the app in the Intune admin center, and adds a line (including the tenant) to `Documentation\<App>-IntuneUploads.log`.

The tenant is shown at the top of the summary. **Check it before answering "Upload now?"**

To deploy the same app to several customers, upload once per tenant. The same package and scripts work everywhere:

```powershell
foreach ($t in 'contoso', 'FAB') { .\New-IntuneApp.ps1 -UploadOnly -Tenant $t -GroupName 'SG-App-7Zip' -Intent required -Supersede -NoPrompt }
```

Each run opens a browser sign-in for that tenant.

Uploading without an assignment, then assigning a **test group** first, is a good habit.

Non-interactive examples:

```powershell
.\New-IntuneApp.ps1 -UploadOnly -Tenant contoso -GroupName 'SG-App-7Zip-Pilot' -Intent required -NoPrompt
.\New-IntuneApp.ps1 -UploadOnly -Tenant FAB -GroupName 'All Staff' -Intent available -Notification showAll -NoSupersede
```

---

## 5. Release a new version of an app

1. Replace the installer in the app's `Source\` folder. Keep only one `.exe` or `.msi`. Replace the logo too if it changed.
2. Run `.\New-IntuneApp.ps1 -NoPrompt -SkipUpload`. It reuses the settings already in the app's scripts. Only the version changes, and the package is rebuilt.
   Run it without `-NoPrompt` to review or change settings.
3. Test on a test device (§4, step 4).
4. Run `.\New-IntuneApp.ps1 -UploadOnly -Tenant <name>` for each tenant that uses the app. It uploads `<Product> <new version>` and offers to **supersede** the previous version in that tenant.

Once you trust an app's updates, rebuild and upload in **one unattended run**:

```powershell
.\New-IntuneApp.ps1 -NoPrompt -Upload -Tenant contoso -Supersede -GroupName 'SG-App-7Zip' -Intent required
```

**How the update reaches devices:**

| Approach | How | When to use |
|---|---|---|
| **Supersedence** (recommended) | The upload sets the new app to *update* the old one. Assign the new app to the same groups. | Normal updates |
| **Separate update app** | Create an app whose install command is `Update-<App>.ps1`, with a requirement rule that the app is already installed | Update only devices that already have the app, without installing it elsewhere |
| **Auto-update for Available apps** | Supersedence plus `-Intent available` | Company Portal installs |

Supersedence only finds older versions named `<Product>` or `<Product> <version>`. For apps first added by hand under another name, set supersedence once in the portal: app > **Properties** > **Supersedence**.

---

## 6. Script reference

### `New-IntuneAppRegistration.ps1` (App Deploy root)

See the parameter table in §3.1.

### `New-IntuneApp.ps1` (app folder)

| Parameter | Description |
|---|---|
| `-AppName` | Name used in script and log file names |
| `-DisplayName` | Add/Remove Programs name pattern |
| `-InstallArgs`, `-InstallLogArg`, `-UninstallArgs` | Silent switches |
| `-ProcessesToClose` | For example `-ProcessesToClose Code,notepad++` |
| `-InstallIfMissing` | The update script installs the app if it's missing |
| `-NoPrompt` | Use parameters, existing settings and detected defaults without asking. Uploads only with `-Upload`. |
| `-SkipPackage` | Only generate scripts and docs. No package, no upload. |

Settings come from, in order: **parameters**, then **settings already in this folder's scripts**, then **detected defaults**.

**What runs:**

| Parameter | Generate + package | Upload |
|---|---|---|
| *(none)* | yes | asks *Upload to Intune now?* |
| `-SkipUpload` | yes | no |
| `-Upload` | yes | yes, without asking first |
| `-UploadOnly` | **no**, uses the existing package | yes |
| `-SkipPackage` | scripts and docs only | no |

**Upload parameters:**

| Parameter | Default | Description |
|---|---|---|
| `-GroupName` / `-GroupId` | asked | Entra ID group to assign. Enter = none. |
| `-Intent` | `required` | `required` or `available` |
| `-Notification` | `hideAll` | `hideAll`, `showReboot`, `showAll` |
| `-Supersede` / `-NoSupersede` | asked | Supersede the newest previous version, or don't |
| `-MinimumOS` | `W10_1809` | Minimum Windows release requirement (`W10_1607` … `W11_22H2`) |
| `-IconPath` | `Source\*.png` | Logo. Otherwise `Documentation\*.png`, otherwise the `.exe` icon. |
| `-Tenant` | asked if several | Tenant to upload to: TenantName, domain or tenant ID from `IntuneTenant-*.json` |
| `-TenantId` + `-ClientId` | | Bypass the tenant files and use these IDs |
| `-NoPrompt` | | Don't ask anything |

---

## 7. What the deployed scripts do on devices

Intune runs the scripts as **SYSTEM**, hidden from the user.

| Script | Behavior |
|---|---|
| `Install-<App>.ps1` | Skips if the same or a newer version is installed. Closes the listed processes and installs silently (`msiexec /i` for an MSI). Checks the app appears in Add/Remove Programs, then deletes the installer from Intune's cache. |
| `Update-<App>.ps1` | Like install, but only when an **older** version is installed (or the app is missing and `InstallIfMissing` is on). Checks the new version afterwards. |
| `Uninstall-<App>.ps1` | Finds the app by ProductCode (MSI) or name. Runs `msiexec /x` or the app's uninstaller silently, and waits until it's gone from Add/Remove Programs. |
| `Detect-<App>.ps1` | Reports "installed" when the Add/Remove Programs entry exists at the packaged version **or newer** |

All of them:

- Relaunch in **64-bit PowerShell** when Intune starts them as 32-bit, so they see the right registry and Program Files.
- Compare versions as numbers padded to four parts (`8.9.8` = `8.9.8.0`). The installer's version comes from its FileVersion (EXE) or ProductVersion (MSI), then from the file name.
- Return **0** = success, **3010** = soft reboot, **1618** = retry (another installation was running), **1** = failed.
- Write logs to `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\`. Intune's *Collect diagnostics* picks them up.

**Things users can notice:**

- Processes in *Processes to close* are **force-closed** without warning, so users can lose unsaved work in apps that don't restore it.
- Intune shows its own install pop-ups unless notifications are `hideAll`, which is the upload's default.
- A 3010 exit code makes Intune prompt for a restart when restart behavior is *based on return codes*.

---

## 8. Installer types and silent switches

`New-IntuneApp.ps1` recognizes the installer type and suggests these switches:

| Installer type | Install | Installer log | Uninstall |
|---|---|---|---|
| Inno Setup | `/VERYSILENT /NORESTART /SUPPRESSMSGBOXES /SP-` | `/LOG="{0}"` | `/VERYSILENT /NORESTART /SUPPRESSMSGBOXES` |
| NSIS | `/S` | none | `/S` |
| WiX Burn bundle | `/quiet /norestart` | `/log "{0}"` | `/uninstall /quiet /norestart` |
| InstallShield | `/s /v"/qn REBOOT=ReallySuppress"` | none | `/s` (usually MSI, handled automatically) |
| Advanced Installer | `/exenoui /qn /norestart` | none | `/exenoui /qn /norestart` |
| **MSI** | `msiexec /i <msi> /qn /norestart` (+ `TRANSFORMS=`) | `/l*v "{0}"` | `msiexec /x {ProductCode} /qn /norestart` |
| **Unknown** | `/S`: **check the vendor's docs** | none | `/S` |

When the type is **Unknown**, wrong switches make the installer wait for input nobody can see, until Intune times out after 60 minutes. Always test these apps (§4, step 4).
If an app's uninstall entry uses `MsiExec.exe /X{GUID}`, the uninstall script uses `msiexec /x /qn` automatically, whatever the installer type.

---

## 9. MSI notes

- The product name, manufacturer, version, ProductCode and UpgradeCode are read from the MSI itself.
- The suggested Add/Remove Programs name is the exact product name, with the version replaced by `*` (for example `7-Zip * (x64 edition)`).
- Add MSI properties at the *Silent install switches* prompt: `/qn /norestart ALLUSERS=1 INSTALLDIR="C:\Apps\X"`.
- Any `.mst` files in `Source\` are added as `TRANSFORMS="a.mst;b.mst"`.
- **Updates** install the new MSI over the old one. This only works if the vendor's MSI supports upgrades (a MajorUpgrade). If an update fails with **1638**, the old version must be uninstalled first; use supersedence type *Replace* in the portal.
- Intune's built-in MSI detection rule (by ProductCode) is an alternative. The custom detection script is preferred because it keeps working when the ProductCode changes between versions.

---

## 10. Intune settings (manual reference)

The upload step sets all of this. To create or check an app by hand, the same values are in `Documentation\<App>-Intune.md`.

| Tab | Setting | Value |
|---|---|---|
| App information | Name / Publisher / Version | `<Product> <Version>` / from the installer / from the installer |
| Program | Install command | `powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File .\Install-<App>.ps1` |
| | Uninstall command | `powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File .\Uninstall-<App>.ps1` |
| | Install behavior | System |
| | Device restart behavior | Determine behavior based on return codes |
| | Return codes | 0 Success, 3010 Soft reboot, 1641 Hard reboot, 1618 Retry |
| Requirements | Architecture / minimum OS | From the installer name (x64, ARM64, or both) / Windows 10 1809 |
| Detection rules | Custom detection script | `Scripts\Detect-<App>.ps1`. Run as 32-bit: **No**. Enforce signature check: **No**. |
| Assignments | End user notifications | Hide all toast notifications (for silent installs) |

---

## 11. Logs and troubleshooting

### Where to look

| Where | What |
|---|---|
| Device: `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\<App>-Install.log` (`-Update`, `-Uninstall`) | The scripts' own step-by-step logs |
| Device: `...\Logs\<App>-Install-Setup.log` | The installer's own log (MSI, Inno, WiX) |
| Device: `...\Logs\IntuneManagementExtension.log`, `AppWorkload.log` | Intune's side: download, command line, exit code, detection |
| Admin PC: `<app folder>\Documentation\<App>-IntuneUploads.log` | Each upload: name, app ID, supersedence, assignment |

### Find the right Add/Remove Programs name

After a test install, run this in PowerShell:

```powershell
Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' |
    Where-Object DisplayName -like '*7-Zip*' |
    Select-Object DisplayName, DisplayVersion, UninstallString, PSChildName
```

If the name doesn't match the pattern, run `.\New-IntuneApp.ps1` again with the right pattern, then rebuild and upload.

### Common problems

| Symptom | Cause / fix |
|---|---|
| `New-IntuneApp.ps1`: *This is the template folder* | Copy and rename `IntuneAppTemplate`, then run the script in the copy |
| *More than one installer found in Source* | Keep exactly one `.exe` or `.msi` in `Source\`; remove the old version |
| *Could not read a version* | Rename the installer to include the version (for example `App-1.2.3.exe`), or enter it when asked |
| *IntuneWinAppUtil.exe not found* | Copy `Tool\` from the template, or run with `-SkipPackage` |
| Installs, but Intune reports *not detected* | The DisplayName pattern doesn't match, or the installer is user-scope (it writes to HKCU, not HKLM). See above. |
| Install hangs until the 60-minute timeout | Wrong silent switches (an invisible prompt). Check the vendor's docs and test with PsExec. |
| Update fails with **1638** (MSI) | That MSI can't upgrade in place. Use supersedence *Replace*, which uninstalls the old version first. |
| `-UploadOnly`: *No Documentation\\*-Intune.json found* or *Package not found* | Run `.\New-IntuneApp.ps1` without `-UploadOnly` first |
| *No tenant configured* | Run `..\New-IntuneAppRegistration.ps1 -TenantId <customer domain>` |
| *Several tenants are configured - pass -Tenant* | `-NoPrompt` was used without `-Tenant`. Add `-Tenant <name>`. |
| *Tenant 'x' not found* | Use a name from the list it shows, or register the tenant first |
| Sign-in goes to the wrong tenant or account | The browser reused another signed-in account. Pick the right account, or use a private browser profile per customer. |
| Sign-in error *AADSTS65001* (consent) | Admin consent is missing. Run `New-IntuneAppRegistration.ps1` as an admin, or grant consent in the portal. |
| Sign-in error *AADSTS50011* (redirect URI) | The app registration is missing `http://localhost` under *Mobile and desktop applications*. Run `New-IntuneAppRegistration.ps1` again. |
| *Group 'X' not found* | Use the exact group display name, or `-GroupId` |
| *already exists in Intune* | That name and version was already uploaded. Upload again only if you mean to. |

---

## 12. Maintaining the template

- **Change behavior for all future apps:** edit `IntuneAppTemplate\Templates\*.ps1.template`. Placeholders such as `__APPNAME__` and `'__INSTALLARGS__'` are filled in by `New-IntuneApp.ps1`.
- **Apply template changes to an existing app:** copy `New-IntuneApp.ps1` and `Templates\` from the template into the app folder, then run `.\New-IntuneApp.ps1 -NoPrompt`. Its settings are kept.
- **App folders from an older template** (no upload step, or an `Add-IntuneApp.ps1` from before the scripts were combined): do the step above, and delete `Add-IntuneApp.ps1`. Tenant details come from the `IntuneTenant-*.json` files, so nothing else is needed.
- **Adding a customer:** run `New-IntuneAppRegistration.ps1 -TenantId <their domain>`. Every app folder can upload to them straight away.
- **Removing a customer:** delete their `IntuneTenant-<name>.json`, and optionally the *Intune App Upload* app registration in their tenant.
- **Don't edit generated scripts in `Scripts\` by hand.** They're replaced the next time `New-IntuneApp.ps1` runs. Change the settings, or the templates, instead.

---

## 13. Known apps

| App | Folder | Installer type | Install switches | Add/Remove Programs name | Processes closed |
|---|---|---|---|---|---|
| Visual Studio Code (system, x64) | `VSCode` | Inno Setup | `/VERYSILENT /NORESTART /SUPPRESSMSGBOXES /SP- /MERGETASKS=!runcode` | `Microsoft Visual Studio Code` | `Code` |
| Notepad++ (x64) | `Notepad++` | NSIS | `/S` | `Notepad++*` | `notepad++` |
| Obsidian | `Obsidian` | NSIS (electron-builder) | `/S` currently, **probably needs `/S /allusers`**: see note | `*Obsidian*` | (none) - consider `Obsidian` |
| PowerShell 7 (x64) | `PowerShell7` | MSI | `/qn /norestart` | `PowerShell 7-x64` | (none) |

**Obsidian note:** electron-builder installers default to a **per-user** install. Run as SYSTEM with only `/S`, the app can end up in the SYSTEM profile and register under HKCU, so detection (which reads HKLM) fails. Re-run `.\New-IntuneApp.ps1` in the Obsidian folder, enter `/S /allusers` at the install switches prompt and `Obsidian` as the process to close, then test with PsExec and check that it registers under HKLM (§11).

Add a row here when you set up a new app, especially when you had to change the suggested switches or name.
