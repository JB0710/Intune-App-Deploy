# Intune App Deploy - GitHub Actions

This folder is a **ready-to-push repository** that builds Intune Win32 app packages and deploys them to one or more customer tenants from **GitHub Actions**, on your own Windows GitHub-hosted runners. It uses the same app-folder layout and `New-IntuneApp.ps1` as the manual process. Here that script also downloads installers, signs in without a person (certificate), and never prompts.

```
 workflow_dispatch: app + tenants (+ new installer URL)
        │
        ▼
 ┌─ build ───────────────────────┐   ┌─ deploy (per tenant) ────────────┐   ┌─ promote (per tenant) ──────────────┐
 │ download installer + SHA256   │──►│ environment intune-<tenant>      │──►│ environment intune-<tenant>-production│
 │ generate scripts + docs       │   │ upload <Product> <Version>       │   │ ⏸ waits for a reviewer's approval    │
 │ build .intunewin              │   │ supersede previous version       │   │ assign the production group         │
 │ commit changes, keep artifact │   │ assign the pilot group           │   └──────────────────────────────────────┘
 └───────────────────────────────┘   └──────────────────────────────────┘
```

---

## Contents

1. [What's in this folder](#1-whats-in-this-folder)
2. [What's different from the manual process](#2-whats-different-from-the-manual-process)
3. [One-time setup](#3-one-time-setup)
4. [Add an app](#4-add-an-app)
5. [Release a new version](#5-release-a-new-version)
6. [Running the workflow](#6-running-the-workflow)
7. [Credentials and security](#7-credentials-and-security)
8. [Using it from your PC](#8-using-it-from-your-pc)
9. [Troubleshooting](#9-troubleshooting)
10. [Apps in this repository](#10-apps-in-this-repository)

---

## 1. What's in this folder

```
GitHub-Runners\                            ← push this folder as the repository root
├── github-workflows\                     ← move to .github\workflows before the first push (§3.1)
│   ├── intune-app.yml                     ← build + deploy + approval-gated production assignment
│   └── validate.yml                       ← on every PR/push: scripts parse, download.json and tenant files are valid
├── .gitignore                             ← keeps installers, packages and credentials out of git
├── README.md                              ← this file
├── New-IntuneAppRegistration.ps1          ← per tenant: interactive sign-in for people (no stored credential)
├── New-IntuneAutomationRegistration.ps1   ← per tenant: app-only sign-in for GitHub Actions (certificate)
├── IntuneTenant-<name>.json               ← per tenant: tenant ID + client IDs (no secrets - safe to commit)
├── IntuneAppTemplate\                     ← copy for every new app
└── VSCode\  Notepad++\  Obsidian\  PowerShell7\
    ├── New-IntuneApp.ps1                  ← the one script per app (CI-aware version)
    ├── Source\download.json               ← where to download the installer + its SHA256
    ├── Source\<logo>.png                  ← app icon in Intune (not packaged)
    ├── Scripts\                           ← generated Install/Uninstall/Update/Detect scripts
    ├── Documentation\                     ← generated <App>-Intune.md / .json
    ├── Templates\  Tool\                  ← script templates, IntuneWinAppUtil.exe
    └── Output\                            ← built .intunewin (not committed)
```

Installers are **not** committed. GitHub rejects files over 100 MB, and VS Code (250 MB) and Obsidian (341 MB) are larger. Each app has `Source\download.json` instead:

```json
{
  "Url": "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.9.8/npp.8.9.8.Installer.x64.exe",
  "FileName": "npp.8.9.8.Installer.x64.exe",
  "Sha256": "7B2A949BF460FB37A3888C9048698F43222A185A48323023DF1C51E78A3CA1C2"
}
```

The build downloads the file and **rejects it if the SHA256 doesn't match**.

---

## 2. What's different from the manual process

The device scripts (`Install-`, `Uninstall-`, `Update-`, `Detect-`) and templates are identical. Only `New-IntuneApp.ps1` gains new features, and it stays fully compatible with running it by hand.

| Feature | How |
|---|---|
| **Installer download** | If `Source` has no installer, it downloads the one in `Source\download.json` and checks the SHA256. `-InstallerUrl <url> -InstallerSha256 <hash>` downloads a new version, replaces the old installer and updates `download.json`. |
| **Never prompts in CI** | When `GITHUB_ACTIONS`, `TF_BUILD` or `CI=true` is set, it behaves as if `-NoPrompt` was given |
| **App-only sign-in** | Used automatically when `INTUNE_CERT_PFX_BASE64` + `INTUNE_CERT_PASSWORD`, `INTUNE_CERT_THUMBPRINT` (or `-CertificateThumbprint`), or `INTUNE_CLIENT_SECRET` is set. The client ID comes from `AutomationClientId` in the tenant file, or from `INTUNE_CLIENT_ID`. Warns when the certificate expires within 30 days. |
| **`-AssignOnly`** | Assigns an already-uploaded `<Product> <Version>` to a group, used for the production step after approval |
| **Job summary and outputs** | Writes a results row (tenant, app, what happened, link to Intune) to the GitHub job summary, and `app_id` / `app_name` as step outputs |
| **Package contents** | `download.json` and logos stay out of the `.intunewin` |

---

## 3. One-time setup

### 3.1 Create the repository

```powershell
cd 'D:\Scripts\Intune\App Deploy\GitHub-Runners'

# GitHub only runs workflows from .github\workflows
New-Item -ItemType Directory -Path .github -Force | Out-Null
Move-Item -Path .\github-workflows -Destination .\.github\workflows

git init -b main
git add .
git commit -m "Intune app deployment"
git remote add origin https://github.com/<org>/<repo>.git
git push -u origin main
```

Check that `git status` shows no `.exe`, `.msi`, `.intunewin` or `.pfx` files. `.gitignore` excludes them.

### 3.2 Point the workflows at your runners

The workflows use `runs-on: ${{ vars.INTUNE_RUNNER || 'windows-latest' }}`. In the repository (or organization), go to **Settings > Secrets and variables > Actions > Variables** and add:

| Variable | Value |
|---|---|
| `INTUNE_RUNNER` | The label of your Windows GitHub-hosted runner, for example `windows-intune` |

Runner requirements:

- **Windows**, because `IntuneWinAppUtil.exe`, MSI reading and icon extraction are Windows-only.
- **PowerShell 7 (`pwsh`)**. GitHub's Windows images include it; check custom images.
- Outbound HTTPS to:
  - `www.powershellgallery.com` (the IntuneWin32App module, installed on each run)
  - `login.microsoftonline.com`, `graph.microsoft.com`
  - `*.blob.core.windows.net` (Intune content upload)
  - your installer download hosts (`github.com`, `release-assets.githubusercontent.com`, `update.code.visualstudio.com`, …)

  This matters if your runners use a private network or firewall.

### 3.3 Allow the build to commit

The build job commits regenerated scripts, docs and `download.json` back to the branch (input `commit_changes`, on by default).

- **Settings > Actions > General > Workflow permissions:** allow *Read and write*, or the job's `contents: write` is refused.
- If `main` is protected, either let `github-actions[bot]` bypass, or run the workflow with `commit_changes` off and commit the changes yourself. The build artifact contains them.

### 3.4 Set up each tenant

Run this from this folder on your Windows PC, once per customer, as a **Global Administrator or Privileged Role Administrator** of that tenant. Granting *application* permissions needs one of those roles. You also need the [GitHub CLI](https://cli.github.com/) installed and signed in (`gh auth login`).

```powershell
.\New-IntuneAutomationRegistration.ps1 -TenantId contoso.onmicrosoft.com  -GitHubRepo <org>/<repo>
.\New-IntuneAutomationRegistration.ps1 -TenantId fabrikam.onmicrosoft.com -GitHubRepo <org>/<repo> -TenantName FAB
```

For each tenant it:

1. Shows the tenant and asks you to confirm.
2. Creates **Intune App Upload (Automation)** with Microsoft Graph **application** permissions `DeviceManagementApps.ReadWrite.All`, `DeviceManagementRBAC.Read.All` and `Group.Read.All`, and grants admin consent.
3. Creates a self-signed **certificate** (RSA 2048, 12 months; `-ValidityMonths` changes this) and adds its public key to the app.
4. Creates the GitHub environments **`intune-<tenant>`** and **`intune-<tenant>-production`**, and stores `INTUNE_CERT_PFX_BASE64` and `INTUNE_CERT_PASSWORD` as secrets in both.
5. Deletes the local copy of the private key, unless you pass `-KeepLocalCopy`.
6. Adds `AutomationClientId` to **`IntuneTenant-<tenant>.json`**.

Then:

- **Commit the tenant file:** `git add IntuneTenant-*.json; git commit -m "Add tenant"; git push`. It only holds IDs; `validate.yml` fails if it ever contains a credential.
- **Add approvers:** in **Settings > Environments > `intune-<tenant>-production`**, add **Required reviewers**. This is the approval gate before production. Optionally limit both environments to the `main` branch under *Deployment branches*.

Prefer a client secret? Use `-CredentialType Secret`, which stores `INTUNE_CLIENT_SECRET` instead. Certificates are recommended. Without `-GitHubRepo`, the credential is written to `$HOME\IntuneAutomationCredentials\<tenant>` for you to add by hand, and then delete.

Tenants already set up for interactive use (`New-IntuneAppRegistration.ps1`) keep working. Both client IDs are stored in the same tenant file.

---

## 4. Add an app

Choosing an app's install switches, process names and Add/Remove Programs name needs a person, so do the first build on your PC:

1. Copy `IntuneAppTemplate` to a new folder named after the app, for example `7-Zip`.
2. Add the logo to `Source\` (a square `.png`, about 256×256).
3. Run, with the vendor's download link and its SHA256:
   ```powershell
   .\New-IntuneApp.ps1 -InstallerUrl 'https://.../7z2408-x64.msi' -InstallerSha256 '<vendor sha256>' -SkipUpload
   ```
   It downloads the installer, writes `Source\download.json`, asks for the settings, and generates the scripts, docs and package. Without `-InstallerSha256` it records the hash it saw and warns you; compare it with the vendor's checksum.
4. Test the package on a test device (see the main README, *Test on a test device*).
5. Commit the app folder: `git add 7-Zip; git commit -m "Add 7-Zip"; git push`. The installer and package are left out automatically.
6. Run the workflow (§6) with `app: 7-Zip`.

---

## 5. Release a new version

No local steps are needed. Run the workflow with:

| Input | Value |
|---|---|
| `app` | `Notepad++` |
| `installer_url` | the new version's download link |
| `installer_sha256` | the SHA256 the vendor publishes |
| `tenants` | `all` (or `contoso,FAB`) |
| `pilot_group` / `production_group` | as usual |

The build downloads and verifies the new installer, regenerates the scripts with the app's existing settings (only the version changes), builds the package, and commits the new `download.json`, scripts and docs. Deploy uploads `Notepad++ <new version>` and supersedes the previous version in each tenant.

To rebuild **without** changing the version (for example after editing templates), leave `installer_url` empty and the installer from `download.json` is used.

---

## 6. Running the workflow

**Actions > Intune app - build and deploy > Run workflow**, or from a terminal:

```powershell
gh workflow run intune-app.yml -f app='Notepad++' -f tenants=all -f pilot_group='SG-Intune-Pilot' -f production_group='SG-Intune-All-Devices'
```

| Input | Default | Description |
|---|---|---|
| `app` | (required) | App folder name |
| `installer_url` / `installer_sha256` | empty | New version (§5). Empty = use `Source\download.json`. |
| `tenants` | `all` | TenantNames, comma separated, or `all` (every tenant file with an `AutomationClientId`) |
| `pilot_group` | empty | Assigned during upload. Empty = no assignment. |
| `production_group` | empty | Assigned after approval. Empty = the *promote* job is skipped. |
| `intent` | `required` | `required` or `available` |
| `supersede` | `true` | Supersede (update in place) the newest previous version in each tenant |
| `commit_changes` | `true` | Commit regenerated files back to the branch |

What happens:

1. **build** checks the app folder and tenants, builds the package, commits changes, and saves the package as an artifact (kept 14 days).
2. **deploy** runs in parallel per tenant, in environment `intune-<tenant>`. One tenant failing doesn't stop the others. If that version is already in a tenant, it's reported and skipped, not uploaded twice.
3. **promote** runs per tenant, in environment `intune-<tenant>-production`. It waits for a reviewer's **approval**, then assigns the production group. It only runs if every tenant's deploy succeeded.

Each job's **summary** lists the tenant, app, result and a link to the app in the Intune admin center. Generated files have a `concurrency` group per app, so two runs for the same app don't overlap.

---

## 7. Credentials and security

- **Separate apps for people and automation.** `Intune App Upload` (interactive, no stored credential) and `Intune App Upload (Automation)` (app-only, certificate) are separate registrations, so the stored credential can be revoked without affecting people.
- **Least privilege.** The automation app gets only the three application permissions it needs.
- **Secrets stay in GitHub environments.** They're per tenant, and the `-production` environment requires approval. Nothing secret is committed. `.gitignore` excludes `*.pfx`, and `validate.yml` checks tenant files.
- **Rotation.** Run `New-IntuneAutomationRegistration.ps1` again for the tenant before the certificate expires. It replaces the certificate on the app and updates both environment secrets in one go. The workflow log warns when fewer than 30 days remain, and `AutomationCredentialExpires` in the tenant file records the date.
- **Downloads are verified.** The build fails on a SHA256 mismatch. Always use the vendor's published checksum for new versions.
- **Inputs are never pasted into scripts.** Workflow inputs reach the scripts as environment variables, so a crafted group name or URL can't run commands.
- **Revoking access to a tenant:** delete the *Intune App Upload (Automation)* app registration in that tenant, its two GitHub environments, and its `IntuneTenant-<name>.json`.

---

## 8. Using it from your PC

The `New-IntuneApp.ps1` in this repository is a drop-in replacement for the manual one. On your PC it behaves exactly as before: prompts, browser sign-in, and the tenant picker. It adds `-InstallerUrl` and `download.json` support. To upload from your PC with the automation app instead of a browser, use a certificate in your store: `-CertificateThumbprint <thumbprint>`.

---

## 9. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| *No tenant is set up for automation* | Run `New-IntuneAutomationRegistration.ps1` and commit `IntuneTenant-<name>.json` |
| *Unknown or not automation-enabled tenant(s)* | Check the `tenants` input against the TenantNames in the tenant files |
| *No app-only credentials in this CI run* | The environment `intune-<tenant>` (or `-production`) is missing `INTUNE_CERT_PFX_BASE64` / `INTUNE_CERT_PASSWORD`. Re-run the registration script with `-GitHubRepo`. |
| Sign-in error *AADSTS700027* / *invalid client assertion* | The certificate in GitHub doesn't match the one on the app (rotated elsewhere, or expired). Re-run the registration script. |
| *Authorization_RequestDenied* / 403 from Graph | Admin consent for the application permissions is missing. Re-run the script as Global Administrator or Privileged Role Administrator. |
| *Checksum mismatch … The download was rejected* | Wrong `installer_sha256`, or the vendor replaced the file. Check the vendor's published hash. |
| *Couldn't tell the installer's file name* | The URL doesn't end in the file name. Add `"FileName"` to `download.json` (or `-InstallerFileName` locally). |
| Download times out or is refused | The runner can't reach the download host. Check your runner network or proxy (§3.2). |
| *Package not found* in deploy | The build failed or the artifact expired (14 days). Re-run the whole workflow. |
| Commit step: *permission denied* / protected branch | See §3.3, or set `commit_changes` to false |
| promote never starts | It waits for approval. Check the run's *Review deployments* button. It's skipped if any tenant's deploy failed, or if `production_group` was empty. |
| *'<App> <Version>' already exists in Intune* | That version is already in that tenant, so it was skipped. Assign it with the promote step or in the portal. |

---

## 10. Apps in this repository

| Folder | Installer | Download source | Notes |
|---|---|---|---|
| `VSCode` | Inno Setup, system x64 | `https://update.code.visualstudio.com/1.141.0/win32-x64/stable` | Check this URL before the first CI run. It's Microsoft's documented pattern, but couldn't be reached from where this was built. The SHA256 comes from your local copy. |
| `Notepad++` | NSIS x64 | GitHub release `notepad-plus-plus/notepad-plus-plus` | |
| `Obsidian` | NSIS (electron-builder) | GitHub release `obsidianmd/obsidian-releases` | **Fix the install switches first:** run `.\New-IntuneApp.ps1 -SkipUpload` locally, enter `/S /allusers` and process `Obsidian`, test, and commit. With `/S` alone it probably installs per-user, and detection fails. |
| `PowerShell7` | MSI x64 | GitHub release `PowerShell/PowerShell` | |

Each `download.json` SHA256 was taken from the installer you already tested locally, so CI builds exactly the same file.
