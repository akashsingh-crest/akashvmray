# VMRay Report Phishing Outlook Add-in — Deployment Guide

This repository contains the necessary components and instructions to deploy the VMRay Report Phishing Add-in for Outlook. This tool allows users to report suspicious emails directly to a VMRay Incident Response (IR) mailbox for automated analysis.

---

## Introduction

### Microsoft Outlook Add-ins

Microsoft Outlook Add-ins are web-based extensions that integrate directly into Microsoft Outlook (Desktop, Web, and Mobile). They allow organizations to extend Outlook's functionality by embedding custom workflows directly inside the mailbox experience.

Using Microsoft 365 Single Sign-On (SSO) and Microsoft Graph API, add-ins can securely interact with user mailbox data without storing credentials, while maintaining enterprise-grade security and compliance.

### About VMRay

VMRay is a leading provider of automated malware analysis and advanced threat detection solutions. Using hypervisor-based sandboxing technology, VMRay delivers deep visibility into sophisticated and evasive cyber threats.

The **VMRay Report Phishing Outlook Add-in** enables users to:

- Report suspicious emails with a single click
- Securely forward the original email (including attachments) to a designated VMRay Incident Response (IR) mailbox
- Automatically move reported emails to a dedicated Outlook folder
- Authenticate seamlessly using Microsoft 365 SSO

This add-in streamlines phishing reporting workflows while maintaining security, transparency, and user simplicity.

---

## Prerequisites

| Requirement | Why |
|---|---|
| **Azure Subscription** | To host the Web App that powers the add-in |
| **Global Administrator** (in your Microsoft 365 tenant) | Required to grant tenant-wide consent during deployment and to upload the manifest to the Microsoft 365 Admin Center |
| **VMRay IR Mailbox** | Designated email address (ending in `ir-mailbox.vmray.com`) that receives reported phishing emails. Provisioned via your VMRay Cloud Portal under **Analysis Settings → IR Mailbox** |
| **PowerShell environment** | Azure Cloud Shell |

---

## Deployment Overview

The entire deployment is driven by a single PowerShell script that handles three phases:

| Phase | What it does | Manual? |
|---|---|---|
| **1. App Registration** | Creates the Azure AD App Registration, adds the required Microsoft Graph permissions, mints a client secret | Automated |
| **1.5. Admin consent** | A Global Administrator must click a one-time URL to grant tenant-wide consent | **Manual click** |
| **2. Web App deployment** | Provisions the Azure Web App via an ARM template, configures environment variables | Automated |
| **3. App Reg configuration** | Wires the App Registration to the deployed Web App's domain (Application ID URI, scope, redirect URIs, pre-authorization) | Automated |
| **4. Manifest download** | Customer downloads a ready-to-upload `manifest.xml` from the deployed Web App | One browser click |
| **5. M365 Admin Center upload** | Global Administrator uploads the manifest in M365 Admin Center to deploy to users | **Manual upload** |

**Total customer-side effort: one PowerShell command + two Global Admin clicks + one file upload.**

---

## Quick Start — Cloud Shell
### Step 1 — Open Cloud Shell

1. Sign in to [Azure Portal](https://portal.azure.com) with your tenant administrator account
2. Click the **`>_`** Cloud Shell icon in the top-right toolbar
3. Choose **PowerShell** if prompted

### Step 2 — Upload the deployment script and ARM template

Click **Manage files → Upload** in the Cloud Shell toolbar and upload these two files from this repo:

- `Scripts/Deploy-VMRayOutlookAddin.ps1`
- `WebApp/azuredeploy.json`

### Step 3 — Run the deployment script

```powershell
./Deploy-VMRayOutlookAddin.ps1 -TemplateFile ~/azuredeploy.json
```

The script is fully interactive — it'll prompt you for everything it needs. Default values appear in brackets; press Enter to accept.

You'll be asked:

| Prompt | What to enter |
|---|---|
| Confirm tenant + subscription | Press Enter to confirm (or N to abort and switch context) |
| Create new or use existing App Registration? | Choose 1 (new) for a first-time deployment |
| Display name for the new App Registration | Press Enter for default (`VMRay-Outlook-Addin-App`) |
| **Open the printed consent URL → click Accept (Global Admin)** | (Manual step) |
| Press Enter once consent is granted | After clicking Accept in the browser |
| Web App name (1-20 chars, alphanumeric+hyphens) | A globally unique name, e.g., `vmray-outlook-acme` |
| Resource group name | Existing RG, or a new name (created if missing) |
| Azure region | Press Enter for default (Central US) |
| App Service SKU | Press Enter for default (S1) |
| IR mailbox email address | Your VMRay IR mailbox, e.g., `xxx@us.ir-mailbox.vmray.com` |
| Move reported emails to a folder? | Press Enter for default (No) |
| Proceed with deployment? | Press Enter to confirm |

The deployment runs for 5-7 minutes. The script handles everything automatically including App Registration configuration after the Web App is deployed.

### Step 4 — Note the manifest URL

When the script finishes, it prints a summary like:

```
Web App URL    : https://vmray-outlook-acme.azurewebsites.net
App Client ID  : 13b429ba-f816-4453-9a62-e9a30ee930f2
Tenant ID      : c91fb25b-dd00-4300-bcfa-9dfd39d1451a

Manifest URL (download for M365 Admin Center):
  https://vmray-outlook-acme.azurewebsites.net/manifest.xml

Diagnostics URL (verify deployment health):
  https://vmray-outlook-acme.azurewebsites.net/diagnostics
```

### Step 5 — Download the manifest

Open the Manifest URL from Step 4 in a browser. It downloads `manifest.xml` automatically. **No editing required** — the file is fully populated with your deployment's values.

### Step 6 — Upload the manifest in Microsoft 365 Admin Center

This step requires Global Administrator access to the Microsoft 365 Admin Center.

1. Sign in to [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. Navigate to **Settings → Integrated apps**
3. Click **Add-ins** tab → **Deploy Add-in**
4. Click **Next** → **Upload custom apps**
5. Choose the `manifest.xml` you downloaded in Step 5
6. Click **Upload**
7. Choose deployment scope:
   - **Everyone** — all users in the tenant
   - **Specific users/groups** — selected users (recommended for staged rollouts)
   - **Just me** — recommended for initial testing
8. Click **Deploy**

Microsoft propagates the add-in to user mailboxes within ~30 minutes (sometimes up to 24 hours).

### Step 7 — Verify

After propagation, users see a **Report Phishing** button in their Outlook ribbon when viewing any email. Click it to forward the message to the VMRay IR mailbox for analysis.

See the [Verification](#verification) section for details.

---


## Re-deployment / Reusing an Existing App Registration

If you already have an App Registration from a previous deployment (e.g., migrating to a different Web App, or sharing one App Reg across dev/staging/prod):

When prompted "How do you want to handle the App Registration?", choose **option 2 — Use an existing App Registration**.

The script will:

1. Look up the App Registration by name or AppId
2. If multiple matches exist, ask you to pick one
3. Ask whether to use your existing client secret or mint a fresh one
4. Verify the required permissions (add any that are missing — preserving any custom permissions you've added)
5. Verify admin consent is already granted (or prompt if not)
6. Continue with Web App deployment as normal

This is also useful for re-running the script after a deployment failure — it's idempotent at every step.

---

## Verification

### Quick health check

Open your Web App's diagnostics URL in a browser:

```
https://YOUR-WEBAPP.azurewebsites.net/diagnostics
```

Expected response:

```json
{
  "status": "ok",
  "timestamp": "...",
  "checks": {
    "env": { "status": "ok", "details": { ... } },
    "recipient": { "status": "ok" },
    "manifestTemplate": { "status": "ok", "path": "..." },
    "moveReportedEmailsToFolder": false
  },
  "manifestUrl": "https://YOUR-WEBAPP.azurewebsites.net/manifest.xml"
}
```

`"status": "ok"` confirms all required env vars are set, the recipient mailbox domain is valid, and the manifest template is in place.

### End-to-end test

1. Sign in to [Outlook on the web](https://outlook.office.com) as a user who has the add-in deployed
2. Open any email
3. Click **Report Phishing** in the ribbon
4. Click **Confirm** in the task pane
5. Wait for the success message: ✅ Email successfully reported to Security Team
6. Open your VMRay portal (e.g., `https://us.cloud.vmray.com`) → **Submissions**
7. Filter: `Interface Type == IR Mailbox`
8. Within 1-2 minutes, your reported email should appear as a new submission

If all of the above succeed, your deployment is fully working.

---

## Troubleshooting

### "Consent verification didn't see all permissions granted"

**What's happening:** After clicking Accept on the consent URL, Microsoft's grant database takes 10-90 seconds to propagate to the Graph API that the script uses to verify consent. The script retries for up to 90 seconds before showing this warning.

**What to do:** The script will offer three options:

1. **Re-open the consent URL** — most common fix. Sometimes the first Accept click doesn't actually register (browser caching, multiple accounts signed in, etc.). Re-clicking usually works. **Tip:** open the URL in an InPrivate/Incognito window for the cleanest session.
2. **Continue anyway** — script proceeds; consent must be granted manually later before users can use the add-in
3. **Abort** — bail out completely

**Recommended:** choose option 1. If it still fails after several re-clicks, verify in Portal → Entra ID → App Registration → API permissions. Look for green checkmarks in the **Status** column.

### "Temporary server issue" when users click Report Phishing

**Most likely cause:** Admin consent wasn't actually granted, even though the script may have continued.

**To diagnose:** Open the Web App's Log stream in Azure Portal. Look for:

```
forwardMail ERROR: Failed to obtain Graph access token
```

This is the signature of missing consent.

**To fix:** Open the consent URL again as a Global Admin:

```
https://login.microsoftonline.com/YOUR-TENANT-ID/adminconsent?client_id=YOUR-APP-CLIENT-ID
```

Click Accept. Wait 1-2 minutes for propagation. Have the user retry.

### Manifest cache on Outlook doesn't update after manifest changes

**What's happening:** Microsoft's M365 caches each tenant's add-in manifest aggressively. After uploading an updated manifest in M365 Admin Center, individual Outlook clients can take 5-30 minutes to refresh their cached copy (sometimes up to 24 hours).

**To force a faster refresh:**
1. Sign out of `outlook.office.com` completely
2. Close all browser tabs
3. Open a fresh InPrivate window
4. Sign in again

For Outlook desktop, close the application completely and reopen.

If the manifest version was bumped (e.g., from `1.0.0.0` to `1.0.1.0`), Microsoft refreshes more aggressively than same-version updates.

### Web App quota errors during deployment

**Symptom:** ARM deployment fails with `SubscriptionIsOverQuotaForSku` or `Total VMs: 0`.

**Why:** Some Azure subscriptions have zero default vCPU quota in unused regions.

**Fix:** Pick a different region when the script prompts. Recommended alternatives:

- Central US
- East US 2
- West US 2
- North Europe

Quotas are per-region — if one fails, try another.

### Script can't find the ARM template

**Symptom:** `ARM template not found at: <path>`

**Cause:** When you upload files individually to Cloud Shell, they land flat in your home directory (no folder structure). The script's default `../WebApp/azuredeploy.json` path doesn't resolve.

**Fix:** Pass the template path explicitly:

```powershell
./Deploy-VMRayOutlookAddin.ps1 -TemplateFile ~/azuredeploy.json
```

This isn't needed when running from a cloned repo with the original folder structure.

### Multiple App Registrations with the same name

**Symptom:** The script finds multiple App Registrations matching your input and asks you to pick one.

**Cause:** Previous test deployments left duplicates in your tenant. Azure AD allows multiple App Regs with the same display name.

**Fix:** The script disambiguates by showing a numbered list with Creation dates and AppIds. Pick the one you want. Cleanup of unused ones can be done manually in Entra ID → App registrations after deployment.

---

## Summary — Comparison with the Original (Manual) Flow

| Step | Original (6 phases) | New (script-driven) |
|---|---|---|
| App Registration setup | ~15 portal clicks | 1 prompt (`./Deploy-VMRayOutlookAddin.ps1`) |
| Admin consent | Manual click | Manual click (unavoidable) |
| Web App deployment | Portal "Deploy to Azure" + manual env var | Auto via script |
| App Reg configuration | ~10 portal clicks | Auto via script |
| Manifest preparation | Manual find/replace in `manifest.xml` | Auto via `/manifest.xml` route |
| M365 Admin Center upload | 1 admin action | 1 admin action (unavoidable) |
| **Total manual interactions** | ~30 portal clicks across multiple pages | **3 actions** (1 PowerShell command + 1 consent click + 1 manifest upload) |

The single-command deployment is the recommended path for all new installations. The legacy step-by-step guide in the original `README.md` remains available for documentation purposes.
