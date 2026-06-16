<#
.SYNOPSIS
  Creates the Azure AD App Registration for the VMRay Report Phishing Outlook Add-in.

.DESCRIPTION
  Automates Phase 1 of the deployment guide:
    - Creates (or reuses) an App Registration in the signed-in tenant
    - Adds the required Microsoft Graph delegated permissions
      (Mail.Send, Mail.ReadWrite, User.Read, openid)
    - Creates a client secret with a configurable lifetime
    - Opens the admin-consent URL for a Global Administrator to approve

  On completion, the script prints the three values needed for the
  "Deploy to Azure" template in Phase 2: Client ID, Tenant ID, Client Secret.

.PARAMETER DisplayName
  Display name of the App Registration. Default: VMRay-Outlook-Addin-App

.PARAMETER SecretLifetimeYears
  How long the client secret should remain valid. Allowed values: 1 or 2.
  Default: 2.

.EXAMPLE
  ./setup.ps1

.EXAMPLE
  ./setup.ps1 -DisplayName "Contoso-VMRay-Outlook-Addin" -SecretLifetimeYears 1
#>

[CmdletBinding()]
param(
  [string]$DisplayName = "VMRay-Outlook-Addin-App",

  [ValidateRange(1, 2)]
  [int]$SecretLifetimeYears = 2
)

$ErrorActionPreference = "Stop"
$ProgressPreference     = "SilentlyContinue"

# --- Constants: Microsoft Graph delegated permission IDs --------------------
# These GUIDs are well-known and identical across every Azure AD tenant.
$MICROSOFT_GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
$Permissions = @(
  @{ Name = "Mail.Send";      Id = "e383f46e-2787-4529-855e-0e479a3ffac0" },
  @{ Name = "Mail.ReadWrite"; Id = "024d486e-b451-40bb-833d-3e66d98c5c73" },
  @{ Name = "User.Read";      Id = "e1fe6dd8-ba31-4d61-89e7-88639da4683d" },
  @{ Name = "openid";         Id = "37f7f235-527c-4136-accd-4a02d197296e" }
)

function Write-Step {
  param([string]$Index, [string]$Message)
  Write-Host ""
  Write-Host "[$Index] $Message" -ForegroundColor Cyan
}

# --- Auth helpers -----------------------------------------------------------
# Tries multiple auth strategies in order so the script "just works" across
# environments: local Windows, local Mac/Linux, and Cloud Shell — including
# tenants where Conditional Access blocks device-code flow.
function Connect-MgGraphSmart {
  param([string[]]$Scopes = @("Application.ReadWrite.All", "Directory.Read.All"))

  # Already connected & token still valid?
  $context = Get-MgContext -ErrorAction SilentlyContinue
  if ($context) {
    try {
      Get-MgApplication -Top 1 -ErrorAction Stop | Out-Null
      Write-Host "  Already connected. Reusing existing session." -ForegroundColor Green
      return
    } catch {
      Write-Host "  Existing session is stale. Re-authenticating..." -ForegroundColor Yellow
      Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
  }

  # Strategy 1: pass-through from an existing Az session
  # (covers Cloud Shell, `az login` users, and tenants with Conditional Access
  # policies that block device-code flow).
  $azContext = $null
  try { $azContext = Get-AzContext -ErrorAction Stop } catch { }
  if ($azContext) {
    try {
      Write-Host "  Detected Az session ($($azContext.Account.Id)). Acquiring Graph token..." -ForegroundColor Gray
      $tokenResult = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com" -ErrorAction Stop
      # Handle both SecureString (Az 14+) and plain-string (older Az) returns.
      $secureToken = if ($tokenResult.Token -is [System.Security.SecureString]) {
        $tokenResult.Token
      } else {
        ConvertTo-SecureString $tokenResult.Token -AsPlainText -Force
      }
      Connect-MgGraph -AccessToken $secureToken -NoWelcome -ErrorAction Stop | Out-Null
      Write-Host "  Connected via Az session." -ForegroundColor Green
      return
    } catch {
      Write-Host "  Az pass-through unavailable. Trying interactive browser..." -ForegroundColor Yellow
    }
  }

  # Strategy 2: interactive browser (default Connect-MgGraph behavior)
  try {
    Write-Host "  Opening browser for sign-in..." -ForegroundColor Gray
    Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop | Out-Null
    Write-Host "  Connected via interactive browser." -ForegroundColor Green
    return
  } catch {
    Write-Host "  Interactive browser failed. Falling back to device code..." -ForegroundColor Yellow
  }

  # Strategy 3: device code (last resort — may be blocked by Conditional Access)
  Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome -ErrorAction Stop | Out-Null
  Write-Host "  Connected via device code." -ForegroundColor Green
}

# Wraps a Graph call. If it fails with an auth error, refreshes the session
# once and retries. Lets the script survive a token expiring mid-run.
function Invoke-WithGraphRetry {
  param(
    [Parameter(Mandatory=$true)][scriptblock]$Block,
    [string]$Description = "Graph operation"
  )
  try {
    return & $Block
  } catch {
    $errText = "$($_.Exception.Message) $($_.ErrorDetails.Message)"
    if ($errText -match "Authentication needed|InvalidAuthenticationToken|401|TokenExpired") {
      Write-Host "  Token expired during '$Description'. Refreshing..." -ForegroundColor Yellow
      Connect-MgGraphSmart
      return & $Block
    }
    throw
  }
}

# --- 1. Ensure the Microsoft Graph PowerShell module is installed ----------
Write-Step "1/6" "Checking for Microsoft.Graph.Applications module..."
$module = Get-Module -ListAvailable -Name Microsoft.Graph.Applications | Select-Object -First 1
if (-not $module) {
  Write-Host "  Module not found. Installing for the current user..." -ForegroundColor Yellow
  Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Applications -ErrorAction Stop

# --- 2. Connect to Microsoft Graph ------------------------------------------
Write-Step "2/6" "Connecting to Microsoft Graph..."
Connect-MgGraphSmart

$context = Get-MgContext
if (-not $context) { throw "Failed to connect to Microsoft Graph." }
$tenantId = $context.TenantId
Write-Host "  Connected to tenant: $tenantId" -ForegroundColor Green

# --- 3. Create or reuse the App Registration -------------------------------
Write-Step "3/6" "Creating App Registration '$DisplayName'..."
$escapedName = $DisplayName.Replace("'", "''")
$existing = Invoke-WithGraphRetry -Description "looking up existing App Registration" -Block {
  Get-MgApplication -Filter "displayName eq '$escapedName'" -ErrorAction SilentlyContinue | Select-Object -First 1
}

if ($existing) {
  Write-Host "  App Registration already exists (AppId: $($existing.AppId)). Reusing." -ForegroundColor Yellow
  $app = $existing
} else {
  # Include a placeholder Web redirect URI so the admin-consent flow has somewhere
  # to redirect after the Global Admin clicks Accept. Microsoft requires at least
  # one redirect URI on the App Reg, otherwise AADSTS500113 ("No reply address
  # registered") fires. The Microsoft-owned 'nativeclient' URL is the standard
  # safe placeholder. finalize.ps1 later adds the real SPA redirect URI; this
  # Web entry stays harmless.
  $app = Invoke-WithGraphRetry -Description "creating App Registration" -Block {
    New-MgApplication `
      -DisplayName $DisplayName `
      -SignInAudience "AzureADMyOrg" `
      -Web @{ RedirectUris = @('https://login.microsoftonline.com/common/oauth2/nativeclient') }
  }
  Write-Host "  Created (AppId: $($app.AppId))" -ForegroundColor Green
}

# --- 4. Set Microsoft Graph delegated permissions --------------------------
Write-Step "4/6" "Setting required Microsoft Graph delegated permissions..."
$resourceAccess = @(
  $Permissions | ForEach-Object {
    @{ Id = $_.Id; Type = "Scope" }   # "Scope" = delegated permission
  }
)
$requiredResourceAccess = @(
  @{
    ResourceAppId  = $MICROSOFT_GRAPH_APP_ID
    ResourceAccess = $resourceAccess
  }
)
Invoke-WithGraphRetry -Description "setting permissions" -Block {
  Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $requiredResourceAccess
}
foreach ($p in $Permissions) {
  Write-Host "  + $($p.Name)" -ForegroundColor Green
}

# --- 5. Create a client secret ----------------------------------------------
Write-Step "5/6" "Creating client secret (valid for $SecretLifetimeYears year(s))..."
$secretParams = @{
  PasswordCredential = @{
    DisplayName = "VMRay add-in secret (created $(Get-Date -Format 'yyyy-MM-dd'))"
    EndDateTime = (Get-Date).AddYears($SecretLifetimeYears)
  }
}
$secret = Invoke-WithGraphRetry -Description "creating client secret" -Block {
  Add-MgApplicationPassword -ApplicationId $app.Id -BodyParameter $secretParams
}
Write-Host "  Secret created (expires: $($secret.EndDateTime.ToString('yyyy-MM-dd')))" -ForegroundColor Green

# --- 6. Open the admin-consent URL -----------------------------------------
Write-Step "6/6" "Opening admin-consent URL in the default browser..."
$consentUrl = "https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$($app.AppId)"
Write-Host "  $consentUrl" -ForegroundColor Gray
Write-Host "  Sign in as a Global Administrator and click 'Accept'." -ForegroundColor Gray
try {
  Start-Process $consentUrl -ErrorAction Stop | Out-Null
} catch {
  Write-Host "  (Could not auto-open a browser in this environment. Copy the URL above and open it manually.)" -ForegroundColor Yellow
}

# --- Persist non-secret state for the finalize step ------------------------
$stateFile = Join-Path $PSScriptRoot "deployment-state.json"
@{
  AppId       = $app.AppId
  AppObjectId = $app.Id
  TenantId    = $tenantId
  DisplayName = $DisplayName
  CreatedAt   = (Get-Date).ToString("o")
} | ConvertTo-Json | Out-File -FilePath $stateFile -Encoding utf8
Write-Host ""
Write-Host "  State written to: $stateFile" -ForegroundColor Gray

# --- Output summary --------------------------------------------------------
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Green
Write-Host "  App Registration is ready. Use these values for the ARM deploy:"     -ForegroundColor Green
Write-Host "========================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Azure Client ID     : $($app.AppId)"
Write-Host "  Azure Tenant ID     : $tenantId"
Write-Host "  Azure Client Secret : $($secret.SecretText)"
Write-Host ""
Write-Host "  IMPORTANT: Save the client secret now. It will not be shown again." -ForegroundColor Yellow
Write-Host ""
Write-Host "  Next: click 'Deploy to Azure' in the README and paste these values," -ForegroundColor Cyan
Write-Host "        then run finalize.ps1 once the Web App is deployed."           -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Green

Disconnect-MgGraph | Out-Null
