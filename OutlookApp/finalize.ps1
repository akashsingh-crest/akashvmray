<#
.SYNOPSIS
  Configures an existing Azure AD App Registration with the values needed
  by the VMRay Report Phishing Outlook Add-in middle-tier (Phase 3 of the
  deployment guide).

.DESCRIPTION
  Run this AFTER setup.ps1 (which created the App Registration) and AFTER
  deploying the Web App via the ARM template (which gives you a domain).

  Performs the four configuration tasks of README Phase 3:
    - Sets the Application ID URI: api://<domain>/<clientId>
    - Adds the access_as_user scope under that URI (admin-consent only)
    - Pre-authorizes the Office host client for that scope
    - Adds the SPA redirect URI: https://<domain>/fallbackauthdialog.html

  All four operations are idempotent — safe to re-run with the same
  parameters. Re-running with a different domain updates the App
  Registration in place.

  After this completes, the next steps are:
    1. Download the rendered manifest from https://<domain>/manifest.xml
    2. Upload it in the Microsoft 365 Admin Center (Phase 5)

.PARAMETER Domain
  The deployed Web App's default domain, e.g. "myapp.azurewebsites.net".
  Required. https:// prefix and trailing slashes are stripped automatically.

.PARAMETER AppId
  Application (Client) ID of the App Registration to configure.
  If omitted, reads from deployment-state.json next to this script
  (written by setup.ps1).

.PARAMETER StateFile
  Path to deployment-state.json. Defaults to a sibling file of this script.

.EXAMPLE
  ./finalize.ps1 -Domain "vmray-outlook-akash.azurewebsites.net"

.EXAMPLE
  ./finalize.ps1 -Domain "myapp.azurewebsites.net" -AppId "abc12345-..."
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$Domain,

  [string]$AppId,

  [string]$StateFile = (Join-Path $PSScriptRoot "deployment-state.json")
)

$ErrorActionPreference = "Stop"
$ProgressPreference     = "SilentlyContinue"

# --- Constants --------------------------------------------------------------
# Microsoft's well-known client ID covering every Office host application
# (Outlook desktop / web / mobile, Word, Excel, etc.). Pre-authorizing this
# client lets Office hosts request our access_as_user scope without showing
# end users a consent dialog.
$OFFICE_CLIENT_ID = "ea5a67f6-b6f3-4338-b240-c655ddc3cc8e"

function Write-Step {
  param([string]$Index, [string]$Message)
  Write-Host ""
  Write-Host "[$Index] $Message" -ForegroundColor Cyan
}

# --- Auth helpers (mirrors setup.ps1) ---------------------------------------
# Picks the best auth strategy for the current environment:
#   1. Reuse an existing valid Microsoft Graph session
#   2. Pass through an existing Az session (Cloud Shell / az login)
#   3. Interactive browser
#   4. Device code (last resort; often blocked by Conditional Access)
function Connect-MgGraphSmart {
  param([string[]]$Scopes = @("Application.ReadWrite.All", "Directory.Read.All"))

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

  $azContext = $null
  try { $azContext = Get-AzContext -ErrorAction Stop } catch { }
  if ($azContext) {
    try {
      Write-Host "  Detected Az session ($($azContext.Account.Id)). Acquiring Graph token..." -ForegroundColor Gray
      $tokenResult = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com" -ErrorAction Stop
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

  try {
    Write-Host "  Opening browser for sign-in..." -ForegroundColor Gray
    Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop | Out-Null
    Write-Host "  Connected via interactive browser." -ForegroundColor Green
    return
  } catch {
    Write-Host "  Interactive browser failed. Falling back to device code..." -ForegroundColor Yellow
  }

  Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome -ErrorAction Stop | Out-Null
  Write-Host "  Connected via device code." -ForegroundColor Green
}

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

# --- Normalize Domain input -------------------------------------------------
# Accept "myapp.azurewebsites.net" or "https://myapp.azurewebsites.net/";
# strip the protocol and any trailing slash.
$Domain = $Domain.Trim()
$Domain = $Domain -replace '^https?://', ''
$Domain = $Domain.TrimEnd('/')

# --- 1. Resolve the App Registration to configure ---------------------------
Write-Step "1/5" "Resolving target App Registration..."

if (-not $AppId) {
  if (-not (Test-Path $StateFile)) {
    throw "AppId not provided and state file '$StateFile' not found. Pass -AppId, or run setup.ps1 first."
  }
  $state = Get-Content $StateFile -Raw | ConvertFrom-Json
  $AppId = $state.AppId
  Write-Host "  Loaded AppId from $StateFile : $AppId" -ForegroundColor Gray
} else {
  Write-Host "  Using provided AppId: $AppId" -ForegroundColor Gray
}

# --- 2. Ensure module + connect to Microsoft Graph --------------------------
Write-Step "2/5" "Connecting to Microsoft Graph..."
$module = Get-Module -ListAvailable -Name Microsoft.Graph.Applications | Select-Object -First 1
if (-not $module) {
  Write-Host "  Microsoft.Graph.Applications module not found. Installing for the current user..." -ForegroundColor Yellow
  Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Applications -ErrorAction Stop

Connect-MgGraphSmart

$context = Get-MgContext
if (-not $context) { throw "Failed to connect to Microsoft Graph." }
$tenantId = $context.TenantId
Write-Host "  Connected to tenant: $tenantId" -ForegroundColor Green

# --- 3. Load the App Registration ------------------------------------------
Write-Step "3/5" "Loading App Registration from tenant..."
$app = Invoke-WithGraphRetry -Description "loading App Registration" -Block {
  Get-MgApplication -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue | Select-Object -First 1
}
if (-not $app) {
  throw "App Registration with AppId '$AppId' not found in tenant $tenantId. Did you run setup.ps1 first?"
}
Write-Host "  Found: $($app.DisplayName)" -ForegroundColor Green

# --- 4. Compute target configuration ----------------------------------------
Write-Step "4/5" "Configuring App Registration for domain '$Domain'..."

# (a) Application ID URI -----------------------------------------------------
$identifierUri = "api://$Domain/$($app.AppId)"
Write-Host "  Identifier URI: $identifierUri" -ForegroundColor Gray

# (b) access_as_user scope ---------------------------------------------------
# Idempotency: if a scope named "access_as_user" already exists on this App
# Reg, reuse its Id so any existing pre-authorized clients keep pointing at
# the same scope. Otherwise mint a fresh GUID.
$existingScope = $app.Api.Oauth2PermissionScopes |
  Where-Object { $_.Value -eq "access_as_user" } |
  Select-Object -First 1

if ($existingScope) {
  $scopeId = $existingScope.Id
  Write-Host "  Reusing existing 'access_as_user' scope (Id: $scopeId)" -ForegroundColor Yellow
} else {
  $scopeId = [guid]::NewGuid().ToString()
  Write-Host "  Will create 'access_as_user' scope (Id: $scopeId)" -ForegroundColor Gray
}

$accessAsUserScope = @{
  Id                      = $scopeId
  Value                   = "access_as_user"
  Type                    = "Admin"   # Admin-only consent; matches README guidance
  IsEnabled               = $true
  AdminConsentDisplayName = "Access VMRay Outlook Add-in API"
  AdminConsentDescription = "Allows the Outlook Add-in to access the middle-tier API on behalf of the signed-in user."
}

# Preserve any other scopes that already exist on this App Reg.
$otherScopes = @($app.Api.Oauth2PermissionScopes | Where-Object { $_.Value -ne "access_as_user" })
$allScopes   = @($accessAsUserScope) + $otherScopes

# (c) Pre-authorize the Office host client -----------------------------------
$officePreAuth = @{
  AppId                  = $OFFICE_CLIENT_ID
  DelegatedPermissionIds = @($scopeId)
}

# Replace any existing Office entry (its DelegatedPermissionIds may have been
# pointing at a stale scope Id from a prior run); preserve unrelated entries.
$otherPreAuth = @($app.Api.PreAuthorizedApplications | Where-Object { $_.AppId -ne $OFFICE_CLIENT_ID })
$allPreAuth   = @($officePreAuth) + $otherPreAuth

# (d) SPA redirect URI -------------------------------------------------------
$spaUri          = "https://$Domain/fallbackauthdialog.html"
$existingSpaUris = @($app.Spa.RedirectUris)
$allSpaUris = if ($existingSpaUris -contains $spaUri) {
  Write-Host "  SPA redirect URI already present" -ForegroundColor Yellow
  $existingSpaUris
} else {
  $existingSpaUris + $spaUri
}

# Apply updates in TWO sequential PATCHes:
#   #1: create the scope + set identifier URI + add SPA redirect.
#   #2: add pre-authorized clients that reference the now-existing scope.
# Microsoft Graph validates preAuthorizedApplications.delegatedPermissionIds
# against the pre-PATCH scope list, so a one-shot PATCH that creates the
# scope AND references it fails with InvalidValue.
$preAuthWithoutOffice = @($app.Api.PreAuthorizedApplications | Where-Object { $_.AppId -ne $OFFICE_CLIENT_ID })

Invoke-WithGraphRetry -Description "updating App Registration (pass 1)" -Block {
  Update-MgApplication -ApplicationId $app.Id `
    -IdentifierUris @($identifierUri) `
    -Api @{
      Oauth2PermissionScopes    = $allScopes
      PreAuthorizedApplications = $preAuthWithoutOffice
    } `
    -Spa @{
      RedirectUris = $allSpaUris
    } `
    -ErrorAction Stop
} | Out-Null

Write-Host "  + Application ID URI"   -ForegroundColor Green
Write-Host "  + access_as_user scope" -ForegroundColor Green
Write-Host "  + SPA redirect URI"     -ForegroundColor Green

Invoke-WithGraphRetry -Description "updating App Registration (pass 2)" -Block {
  Update-MgApplication -ApplicationId $app.Id `
    -Api @{
      Oauth2PermissionScopes    = $allScopes
      PreAuthorizedApplications = $allPreAuth
    } `
    -ErrorAction Stop
} | Out-Null

Write-Host "  + Office client pre-auth" -ForegroundColor Green

# --- 5. Done — print next steps ---------------------------------------------
Write-Step "5/5" "Configuration complete."

$manifestUrl    = "https://$Domain/manifest.xml"
$diagnosticsUrl = "https://$Domain/diagnostics"

Write-Host ""
Write-Host "========================================================================" -ForegroundColor Green
Write-Host "  App Registration is fully configured for '$Domain'."                  -ForegroundColor Green
Write-Host "========================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Manifest URL (download for upload in M365 Admin Center):"
Write-Host "    $manifestUrl" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Diagnostics URL (verify the deployed Web App is healthy):"
Write-Host "    $diagnosticsUrl" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor Cyan
Write-Host "    1. Open the manifest URL in a browser and save the downloaded file."
Write-Host "    2. Upload manifest.xml in the Microsoft 365 Admin Center"
Write-Host "       (Settings > Integrated apps > Deploy Add-in)."
Write-Host "       This step requires a Global Administrator."
Write-Host "========================================================================" -ForegroundColor Green
