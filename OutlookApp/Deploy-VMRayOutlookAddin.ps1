<#
.SYNOPSIS
  End-to-end deployment of the VMRay Report Phishing Outlook Add-in.

.DESCRIPTION
  Single interactive script that handles all three deployment phases:

    Phase 1: Azure AD App Registration (create new OR use existing)
    Phase 2: Azure Web App deployment via ARM template
    Phase 3: App Registration configuration (Application ID URI, scopes,
             pre-authorization, SPA redirect)

  All inputs are prompted interactively if not provided as parameters. A
  Global Administrator must click the admin-consent URL once during Phase 1
  to authorize the App Registration's permissions - that's the only manual
  step the script cannot automate.

  The individual scripts (setup.ps1 / deploy.ps1 / finalize.ps1) remain
  available for advanced / CI scenarios.

.PARAMETER DisplayName
  Display name for a NEW App Registration. Ignored if using existing.

.PARAMETER AppId
  Application (Client) ID of an EXISTING App Registration. If provided,
  the script uses this instead of creating a new one. Setting this
  implies "Use existing".

.PARAMETER ClientSecret
  Existing client secret value (SecureString). Only used if AppId is
  provided. If you pass AppId but not ClientSecret, the script will offer
  to mint a new one on the existing App Reg.

.PARAMETER WebAppName
  Globally unique name for the new Azure Web App (1-20 chars, alphanumeric+hyphens).

.PARAMETER ResourceGroup
  Resource group to deploy into. Created if it doesn't exist.

.PARAMETER Region
  Azure region for the resource group. Default: "Central US".

.PARAMETER Sku
  App Service pricing tier. Default: "S1".

.PARAMETER Recipient
  IR mailbox email address (must end in ir-mailbox.vmray.com).

.PARAMETER MoveReportedPhishingEmailsToFolder
  "true" or "false". Default: "false".

.PARAMETER TemplateFile
  Path to ARM template. Defaults to ../WebApp/azuredeploy.json next to this script.

.PARAMETER SkipSetup
  Skip Phase 1. Requires -AppId and -ClientSecret.

.PARAMETER SkipDeploy
  Skip Phase 2. Requires -WebAppName (uses existing Web App).

.PARAMETER SkipFinalize
  Skip Phase 3.

.EXAMPLE
  # Fully interactive
  ./Deploy-VMRayOutlookAddin.ps1

.EXAMPLE
  # Re-deploy reusing an existing App Reg
  ./Deploy-VMRayOutlookAddin.ps1 -AppId "abc12345-..."
#>

[CmdletBinding()]
param(
  # Phase 1 - App Registration
  [string]$DisplayName,
  [string]$AppId,
  [SecureString]$ClientSecret,
  [ValidateRange(1, 2)][int]$SecretLifetimeYears = 2,

  # Phase 2 - Deploy
  [string]$WebAppName,
  [string]$ResourceGroup,
  [string]$Region = "Central US",
  [ValidateSet("B1","B2","B3","S1","S2","S3","P1v3","P2v3")]
  [string]$Sku = "S1",
  [string]$Recipient,
  [ValidateSet("true","false")]
  [string]$MoveReportedPhishingEmailsToFolder = "false",

  # Other
  [string]$TemplateFile = (Join-Path $PSScriptRoot "../WebApp/azuredeploy.json"),

  # Skip flags
  [switch]$SkipSetup,
  [switch]$SkipDeploy,
  [switch]$SkipFinalize
)

$ErrorActionPreference = "Stop"
$ProgressPreference     = "SilentlyContinue"

# Microsoft Graph delegated permission GUIDs (well-known, identical across tenants)
$MICROSOFT_GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
$REQUIRED_PERMISSIONS = @(
  @{ Name = "Mail.Send";      Id = "e383f46e-2787-4529-855e-0e479a3ffac0" },
  @{ Name = "Mail.ReadWrite"; Id = "024d486e-b451-40bb-833d-3e66d98c5c73" },
  @{ Name = "User.Read";      Id = "e1fe6dd8-ba31-4d61-89e7-88639da4683d" },
  @{ Name = "openid";         Id = "37f7f235-527c-4136-accd-4a02d197296e" }
)

# Microsoft's well-known Office host client ID (for pre-authorization)
$OFFICE_CLIENT_ID = "ea5a67f6-b6f3-4338-b240-c655ddc3cc8e"

# ===========================================================================
#  Display helpers
# ===========================================================================
function Write-Banner {
  param([string]$Title)
  Write-Host ""
  Write-Host "===========================================================================" -ForegroundColor Magenta
  Write-Host "  $Title" -ForegroundColor Magenta
  Write-Host "===========================================================================" -ForegroundColor Magenta
}

function Write-Phase {
  param([string]$Number, [string]$Title)
  Write-Host ""
  Write-Host ">>> Phase $Number - $Title" -ForegroundColor Cyan
  Write-Host ""
}

function Write-Step {
  param([string]$Index, [string]$Message)
  Write-Host ""
  Write-Host "[$Index] $Message" -ForegroundColor Cyan
}

# ===========================================================================
#  Interactive prompt helpers
# ===========================================================================
function Read-Choice {
  param(
    [string]$Prompt,
    [string[]]$Options,
    [int]$Default = 1
  )
  Write-Host ""
  Write-Host $Prompt -ForegroundColor Yellow
  for ($i = 0; $i -lt $Options.Length; $i++) {
    $marker = if (($i + 1) -eq $Default) { " (default)" } else { "" }
    Write-Host "  [$($i + 1)] $($Options[$i])$marker"
  }
  do {
    $input = Read-Host "  Choice"
    if ([string]::IsNullOrWhiteSpace($input)) { return $Default }
    $num = 0
    if ([int]::TryParse($input, [ref]$num) -and $num -ge 1 -and $num -le $Options.Length) {
      return $num
    }
    Write-Host "    Invalid. Enter a number between 1 and $($Options.Length)." -ForegroundColor Red
  } while ($true)
}

function Read-Text {
  param(
    [string]$Prompt,
    [string]$Default = $null,
    [string]$ValidationPattern = $null,
    [string]$ValidationMessage = "Invalid input. Try again."
  )
  do {
    $defaultHint = if ($Default) { " [default: $Default]" } else { "" }
    Write-Host ""
    $value = Read-Host "  $Prompt$defaultHint"
    if ([string]::IsNullOrWhiteSpace($value) -and $Default) {
      return $Default
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
      Write-Host "    Required. Please enter a value." -ForegroundColor Red
      continue
    }
    if (-not $ValidationPattern -or $value -match $ValidationPattern) {
      return $value
    }
    Write-Host "    $ValidationMessage" -ForegroundColor Red
  } while ($true)
}

function Confirm-Action {
  param([string]$Prompt, [bool]$Default = $true)
  $defaultStr = if ($Default) { "Y/n" } else { "y/N" }
  do {
    Write-Host ""
    $input = (Read-Host "  $Prompt [$defaultStr]").Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($input)) { return $Default }
    if ($input -in @("y","yes")) { return $true }
    if ($input -in @("n","no"))  { return $false }
    Write-Host "    Please answer y or n." -ForegroundColor Red
  } while ($true)
}

function Wait-ForEnter {
  param([string]$Message = "Press ENTER when ready to continue, or Ctrl+C to abort")
  Write-Host ""
  Write-Host "  $Message" -ForegroundColor Yellow
  Read-Host | Out-Null
}

# ===========================================================================
#  Auth helpers (mirrors setup.ps1 / finalize.ps1 patterns)
# ===========================================================================
function Connect-MgGraphSmart {
  param([string[]]$Scopes = @("Application.ReadWrite.All", "Directory.Read.All"))

  $context = Get-MgContext -ErrorAction SilentlyContinue
  if ($context) {
    try {
      Get-MgApplication -Top 1 -ErrorAction Stop | Out-Null
      Write-Host "  Already connected to Microsoft Graph. Reusing session." -ForegroundColor Green
      return
    } catch {
      Write-Host "  Existing Graph session is stale. Re-authenticating..." -ForegroundColor Yellow
      Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
  }

  $azContext = $null
  try { $azContext = Get-AzContext -ErrorAction Stop } catch { }
  if ($azContext) {
    try {
      Write-Host "  Using Az session to acquire Graph token..." -ForegroundColor Gray
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
    Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop | Out-Null
    Write-Host "  Connected via interactive browser." -ForegroundColor Green
    return
  } catch {
    Write-Host "  Interactive browser failed. Falling back to device code..." -ForegroundColor Yellow
  }

  Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome -ErrorAction Stop | Out-Null
  Write-Host "  Connected via device code." -ForegroundColor Green
}

function Connect-AzSmart {
  $azContext = Get-AzContext -ErrorAction SilentlyContinue
  if ($azContext) {
    Write-Host "  Azure session: $($azContext.Account.Id) (tenant $($azContext.Tenant.Id))" -ForegroundColor Green
    return
  }
  Write-Host "  No Azure session. Starting sign-in..." -ForegroundColor Gray
  Connect-AzAccount -ErrorAction Stop | Out-Null
  $azContext = Get-AzContext
  Write-Host "  Connected as $($azContext.Account.Id)." -ForegroundColor Green
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

# ===========================================================================
#  Module loaders
# ===========================================================================
function Ensure-Module {
  param([string]$Name)
  $module = Get-Module -ListAvailable -Name $Name | Select-Object -First 1
  if (-not $module) {
    Write-Host "  Installing $Name (current user, one-time)..." -ForegroundColor Yellow
    Install-Module $Name -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module $Name -ErrorAction Stop
}

# ===========================================================================
#  Phase 1 helpers - App Registration
# ===========================================================================
function Find-AppRegistration {
  param([string]$NameOrAppId)

  # Try as AppId (GUID) first
  if ($NameOrAppId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
    $byAppId = Invoke-WithGraphRetry -Description "looking up App Reg by AppId" -Block {
      Get-MgApplication -Filter "appId eq '$NameOrAppId'" -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if ($byAppId) { return @($byAppId) }
  }

  # Otherwise, by display name (can return multiple)
  $escapedName = $NameOrAppId.Replace("'", "''")
  $byName = Invoke-WithGraphRetry -Description "looking up App Reg by name" -Block {
    @(Get-MgApplication -Filter "displayName eq '$escapedName'" -ErrorAction SilentlyContinue)
  }
  return $byName
}

function Select-AppRegistration {
  param([array]$Candidates)
  if ($Candidates.Count -eq 1) { return $Candidates[0] }

  Write-Host ""
  Write-Host "  Found $($Candidates.Count) App Registrations with that name:" -ForegroundColor Yellow
  for ($i = 0; $i -lt $Candidates.Count; $i++) {
    $c = $Candidates[$i]
    $created = if ($c.CreatedDateTime) { $c.CreatedDateTime.ToString("yyyy-MM-dd") } else { "unknown" }
    Write-Host ("    [{0}] AppId: {1}  Created: {2}" -f ($i + 1), $c.AppId, $created)
  }

  do {
    $input = Read-Host "  Which one do you want to use? [1-$($Candidates.Count)]"
    $num = 0
    if ([int]::TryParse($input, [ref]$num) -and $num -ge 1 -and $num -le $Candidates.Count) {
      return $Candidates[$num - 1]
    }
    Write-Host "    Invalid. Enter a number between 1 and $($Candidates.Count)." -ForegroundColor Red
  } while ($true)
}

function Add-MissingPermissions {
  param($App)

  $currentResourceAccess = @($App.RequiredResourceAccess)
  $graphEntry = $currentResourceAccess | Where-Object { $_.ResourceAppId -eq $MICROSOFT_GRAPH_APP_ID } | Select-Object -First 1

  $existingScopeIds = if ($graphEntry) {
    @($graphEntry.ResourceAccess | Where-Object { $_.Type -eq "Scope" } | ForEach-Object { $_.Id })
  } else { @() }

  $missing = $REQUIRED_PERMISSIONS | Where-Object { $existingScopeIds -notcontains $_.Id }

  if ($missing.Count -eq 0) {
    Write-Host "  All 4 required permissions already present." -ForegroundColor Green
    return
  }

  Write-Host "  Adding missing permissions:" -ForegroundColor Yellow
  $missing | ForEach-Object { Write-Host "    + $($_.Name)" -ForegroundColor Yellow }

  # Build the merged Graph entry (keep existing scopes, add missing ones)
  $allGraphScopes = @($existingScopeIds + ($missing | ForEach-Object { $_.Id })) | Select-Object -Unique
  $newGraphResourceAccess = @($allGraphScopes | ForEach-Object {
    @{ Id = $_; Type = "Scope" }
  })

  # Preserve non-Scope (e.g., Role) entries on Graph
  if ($graphEntry) {
    $nonScope = @($graphEntry.ResourceAccess | Where-Object { $_.Type -ne "Scope" })
    $newGraphResourceAccess = $newGraphResourceAccess + ($nonScope | ForEach-Object { @{ Id = $_.Id; Type = $_.Type } })
  }

  # Preserve non-Graph resource entries
  $nonGraphEntries = @($currentResourceAccess | Where-Object { $_.ResourceAppId -ne $MICROSOFT_GRAPH_APP_ID } |
    ForEach-Object {
      @{
        ResourceAppId  = $_.ResourceAppId
        ResourceAccess = @($_.ResourceAccess | ForEach-Object { @{ Id = $_.Id; Type = $_.Type } })
      }
    })

  $newRequiredResourceAccess = @(
    @{
      ResourceAppId  = $MICROSOFT_GRAPH_APP_ID
      ResourceAccess = $newGraphResourceAccess
    }
  ) + $nonGraphEntries

  Invoke-WithGraphRetry -Description "adding missing permissions" -Block {
    Update-MgApplication -ApplicationId $App.Id -RequiredResourceAccess $newRequiredResourceAccess
  }
}

function Test-AdminConsentGranted {
  param([string]$AppIdToCheck)

  try {
    $sp = Invoke-WithGraphRetry -Description "looking up service principal" -Block {
      Get-MgServicePrincipal -Filter "appId eq '$AppIdToCheck'" -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if (-not $sp) { return $false }   # No SP yet = no consent

    $grants = Invoke-WithGraphRetry -Description "listing OAuth2 grants" -Block {
      Get-MgOauth2PermissionGrant -Filter "clientId eq '$($sp.Id)'" -ErrorAction SilentlyContinue
    }
    if (-not $grants) { return $false }

    # Look for a grant for Microsoft Graph that includes all our required scopes
    $graphSp = Invoke-WithGraphRetry -Description "looking up Graph SP" -Block {
      Get-MgServicePrincipal -Filter "appId eq '$MICROSOFT_GRAPH_APP_ID'" -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if (-not $graphSp) { return $false }

    $grantsForGraph = @($grants | Where-Object { $_.ResourceId -eq $graphSp.Id })
    if ($grantsForGraph.Count -eq 0) { return $false }

    $grantedScopes = @()
    foreach ($g in $grantsForGraph) {
      if ($g.Scope) { $grantedScopes += ($g.Scope -split ' ') }
    }
    $grantedScopes = $grantedScopes | Where-Object { $_ } | Select-Object -Unique

    $requiredScopeNames = $REQUIRED_PERMISSIONS | ForEach-Object { $_.Name }
    foreach ($req in $requiredScopeNames) {
      if ($grantedScopes -notcontains $req) { return $false }
    }
    return $true
  } catch {
    return $false   # If anything errors, assume not granted (safer)
  }
}

# Retry the consent verification up to N times with delay between attempts.
# Returns $true once consent is detected on all 4 scopes; $false after timeout.
# Used to absorb the 10-90s propagation lag between Microsoft's consent DB
# and Graph API's read-side.
function Wait-ForConsentToPropagate {
  param(
    [Parameter(Mandatory=$true)][string]$AppIdToCheck,
    [int]$MaxRetries = 6,
    [int]$RetryDelay = 15
  )

  Write-Host "  Verifying consent..." -ForegroundColor Gray

  for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
    Start-Sleep -Seconds $RetryDelay

    if (Test-AdminConsentGranted -AppIdToCheck $AppIdToCheck) {
      Write-Host "  Consent confirmed for all 4 permissions." -ForegroundColor Green
      return $true
    }

    if ($attempt -lt $MaxRetries) {
      $elapsed = $attempt * $RetryDelay
      Write-Host "  Still propagating (${elapsed}s elapsed), retrying..." -ForegroundColor Gray
    }
  }

  return $false
}

# ===========================================================================
#                              SCRIPT START
# ===========================================================================

Write-Banner "VMRay Outlook Add-in - End-to-End Deployment"

Write-Host ""
Write-Host "  This script will:" -ForegroundColor Gray
Write-Host "    1. Create or reuse an Azure AD App Registration" -ForegroundColor Gray
Write-Host "    2. Deploy the Azure Web App via ARM template" -ForegroundColor Gray
Write-Host "    3. Configure the App Registration with the new Web App's domain" -ForegroundColor Gray
Write-Host ""
Write-Host "  Phase 1 includes one manual checkpoint (Global Admin clicks consent URL)." -ForegroundColor Gray

# ---------------------------------------------------------------------------
# Pre-flight: connect + show context
# ---------------------------------------------------------------------------
Write-Phase "0" "Pre-flight"

Write-Step "1/2" "Loading PowerShell modules..."
Ensure-Module Microsoft.Graph.Applications
Ensure-Module Az.Resources

Write-Step "2/2" "Connecting to Microsoft Graph and Azure..."
Connect-MgGraphSmart
Connect-AzSmart

$mgContext = Get-MgContext
$azContext = Get-AzContext
$tenantId  = $mgContext.TenantId

Write-Host ""
Write-Host "  Tenant       : $tenantId" -ForegroundColor White
Write-Host "  Subscription : $($azContext.Subscription.Name) ($($azContext.Subscription.Id))" -ForegroundColor White
Write-Host ""
if (-not (Confirm-Action "Continue with this tenant + subscription?")) {
  throw "Aborted by user. Switch context (Connect-AzAccount / Connect-MgGraph) and retry."
}

# ===========================================================================
# PHASE 1 - App Registration
# ===========================================================================
$skipPhase1 = $SkipSetup.IsPresent

if ($skipPhase1) {
  Write-Phase "1" "App Registration - SKIPPED (--SkipSetup)"
  if (-not $AppId)        { throw "-SkipSetup requires -AppId" }
  if (-not $ClientSecret) { throw "-SkipSetup requires -ClientSecret" }
  $app = Invoke-WithGraphRetry -Description "loading App Reg" -Block {
    Get-MgApplication -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue | Select-Object -First 1
  }
  if (-not $app) { throw "App Registration with AppId '$AppId' not found." }
  $appClientId       = $app.AppId
  $appClientSecret   = $ClientSecret
}
else {
  Write-Phase "1" "App Registration"

  # Determine NEW vs EXISTING
  $appChoice = if ($AppId) { 2 }
               else { Read-Choice -Prompt "How do you want to handle the App Registration?" `
                                  -Options @("Create a new App Registration",
                                             "Use an existing App Registration") `
                                  -Default 1 }

  if ($appChoice -eq 1) {
    # ----- NEW App Reg path -----
    if (-not $DisplayName) {
      $DisplayName = Read-Text -Prompt "Display name for the new App Registration" `
                                -Default "VMRay-Outlook-Addin-App"
    }

    # Warn on duplicates
    $existing = Invoke-WithGraphRetry -Description "checking for duplicates" -Block {
      $escName = $DisplayName.Replace("'", "''")
      @(Get-MgApplication -Filter "displayName eq '$escName'")
    }
    if ($existing.Count -gt 0) {
      Write-Host ""
      Write-Host "  WARNING: $($existing.Count) App Registration(s) named '$DisplayName' already exist:" -ForegroundColor Yellow
      $existing | ForEach-Object { Write-Host "    - AppId: $($_.AppId)" -ForegroundColor Yellow }
      if (-not (Confirm-Action "Create another with the same name?" -Default $false)) {
        throw "Aborted. Re-run with a different -DisplayName or choose 'Use existing'."
      }
    }

    Write-Step "1/3" "Creating App Registration '$DisplayName'..."
    # https://portal.azure.com/ as the placeholder Web redirect URI lands the
    # Global Admin on the Azure portal home page after consent - friendlier than
    # the older 'nativeclient' URL which produced a "Not the right page" message.
    $app = Invoke-WithGraphRetry -Description "creating App Reg" -Block {
      New-MgApplication `
        -DisplayName $DisplayName `
        -SignInAudience "AzureADMyOrg" `
        -Web @{ RedirectUris = @('https://portal.azure.com/') }
    }
    Write-Host "  Created (AppId: $($app.AppId))" -ForegroundColor Green

    Write-Step "2/3" "Adding required Microsoft Graph permissions..."
    $resourceAccess = @($REQUIRED_PERMISSIONS | ForEach-Object { @{ Id = $_.Id; Type = "Scope" } })
    Invoke-WithGraphRetry -Description "setting permissions" -Block {
      Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @(
        @{ ResourceAppId = $MICROSOFT_GRAPH_APP_ID; ResourceAccess = $resourceAccess }
      )
    }
    $REQUIRED_PERMISSIONS | ForEach-Object { Write-Host "    + $($_.Name)" -ForegroundColor Green }

    Write-Step "3/3" "Creating client secret (valid for $SecretLifetimeYears year(s))..."
    $secretParams = @{
      PasswordCredential = @{
        DisplayName = "VMRay add-in secret (created $(Get-Date -Format 'yyyy-MM-dd'))"
        EndDateTime = (Get-Date).AddYears($SecretLifetimeYears)
      }
    }
    $secret = Invoke-WithGraphRetry -Description "creating secret" -Block {
      Add-MgApplicationPassword -ApplicationId $app.Id -BodyParameter $secretParams
    }
    Write-Host "  Secret created (expires: $($secret.EndDateTime.ToString('yyyy-MM-dd')))" -ForegroundColor Green
    $appClientSecret = ConvertTo-SecureString $secret.SecretText -AsPlainText -Force

    # ----- Save secret for customer record -----
    Write-Host ""
    Write-Host "  +----------------------------------------------------------------------+" -ForegroundColor Yellow
    Write-Host "  |  CLIENT SECRET (visible only once - save it now if you need it):     |" -ForegroundColor Yellow
    Write-Host "  |                                                                      |" -ForegroundColor Yellow
    Write-Host "  |  $($secret.SecretText)" -ForegroundColor White
    Write-Host "  |                                                                      |" -ForegroundColor Yellow
    Write-Host "  +----------------------------------------------------------------------+" -ForegroundColor Yellow
  }
  else {
    # ----- EXISTING App Reg path -----
    if (-not $AppId) {
      $nameOrId = Read-Text -Prompt "Enter App Registration name OR Application (Client) ID"
    } else {
      $nameOrId = $AppId
    }

    $candidates = Find-AppRegistration -NameOrAppId $nameOrId
    if ($candidates.Count -eq 0) {
      throw "No App Registration found matching '$nameOrId'. Run again with the correct name/AppId, or choose 'Create new'."
    }

    $app = Select-AppRegistration -Candidates $candidates
    Write-Host ""
    Write-Host "  Selected: $($app.DisplayName)  (AppId: $($app.AppId))" -ForegroundColor Green

    Write-Step "1/2" "Verifying and adding missing permissions..."
    Add-MissingPermissions -App $app

    Write-Step "2/2" "Resolving client secret..."
    if ($ClientSecret) {
      Write-Host "  Using ClientSecret provided via parameter." -ForegroundColor Green
      $appClientSecret = $ClientSecret
    } else {
      $secretChoice = Read-Choice -Prompt "Do you have the client secret for this App Registration?" `
                                  -Options @("Yes - I'll paste it now",
                                             "No - generate a new client secret on this App Reg") `
                                  -Default 2
      if ($secretChoice -eq 1) {
        Write-Host ""
        Write-Host "  Paste the client secret value (input hidden):" -ForegroundColor Yellow
        $appClientSecret = Read-Host -AsSecureString "  Client Secret"
      } else {
        Write-Host "  Creating a fresh client secret (existing secrets stay valid)..." -ForegroundColor Yellow
        $secretParams = @{
          PasswordCredential = @{
            DisplayName = "VMRay add-in secret (created $(Get-Date -Format 'yyyy-MM-dd'))"
            EndDateTime = (Get-Date).AddYears($SecretLifetimeYears)
          }
        }
        $secret = Invoke-WithGraphRetry -Description "minting new secret" -Block {
          Add-MgApplicationPassword -ApplicationId $app.Id -BodyParameter $secretParams
        }
        Write-Host "  Secret created (expires: $($secret.EndDateTime.ToString('yyyy-MM-dd')))" -ForegroundColor Green
        $appClientSecret = ConvertTo-SecureString $secret.SecretText -AsPlainText -Force
        Write-Host ""
        Write-Host "  +----------------------------------------------------------------------+" -ForegroundColor Yellow
        Write-Host "  |  NEW CLIENT SECRET (visible only once):                              |" -ForegroundColor Yellow
        Write-Host "  |                                                                      |" -ForegroundColor Yellow
        Write-Host "  |  $($secret.SecretText)" -ForegroundColor White
        Write-Host "  |                                                                      |" -ForegroundColor Yellow
        Write-Host "  +----------------------------------------------------------------------+" -ForegroundColor Yellow
      }
    }
  }

  $appClientId = $app.AppId

  # ----- Admin consent -----
  Write-Step "Consent" "Verifying admin consent status..."
  $consentAlreadyGranted = Test-AdminConsentGranted -AppIdToCheck $appClientId

  if ($consentAlreadyGranted) {
    Write-Host "  Admin consent already granted for all 4 permissions. Skipping." -ForegroundColor Green
  } else {
    $consentUrl = "https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$appClientId"
    Write-Host ""
    Write-Host "  Admin consent is REQUIRED before the add-in will work for users." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  A Global Administrator must:" -ForegroundColor Yellow
    Write-Host "    1. Open this URL:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    $consentUrl" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    2. Sign in as Global Admin." -ForegroundColor Yellow
    Write-Host "    3. Click 'Accept'." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  After clicking Accept, you'll be redirected to the Azure portal home page." -ForegroundColor Gray
    Write-Host "  That confirms consent was granted - return here and press ENTER." -ForegroundColor Gray
    try { Start-Process $consentUrl -ErrorAction Stop | Out-Null } catch { }

    Wait-ForEnter "Press ENTER once consent has been granted"

    # Verify with retries - usually 10-30s, can take up to ~90s.
    Write-Host ""
    $consentVerified = Wait-ForConsentToPropagate -AppIdToCheck $appClientId

    # If verification failed after 90s, offer the user a recovery loop:
    # most "first click didn't register" cases are fixed by clicking the
    # consent URL again. Loop until verified, continue-anyway, or abort.
    $continueWithoutConsent = $false
    while (-not $consentVerified -and -not $continueWithoutConsent) {
      Write-Host ""
      Write-Host "  WARNING: Consent verification didn't see all permissions granted after 90 seconds." -ForegroundColor Yellow
      Write-Host "  Either propagation is unusually slow, OR the Accept click didn't actually register." -ForegroundColor Yellow
      Write-Host "  Verify in Portal: Entra ID -> App Reg -> API permissions -> Status column." -ForegroundColor Gray

      $recoveryChoice = Read-Choice -Prompt "What would you like to do?" -Options @(
        "Re-open the consent URL and try again (recommended)",
        "Continue with deployment anyway (add-in won't work until consent is fixed later)",
        "Abort"
      ) -Default 1

      switch ($recoveryChoice) {
        1 {
          Write-Host ""
          Write-Host "  Re-opening consent URL:" -ForegroundColor Cyan
          Write-Host "    $consentUrl" -ForegroundColor Gray
          Write-Host "  Tip: try an InPrivate / Incognito window, or sign out of other Microsoft accounts first." -ForegroundColor Gray
          try { Start-Process $consentUrl -ErrorAction Stop | Out-Null } catch { }
          Wait-ForEnter "After clicking Accept again, press ENTER"
          $consentVerified = Wait-ForConsentToPropagate -AppIdToCheck $appClientId
        }
        2 {
          Write-Host "  Continuing without verified consent. Add-in won't work until consent is granted manually." -ForegroundColor Yellow
          $continueWithoutConsent = $true
        }
        3 {
          throw "Aborted by user. Re-run after consent propagates."
        }
      }
    }
  }
}

# ===========================================================================
# PHASE 2 - ARM Deploy
# ===========================================================================
if ($SkipDeploy.IsPresent) {
  Write-Phase "2" "ARM Deploy - SKIPPED (--SkipDeploy)"
  if (-not $WebAppName)    { $WebAppName    = Read-Text -Prompt "Existing Web App name (without .azurewebsites.net)" `
                                                        -ValidationPattern '^[a-zA-Z0-9-]{1,40}$' }
  $domain = "$WebAppName.azurewebsites.net"
  Write-Host "  Using existing Web App at: $domain" -ForegroundColor Yellow
}
else {
  Write-Phase "2" "Deploy Azure Web App"

  if (-not $WebAppName) {
    $WebAppName = Read-Text -Prompt "Web App name (1-20 chars, alphanumeric+hyphens)" `
                            -ValidationPattern '^[a-zA-Z0-9-]{1,20}$' `
                            -ValidationMessage "Must be 1-20 chars, letters/numbers/hyphens only."
  }

  if (-not $ResourceGroup) {
    $ResourceGroup = Read-Text -Prompt "Resource group name (will be created if it doesn't exist)"
  }

  if (-not $PSBoundParameters.ContainsKey('Region')) {
    $Region = Read-Text -Prompt "Azure region" -Default $Region
  }

  if (-not $PSBoundParameters.ContainsKey('Sku')) {
    $Sku = Read-Text -Prompt "App Service SKU (B1/B2/B3/S1/S2/S3/P1v3/P2v3)" -Default $Sku
  }

  if (-not $Recipient) {
    $Recipient = Read-Text -Prompt "IR mailbox email address" `
                           -ValidationPattern '^[^\s@]+@[^\s@]+\.[^\s@]+$' `
                           -ValidationMessage "Must be a valid email address."
  }

  if (-not $PSBoundParameters.ContainsKey('MoveReportedPhishingEmailsToFolder')) {
    $moveChoice = Read-Choice -Prompt "Move reported phishing emails to a 'Phishing Report' folder?" `
                              -Options @("No (just forward, keep in inbox)",
                                         "Yes (forward AND move to a folder)") `
                              -Default 1
    $MoveReportedPhishingEmailsToFolder = if ($moveChoice -eq 2) { "true" } else { "false" }
  }

  # Confirm before destructive action
  Write-Host ""
  Write-Host "  About to deploy with:" -ForegroundColor Yellow
  Write-Host "    Web App Name : $WebAppName" -ForegroundColor White
  Write-Host "    Resource Grp : $ResourceGroup" -ForegroundColor White
  Write-Host "    Region       : $Region" -ForegroundColor White
  Write-Host "    SKU          : $Sku" -ForegroundColor White
  Write-Host "    Recipient    : $Recipient" -ForegroundColor White
  Write-Host "    Move to fldr : $MoveReportedPhishingEmailsToFolder" -ForegroundColor White
  Write-Host "    Template     : $TemplateFile" -ForegroundColor White
  if (-not (Confirm-Action "Proceed with deployment?")) {
    throw "Aborted by user."
  }

  # Resource group
  Write-Step "1/3" "Ensuring resource group exists..."
  $rg = Get-AzResourceGroup -Name $ResourceGroup -ErrorAction SilentlyContinue
  if ($rg) {
    Write-Host "  Exists in region '$($rg.Location)'." -ForegroundColor Green
  } else {
    Write-Host "  Creating resource group in '$Region'..." -ForegroundColor Yellow
    New-AzResourceGroup -Name $ResourceGroup -Location $Region | Out-Null
    Write-Host "  Created." -ForegroundColor Green
  }

  # Validate template
  Write-Step "2/3" "Validating ARM template..."
  if (-not (Test-Path $TemplateFile)) { throw "ARM template not found at: $TemplateFile" }

  # Use splatting (named dynamic parameters) instead of -TemplateParameterObject.
  # The Object form serializes the hashtable to JSON, which breaks on SecureString.
  # Splatting binds via the cmdlet's dynamic parameters and handles SecureString correctly.
  $testParams = @{
    ResourceGroupName                   = $ResourceGroup
    TemplateFile                        = $TemplateFile
    WebAppName                          = $WebAppName
    AzureClientID                       = $appClientId
    AzureClientSecret                   = $appClientSecret
    AzureTenantID                       = $tenantId
    Sku                                 = $Sku
    Recipient                           = $Recipient
    MoveReportedPhishingEmailsToFolder  = $MoveReportedPhishingEmailsToFolder
    ErrorAction                         = "Stop"
  }
  try {
    $validationResult = Test-AzResourceGroupDeployment @testParams
    if ($validationResult) {
      Write-Host "  Validation FAILED:" -ForegroundColor Red
      $validationResult | ForEach-Object { Write-Host "    - $($_.Message)" -ForegroundColor Red }
      throw "Template validation failed."
    }
    Write-Host "  Validation passed." -ForegroundColor Green
  } catch {
    if ($_.Exception.Message -match "serialize secure string|SecureString") {
      Write-Host "  Pre-validation skipped (Az SecureString serialization quirk - harmless)." -ForegroundColor Yellow
      Write-Host "  The actual deployment in the next step will surface any template errors." -ForegroundColor Gray
    } else {
      throw
    }
  }

  # Deploy
  Write-Step "3/3" "Deploying (this takes 3-7 minutes)..."
  $deploymentName = "vmray-outlook-$(Get-Date -Format 'yyyyMMddHHmmss')"
  Write-Host "  Deployment name: $deploymentName" -ForegroundColor Gray

  $deployParams = @{
    Name                                = $deploymentName
    ResourceGroupName                   = $ResourceGroup
    TemplateFile                        = $TemplateFile
    WebAppName                          = $WebAppName
    AzureClientID                       = $appClientId
    AzureClientSecret                   = $appClientSecret
    AzureTenantID                       = $tenantId
    Sku                                 = $Sku
    Recipient                           = $Recipient
    MoveReportedPhishingEmailsToFolder  = $MoveReportedPhishingEmailsToFolder
    ErrorAction                         = "Stop"
  }
  $deployment = New-AzResourceGroupDeployment @deployParams
  if ($deployment.ProvisioningState -ne "Succeeded") {
    throw "Deployment finished with state '$($deployment.ProvisioningState)'."
  }
  Write-Host "  Deployment succeeded." -ForegroundColor Green

  # Use the actual hostname Azure assigned (handles unique-hostname feature, sovereign clouds)
  $domain = if ($deployment.Outputs.WebAppURL) {
    ($deployment.Outputs.WebAppURL.Value -replace '^https?://', '').TrimEnd('/')
  } else {
    "$WebAppName.azurewebsites.net"
  }
  Write-Host "  Web App URL: https://$domain" -ForegroundColor Green
}

# ===========================================================================
# PHASE 3 - Finalize App Registration
# ===========================================================================
if ($SkipFinalize.IsPresent) {
  Write-Phase "3" "Configure App Registration - SKIPPED (--SkipFinalize)"
}
else {
  Write-Phase "3" "Configure App Registration for domain '$domain'"

  # Reload app fresh (state may have changed since Phase 1)
  $app = Invoke-WithGraphRetry -Description "reloading App Reg" -Block {
    Get-MgApplication -Filter "appId eq '$appClientId'" -ErrorAction SilentlyContinue | Select-Object -First 1
  }

  # Compute identifier URI
  $identifierUri = "api://$domain/$appClientId"
  Write-Host "  Identifier URI: $identifierUri" -ForegroundColor Gray

  # Compute scope (reuse existing access_as_user Id if present)
  $existingScope = $app.Api.Oauth2PermissionScopes | Where-Object { $_.Value -eq "access_as_user" } | Select-Object -First 1
  $scopeId = if ($existingScope) {
    Write-Host "  Reusing existing 'access_as_user' scope (Id: $($existingScope.Id))" -ForegroundColor Yellow
    $existingScope.Id
  } else {
    Write-Host "  Creating 'access_as_user' scope" -ForegroundColor Gray
    [guid]::NewGuid().ToString()
  }

  $accessAsUserScope = @{
    Id                      = $scopeId
    Value                   = "access_as_user"
    Type                    = "Admin"
    IsEnabled               = $true
    AdminConsentDisplayName = "Access VMRay Outlook Add-in API"
    AdminConsentDescription = "Allows the Outlook Add-in to access the middle-tier API on behalf of the signed-in user."
  }
  $otherScopes = @($app.Api.Oauth2PermissionScopes | Where-Object { $_.Value -ne "access_as_user" })
  $allScopes   = @($accessAsUserScope) + $otherScopes

  # Office pre-auth
  $officePreAuth = @{ AppId = $OFFICE_CLIENT_ID; DelegatedPermissionIds = @($scopeId) }
  $otherPreAuth  = @($app.Api.PreAuthorizedApplications | Where-Object { $_.AppId -ne $OFFICE_CLIENT_ID })
  $allPreAuth    = @($officePreAuth) + $otherPreAuth

  # SPA URI
  $spaUri = "https://$domain/fallbackauthdialog.html"
  $existingSpaUris = @($app.Spa.RedirectUris)
  $allSpaUris = if ($existingSpaUris -contains $spaUri) { $existingSpaUris } else { $existingSpaUris + $spaUri }

  # Two-pass PATCH (Graph validates pre-auth against pre-PATCH scope list)
  $preAuthWithoutOffice = @($app.Api.PreAuthorizedApplications | Where-Object { $_.AppId -ne $OFFICE_CLIENT_ID })

  Invoke-WithGraphRetry -Description "Phase 3 pass 1" -Block {
    Update-MgApplication -ApplicationId $app.Id `
      -IdentifierUris @($identifierUri) `
      -Api @{ Oauth2PermissionScopes = $allScopes; PreAuthorizedApplications = $preAuthWithoutOffice } `
      -Spa @{ RedirectUris = $allSpaUris } `
      -ErrorAction Stop
  } | Out-Null

  Write-Host "  + Application ID URI"   -ForegroundColor Green
  Write-Host "  + access_as_user scope" -ForegroundColor Green
  Write-Host "  + SPA redirect URI"     -ForegroundColor Green

  Invoke-WithGraphRetry -Description "Phase 3 pass 2" -Block {
    Update-MgApplication -ApplicationId $app.Id `
      -Api @{ Oauth2PermissionScopes = $allScopes; PreAuthorizedApplications = $allPreAuth } `
      -ErrorAction Stop
  } | Out-Null

  Write-Host "  + Office client pre-auth" -ForegroundColor Green
}

# ===========================================================================
# FINAL SUMMARY
# ===========================================================================
Write-Banner "Deployment complete"

$manifestUrl    = "https://$domain/manifest.xml"
$diagnosticsUrl = "https://$domain/diagnostics"

Write-Host ""
Write-Host "  Web App URL    : https://$domain" -ForegroundColor White
Write-Host "  App Client ID  : $appClientId" -ForegroundColor White
Write-Host "  Tenant ID      : $tenantId" -ForegroundColor White
Write-Host ""
Write-Host "  Manifest URL (download for M365 Admin Center):" -ForegroundColor Cyan
Write-Host "    $manifestUrl" -ForegroundColor White
Write-Host ""
Write-Host "  Diagnostics URL (verify deployment health):" -ForegroundColor Cyan
Write-Host "    $diagnosticsUrl" -ForegroundColor White
Write-Host ""
Write-Host "  Next steps (manual - only 1 step left!):" -ForegroundColor Yellow
Write-Host "    1. Open the Manifest URL in a browser - saves manifest.xml" -ForegroundColor Yellow
Write-Host "    2. Microsoft 365 Admin Center -> Integrated apps -> Deploy Add-in" -ForegroundColor Yellow
Write-Host "       (or Update Add-in if a previous deployment exists)" -ForegroundColor Yellow
Write-Host "       Upload manifest.xml" -ForegroundColor Yellow
Write-Host "       This step requires a Global Administrator." -ForegroundColor Yellow
Write-Host ""
Write-Host "===========================================================================" -ForegroundColor Magenta
