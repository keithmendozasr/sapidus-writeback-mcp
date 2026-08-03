<#
.SYNOPSIS
    Registers (or reuses) the boundary-7a "Drive Writeback MCP" Entra app registration.

.DESCRIPTION
    Idempotent script version of the steps documented in ../DEPLOYMENT.md's
    "Register the boundary-7a Entra app" section — read that doc first, this
    script exists so you don't have to hand-type the same commands on a re-run,
    not as a replacement for understanding what each step does.

    OPSEC posture (see ../docs/active/PRD-drive-write.md phase-0 planning notes):
      - No secrets are created, embedded, or logged by this script. It only
        creates an Entra app registration and grants delegated Graph scopes —
        it never touches a client secret, refresh token, or Key Vault value.
      - Runs entirely under your own interactive `az login` session. It does
        not create or store a non-interactive credential for itself.
      - Every identifying value is a parameter with a generic default, not a
        hardcoded literal — safe to hand to someone deploying this to their
        own tenant.
      - Safe to re-run: checks for an existing app/permissions/redirect URI
        before creating them, the same "already exists, continue" posture
        this server's own create_folder tool will use.

.PARAMETER DisplayName
    Entra app display name. Defaults to "Drive Writeback MCP", per
    REPO-CONVENTIONS.md §5's "<Server Name> MCP" naming pattern.

.PARAMETER TenantId
    Target tenant ID. Defaults to the tenant of the current `az login` session.

.EXAMPLE
    ./Register-EntraApp.ps1
    Registers (or reuses) the app in whatever tenant `az login` is currently
    signed into.
#>

[CmdletBinding()]
param(
    [string]$DisplayName = "Drive Writeback MCP",

    [string]$TenantId
)

$ErrorActionPreference = "Stop"

$graphResourceAppId = "00000003-0000-0000-c000-000000000000"
$requiredScopes = @("Files.ReadWrite.All", "Sites.Read.All")

function Get-CurrentTenantId
{
    az account show --query tenantId -o tsv
}

function Find-ExistingApp
{
    param([string]$DisplayName)

    $apps = az ad app list --display-name $DisplayName --output json | ConvertFrom-Json

    if ($apps.Count -gt 0)
    {
        return $apps[0]
    }

    return $null
}

function New-SingleTenantApp
{
    param([string]$DisplayName)

    $appJson = az ad app create `
        --display-name $DisplayName `
        --sign-in-audience AzureADMyOrg `
        --output json

    return $appJson | ConvertFrom-Json
}

function Confirm-ServicePrincipal
{
    param([string]$AppId)

    $existing = az ad sp show --id $AppId --output json 2>$null

    if ($existing)
    {
        return $existing | ConvertFrom-Json
    }

    $spJson = az ad sp create --id $AppId --output json

    return $spJson | ConvertFrom-Json
}

function Resolve-GraphScopeIds
{
    param([string[]]$ScopeNames)

    $filter = "appId eq '$graphResourceAppId'"
    $scopesJson = az rest --method GET `
        --url "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=$filter&`$select=oauth2PermissionScopes" `
        --query "value[0].oauth2PermissionScopes" `
        --output json

    $allScopes = $scopesJson | ConvertFrom-Json
    $resolved = @{}

    foreach ($name in $ScopeNames)
    {
        $match = $allScopes | Where-Object { $_.value -eq $name }

        if (-not $match)
        {
            throw "Could not resolve Graph delegated scope '$name' — check the name against https://learn.microsoft.com/graph/permissions-reference"
        }

        $resolved[$name] = $match.id
    }

    return $resolved
}

function Add-MissingGraphPermissions
{
    param(
        [string]$AppId,
        [hashtable]$ScopeIds
    )

    $currentJson = az ad app permission list --id $AppId --output json
    $current = $currentJson | ConvertFrom-Json
    $currentGraphAccess = $current | Where-Object { $_.resourceAppId -eq $graphResourceAppId }
    $currentScopeIds = @()

    if ($currentGraphAccess)
    {
        $currentScopeIds = $currentGraphAccess.resourceAccess.id
    }

    $missing = $ScopeIds.GetEnumerator() | Where-Object { $currentScopeIds -notcontains $_.Value }

    if (-not $missing)
    {
        Write-Host "All required Graph delegated scopes already granted, skipping."
        return
    }

    $permissionArgs = $missing | ForEach-Object { "$($_.Value)=Scope" }

    az ad app permission add `
        --id $AppId `
        --api $graphResourceAppId `
        --api-permissions $permissionArgs
}

function Grant-AndVerifyAdminConsent
{
    param(
        [string]$AppId,
        [string]$ServicePrincipalObjectId,
        [hashtable]$ScopeIds
    )

    az ad app permission admin-consent --id $AppId

    $grantsJson = az rest --method GET `
        --url "https://graph.microsoft.com/v1.0/servicePrincipals/$ServicePrincipalObjectId/oauth2PermissionGrants" `
        --output json
    $grants = $grantsJson | ConvertFrom-Json
    $grantedScopeNames = @()

    foreach ($grant in $grants)
    {
        $grantedScopeNames += ($grant.scope -split " ")
    }

    $stillMissing = $ScopeIds.Keys | Where-Object { $grantedScopeNames -notcontains $_ }

    if ($stillMissing)
    {
        throw "Admin consent did not take for: $($stillMissing -join ', '). Re-run as a Global Administrator or Privileged Role Administrator on this tenant."
    }

    Write-Host "Admin consent confirmed for: $($ScopeIds.Keys -join ', ')"
}

function Set-PublicClientRedirectUri
{
    param([string]$AppId)

    az ad app update --id $AppId --public-client-redirect-uris "http://localhost"
}

if (-not $TenantId)
{
    $TenantId = Get-CurrentTenantId
}

Write-Host "Target tenant: $TenantId"

$app = Find-ExistingApp -DisplayName $DisplayName

if ($app)
{
    Write-Host "Reusing existing app '$DisplayName' (appId $($app.appId))"
}
else
{
    Write-Host "Creating app '$DisplayName'..."
    $app = New-SingleTenantApp -DisplayName $DisplayName
}

$clientId = $app.appId

$servicePrincipal = Confirm-ServicePrincipal -AppId $clientId
$scopeIds = Resolve-GraphScopeIds -ScopeNames $requiredScopes

Add-MissingGraphPermissions -AppId $clientId -ScopeIds $scopeIds
Grant-AndVerifyAdminConsent -AppId $clientId -ServicePrincipalObjectId $servicePrincipal.id -ScopeIds $scopeIds
Set-PublicClientRedirectUri -AppId $clientId

Write-Host ""
Write-Host "Done. Env vars for the Phase 0 E2E spike tests:"
Write-Host "  DRIVE_WRITEBACK_TENANT_ID=$TenantId"
Write-Host "  DRIVE_WRITEBACK_CLIENT_ID=$clientId"
