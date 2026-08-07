# Deployment — drive-writeback

Checklist of `az` commands to provision this server, kept as a runbook rather than IaC — same rationale as `outlook-writeback/DEPLOYMENT.md`: proportionate for a single-resource-group, single-user deployment, worth converting to Bicep/Terraform only if this pattern gets repeated many times over.

**This is meant to work as a full disaster-recovery runbook.** Angle-bracket values (`<tenant-id>`, `<client-id>`, `<sub-id>`, etc.) are placeholders for values specific to your own deployment — substitute your own tenant/subscription IDs as you go. `<server>` stands for this server's folder name (`drive-writeback`), written generically the same way `outlook-writeback/DEPLOYMENT.md` is, so this doc reads correctly whether you're standing this up in the original deployment tenant or your own tenant on a fresh OSS deployment.

**Phases 1 and 2 are implemented and deployed.** The resource group, storage account, Key Vault, and Function App below all exist in Azure, and boundary-7b (Easy Auth + the "Drive Writeback MCP Connector" Entra app, serving Claude Desktop and claude.ai) is live too — see the "Multi-client OAuth via Easy Auth + Entra ID" section below. This doc remains the disaster-recovery runbook: if any of it were deleted, following every section below top to bottom reproduces it.

**This doc is the single source of truth for provisioning/rebuilding this server** — same posture as `outlook-writeback/DEPLOYMENT.md`, no separate wrapper script. (An earlier revision of this doc pointed at `scripts/Register-EntraApp.ps1` for repeatable execution; that script had a real bug — see the admin-consent note in step 5 below — and rather than fix a second copy of this logic to maintain in lockstep with the doc, it's been removed. Hand-typing six `az` commands on the rare occasion this needs re-running is cheaper than keeping a script in sync.)

## Prerequisites

- `az login` as an account with access to the target Azure subscription and tenant. This is a **delegated, interactive** session — nothing in this doc stores a credential for itself.
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli) with the `az ad` command group (no separate extension needed for app registration).

## Register the boundary-7a Entra app ("Drive Writeback MCP")

**Do this first** — every step after this one refers to "the app" as though it already exists. This is boundary 7a (MCP server → Microsoft Graph), unrelated to any future "Drive Writeback MCP Connector" app for boundary 7b (Claude → MCP server), which doesn't get registered until multi-client OAuth is actually being built (see `outlook-writeback/DEPLOYMENT.md`'s "Multi-client OAuth" section for what that phase looks like when it arrives here).

1. Create the app, single-tenant, following `REPO-CONVENTIONS.md` §5's `"<Server Name> MCP"` naming pattern:
   ```
   az ad app create --display-name "Drive Writeback MCP" --sign-in-audience AzureADMyOrg --output json
   ```
   Note the returned `appId` — this is `<client-id>` for the rest of this doc.
2. Create a service principal for it in this tenant (an app registration alone can't be signed into or consented to until a service principal exists locally):
   ```
   az ad sp create --id <client-id>
   ```
3. Resolve the two delegated Graph permission IDs this server needs — `Files.ReadWrite.All` and `Sites.Read.All` (PRD §6; no other scopes belong here). **Resolved dynamically here, not hardcoded** — unlike `outlook-writeback`'s doc, these two GUIDs weren't independently verified against a live tenant while writing this doc, and a wrong hardcoded permission GUID is a worse failure mode (silently requesting the wrong scope) than one extra `az rest` call:
   PowerShell:
   ```powershell
   az rest --method GET `
     --url "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=oauth2PermissionScopes" `
     --query "value[0].oauth2PermissionScopes[?value=='Files.ReadWrite.All' || value=='Sites.Read.All'].{value:value,id:id}"
   ```
   bash/Git Bash:
   ```bash
   az rest --method GET \
     --url "https://graph.microsoft.com/v1.0/servicePrincipals?\$filter=appId eq '00000003-0000-0000-c000-000000000000'&\$select=oauth2PermissionScopes" \
     --query "value[0].oauth2PermissionScopes[?value=='Files.ReadWrite.All' || value=='Sites.Read.All'].{value:value,id:id}"
   ```
   (`00000003-0000-0000-c000-000000000000` is Microsoft Graph's well-known resource `appId` — the same in every tenant. The `` `$ ``/`\$` escaping before `filter`/`select` stops each shell from treating them as variable interpolation — PowerShell uses a backtick escape, bash/Git Bash uses a backslash; dropping it silently produces a broken query string rather than an error.) Note the two returned `id` values as `<files-readwrite-all-id>` and `<sites-read-all-id>`.
4. Add the two permissions:

   PowerShell:
   ```powershell
   az ad app permission add --id <client-id> --api 00000003-0000-0000-c000-000000000000 `
     --api-permissions <files-readwrite-all-id>=Scope <sites-read-all-id>=Scope
   ```
   bash/Git Bash:
   ```bash
   az ad app permission add --id <client-id> --api 00000003-0000-0000-c000-000000000000 \
     --api-permissions <files-readwrite-all-id>=Scope <sites-read-all-id>=Scope
   ```
   **Note:** if you create the registration through the Azure Portal instead of the CLI, it adds Graph's `User.Read` by default — not a scope this server needs. Remove it, or use the CLI path above and never acquire it (same note as `outlook-writeback/DEPLOYMENT.md`).
5. Grant admin consent (requires Global Administrator or Privileged Role Administrator on the target tenant):
   ```
   az ad app permission admin-consent --id <client-id>
   ```
   **Confirm it actually took** — this call is silent on success and can appear to succeed while leaving consent unresolved if the caller isn't privileged enough:
   ```
   az rest --method GET --url "https://graph.microsoft.com/v1.0/servicePrincipals/<sp-object-id>/oauth2PermissionGrants"
   ```
   Expect one grant with `"consentType": "AllPrincipals"` and `"scope"` containing both `Files.ReadWrite.All` and `Sites.Read.All` (plus `User.Read` if step 4's note above wasn't followed).

   **Known issue, observed live during Phase 0 testing:** the GET immediately after `admin-consent` can come back empty even though the caller *is* a Global Administrator and the grant genuinely succeeded — this is eventual-consistency lag between `admin-consent` writing the grant and it becoming visible to a `GET`, not a permissions problem. If the GET comes back empty right after granting, wait ~10-30 seconds and re-run the GET before concluding the caller isn't privileged enough. This was the root cause the one time an earlier `Register-EntraApp.ps1` wrapper script threw "Admin consent did not take" here — the grant was already correct, the script's check just ran with no retry.
6. Register the public-client redirect URIs. Two are needed as of Phase 1, and they are not interchangeable: the `Category=E2E` test tier's `InteractiveBrowserCredential` needs bare `http://localhost`, while the new `DriveWriteback.Bootstrap` console tool (one-time refresh-token seeding, below) runs its own loopback listener on a fixed port — `RedirectPort` in `DriveWriteback.Bootstrap/Program.cs`, `8500` by default (a different port than `outlook-writeback`'s `8400`, so both bootstrap tools can coexist in a dev environment) — and needs that exact `http://localhost:8500/` form registered:
   ```
   az ad app update --id <client-id> --public-client-redirect-uris "http://localhost" "http://localhost:8500/"
   ```
   If you change `RedirectPort`, change the registered URI to match. No `web.redirectUris` or client secret is needed for this app — it's used exclusively as a public client via delegated auth-code + PKCE.

At this point you have everything the Phase 0 E2E spike tests need:
```
DRIVE_WRITEBACK_TENANT_ID=<tenant-id>
DRIVE_WRITEBACK_CLIENT_ID=<client-id>
```
Run `dotnet test --filter Category=E2E` from the repo root with those two set — it'll open a real browser sign-in the first time.

---

**Everything below this line has been executed against the real deployment** (resource group, storage account, Key Vault, Function App, boundary-7a app, boundary-7b Connector app, Easy Auth). Kept as runbook text for disaster recovery — re-run a section only if its resource needs to be rebuilt from scratch.

## Resources provisioned (in order)

Mirrors `outlook-writeback/DEPLOYMENT.md`'s own "Resources provisioned" section exactly — same gotchas, same workarounds, same reasoning — with names and Graph scopes swapped for this server. Read that section for the full detail (globally-unique-name warning, the `az role assignment create` `MissingSubscription` workaround, the PowerShell `--%` app-settings gotcha); only the differences are called out below.

1. Resource group: `az group create -n <server>-rg -l <region> --tags project=sapidus-writeback-mcp`.
2. Storage account (Flex Consumption's required host storage): `az storage account create -g <server>-rg -n <server>sa -l <region> --sku Standard_LRS --tags project=sapidus-writeback-mcp`.
3. Function App on Flex Consumption, .NET 10 isolated worker, capped at max instance count 1 (same single-user-refresh-token-races-itself reasoning as `outlook-writeback` — see `DriveWriteback.Graph/Auth/SilentGraphCredential.cs`):

   PowerShell:
   ```powershell
   az functionapp create `
     --resource-group <server>-rg `
     --name <server>-func `
     --storage-account <server>sa `
     --flexconsumption-location <region> `
     --runtime dotnet-isolated `
     --runtime-version 10.0 `
     --maximum-instance-count 1 `
     --assign-identity "[system]" `
     --tags project=sapidus-writeback-mcp
   ```
   bash/Git Bash:
   ```bash
   az functionapp create \
     --resource-group <server>-rg \
     --name <server>-func \
     --storage-account <server>sa \
     --flexconsumption-location <region> \
     --runtime dotnet-isolated \
     --runtime-version 10.0 \
     --maximum-instance-count 1 \
     --assign-identity "[system]" \
     --tags project=sapidus-writeback-mcp
   ```
4. Key Vault, RBAC-mode: `az keyvault create -g <server>-rg -n <server>-kv -l <region> --enable-rbac-authorization true --tags project=sapidus-writeback-mcp`.
5. Grant the Function App's system-assigned managed identity **Key Vault Secrets Officer** on the vault (read *and* write — the app rotates the stored refresh token in place). If `az role assignment create` fails with `MissingSubscription`, use the `az rest` ARM-REST workaround documented in `outlook-writeback/DEPLOYMENT.md` step 5.
6. Grant your own Entra user the same **Key Vault Secrets Officer** role, scoped for the one-time bootstrap write below.
7. **Confirmation-signing-key secret**, for `delete_item`'s two-step confirmation token (Phase 2) — same shape as `outlook-writeback`'s `delete-confirmation-signing-key`, generated fresh, not reused across servers.

   PowerShell — generate the key with .NET's crypto RNG rather than relying on `openssl` being installed:
   ```powershell
   $signingKey = [Convert]::ToBase64String([System.Security.Cryptography.RandomNumberGenerator]::GetBytes(32))
   az keyvault secret set --vault-name <server>-kv --name delete-confirmation-signing-key --value $signingKey
   ```
   bash/Git Bash:
   ```bash
   az keyvault secret set --vault-name <server>-kv --name delete-confirmation-signing-key \
     --value "$(openssl rand -base64 32)"
   ```
8. Function App application settings.

   **PowerShell gotcha, and it fails silently** — same one `outlook-writeback/DEPLOYMENT.md` step 8 documents in full: a plain backtick-continued `az functionapp config appsettings set` command silently truncates the trailing `)` off a Key Vault-reference app-setting value, leaving something `Convert.FromBase64String` can't parse and crashing the isolated worker at startup before any function registers. Now that this server has one Key Vault-reference setting too (`DRIVE_WRITEBACK_CONFIRMATION_SIGNING_KEY`), the same `--%` stop-parsing workaround applies — **everything from `--%` onward must be on a single physical line**, no backtick continuations after it:

   PowerShell:
   ```powershell
   az functionapp config appsettings set `
     --resource-group <server>-rg `
     --name <server>-func `
     --settings --% DRIVE_WRITEBACK_TENANT_ID=<tenant id> DRIVE_WRITEBACK_CLIENT_ID=<"Drive Writeback MCP" app registration's client id> DRIVE_WRITEBACK_KEY_VAULT_URI=https://<server>-kv.vault.azure.net/ DRIVE_WRITEBACK_DRY_RUN=true DRIVE_WRITEBACK_MAX_CONTENT_BYTES=1048576 DRIVE_WRITEBACK_CONFIRMATION_SIGNING_KEY="@Microsoft.KeyVault(SecretUri=https://<server>-kv.vault.azure.net/secrets/delete-confirmation-signing-key/)"
   ```
   After running it, verify the value came through intact — the failure mode is silent, not an error PowerShell surfaces:
   ```powershell
   az functionapp config appsettings list --resource-group <server>-rg --name <server>-func --query "[?name=='DRIVE_WRITEBACK_CONFIRMATION_SIGNING_KEY'].value" -o tsv
   ```
   The value should end in `secrets/delete-confirmation-signing-key/)` — closing paren included.

   bash/Git Bash:
   ```bash
   az functionapp config appsettings set \
     --resource-group <server>-rg \
     --name <server>-func \
     --settings \
       DRIVE_WRITEBACK_TENANT_ID=<tenant id> \
       DRIVE_WRITEBACK_CLIENT_ID=<"Drive Writeback MCP" app registration's client id> \
       DRIVE_WRITEBACK_KEY_VAULT_URI=https://<server>-kv.vault.azure.net/ \
       DRIVE_WRITEBACK_DRY_RUN=true \
       DRIVE_WRITEBACK_MAX_CONTENT_BYTES=1048576 \
       DRIVE_WRITEBACK_CONFIRMATION_SIGNING_KEY="@Microsoft.KeyVault(SecretUri=https://<server>-kv.vault.azure.net/secrets/delete-confirmation-signing-key/)"
   ```
   The truncation bug above is specific to PowerShell's backtick continuation, but it costs nothing to verify here too — same check, works identically in either shell:
   ```bash
   az functionapp config appsettings list --resource-group <server>-rg --name <server>-func --query "[?name=='DRIVE_WRITEBACK_CONFIRMATION_SIGNING_KEY'].value" -o tsv
   ```
   The value should end in `secrets/delete-confirmation-signing-key/)` — closing paren included.

   Leave `DRIVE_WRITEBACK_DRY_RUN=true` until you've confirmed writes behave as expected against the live tenant, then flip it to `false` (no code change needed — see `Program.cs`). **On the current production deployment this has already been done** — `DRIVE_WRITEBACK_DRY_RUN` is `false`, real writes confirmed working end to end from Claude Desktop and claude.ai (see `CLAUDE.md`'s Status section); this step's `true` starting value only matters when standing the server up fresh. Aside from the confirmation-signing-key reference, no other Key Vault-reference settings are used — the refresh token is read/written live via the Key Vault Secrets SDK (see `Auth/KeyVaultRefreshTokenStore.cs`). This also means the Function App's managed identity needs **Key Vault Secrets User** (read-only is enough here) in addition to the Secrets Officer grant from step 5, if not already covered by it.
9. Application Insights — created and linked automatically by `az functionapp create`; no separate step needed.

## One-time bootstrap (seed the initial refresh token)

Run once, locally, after the resources above exist:

PowerShell:
```powershell
cd drive-writeback
$env:DRIVE_WRITEBACK_TENANT_ID = "<tenant id>"
$env:DRIVE_WRITEBACK_CLIENT_ID = "<client id>"
$env:DRIVE_WRITEBACK_KEY_VAULT_URI = "https://<server>-kv.vault.azure.net/"
dotnet run --project DriveWriteback.Bootstrap
```
bash/Git Bash:
```bash
cd drive-writeback
DRIVE_WRITEBACK_TENANT_ID=<tenant id> \
DRIVE_WRITEBACK_CLIENT_ID=<client id> \
DRIVE_WRITEBACK_KEY_VAULT_URI=https://<server>-kv.vault.azure.net/ \
dotnet run --project DriveWriteback.Bootstrap
```
Opens a browser for a one-time sign-in against the "Drive Writeback MCP" Entra app registered above, then writes the resulting refresh token directly to Key Vault. Re-run only as a recovery step if the stored token ever goes bad — see `outlook-writeback/DEPLOYMENT.md`'s bootstrap section for the reasoning (same 90-day sliding-window behavior, same `SilentGraphCredential` mechanics).

## Deploy

```
cd drive-writeback
func azure functionapp publish <server>-func --dotnet-isolated
```
Manual, not CI/CD, same rationale as `outlook-writeback`. The `--dotnet-isolated` flag is required — `func` can't otherwise determine the project language when multiple sibling `.csproj` files share this directory.

## Local smoke test (before touching Azure at all)

```
cd drive-writeback
cp local.settings.json.example local.settings.json
func start
```
Copy the example first — `local.settings.json` is gitignored and won't exist otherwise, and `func start` fails immediately without one. Point the copy's `AzureWebJobsStorage` at the real storage account (or run Azurite) and set `DRIVE_WRITEBACK_TENANT_ID`/`CLIENT_ID`/`KEY_VAULT_URI`/`DRY_RUN`/`MAX_CONTENT_BYTES`/`CONFIRMATION_SIGNING_KEY` to exercise the real Key Vault + Entra app from a local run, via your own `az login` session (`DefaultAzureCredential` picks it up automatically). `CONFIRMATION_SIGNING_KEY` doesn't have to be the same value as the deployed secret for a local smoke test — any base64 string works, e.g. `openssl rand -base64 32` (bash) or `[Convert]::ToBase64String([System.Security.Cryptography.RandomNumberGenerator]::GetBytes(32))` (PowerShell). Leave `DRIVE_WRITEBACK_DRY_RUN` at `true` for this first run — confirm the write tools resolve and validate correctly before ever flipping it to `false` against real data.

**This is now the standard dev-loop path, not just a first-run check.** This server has no CLI registration against the deployed Function App (see "Multi-client OAuth" below for why) — iterate locally against `func start` + Azurite/the real storage account instead of round-tripping through a deployed CLI connection.

## Multi-client OAuth via Easy Auth + Entra ID

This is the boundary-7b auth path for this server, serving Claude Desktop and claude.ai. Unlike `outlook-writeback`, this server has **no CLI registration** — Claude Code CLI is not wired up against the deployed Function App at all; local iteration goes through `func start` + Azurite instead (see "Local smoke test" above). If a future need arises to add the CLI back, mirror `outlook-writeback/DEPLOYMENT.md`'s CLI section against the same Connector app below (its `publicClient.redirectUris` are simply unset here today).

### `host.json`: Easy Auth is the sole gate

`host.json` sets `extensions.mcp.system.webhookAuthorizationLevel: "Anonymous"` — required once Easy Auth requires authentication, since the MCP extension's own separate `mcp_extension` system-key check would otherwise 403 a valid Easy-Auth-validated bearer token underneath it. Only the `host.json` literal has any effect (the app-setting equivalent doesn't, per `outlook-writeback/DEPLOYMENT.md`'s note on this same gotcha) — republish via `func azure functionapp publish` after any change here. Once this setting is live, the system key stops providing any real protection on its own; Easy Auth is the only enforcing layer.

### Boundary-7b Entra app: "Drive Writeback MCP Connector"

Separate from "Drive Writeback MCP" (boundary 7a, Graph delegation, registered above) — this one represents the MCP server's own audience for Claude-side OAuth:
```
az ad app create --display-name "Drive Writeback MCP Connector" --sign-in-audience AzureADMyOrg --output json
```
Then, against the returned object id, via `az rest --method PATCH` against `https://graph.microsoft.com/v1.0/applications/<object-id>` — **each of the four calls below needs `--headers "Content-Type=application/json"` alongside `--body`**, or Graph rejects it with `BadRequest: Write requests (excluding DELETE) must contain the Content-Type header declaration` even though the body is valid JSON. On PowerShell, the header alone isn't enough — `az.cmd` (a batch file, reparsed by `cmd.exe`) mangles a quoted JSON argument, so every PowerShell form below writes the body to a temp file and passes `--body @<file>` instead (the same confirmed-working pattern the "Custom domain" section below relies on). bash/Git Bash doesn't share that quoting failure mode, so an inline `--body '...'` is fine there.

**1. `requestedAccessTokenVersion` — must land first, before `identifierUris`**, or Entra rejects a same-tenant HTTPS App ID URI with `InvalidUniqueTenantIdentifierAsPerAppPolicy`:

PowerShell:
```powershell
$body = '{"api": {"requestedAccessTokenVersion": 2}}'
$bodyFile = New-TemporaryFile
Set-Content -Path $bodyFile -Value $body -NoNewline -Encoding utf8NoBOM

az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/<connector-app-object-id>" --headers "Content-Type=application/json" --body "@$bodyFile"
```
bash/Git Bash:
```bash
az rest --method PATCH \
  --url "https://graph.microsoft.com/v1.0/applications/<connector-app-object-id>" \
  --headers "Content-Type=application/json" \
  --body '{"api": {"requestedAccessTokenVersion": 2}}'
```

**2. `identifierUris`** — both `https://<server>-func.azurewebsites.net` and `.../runtime/webhooks/mcp` — the MCP extension resource-checks against the full route, not just the host:

PowerShell:
```powershell
$body = '{"identifierUris":["https://<server>-func.azurewebsites.net","https://<server>-func.azurewebsites.net/runtime/webhooks/mcp"]}'
$bodyFile = New-TemporaryFile
Set-Content -Path $bodyFile -Value $body -NoNewline -Encoding utf8NoBOM

az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/<connector-app-object-id>" --headers "Content-Type=application/json" --body "@$bodyFile"
```
bash/Git Bash:
```bash
az rest --method PATCH \
  --url "https://graph.microsoft.com/v1.0/applications/<connector-app-object-id>" \
  --headers "Content-Type=application/json" \
  --body '{"identifierUris":["https://<server>-func.azurewebsites.net","https://<server>-func.azurewebsites.net/runtime/webhooks/mcp"]}'
```

**3. An "Expose an API" scope.** Needs a fresh GUID for the scope `id` — generate one however's convenient (PowerShell's `[guid]::NewGuid()`; on bash/Git Bash, `python3 -c "import uuid; print(uuid.uuid4())"` if Python's available, or any other UUID generator):

PowerShell:
```powershell
$scopeId = [guid]::NewGuid()
$body = "{`"api`": {`"oauth2PermissionScopes`": [{`"type`": `"User`", `"value`": `"access_as_user`", `"isEnabled`": true, `"id`": `"$scopeId`", `"adminConsentDisplayName`": `"Access <server> MCP as the signed-in user`", `"adminConsentDescription`": `"Allows the app to call <server> MCP tools on behalf of the signed-in user.`", `"userConsentDisplayName`": `"Access <server> MCP as you`", `"userConsentDescription`": `"Allows the app to call <server> MCP tools on your behalf.`"}]}}"
$bodyFile = New-TemporaryFile
Set-Content -Path $bodyFile -Value $body -NoNewline -Encoding utf8NoBOM

az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/<connector-app-object-id>" --headers "Content-Type=application/json" --body "@$bodyFile"
```
bash/Git Bash:
```bash
scope_id=$(python3 -c "import uuid; print(uuid.uuid4())")
az rest --method PATCH \
  --url "https://graph.microsoft.com/v1.0/applications/<connector-app-object-id>" \
  --headers "Content-Type=application/json" \
  --body "{\"api\": {\"oauth2PermissionScopes\": [{\"type\": \"User\", \"value\": \"access_as_user\", \"isEnabled\": true, \"id\": \"$scope_id\", \"adminConsentDisplayName\": \"Access <server> MCP as the signed-in user\", \"adminConsentDescription\": \"Allows the app to call <server> MCP tools on behalf of the signed-in user.\", \"userConsentDisplayName\": \"Access <server> MCP as you\", \"userConsentDescription\": \"Allows the app to call <server> MCP tools on your behalf.\"}]}}"
```

**4. `web.redirectUris`** — required for both Desktop/Cowork and claude.ai's connector UI, which both present a client secret at token exchange (confidential client, hence `web`, not `publicClient`):

PowerShell:
```powershell
$body = '{"web": {"redirectUris": ["https://claude.ai/api/mcp/auth_callback"]}}'
$bodyFile = New-TemporaryFile
Set-Content -Path $bodyFile -Value $body -NoNewline -Encoding utf8NoBOM

az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/<connector-app-object-id>" --headers "Content-Type=application/json" --body "@$bodyFile"
```
bash/Git Bash:
```bash
az rest --method PATCH \
  --url "https://graph.microsoft.com/v1.0/applications/<connector-app-object-id>" \
  --headers "Content-Type=application/json" \
  --body '{"web": {"redirectUris": ["https://claude.ai/api/mcp/auth_callback"]}}'
```

No `publicClient.redirectUris` entry — that's only needed if the CLI is ever wired up against this app, which it currently isn't.

### Client secret

```
az ad app credential reset --id <connector-app-id> --append --display-name "<server>-connector-secret" --years 1 --output json
```
Store the resulting `password` as the `<server>-connector-secret` Key Vault secret, referenced from a `MICROSOFT_PROVIDER_AUTHENTICATION_SECRET` app setting via `@Microsoft.KeyVault(SecretUri=...)` — same pattern as `DRIVE_WRITEBACK_CONFIRMATION_SIGNING_KEY`. This secret is for Easy Auth's own server-side token exchange (and for pasting into Claude Desktop's/claude.ai's connector setup, which need it directly since Entra has no dynamic client registration) — it is not a CLI secret, since no CLI is registered against this app.

Also set `WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES=https://<server>-func.azurewebsites.net/runtime/webhooks/mcp/access_as_user` — without it, the `401` response's `WWW-Authenticate` header comes back bare (no `scope`, no `resource_metadata`), breaking RFC 9728 discovery for any client relying on that header rather than probing the PRM endpoint directly.

**If you're also setting up a custom domain, read the "Custom domain" section below first** and register `identifierUris`/`WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` against the final hostname once, here, rather than the `azurewebsites.net` one — that avoids the swap dance that section otherwise walks through.

### `authsettingsV2` via ARM REST PUT

`az webapp auth update` fails on an app already on auth v1 (`Cannot use auth v2 commands when the app is using auth v1`) — go straight to the ARM REST PUT:

PowerShell:
```powershell
az rest --method PUT `
  --url "https://management.azure.com/subscriptions/<sub-id>/resourceGroups/<server>-rg/providers/Microsoft.Web/sites/<server>-func/config/authsettingsV2?api-version=2022-03-01" `
  --body '{
    "properties": {
      "platform": { "enabled": true },
      "globalValidation": { "requireAuthentication": true, "unauthenticatedClientAction": "Return401" },
      "identityProviders": {
        "azureActiveDirectory": {
          "enabled": true,
          "registration": {
            "clientId": "<connector-app-id>",
            "clientSecretSettingName": "MICROSOFT_PROVIDER_AUTHENTICATION_SECRET",
            "openIdIssuer": "https://login.microsoftonline.com/<tenant-id>/v2.0"
          },
          "validation": { "defaultAuthorizationPolicy": { "allowedPrincipals": {} }, "jwtClaimChecks": {} }
        },
        "facebook": {"enabled": false}, "gitHub": {"enabled": false}, "google": {"enabled": false},
        "legacyMicrosoftAccount": {"enabled": false}, "twitter": {"enabled": false}, "apple": {"enabled": false}
      }
    }
  }'
```
bash/Git Bash:
```bash
az rest --method PUT \
  --url "https://management.azure.com/subscriptions/<sub-id>/resourceGroups/<server>-rg/providers/Microsoft.Web/sites/<server>-func/config/authsettingsV2?api-version=2022-03-01" \
  --body '{
    "properties": {
      "platform": { "enabled": true },
      "globalValidation": { "requireAuthentication": true, "unauthenticatedClientAction": "Return401" },
      "identityProviders": {
        "azureActiveDirectory": {
          "enabled": true,
          "registration": {
            "clientId": "<connector-app-id>",
            "clientSecretSettingName": "MICROSOFT_PROVIDER_AUTHENTICATION_SECRET",
            "openIdIssuer": "https://login.microsoftonline.com/<tenant-id>/v2.0"
          },
          "validation": { "defaultAuthorizationPolicy": { "allowedPrincipals": {} }, "jwtClaimChecks": {} }
        },
        "facebook": {"enabled": false}, "gitHub": {"enabled": false}, "google": {"enabled": false},
        "legacyMicrosoftAccount": {"enabled": false}, "twitter": {"enabled": false}, "apple": {"enabled": false}
      }
    }
  }'
```
Leave `allowedAudiences` and `defaultAuthorizationPolicy.allowedApplications` out deliberately — they don't fix the bare-`WWW-Authenticate` symptom (that's `WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES`'s job) and add nothing here. Explicitly disabling the other built-in identity providers does matter — it stops the ARM API silently defaulting them into an unconfigured-but-enabled state.

Restart after applying either of the above: `az functionapp restart --resource-group <server>-rg --name <server>-func`.

### PRM endpoint — no custom code needed

`GET /.well-known/oauth-protected-resource/runtime/webhooks/mcp` is served natively by Easy Auth v2 once the above is live:
```json
{"resource":"https://<server>-func.azurewebsites.net/runtime/webhooks/mcp","authorization_servers":["https://login.microsoftonline.com/<tenant-id>/v2.0"],"scopes_supported":["https://<server>-func.azurewebsites.net/runtime/webhooks/mcp/access_as_user"]}
```

### Validation checks

1. No `Authorization` header, correct `Accept: application/json, text/event-stream` → `401` plus `WWW-Authenticate: Bearer ... resource_metadata="https://<server>-func.azurewebsites.net/.well-known/oauth-protected-resource/runtime/webhooks/mcp"`.
2. `GET /.well-known/oauth-protected-resource/runtime/webhooks/mcp` → the PRM document above.
3. Add the connector in Claude Desktop / claude.ai with the server URL, Connector app's `appId` as Client ID, and the secret from above as Client Secret — confirm a real Entra sign-in/consent redirect completes and the connector shows "Connected." A `get_item` or dry-run write tool call end to end is the most reliable confirmation of a valid bearer token reaching the MCP extension — a raw `curl` can't easily manufacture one.

## Custom domain

Same phase `outlook-writeback` already went through — see `outlook-writeback/DEPLOYMENT.md`'s own "Custom domain" section for the original writeup this one mirrors. **No application source code change is required for any of this** — `drive-writeback`'s C# source has no hostname string literals anywhere (confirmed by grepping `Program.cs`, `DriveWriteback.Bootstrap/Program.cs`, `host.json`, and everything under `DriveWriteback.Graph/` and `shared/Sapidus.Writeback.Shared/`); the MCP extension's OAuth resource/audience/issuer/PRM-endpoint/`WWW-Authenticate` behavior is entirely handled by Azure's Easy Auth v2 platform, configured via the app settings and Entra app fields below, not by any code in this repo. This section is Azure/Entra config only.

**Naming convention:** each server in this monorepo is already fully isolated per `../REPO-CONVENTIONS.md` §3 — own resource group, own Function App, own Entra app. A custom domain doesn't change that: the convention is a **flat subdomain per server**, mirroring the `<server>-rg`/`<server>-func` pattern — e.g. `drive-writeback.example.com`, never nested under another server's domain.

**Expect a cutover, not a coexistence.** Once the custom domain is working end to end (DNS, managed certificate, Easy Auth discovery, and a fresh OAuth authorize), `<server>-func.azurewebsites.net` stops working for OAuth-authenticated MCP clients — it still resolves and serves traffic, but a fresh authorize against it hits the same class of error this work is fixing, mirrored. Plan to migrate every connected client in one pass: for this server that's just Claude Desktop and claude.ai (no CLI registration exists against this Function App — see "Multi-client OAuth" above).

**DNS + managed certificate (Azure portal):** add a CNAME (`<subdomain>` → `<server>-func.azurewebsites.net`) plus an `asuid.<subdomain>` TXT record (the domain-verification ID, from the "Add custom domain" dialog) at your DNS provider, then in the portal: `<server>-func` → **Settings** → **Custom domains** → **Add custom domain** → domain provider **All other domain services** (or your provider if listed) → enter the hostname → **TLS/SSL certificate: App Service Managed Certificate**, **SNI SSL** → **Validate** once both records show green → **Add**. Wait up to ~10 minutes for the managed cert to bind. No Azure CLI support for Flex Consumption's site-scoped certificate model as of the docs current at research time (2026-05-18) — this has to go through the portal or an ARM/Bicep template.

**Reconciling with Easy Auth — swap, not add, in the same pass.** The "Drive Writeback MCP Connector" app is self-referencing the same way `outlook-writeback`'s is, even without a CLI registration: Claude Desktop/claude.ai's confidential-client flow (`web.redirectUris` + client secret) still requests a token whose audience is the Connector app's own `identifierUri`. Entra only resolves that unambiguously when there's exactly one candidate `identifierUri` per requested scope, which rules out the two obvious approaches before the working one:

1. **Wrong — add the new hostname's `identifierUris` alongside the old ones:** breaks the already-working connection with `AADSTS90009: Application '<app-id>' is requesting a token for itself. This scenario is supported only if resource is specified using the GUID based Application ID URI.`
2. **Wrong — leave `identifierUris` and `WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` alone entirely:** a `401` against the new hostname still advertises the old hostname's scope (`WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` is a single static value, not derived per-request — only the PRM document's `resource` field is dynamic, from the request's Host header). A cached-token reconnect can work anyway, masking the problem; a later fresh authorize fails with `OAuth error: invalid_target - AADSTS9010010: The resource parameter provided in the request doesn't match with the requested scopes`.
3. **Right — swap, not add, in the same pass:**

   PowerShell — **write the body to a file and pass `--body @file` rather than an inline quoted string.** Both `--headers "Content-Type=application/json"` alone (Graph needs it — see below) and a plain multi-line `'...'` string spliced across backtick-continued lines (fragile in this exact way — the console can get stuck waiting on an unterminated line, and even when it does submit, Graph reports it can't parse the JSON) were tried and failed live against this tenant; the file-based form below is the one that actually worked:
   ```powershell
   $body = '{"identifierUris":["https://<subdomain>.example.com","https://<subdomain>.example.com/runtime/webhooks/mcp"]}'
   $bodyFile = New-TemporaryFile
   Set-Content -Path $bodyFile -Value $body -NoNewline -Encoding utf8NoBOM

   az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/<connector-app-object-id>" --headers "Content-Type=application/json" --body "@$bodyFile"

   az functionapp config appsettings set --name <server>-func --resource-group <server>-rg --settings WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES=https://<subdomain>.example.com/runtime/webhooks/mcp/access_as_user
   ```
   bash/Git Bash — untested against a live tenant during this pass, but kept as the multi-line form since bash doesn't share PowerShell's `az.cmd`/console quoting failure modes documented above:
   ```bash
   az rest --method PATCH \
     --url "https://graph.microsoft.com/v1.0/applications/<connector-app-object-id>" \
     --headers "Content-Type=application/json" \
     --body '{"identifierUris":[
       "https://<subdomain>.example.com",
       "https://<subdomain>.example.com/runtime/webhooks/mcp"
     ]}'

   az functionapp config appsettings set \
     --name <server>-func \
     --resource-group <server>-rg \
     --settings WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES=https://<subdomain>.example.com/runtime/webhooks/mcp/access_as_user
   ```
   **`--headers "Content-Type=application/json"` is required** — Microsoft Graph rejects a PATCH/PUT with a `--body` but no explicit `Content-Type` (`BadRequest: Write requests (excluding DELETE) must contain the Content-Type header declaration`), even though `az rest` sends valid JSON. ARM (`management.azure.com`) calls elsewhere in this doc haven't been observed to need this, only Graph (`graph.microsoft.com`) ones. On PowerShell, that header alone isn't sufficient either — confirmed live: adding just the header still failed with `BadRequest: Unable to read JSON request payload`, tracing back to how `az.cmd` (a batch file, reparsed by `cmd.exe`) handles a multi-line quoted argument. Writing the JSON to a file first and passing `--body @<file>` sidesteps the whole quoting chain and is the form confirmed working.

   Both changes have to land together — updating only `identifierUris` would leave the advertised scope pointing at a URI the app no longer registers, breaking auth for everyone until the second command lands. There's no way to make both hostnames work for OAuth at once on this self-referencing app; true dual-hostname support would need a non-self-referencing app split (separate app for the API vs. the client), not attempted here.

**How to confirm it worked:** an unauthenticated request against the new hostname should return `401` with `WWW-Authenticate`'s `scope` and the PRM document's `resource`/`scopes_supported` all reading the new hostname — they must agree. Then run a fresh (not reconnect) authorize on both client surfaces this server has: Claude Desktop (a confidential client — this needs a remove-and-re-add of the connector, an in-place URL edit isn't enough) and claude.ai's connector UI. The old hostname's `401` will now mirror the new hostname's scope back, so a fresh authorize there fails with the same `AADSTS9010010` in reverse — expected, per the cutover tradeoff above.
