# Deployment — outlook-writeback

Checklist of `az` commands to provision and deploy this server, kept as a runbook rather than IaC — proportionate for a single-resource-group, single-user deployment. Worth converting to Bicep/Terraform if you end up running several of these.

**This is meant to work as a full disaster-recovery runbook** — if `<server>-rg`, its Entra apps, or both were deleted outright, following every section below top to bottom against a clean subscription/tenant should reproduce a working deployment with no undocumented manual steps or tribal knowledge. One section (**Wire up Claude Code CLI**) is marked superseded — skip it on a fresh build and use **Migrate Claude Code CLI to OAuth** instead; it's kept only because it explains the system-key mechanism the OAuth path replaced.

Angle-bracket values (`<tenant id>`, `<client id>`, `<sub-id>`, etc.) and `example.com` are all placeholders for values specific to your own deployment — substitute your own tenant/subscription IDs and your own domain as you go. `<server>` specifically stands for this server's folder name (`outlook-writeback`, per `../REPO-CONVENTIONS.md`'s `<server>-rg` / `<server>-func` naming table) — resource names below are written generically so this doc can be followed for any server in this monorepo, not just this one.

## Prerequisites

- `az login` as an account with access to the target Azure subscription and tenant.
- [Azure Functions Core Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local) (`npm install -g azure-functions-core-tools@4 --unsafe-perm true`).
- Confirm .NET 10 is available as a Flex Consumption runtime stack in your target region before provisioning — this changes the hosting-plan decision if it's not:
  ```
  az functionapp list-flexconsumption-runtimes --location <region> --runtime dotnet-isolated
  ```

## Register the boundary-7a Entra app ("Outlook Writeback MCP")

**Do this first.** Every step after this one refers to "the 'Outlook Writeback MCP' app" as though it already exists; on a from-scratch deployment it doesn't. The display name follows `../REPO-CONVENTIONS.md` §5's `"<Server Name> MCP"` pattern — you can name it anything, but keep it distinguishable from the first-party Microsoft and Anthropic entries already in your tenant's Enterprise Applications list.

This is boundary 7a (MCP server → Microsoft Graph) — separate from and unrelated to the "Outlook Writeback MCP Connector" app registered later, in the multi-client OAuth section, for boundary 7b (Claude → MCP server).

1. Create the app, single-tenant:
   ```
   az ad app create --display-name "Outlook Writeback MCP" --sign-in-audience AzureADMyOrg --output json
   ```
   Note the returned `appId` — this is `<client-id>` referenced throughout the rest of this doc (the `OUTLOOK_WRITEBACK_CLIENT_ID` app setting, the Bootstrap tool's env var, etc.).
2. Create a service principal for it in this tenant (an app registration alone isn't enough — nothing can consent to or sign into it until a service principal exists locally):
   ```
   az ad sp create --id <client-id>
   ```
3. Add the two Graph delegated permissions this server actually uses — `Mail.ReadWrite` and `Calendars.ReadWrite` (per the minimum-scopes requirement in `docs/archive/`; no other scopes should be added here):
   ```
   az ad app permission add --id <client-id> --api 00000003-0000-0000-c000-000000000000 \
     --api-permissions 024d486e-b451-40bb-833d-3e66d98c5c73=Scope 1ec239c2-d7c9-4623-a91a-a9775856bb36=Scope
   ```
   (`00000003-0000-0000-c000-000000000000` is Microsoft Graph's well-known resource `appId`, the same in every tenant. The two permission GUIDs are Graph's `Mail.ReadWrite` and `Calendars.ReadWrite` delegated-scope IDs respectively — also tenant-invariant; resolved via `az ad sp show --id 00000003-0000-0000-c000-000000000000 --query "oauth2PermissionScopes[?value=='Mail.ReadWrite' || value=='Calendars.ReadWrite']"` if they ever need re-confirming.)

   **Note:** if you create the registration through the Azure Portal instead of the CLI, it adds Graph's `User.Read` (`e1fe6dd8-ba31-4d61-89e7-88639da4683d`) by default. That isn't a scope this server needs — nothing in this codebase calls anything requiring it — so remove it, or use the CLI path above and never acquire it.
4. Grant admin consent (requires Global Administrator or Privileged Role Administrator on the target tenant):
   ```
   az ad app permission admin-consent --id <client-id>
   ```
   Confirm it actually took (this call is silent on success and can appear to succeed while leaving consent unresolved if the caller isn't privileged enough):
   ```
   az rest --method GET --url "https://graph.microsoft.com/v1.0/servicePrincipals/<sp-object-id>/oauth2PermissionGrants"
   ```
   Expect one grant with `"consentType": "AllPrincipals"` and `"scope": "Calendars.ReadWrite Mail.ReadWrite"` (plus `User.Read` if step 3's note above wasn't skipped).
5. Register the public-client redirect URIs. Two are needed, and they are not interchangeable: the `Category=E2E` test tier uses `InteractiveBrowserCredential`, which needs bare `http://localhost`, while the `OutlookWriteback.Bootstrap` console tool (one-time refresh-token seeding, below) runs its own loopback listener on a fixed port — `RedirectPort` in `OutlookWriteback.Bootstrap/Program.cs`, `8400` by default — and needs that exact `http://localhost:8400/` form registered. Omitting the second is what makes bootstrap fail on an otherwise correctly configured tenant:
   ```
   az ad app update --id <client-id> --public-client-redirect-uris "http://localhost" "http://localhost:8400/"
   ```
   If you change `RedirectPort`, change the registered URI to match. No `web.redirectUris` or client secret is needed for this app — it's used exclusively as a public client via delegated auth-code + PKCE, never as a confidential client, and it never runs non-interactively itself (that's what `SilentGraphCredential`'s stored refresh token is for).

## Resources provisioned (in order)

**Before you start — three of these names are globally unique.** Function App, storage account, and Key Vault names each live in a *global* Azure namespace, not one scoped to your subscription, so the bare `<server>-func` / `<server>sa` / `<server>-kv` convention defaults will often already be taken by someone else. Pick your own distinguishing suffix (a short tenant nickname, initials, a few random characters — `<server>-func-a1b2` and so on) and use it consistently everywhere below. Resource group names are only unique within a subscription, so `<server>-rg` can be used as-is.

1. Resource group:
   ```
   az group create -n <server>-rg -l <region> --tags project=sapidus-writeback-mcp
   ```
2. Storage account (Flex Consumption's required host storage, `AzureWebJobsStorage`):
   ```
   az storage account create -g <server>-rg -n <server>sa -l <region> --sku Standard_LRS --tags project=sapidus-writeback-mcp
   ```
3. Function App on Flex Consumption, .NET 10 isolated worker, capped at max instance count 1 (low, bursty single-user volume — no need for concurrency, and it avoids a cross-instance race on the Key Vault-backed refresh token; see `OutlookWriteback.Graph/Auth/SilentGraphCredential.cs`). Raise this if you expect concurrent callers, but read `SilentGraphCredential`'s doc comments first — the refresh-token rotation is what constrains it:

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
4. Key Vault, RBAC-mode (not legacy access policies):
   ```
   az keyvault create -g <server>-rg -n <server>-kv -l <region> --enable-rbac-authorization true --tags project=sapidus-writeback-mcp
   ```
5. Grant the Function App's system-assigned managed identity **Key Vault Secrets Officer** on the vault — needs read *and* write, since the app rotates the stored refresh token in place; `Key Vault Secrets User` alone isn't enough.

   **Known issue (observed with az CLI 2.88.0):** `az role assignment create` can fail with `(MissingSubscription) The request did not have a subscription or a valid tenant level resource provider`, even though `az account show`, `az group show`, and other ARM commands work fine, and `az rest` against the same ARM role-assignment REST endpoint succeeds. That signature points at an `az role assignment` command-group bug rather than a real permissions or subscription problem. If you hit it, work around it by calling the ARM REST API directly via `az rest` instead:

   PowerShell:
   ```powershell
   az rest --method put `
     --url "https://management.azure.com/subscriptions/<sub-id>/resourceGroups/<server>-rg/providers/Microsoft.KeyVault/vaults/<server>-kv/providers/Microsoft.Authorization/roleAssignments/<new-guid>?api-version=2022-04-01" `
     --body '{"properties":{"roleDefinitionId":"/subscriptions/<sub-id>/providers/Microsoft.Authorization/roleDefinitions/b86a8fe4-44ce-4948-aee5-eccb2c155cd7","principalId":"<function-app-principal-id>","principalType":"ServicePrincipal"}}'
   ```
   bash/Git Bash:
   ```bash
   az rest --method put \
     --url "https://management.azure.com/subscriptions/<sub-id>/resourceGroups/<server>-rg/providers/Microsoft.KeyVault/vaults/<server>-kv/providers/Microsoft.Authorization/roleAssignments/<new-guid>?api-version=2022-04-01" \
     --body '{"properties":{"roleDefinitionId":"/subscriptions/<sub-id>/providers/Microsoft.Authorization/roleDefinitions/b86a8fe4-44ce-4948-aee5-eccb2c155cd7","principalId":"<function-app-principal-id>","principalType":"ServicePrincipal"}}'
   ```
   (`b86a8fe4-44ce-4948-aee5-eccb2c155cd7` is the built-in Key Vault Secrets Officer role definition ID — same in every tenant.) If `az role assignment create` starts working again on a future CLI version, prefer it; this is a workaround, not the preferred path.
6. Grant your own Entra user the same **Key Vault Secrets Officer** role on the vault (same `az rest` workaround), scoped for the one-time bootstrap write in the next section.
7. Generate and store the `delete_event` confirmation-token signing key. Unlike the Graph refresh token, this key is static (no rotation, no write-back), so — unlike every other secret this project uses — it's provisioned as a plain **Key Vault-reference app setting**.

   The `\` line continuations and `$(openssl rand -base64 32)` substitution below are bash syntax (this repo's other `az` commands are written for bash/Git Bash). In PowerShell, generate the key with .NET's crypto RNG instead of relying on `openssl` being installed, then pass it as a variable:
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

   **PowerShell gotcha, and it fails silently:** a plain backtick-continued `az functionapp config appsettings set` command (one `--settings` value per line, each ending in a backtick) silently dropped the trailing `)` from the Key Vault reference value below, leaving a syntactically invalid setting that Azure couldn't resolve. `Program.cs` then fed that literal, unresolved string straight into `Convert.FromBase64String`, which isn't valid base64 either, and threw before any function registered — the isolated worker process crashed at startup, which is what a Function App surfaces in the portal as "Encountered an error (BadGateway) from host runtime" when it tries to list functions. Use the exact command shape below, not the naive one-continuation-per-line version.

   **Array splatting does NOT fix this**, despite looking like it should — confirmed by reproducing the exact same truncation with `& az @settingsArray` against a live throwaway app setting. Whatever mangles the trailing `)` happens inside `az`'s own argument handling for `functionapp config appsettings set` specifically, not in how PowerShell hands the process its argv, so no PowerShell-side workaround that changes *how the array is built* fixes it.

   The `--%` stop-parsing token is what actually works, confirmed live — but only under one strict condition: **everything from `--%` onward must be on a single physical line.** Backtick line-continuation is fine on the lines *before* `--%` (PowerShell hasn't stopped parsing yet at that point), but a backtick placed *after* `--%` is no longer treated as a continuation character — it's passed to `az` as a literal character, the command silently ends at that line, and every following line gets parsed as its own separate (invalid) command. Reformatting the one-liner below back into multiple backtick-continued lines is exactly what triggers it:
   ```powershell
   az functionapp config appsettings set `
     --resource-group <server>-rg `
     --name <server>-func `
     --settings --% OUTLOOK_WRITEBACK_TENANT_ID=<tenant id> OUTLOOK_WRITEBACK_CLIENT_ID=<"Outlook Writeback MCP" app registration's client id> OUTLOOK_WRITEBACK_KEY_VAULT_URI=https://<server>-kv.vault.azure.net/ OUTLOOK_WRITEBACK_CONFIRMATION_SIGNING_KEY="@Microsoft.KeyVault(SecretUri=https://<server>-kv.vault.azure.net/secrets/delete-confirmation-signing-key/)"
   ```
   Note the shape: backtick-continued lines for `--resource-group` and `--name` are fine, but the moment `--% ` appears, every `--settings` value has to be crammed onto that same line with no further backticks. If you need to re-run this, do not "clean it up" by re-wrapping the `--%`-and-after portion across multiple lines — that's the specific thing that breaks it.

   After running it, always verify the value came through intact — the failure mode here is silent, not an error PowerShell surfaces:
   ```powershell
   az functionapp config appsettings list --resource-group <server>-rg --name <server>-func --query "[?name=='OUTLOOK_WRITEBACK_CONFIRMATION_SIGNING_KEY'].value" -o tsv
   ```
   The value should end in `secrets/delete-confirmation-signing-key/)` — closing paren included.

   bash/Git Bash:
   ```
   az functionapp config appsettings set \
     --resource-group <server>-rg \
     --name <server>-func \
     --settings \
       OUTLOOK_WRITEBACK_TENANT_ID=<tenant id> \
       OUTLOOK_WRITEBACK_CLIENT_ID=<"Outlook Writeback MCP" app registration's client id> \
       OUTLOOK_WRITEBACK_KEY_VAULT_URI=https://<server>-kv.vault.azure.net/ \
       OUTLOOK_WRITEBACK_CONFIRMATION_SIGNING_KEY="@Microsoft.KeyVault(SecretUri=https://<server>-kv.vault.azure.net/secrets/delete-confirmation-signing-key/)"
   ```
   Aside from that one Key Vault reference, no others are used anywhere — the refresh token is read/written live via the Key Vault Secrets SDK (see `Auth/KeyVaultRefreshTokenStore.cs`), and the Claude-facing key (below) is a platform-managed system key, never stored by this project. This also means the Function App's managed identity needs **Key Vault Secrets User** (read-only is enough here) in addition to the Secrets Officer grant from step 5, if not already covered by it.
9. Application Insights — created and linked automatically by `az functionapp create`; no separate step needed.

## Hardening: budget alert and custom domain

Budget alert: set a low monthly Cost Management budget on `<server>-rg`, sized to your expected volume, via the Azure portal (Cost Management + Billing → Budgets) rather than `az consumption budget create` — that CLI command group is still preview and needs a more involved notification-contact JSON payload than a one-line call covers cleanly, so the portal is simpler for a one-time setup. Note it's alert-only: crossing the threshold sends a notification and never throttles or disables the Function App, so a tight threshold can't interrupt a legitimate call mid-month.

### Custom domain

Each server in this monorepo is already fully isolated per `../REPO-CONVENTIONS.md` §3 — own resource group, own Function App, own Entra app. A custom domain doesn't change that: each server's hostname is its own CNAME target and its own TLS certificate, nothing shared. So instead of nesting future servers under one shared front door, the convention is a **flat subdomain per server**, mirroring the existing `<server>-rg` / `<server>-func` pattern — e.g. `outlook-writeback.example.com`, and a future `onedrive-writeback` server would get `onedrive-writeback.example.com`, never nested under this one.

**Expect a cutover, not a coexistence.** Once the custom domain is working end to end (DNS, managed certificate, Easy Auth discovery, and a fresh OAuth authorize), `<server>-func.azurewebsites.net` stops working for OAuth-authenticated MCP clients — it still resolves and serves traffic, but a fresh authorize against it hits the same class of error the custom-domain work was fixing, mirrored. That tradeoff is acceptable when you have a small number of clients to repoint; if you need both hostnames to keep working, you need the non-self-referencing app split noted at the end of this section. Plan to migrate every client in one pass.

**DNS + managed certificate (Azure portal):** standard Azure custom-domain flow — add a CNAME (`<subdomain>` → `<server>-func.azurewebsites.net`) plus an `asuid.<subdomain>` TXT record (the domain-verification ID, from the "Add custom domain" dialog) at your DNS provider, then in the portal: `<server>-func` → **Settings** → **Custom domains** → **Add custom domain** → domain provider **All other domain services** (or your provider if listed) → enter the hostname → **TLS/SSL certificate: App Service Managed Certificate**, **SNI SSL** → **Validate** once both records show green → **Add**. Wait up to ~10 minutes for the managed cert to bind. No Azure CLI support for Flex Consumption's site-scoped certificate model as of the docs current at research time (2026-05-18) — this has to go through the portal or an ARM/Bicep template.

**Reconciling with Easy Auth — three attempts, not one, because of the self-referencing app.** This Entra app is self-referencing (per the multi-client OAuth section above, the same App ID is both the public-client CLI registration *and* the exposed API), and Entra only resolves that unambiguously when there's exactly one candidate `identifierUri` per requested scope. That constraint rules out the two obvious approaches before the working one:

1. **Wrong — add the new hostname's `identifierUris` alongside the old ones:** a PATCH to 4 entries (2 old forms + 2 new) broke the already-working connection with `AADSTS90009: Application '<app-id>' is requesting a token for itself. This scenario is supported only if resource is specified using the GUID based Application ID URI.` Fix at the time: revert to exactly the original 2 entries.
2. **Wrong — leave `identifierUris` and `WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` alone entirely:** with `identifierUris` back to the old hostname only, a `401` against the *new* hostname still advertised the *old* hostname's scope (`WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` is a single static value, not derived per-request — only the PRM document's `resource` field is dynamic, from the request's Host header). A cached-token reconnect worked anyway, masking the problem; a later fresh authorize failed with `OAuth error: invalid_target - AADSTS9010010: The resource parameter provided in the request doesn't match with the requested scopes` — a client sending both a `resource` (matching the new hostname) and a `scope` (matching the old one) is asking Entra for two things that disagree.
3. **Right — swap, not add, in the same pass:**

   bash/Git Bash (as originally run for this server):
   ```bash
   az rest --method PATCH \
     --url "https://graph.microsoft.com/v1.0/applications/<object-id>" \
     --headers "Content-Type=application/json" \
     --body '{"identifierUris":[
       "https://<server>.example.com",
       "https://<server>.example.com/runtime/webhooks/mcp"
     ]}'

   az functionapp config appsettings set \
     --name <server>-func \
     --resource-group <server>-rg \
     --settings WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES=https://<server>.example.com/runtime/webhooks/mcp/access_as_user
   ```
   PowerShell — **write the body to a file and pass `--body @file`, not an inline multi-line quoted string.** Confirmed live on `drive-writeback`'s identical PATCH (see that server's `DEPLOYMENT.md` "Custom domain" section): a plain `'...'` body spanning multiple backtick-continued lines either leaves the console stuck on an unterminated line, or — once it does submit — gets mangled by `az.cmd` (a batch file, reparsed by `cmd.exe`) into something Graph reports as unparseable JSON, even with the `Content-Type` header set correctly. This has not been re-run against `outlook-writeback`'s own tenant, but the same `az.cmd` mechanics apply:
   ```powershell
   $body = '{"identifierUris":["https://<server>.example.com","https://<server>.example.com/runtime/webhooks/mcp"]}'
   $bodyFile = New-TemporaryFile
   Set-Content -Path $bodyFile -Value $body -NoNewline -Encoding utf8NoBOM

   az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/<object-id>" --headers "Content-Type=application/json" --body "@$bodyFile"

   az functionapp config appsettings set --name <server>-func --resource-group <server>-rg --settings WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES=https://<server>.example.com/runtime/webhooks/mcp/access_as_user
   ```
   **`--headers "Content-Type=application/json"` is required** — Microsoft Graph rejects a PATCH/PUT that has a `--body` but no explicit `Content-Type` (`BadRequest: Write requests (excluding DELETE) must contain the Content-Type header declaration`), even though `az rest` sends valid JSON. ARM (`management.azure.com`) calls elsewhere in this doc haven't been observed to need this, only Graph (`graph.microsoft.com`) ones.

   Both changes have to land together — updating only `identifierUris` would leave the advertised scope pointing at a URI the app no longer registers, breaking auth for everyone until the second command lands. There's no way to make both hostnames work for OAuth at once on this self-referencing app. A future server needing true dual-hostname support would need a non-self-referencing app split (separate app for the API vs. the public client) instead — not attempted here.

**How to confirm it worked, across all three client surfaces:** an unauthenticated request against the new hostname should return `401` with `WWW-Authenticate`'s `scope` and the PRM document's `resource`/`scopes_supported` all reading the new hostname — they must agree. Then run a fresh (not reconnect) authorize on each client you use: Claude Code CLI, Claude Desktop (a confidential client — this needs a remove-and-re-add, an in-place URL edit isn't enough), and claude.ai's connector UI. The old hostname's `401` will now mirror the new hostname's scope back, so a fresh authorize there fails with the same `AADSTS9010010` in reverse. That's expected, per the cutover tradeoff above.

## One-time bootstrap (seed the initial refresh token)

Run once, locally, after the resources above exist:

PowerShell (inline `VAR=value dotnet run` prefixing is bash-only syntax — PowerShell has no equivalent, so set each variable first):
```powershell
cd outlook-writeback
$env:OUTLOOK_WRITEBACK_TENANT_ID = "<tenant id>"
$env:OUTLOOK_WRITEBACK_CLIENT_ID = "<client id>"
$env:OUTLOOK_WRITEBACK_KEY_VAULT_URI = "https://<server>-kv.vault.azure.net/"
dotnet run --project OutlookWriteback.Bootstrap
```
bash/Git Bash:
```bash
cd outlook-writeback
OUTLOOK_WRITEBACK_TENANT_ID=<tenant id> \
OUTLOOK_WRITEBACK_CLIENT_ID=<client id> \
OUTLOOK_WRITEBACK_KEY_VAULT_URI=https://<server>-kv.vault.azure.net/ \
dotnet run --project OutlookWriteback.Bootstrap
```

Opens a browser for a one-time sign-in against the "Outlook Writeback MCP" Entra app registered at the top of this doc, then writes the resulting refresh token directly to Key Vault. Re-run only as a recovery step if the stored token ever goes bad — Entra doesn't immediately revoke the previous refresh token on redemption (90-day sliding window for this app's redirect type), so a missed rotation is a degraded/wasted-redemption condition, not an outage; see `SilentGraphCredential`'s doc comments for the full reasoning.

## Deploy

```
cd outlook-writeback
func azure functionapp publish <server>-func --dotnet-isolated
```

Manual, not CI/CD — reasonable at low deploy frequency for a single-user tool; wire up a pipeline if you deploy often. The `--dotnet-isolated` flag is required; `func` can't otherwise determine the project language when multiple sibling `.csproj` files share this directory.

## Wire up Claude Code CLI via system key (superseded — do not use for a fresh deployment)

**Skip this section on a from-scratch deployment.** This was the original auth path (the MCP extension's own `x-functions-key` system key) and it stops working once `host.json`'s `webhookAuthorizationLevel: "Anonymous"` setting is in place (see the multi-client OAuth section below) — that setting is already committed to this repo, so a fresh clone-and-deploy publishes with it from the start, and the key stops being a real gate the moment Easy Auth is configured. Go straight to **Migrate Claude Code CLI to OAuth** instead. This section is kept only because it explains the mechanism the OAuth path had to route around.

Retrieve the MCP extension's platform-managed system key (not a secret this project generates or stores — issued and rotated by the Functions platform):

```
az functionapp keys list --resource-group <server>-rg --name <server>-func --query systemKeys.mcp_extension --output tsv
```

Then:

```
claude mcp add --transport http outlook-write https://<server>-func.azurewebsites.net/runtime/webhooks/mcp --header "x-functions-key: <key>"
```

## Multi-client OAuth via Easy Auth + Entra ID

This is the auth path to use. It replaces the superseded `x-functions-key` section above and serves all three client surfaces (Claude Code CLI, Desktop/Cowork, claude.ai) through one Entra app.

If you want to try this out before touching a working deployment, do it in a throwaway resource group (`<server>-spike-rg` / `<server>-spike-func`) kept isolated per `../REPO-CONVENTIONS.md`'s per-server invariant, and delete it afterwards. See `docs/archive/` for why Easy Auth was chosen over a hand-rolled OAuth 2.1 server, and for the coexistence finding that forces the Claude Code CLI migration below.

Three steps below are easy to get wrong and fail in non-obvious ways — a config check that only shows up against a real client app, not against `curl`. They're called out inline where they occur.

### `host.json`: let Easy Auth be the sole gate

Once Easy Auth requires authentication, the MCP extension's own separate `mcp_extension` system-key check still runs underneath it and returns 403 for a valid Easy-Auth-validated bearer token unless relaxed:
```json
"extensions": {
  "mcp": {
    "system": {
      "webhookAuthorizationLevel": "Anonymous"
    }
  }
}
```
Don't reach for the app-setting equivalent (`AzureFunctionsJobHost__extensions__mcp__system__webhookAuthorizationLevel=Anonymous`) — it persists correctly and shows up in `az functionapp config appsettings list`, but has no observable effect even after a full stop/start cycle. Only the `host.json` literal, republished via `func azure functionapp publish`, actually changes behavior. Every instance you deploy builds from the same `outlook-writeback/host.json`, so this setting is already in place on a fresh clone. Once it is, the system key stops providing any real protection on its own — Easy Auth becomes the only enforcing layer, which is why the Claude Code CLI has to move off `x-functions-key` entirely rather than keeping both mechanisms alive side by side.

### Register the boundary-7b Entra app

Separate from the "Outlook Writeback MCP" app (boundary 7a, Graph delegation) — this one represents the MCP server's own audience:
```
az ad app create --display-name "Outlook Writeback MCP Connector" --sign-in-audience AzureADMyOrg --output json
```
Then, against the returned object ID (not the appId) — `az ad app update` doesn't cover all of these fields, so use direct Graph PATCH via `az rest --method PATCH` against `https://graph.microsoft.com/v1.0/applications/<object-id>` — **each of the PATCH calls below needs `--headers "Content-Type=application/json"` alongside `--body`**, or Graph rejects it with `BadRequest: Write requests (excluding DELETE) must contain the Content-Type header declaration` even though the body is valid JSON:

1. Set `{"api": {"requestedAccessTokenVersion": 2}}` **before** setting `identifierUris` — Entra rejects a same-tenant HTTPS App ID URI with `InvalidUniqueTenantIdentifierAsPerAppPolicy` otherwise.
2. Set `identifierUris` to your server's hostname, both with and without the `/runtime/webhooks/mcp` suffix (`https://<server>-func.azurewebsites.net` and `.../runtime/webhooks/mcp`) — the MCP extension resource-checks against the full route, not just the host, so both forms are needed to avoid `AADSTS9010010`. If you're also setting up a custom domain, read that section first and register the final hostname once rather than swapping it later.
3. Add an "Expose an API" scope via `{"api": {"oauth2PermissionScopes": [{"type": "User", "value": "access_as_user", "isEnabled": true, "id": "<new-guid>", "adminConsentDisplayName": "...", "adminConsentDescription": "...", "userConsentDisplayName": "...", "userConsentDescription": "..."}]}}`.
4. Register Claude Code CLI's loopback redirect URIs under `publicClient.redirectUris` (`http://localhost/callback`, `http://127.0.0.1/callback`, plus any fixed `--callback-port` used) with `"isFallbackPublicClient": true` — needed once the CLI migrates to this same OAuth flow (`AADSTS500113` otherwise).
5. **Register Desktop/Cowork's redirect URI too — this is required, not optional.** It's tempting to assume Anthropic's client handles its own redirect and needs nothing registered; it doesn't, and real Claude Desktop fails with `AADSTS50011: The redirect URI 'https://claude.ai/api/mcp/auth_callback' ... does not match the redirect URIs configured for the application`. Add `https://claude.ai/api/mcp/auth_callback` under **`web.redirectUris`** (not `publicClient.redirectUris`) — Desktop/Cowork presents the Client Secret during token exchange, making it a confidential client, which Entra requires registering under the `web` platform rather than `publicClient`:
   ```
   az rest --method PATCH \
     --url "https://graph.microsoft.com/v1.0/applications/<object-id>" \
     --headers "Content-Type=application/json" \
     --body '{"web":{"redirectUris":["https://claude.ai/api/mcp/auth_callback"]}}'
   ```

**One boundary-7b app per server, per `../REPO-CONVENTIONS.md` §3.** If you spiked this in a throwaway resource group first, repoint that same app's `identifierUris` at your real hostname (`az rest --method PATCH`, dropping the spike entries) rather than registering a second one — a spike is a measurement, not a second server.

### Client secret

```
az ad app credential reset --id <connector-app-id> --append --display-name "<server>-connector-secret" --years 1 --output json
```
Store the resulting `password` the way every other secret in this project is stored — **Key Vault**, not a raw Function App setting, referenced via `clientSecretSettingName` in the `authsettingsV2` payload below. Store it as the `<server>-connector-secret` Key Vault secret and reference it from a `@Microsoft.KeyVault(SecretUri=...)` app setting named `MICROSOFT_PROVIDER_AUTHENTICATION_SECRET`, matching `OUTLOOK_WRITEBACK_CONFIRMATION_SIGNING_KEY`'s pattern in step 8 above — Easy Auth's `clientSecretSettingName` field just needs an app-setting *name* to resolve at runtime, so a Key Vault-reference setting works transparently there too. Delete any throwaway spike credential from the app once the real one is in place.

**This secret is for Easy Auth's own server-side token exchange only — it is not the Claude Code CLI's secret.** The Entra app has two client registrations layered on one App ID: a confidential one (`web.redirectUris`, uses this secret — Easy Auth itself, and Desktop/Cowork) and a public one (`publicClient.redirectUris`, PKCE only, `isFallbackPublicClient: true` — the CLI). Passing this secret to `claude mcp add --client-secret` produces `AADSTS700025: Client is public so neither 'client_assertion' nor 'client_secret' should be presented` at token-exchange time — the CLI must be registered with **`--client-id` only, no `--client-secret`** (see the CLI section below).

### `authsettingsV2` via ARM REST PUT

`az webapp auth update` fails outright on an app already on auth v1 (`Cannot use auth v2 commands when the app is using auth v1` — the same class of CLI-command-group gap as the `az role assignment create` bug in step 5 above), so go straight to the ARM REST PUT:
```
az rest --method PUT \
  --url "https://management.azure.com/subscriptions/<sub-id>/resourceGroups/<server>-rg/providers/Microsoft.Web/sites/<server>-func/config/authsettingsV2?api-version=2022-03-01" \
  --body '{
    "properties": {
      "platform": { "enabled": true },
      "globalValidation": {
        "requireAuthentication": true,
        "unauthenticatedClientAction": "Return401"
      },
      "identityProviders": {
        "azureActiveDirectory": {
          "enabled": true,
          "registration": {
            "clientId": "<connector-app-id>",
            "clientSecretSettingName": "MICROSOFT_PROVIDER_AUTHENTICATION_SECRET",
            "openIdIssuer": "https://login.microsoftonline.com/<tenant-id>/v2.0"
          },
          "validation": {
            "defaultAuthorizationPolicy": { "allowedPrincipals": {} },
            "jwtClaimChecks": {}
          }
        },
        "facebook": {"enabled": false}, "gitHub": {"enabled": false}, "google": {"enabled": false},
        "legacyMicrosoftAccount": {"enabled": false}, "twitter": {"enabled": false}, "apple": {"enabled": false}
      }
    }
  }'
```
**Leave `allowedAudiences` and `defaultAuthorizationPolicy.allowedApplications` out** — the payload above omits both deliberately. Adding them looks like it should tighten validation, but it doesn't fix the bare-`WWW-Authenticate` symptom described below (that's the app setting's job) and it adds nothing here. Explicitly disabling the other built-in identity providers, on the other hand, does matter: it stops the ARM API silently defaulting them into an unconfigured-but-enabled state.

**Required app setting, and poorly documented upstream:** `WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` must be set to `https://<server>-func.azurewebsites.net/runtime/webhooks/mcp/access_as_user` (the app's own resource URI + its `access_as_user` scope). Without it, Easy Auth still enforces auth correctly and the PRM document (`/.well-known/oauth-protected-resource`) still serves fine on its own, but the `401` response's `WWW-Authenticate` header comes back bare — no `scope`, no `resource_metadata` — which breaks RFC 9728 discovery for any client relying on that header rather than probing the PRM endpoint directly. This is the single easiest thing to miss in this whole section, and auth appears to work without it:
```
az functionapp config appsettings set \
  --resource-group <server>-rg \
  --name <server>-func \
  --settings WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES="https://<server>-func.azurewebsites.net/runtime/webhooks/mcp/access_as_user"
```
Restart after applying either of the above:
```
az functionapp restart --resource-group <server>-rg --name <server>-func
```

### PRM endpoint — no custom code needed

`GET /.well-known/oauth-protected-resource` (and its `/runtime/webhooks/mcp`-suffixed sibling) is served natively by Easy Auth v2 once the above is live. It should return:
```json
{"resource":"https://<server>-func.azurewebsites.net","authorization_servers":["https://login.microsoftonline.com/<tenant-id>/v2.0"]}
```
No hand-rolled RFC 9728 shim endpoint is needed — the fallback considered in `docs/archive/` turned out to be unnecessary.

### Validation checks

The `Accept` header is checked *after* auth, so an unauthenticated 406 doesn't confirm anything about auth — always pair with the auth header when interpreting results:
1. No `Authorization` header, correct `Accept: application/json, text/event-stream` → `401` plus `WWW-Authenticate: Bearer ... resource_metadata="https://<server>-func.azurewebsites.net/.well-known/oauth-protected-resource/runtime/webhooks/mcp"` (only once `WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` is set — see above).
2. Valid bearer token, missing/wrong `Accept` header → `406` (proves the request passed auth and reached content negotiation).
3. Valid bearer token, correct `Accept` header → `200` plus a full MCP `initialize` JSON-RPC response. In practice the easiest way to confirm this one is a real `create_draft` call end to end from a connected client rather than a raw `curl` — the CLI's OAuth SDK manages the token internally and there's no simple way to extract it for a manual request.

### Migrate Claude Code CLI to OAuth

Remove any existing `x-functions-key`-based registration, then re-add — **no `--client-secret`**, since the CLI uses the app's public-client (PKCE) registration, not the confidential one Easy Auth itself uses (see the client secret section above; passing one here throws `AADSTS700025`):
```
claude mcp remove outlook-write
claude mcp add --transport http outlook-write https://<server>-func.azurewebsites.net/runtime/webhooks/mcp \
  --client-id <connector-app-id> --callback-port <port>
```
Pick any free high port for `--callback-port`; it just has to match a port you already registered under `publicClient.redirectUris` (both `http://localhost:<port>/callback` and `http://127.0.0.1:<port>/callback`).

Running `/mcp` in the CLI attempts an automatic browser-redirect sign-in. If it fails with `SDK auth failed: Incompatible auth server: does not support dynamic client registration`, the CLI's normal flow couldn't complete (Entra has no DCR support, which is why `--client-id` is supplied manually in the first place) — the CLI falls back to exposing two synthetic tools on the server, `authenticate`/`complete_authentication`, for a manual auth-code flow: `authenticate` returns an `https://login.microsoftonline.com/.../authorize` URL to open in a browser; if the `localhost` redirect page doesn't load after sign-in (expected — nothing is listening there outside the CLI's own OAuth flow, depending on client version), copy the full URL from the address bar and pass it to `complete_authentication`.

### Cleanup

If you spiked this first, tear the spike down once the real deployment validates against your clients:

- Remove the spike's throwaway credential from the connector app, leaving only the `<server>-connector-secret` Key Vault entry.
- Delete the whole spike resource group — Function App, storage account, App Service plan, and Log Analytics workspace go with it:
  ```
  az group delete --name <server>-spike-rg --yes
  ```

## Local smoke test (before touching Azure at all)

```
cd outlook-writeback
func start
```

Prints the discovered function list on startup (confirms tool discovery instantly, rather than waiting on a deploy-and-poll loop) and exposes `http://localhost:7071/runtime/webhooks/mcp` for a full local transport check. Point `local.settings.json`'s `AzureWebJobsStorage` at the real storage account (or run Azurite) and set all four `OUTLOOK_WRITEBACK_*` values — `TENANT_ID`, `CLIENT_ID`, `KEY_VAULT_URI`, and `CONFIRMATION_SIGNING_KEY` — to exercise the real Key Vault + Entra app from a local run, via your own `az login` session (`DefaultAzureCredential` picks it up automatically). `CONFIRMATION_SIGNING_KEY` doesn't have to be the same value as the deployed secret for a local smoke test — any base64 string works, e.g. `openssl rand -base64 32` (bash) or `[Convert]::ToBase64String([System.Security.Cryptography.RandomNumberGenerator]::GetBytes(32))` (PowerShell).

## Known issue: Claude-side OAuth refresh gap for external-IdP connectors

This server's boundary-7b OAuth (Easy Auth + Entra ID, per the multi-client OAuth section above) is affected by an open, currently-unfixed Claude-side bug — [anthropics/claude-ai-mcp#228](https://github.com/anthropics/claude-ai-mcp/issues/228) — where custom connectors backed by an external IdP (Entra ID is the case cited in the issue) don't get Claude's documented reactive/proactive token refresh applied. In practice the connector shows "Connected," but the access token forwarded on the next tool call has expired, and the call fails until the user manually reconnects it. Entra access tokens default to roughly a one-hour lifetime, so any invocation after an idle gap longer than that is likely to hit this — the less often you call the server, the more often you'll see it. See `docs/archive/` for the full writeup and the mitigation adopted (tool-level guidance to prompt a reconnect, plus an open action item to check the connector app's token-lifetime/Conditional Access configuration).

## Known dependency pin

`Microsoft.ApplicationInsights.WorkerService` is pinned to `2.23.0` in `OutlookWriteback.csproj`, not the newer `3.1.2` major release. `Microsoft.Azure.Functions.Worker.ApplicationInsights` 2.50.0 transitively depends on `Microsoft.ApplicationInsights.PerfCounterCollector` 2.23.0 (AI SDK 2.x) — forcing the 3.x line loads an incompatible `ITelemetryInitializer` and crashes the isolated worker process with a `TypeLoadException` at startup, silently preventing the Function App from ever specializing. Revisit this pin once the Functions ecosystem catches up to AI SDK 3.x.
