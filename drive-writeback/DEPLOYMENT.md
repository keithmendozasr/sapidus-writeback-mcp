# Deployment — drive-writeback

Checklist of `az` commands to provision this server, kept as a runbook rather than IaC — same rationale as `outlook-writeback/DEPLOYMENT.md`: proportionate for a single-resource-group, single-user deployment, worth converting to Bicep/Terraform only if this pattern gets repeated many times over.

**This is meant to work as a full disaster-recovery runbook.** Angle-bracket values (`<tenant-id>`, `<client-id>`, `<sub-id>`, etc.) are placeholders for values specific to your own deployment — substitute your own tenant/subscription IDs as you go. `<server>` stands for this server's folder name (`drive-writeback`), written generically the same way `outlook-writeback/DEPLOYMENT.md` is, so this doc reads correctly whether you're standing this up in the original deployment tenant or your own tenant on a fresh OSS deployment.

**Phase 1 code is implemented; none of the infrastructure below has been provisioned yet.** The Functions host project, `DriveWriteback.Bootstrap`, and their runbook sections now exist (this doc's "Resources provisioned" section onward), but no resource group, Function App, or Key Vault has actually been created — implementing Phase 1's code, tests, and this runbook was explicitly in scope for that pass; running `az`/`func` commands against a real subscription was explicitly not (see `drive-writeback/CLAUDE.md`'s Status section). You run each section below yourself when you're ready to stand this up.

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
   ```
   az rest --method GET \
     --url "https://graph.microsoft.com/v1.0/servicePrincipals?\$filter=appId eq '00000003-0000-0000-c000-000000000000'&\$select=oauth2PermissionScopes" \
     --query "value[0].oauth2PermissionScopes[?value=='Files.ReadWrite.All' || value=='Sites.Read.All'].{value:value,id:id}"
   ```
   (`00000003-0000-0000-c000-000000000000` is Microsoft Graph's well-known resource `appId` — the same in every tenant.) Note the two returned `id` values as `<files-readwrite-all-id>` and `<sites-read-all-id>`.
4. Add the two permissions:
   ```
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

**Everything below this line is Phase 1 runbook text — none of it has been executed as part of this implementation pass.** Per that pass's scope boundary (code, tests, and this runbook only), no `az`/`func` commands were run and no live Azure resources were provisioned; you run each section below yourself once the code is ready. The one-time interactive Bootstrap sign-in also needs your own browser session, so it can't happen in an implementation pass regardless.

## Resources provisioned (in order)

Mirrors `outlook-writeback/DEPLOYMENT.md`'s own "Resources provisioned" section exactly — same gotchas, same workarounds, same reasoning — with names and Graph scopes swapped for this server. Read that section for the full detail (globally-unique-name warning, the `az role assignment create` `MissingSubscription` workaround, the PowerShell `--%` app-settings gotcha); only the differences are called out below.

1. Resource group: `az group create -n <server>-rg -l <region> --tags project=sapidus-writeback-mcp`.
2. Storage account (Flex Consumption's required host storage): `az storage account create -g <server>-rg -n <server>sa -l <region> --sku Standard_LRS --tags project=sapidus-writeback-mcp`.
3. Function App on Flex Consumption, .NET 10 isolated worker, capped at max instance count 1 (same single-user-refresh-token-races-itself reasoning as `outlook-writeback` — see `DriveWriteback.Graph/Auth/SilentGraphCredential.cs`):
   ```
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
7. **No confirmation-signing-key secret** — unlike `outlook-writeback`'s `delete-confirmation-signing-key`, this server has no delete surface yet (Phase 2), so there's nothing analogous to provision here.
8. Function App application settings — no Key Vault-reference setting needed here (unlike `outlook-writeback`'s confirmation-signing-key), so the PowerShell `--%` gotcha that doc describes doesn't apply to this list; a plain backtick-continued command is fine:
   ```powershell
   az functionapp config appsettings set `
     --resource-group <server>-rg `
     --name <server>-func `
     --settings `
       DRIVE_WRITEBACK_TENANT_ID=<tenant id> `
       DRIVE_WRITEBACK_CLIENT_ID=<"Drive Writeback MCP" app registration's client id> `
       DRIVE_WRITEBACK_KEY_VAULT_URI=https://<server>-kv.vault.azure.net/ `
       DRIVE_WRITEBACK_DRY_RUN=true `
       DRIVE_WRITEBACK_MAX_CONTENT_BYTES=1048576
   ```
   Leave `DRIVE_WRITEBACK_DRY_RUN=true` until you've confirmed writes behave as expected against the live tenant, then flip it to `false` (no code change needed — see `Program.cs`).
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
func start
```
Point `local.settings.json`'s `AzureWebJobsStorage` at the real storage account (or run Azurite) and set `DRIVE_WRITEBACK_TENANT_ID`/`CLIENT_ID`/`KEY_VAULT_URI`/`DRY_RUN`/`MAX_CONTENT_BYTES` to exercise the real Key Vault + Entra app from a local run, via your own `az login` session (`DefaultAzureCredential` picks it up automatically). Leave `DRIVE_WRITEBACK_DRY_RUN` at `true` for this first run — confirm `create_folder`/`create_file` resolve and validate correctly before ever flipping it to `false` against real data.

## Not yet added: Easy Auth / boundary-7b Connector app

Deliberately absent from this doc, same "append once actually needed" posture the doc already states for itself above. Phase 1 has no boundary-7b Connector app and no Easy Auth — `host.json` correspondingly omits `system.webhookAuthorizationLevel: "Anonymous"` (see that file's own comment history / the commit that added it): relaxing that setting is only safe once Easy Auth is in front of it, per the finding already documented in `outlook-writeback/DEPLOYMENT.md`'s multi-client OAuth section. Until boundary-7b lands here, the MCP extension's own system-key check is the only gate — mirror `outlook-writeback/DEPLOYMENT.md`'s "Wire up Claude Code CLI via system key" section (`az functionapp keys list ... systemKeys.mcp_extension`, `claude mcp add ... --header "x-functions-key: <key>"`) as the interim path, and build out the full multi-client OAuth section here (mirroring that same doc) once this server needs multiple callers.
