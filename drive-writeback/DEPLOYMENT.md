# Deployment — drive-writeback

Checklist of `az` commands to provision this server, kept as a runbook rather than IaC — same rationale as `outlook-writeback/DEPLOYMENT.md`: proportionate for a single-resource-group, single-user deployment, worth converting to Bicep/Terraform only if this pattern gets repeated many times over.

**This is meant to work as a full disaster-recovery runbook.** Angle-bracket values (`<tenant-id>`, `<client-id>`, `<sub-id>`, etc.) are placeholders for values specific to your own deployment — substitute your own tenant/subscription IDs as you go. `<server>` stands for this server's folder name (`drive-writeback`), written generically the same way `outlook-writeback/DEPLOYMENT.md` is, so this doc reads correctly whether you're standing this up in the original deployment tenant or your own tenant on a fresh OSS deployment.

**Phase 0 scope only, for now.** This server currently has no Functions host project, no resource group, no Function App, no Key Vault — only the Graph client library and its validation-spike tests (see `docs/active/PRD-drive-write.md` §11). The one thing Phase 0 actually needs from Azure is the boundary-7a Entra app registration below, so that's the only section here. Resource group / Function App / Key Vault / Bootstrap / boundary-7b Connector app sections get appended once Phase 1 actually needs them — mirror the corresponding sections of `outlook-writeback/DEPLOYMENT.md` when that time comes, adjusting scopes and naming for this server.

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
6. Register the public-client redirect URI. Phase 0 only needs one — `http://localhost`, for the `Category=E2E` test tier's `InteractiveBrowserCredential` — unlike `outlook-writeback`, there's no second fixed-port redirect URI yet because there's no `DriveWriteback.Bootstrap` console tool in Phase 0 (that's Phase 1, once a deployed non-interactive service needs a seeded refresh token):
   ```
   az ad app update --id <client-id> --public-client-redirect-uris "http://localhost"
   ```
   No `web.redirectUris` or client secret is needed — this app is used exclusively as a public client via delegated auth-code + PKCE.

At this point you have everything the Phase 0 E2E spike tests need:
```
DRIVE_WRITEBACK_TENANT_ID=<tenant-id>
DRIVE_WRITEBACK_CLIENT_ID=<client-id>
```
Run `dotnet test --filter Category=E2E` from the repo root with those two set — it'll open a real browser sign-in the first time.
