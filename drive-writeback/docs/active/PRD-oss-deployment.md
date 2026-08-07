# PRD — Multi-Tenant / OSS Deployability for `drive-writeback`

**Owner:** Keith Mendoza
**Status:** Not started
**Target platform:** Azure
**Language:** C#
**Target clients:** Claude Desktop, Claude Cowork, claude.ai
**Repo location:** `sapidus-writeback-mcp/drive-writeback/` (one server folder in the `sapidus-writeback-mcp` monorepo — see `../REPO-CONVENTIONS.md` for the cross-server rules this PRD inherits)
**Extracted from:** `docs/active/PRD-drive-write.md` ("v1" below) — v1 was designed and deployed for a single, admin-owned tenant. This doc collects everything v1 flagged as relevant only to a future OSS release, where the deployer is not necessarily a tenant admin and may run with partial Graph consent. None of it is built. Every section of v1 not called out here is unchanged and unaffected by this PRD.

---

## 1. Problem statement

v1's two required Graph scopes (`Files.ReadWrite.All`, `Sites.Read.All`) both require **tenant admin consent**. On the tenant this server actually runs on, that's a non-issue — the owner is the admin. But if this repo is ever released as OSS (the stated long-term intent — see root `CLAUDE.md`), a deployer who is not a tenant admin cannot consent, and today the server would fail at startup rather than degrade gracefully or explain why.

A second, smaller gap in the same theme: this server's Entra app registrations are single-tenant, work/school accounts only (`REPO-CONVENTIONS.md`). Personal Microsoft accounts (MSA) are a different identity type entirely and would need a different account-type registration — currently unexplored.

## 2. Scope

### 2.1 Startup scope detection (the cheap half)

At startup, after redeeming the seeded refresh token, inspect the granted-scopes claim and log a clear, specific entry (Application Insights, same pipeline as the structured audit log — v1 §9 item 7) naming exactly which required scope is missing — e.g. "Sites.Read.All not granted; SharePoint tools will fail." This does not change runtime behavior: a SharePoint call still fails with Graph's own 403 if the scope is missing. The point is only that the *reason* is always discoverable in the log, never a mystery.

Verified not yet built: no code path in this repo inspects granted scopes or logs a missing-scope warning today.

### 2.2 Full graceful degradation (the real work)

Detect missing scopes and conditionally register tools, or have affected tools return a clear "scope not consented; some functionality disabled" error surfaced to the model at call time, rather than a raw Graph 403 — with the missing-scope state also reflected in tool descriptions at registration time so the model doesn't attempt calls that cannot succeed.

This is real work (scope introspection driving conditional tool registration) with zero benefit on an admin-owned tenant — it matters only for a future OSS deployer running with partial consent. **Build when an actual OSS release is imminent, not before.**

### 2.3 Personal Microsoft account (MSA) support

Currently out of scope entirely — this server's Entra app registrations are single-tenant (`AzureADMyOrg`), work/school accounts only. Supporting a personal OneDrive under a personal Microsoft account would require a different account-type registration (`signInAudience`) and has not been explored. Open question, not a committed goal of this PRD — flagged here so it isn't lost, not because it's scheduled.

## 3. Related, not yet in the doc: Connector app sign-in restriction

By default, Entra allows any tenant user to sign in to an app registration once it exists, unless sign-in is restricted. Since all Graph calls run under the single seeded Server-app refresh token (not per-caller delegation), an unrestricted Connector app would let any tenant user drive this server with the owner's own Graph access.

**Action (already resolved as feasible, not yet applied as an ongoing checklist item):** set "Assignment required" = Yes on the Connector app's Enterprise Application object and assign only the intended user. Confirmed during v1 Phase 0 that this works on the tenant's Entra ID Free license tier without a P1/P2 upgrade. This is a real security control worth re-verifying is in place any time the Connector app is touched (e.g. the boundary-7b Easy Auth work already deployed) — not just an OSS-release concern, though it becomes load-bearing the moment sign-in isn't restricted to one known user. Applies to `outlook-writeback` too; candidate for a `REPO-CONVENTIONS.md` callout regardless of when this PRD itself is picked up.

## 4. Success criteria

- A non-admin deployer running with partial consent gets a specific, actionable log entry at startup (§2.1) rather than a silent or generic failure.
- Tools affected by a missing scope either don't register, or fail with a clear "not consented" error distinguishable from every other failure mode (§2.2).
- "Assignment required" verified set on the Connector app as a standing check, independent of this PRD's own timeline (§3).

## 5. References

- Microsoft Graph — permissions reference: https://learn.microsoft.com/en-us/graph/permissions-reference
- Repo-internal: `../REPO-CONVENTIONS.md`, `docs/active/PRD-drive-write.md`
