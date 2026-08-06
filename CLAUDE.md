# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Servers

- **outlook-writeback** — all phases (spike, email MVP, calendar, multi-client OAuth, custom domain + cost hardening) are complete. See `outlook-writeback/CLAUDE.md` for build/test commands, runtime specifics, and pointers to that server's own design docs — this file doesn't duplicate them.
- **drive-writeback** — Phases 1 and 2 (safe creates plus the mutation surface: `update_file_content`, `rename_item`, `move_item`, `delete_item`) are implemented and deployed, including boundary-7b OAuth for Claude Desktop/claude.ai. See `drive-writeback/CLAUDE.md` for build/test commands, runtime specifics, and pointers to that server's own design docs — this file doesn't duplicate them.

## What this repo is

`sapidus-writeback-mcp` is a monorepo of narrow, single-purpose **write-only** MCP servers, each covering a slice of Microsoft Graph that Anthropic's official M365 connector doesn't expose on a Pro plan (write access there is a Team/Enterprise-only admin surface). Each server:

- Handles **writes only** — read/search stays with the M365 connector and is never duplicated here.
- Requests the **minimum Graph scopes** needed for its own job, nothing else.
- Is independently deployable and independently revocable.

The repo is one codebase for developer convenience only. It is deliberately **not** one deployed identity, one resource group, or one set of Graph scopes across capabilities — see `REPO-CONVENTIONS.md` for the full rules this implies.

## Directory layout (target shape)

```
sapidus-writeback-mcp/
├── REPO-CONVENTIONS.md          # cross-server rules, read this first
├── shared/                       # capability-agnostic code only — 7b transport/auth, plus generic non-Graph primitives like confirmation tokens
├── outlook-writeback/
│   ├── CLAUDE.md                 # build/test commands, runtime specifics for this server
│   ├── docs/
│   │   ├── active/                # in-progress specs, not yet shipped
│   │   └── archive/                # specs for shipped/superseded work
│   ├── DEPLOYMENT.md
│   ├── OutlookWriteback.csproj
│   ├── Program.cs
│   ├── host.json
│   ├── local.settings.json       # gitignored, local dev only
│   └── Functions/                # one class per MCP tool
└── <future-server>/              # same shape: C# Azure Functions project + CLAUDE.md + docs/
```

One folder per server. Each folder is meant to be a complete, independently deployable unit with its own design docs under `docs/`. Before adding a new server folder, follow the checklist in `REPO-CONVENTIONS.md` §8.

## The invariant: per-server isolation

For every server folder, the following are separate and never shared across servers — this is treated as a hard invariant, not a convenience to relax later (see `REPO-CONVENTIONS.md` §3 for the full rationale):

| Layer | Convention | Example |
|---|---|---|
| Entra App Registration | one per server, scoped to only that server's Graph permissions | "Outlook Writeback MCP" — `Mail.ReadWrite`, `Calendars.ReadWrite` only |
| Resource group | `<server>-rg` | `onedrive-writeback-rg` |
| Function App | `<server>-func` | `onedrive-writeback-func` |
| Key Vault secrets | own entries per server | never reused across servers |

Adding a new capability always means a new folder + new Entra app + new resource group + new Function App — never granting an existing app more scopes. Naming is traceable end-to-end: folder name → resource group → Function App → Entra display name, all built from the same `<server>` string.

**Entra display name rule:** never include "Claude" in a server's Entra app display name (pattern is `"<Server Name> MCP"`, e.g. "Outlook Writeback MCP"). The tenant already has Anthropic's own official service principals ("M365 MCP Client for Claude" / "M365 MCP Server for Claude"); a self-registered app named with "Claude" both collides visually with those and misleadingly implies Anthropic affiliation. If it's useful to note which client calls a server, put that in the app's description/publisher-notes field, not the display name. This is one instance of the broader agent-agnostic principle below — the naming rule is about keeping the architecture client-neutral, not about avoiding "Claude" in docs generally.

## MCP servers must stay agent-agnostic

Per the MCP spec, a server has no concept of which AI agent is calling it — it implements the protocol (JSON-RPC over stdio/HTTP, `tools/list`, `tools/call`, etc.) for any conforming client. This repo's servers must preserve that property (see `REPO-CONVENTIONS.md` §9):

- Tool schemas, transport choice, and protocol behavior must never assume Claude specifically — design and test against the MCP spec, not one client's quirks. A tool that only works when called by Claude is a bug, not an acceptable simplification.
- Mentioning Claude in docs is fine where it's purely descriptive — target clients, setup instructions (`claude mcp add ...`), examples. That documents who's using the server today; it isn't a design constraint.
- The line: anything baked into the architecture (Entra naming, auth flows, tool contracts, transport behavior) stays agent-agnostic; anything just describing current usage may reference Claude freely.

## What belongs in `shared/`

`shared/` is for genuinely capability-agnostic code. Two categories qualify so far:

1. **The Claude-to-MCP-server auth/transport layer** (each server's PRD calls this "boundary 7b": the bearer token or self-issued OAuth 2.1 layer deciding *who may call this server's tools*, as distinct from *what the server may do in Graph*).
2. **Generic, non-Graph safety primitives with no server identity baked in** — e.g. `ConfirmationTokenService`, the stateless HMAC-signed token that gates a destructive tool's second, confirming call (originally written for `outlook-writeback`'s `delete_event`, extracted when `drive-writeback`'s `delete_item` needed the identical mechanism). It never calls Graph, never references a Graph model, and each server still supplies its own signing key from its own Key Vault secret — only the mechanism is shared, not runtime state or credentials.

**Never in `shared/`:** Graph client secrets/certificates/tokens for any specific server, business logic specific to one server's tools, or anything that would require one server's Key Vault access to reach another's.

Test for whether something belongs in `shared/`: *would this code need to change if a new server's Graph scopes changed?* If yes, it's server-specific.

**This test alone isn't sufficient, though — see `REPO-CONVENTIONS.md` §4.** Boundary-7a auth code (redeem-refresh-token/cache-access-token/rotate-refresh-token) passes this same test — it's fully parameterized and holds no server identity either — yet is deliberately *not* shared, because doing so would couple every server's Graph-auth failure modes together, which is exactly what this repo's independent-deployability invariant exists to prevent. The distinguishing question for a borderline case: does this code touch Graph or credentials at all? `ConfirmationTokenService` doesn't (it signs an opaque id string and a timestamp, nothing else), so the 7a coupling risk doesn't transfer to it; boundary-7a auth code by definition does.

## Two independent auth boundaries — don't conflate them

Every server has two separate auth layers (see that server's own design docs under `docs/` for the full definitions):

- **7a — MCP server → Microsoft Graph.** Each server owns its own single-tenant Entra App Registration and OAuth 2.0 Authorization Code + PKCE flow to Graph, with only the delegated scopes that server needs. This is never shared across servers and never lives in `shared/`.
- **7b — Claude → MCP server.** Governs who's allowed to call the server's tools at all (in the Outlook server's case, a single user). Prefer a static bearer token (`claude mcp add --transport http ... --header "Authorization: Bearer <token>"`) over a full OAuth server unless the client surface requires OAuth. This boundary's code is one candidate for `shared/` — see `## What belongs in shared/` above for the other (generic, non-Graph safety primitives like confirmation tokens).

## Secrets handling

- No `local.settings.json` in production, ever — it's local-dev-only, listed in `.gitignore` from a server folder's first commit, with a `local.settings.json.example` placeholder alongside it.
- Production secrets live in Azure Key Vault, one secret set per server. Function App settings hold Key Vault references (`@Microsoft.KeyVault(SecretUri=...)`), never raw values.
- Each Function App reaches Key Vault via its own Managed Identity.
- Logging never includes secret values or mailbox/event content — tool invocation metadata only (name, timestamp, success/failure).

## Cost tracking

Every resource in every server's resource group gets the tag `project: sapidus-writeback-mcp`, so the family rolls up in Azure Cost Management despite the resource groups themselves staying isolated.

## Document maintenance

- Keep the documentation in sync with changes made as necessary. At minimum, this means both this root `CLAUDE.md` and every per-server `CLAUDE.md` — update each as its own scope changes, not just at milestone boundaries.
- Each server folder carries its own `CLAUDE.md` for build/test commands and runtime specifics scoped to that server. Root `CLAUDE.md` stays a cross-server rules doc plus a one-line-per-server pointer in `## Servers` — don't let per-server implementation detail accumulate back into this file.
- Every server gets its own release-please PR, never a shared repo-wide one — see `REPO-CONVENTIONS.md` §8 for the registration steps required when adding a new server folder, and that server's own `CLAUDE.md` for its specific versioning details.
- A PRD moves from `docs/active/` to `docs/archive/` (and its `Status` field only says `Completed`) once its work is actually released — merged **and** deployed to production — not merely code-complete or merged to `main`. While work is implemented but not yet deployed, keep the doc in `docs/active/` with a status that says so (e.g., `Implemented — pending deployment`).
