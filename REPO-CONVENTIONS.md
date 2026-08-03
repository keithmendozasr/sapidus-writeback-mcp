# sapidus-writeback-mcp: Repo Conventions

**Applies to:** every server folder in this repo, current and future.
**Owner:** Keith Mendoza
**Status:** Active

This document exists so these rules are defined once. Individual server design docs (each server's own `docs/` folder) link back here instead of repeating them.

---

## 1. What this repo is

`sapidus-writeback-mcp` is a monorepo of narrow, single-purpose **write-only** MCP servers, each covering a slice of Microsoft Graph that Anthropic's official M365 connector doesn't cover on a Pro plan (write access is a Team/Enterprise-only surface — see any server's own design docs for the full context).

Each server:
- Handles **writes only**. Read/search stays with the M365 connector — never duplicated here.
- Requests the **minimum Graph scopes** needed for its own job, and nothing else.
- Is independently deployable and independently revocable.

The repo is one codebase for developer convenience. It is deliberately **not** one deployed identity, one resource group, or one set of Graph scopes across capabilities.

## 2. Directory layout

```
sapidus-writeback-mcp/
├── REPO-CONVENTIONS.md          # this file
├── shared/                       # 7b transport/auth code ONLY — see §4
├── outlook-writeback/
│   ├── CLAUDE.md                 # build/test commands, runtime specifics for this server
│   ├── docs/
│   │   ├── active/                # in-progress specs, not yet shipped
│   │   └── archive/                # specs for shipped/superseded work
│   ├── OutlookWriteback.csproj
│   ├── Program.cs
│   ├── host.json
│   ├── local.settings.json       # gitignored, local dev only
│   └── Functions/
│       └── ...                   # one class per MCP tool
├── onedrive-writeback/           # future — same shape
│   ├── CLAUDE.md
│   ├── docs/
│   │   ├── active/
│   │   └── archive/
│   ├── OnedriveWriteback.csproj
│   ├── Program.cs
│   ├── host.json
│   ├── local.settings.json
│   └── Functions/
│       └── ...
└── ...
```

Each server is a C# Azure Functions project (isolated worker model).

One folder per server. Each folder is a complete, independently deployable unit with its own design docs under `docs/`.

## 3. Per-server isolation (the invariant)

For every server folder, all of the following are separate and never shared across servers:

| Layer | Convention | Example |
|---|---|---|
| Entra App Registration | one per server, scoped to only that server's Graph permissions | "Outlook Writeback MCP" — `Mail.ReadWrite`, `Calendars.ReadWrite` only |
| Resource group | `<server>-rg` | `onedrive-writeback-rg` |
| Function App | `<server>-func` | `onedrive-writeback-func` |
| Key Vault secrets | own entries per server, never reused across servers | Outlook's Graph client secret is a separate Key Vault secret from any future server's |

**One Function App per MCP server, matching the one-Entra-app-per-server invariant above.** Easy Auth (App Service Authentication v2) is configured at the Function App level, not per-route — there is no way to point different paths within one Function App at different identity providers. Sharing a Function App across two servers would force them to share one boundary-7b Connector app too, which breaks the invariant this section exists to protect. Each server gets its own subdomain (`<server>.<your-domain>`) and its own Flex Consumption plan, never a shared one.

**Why this is the rule and not a starting point to relax later:** the whole reason these servers exist as separate write-only tools instead of one broad-access app is to keep blast radius small — a leaked credential or a buggy handler in one server should never expose Graph scopes belonging to another. Adding a new capability by granting an *existing* app more scopes would quietly undo that property. New capability = new folder = new Entra app = new resource group = new Function App, every time.

Naming stays traceable end-to-end: folder name → resource group → Function App → Entra display name, all built from the same `<server>` string, so an incident in the Azure Portal can be traced back to exactly which Entra app's scopes are exposed without cross-referencing a spreadsheet.

## 4. What's allowed in `shared/`

`shared/` is for code that is genuinely capability-agnostic — concretely, the Claude-to-MCP-server auth/transport layer (what each server's PRD calls "boundary 7b": the bearer token or self-issued OAuth 2.1 layer that decides *who's allowed to call this server's tools*, as opposed to *what the server is allowed to do in Graph*).

**Never allowed in `shared/`:**
- Graph client secrets, certificates, or tokens for any specific server
- Business logic specific to one server's tools (drafting, calendar math, etc.)
- Anything that would require one server's Key Vault access to reach another server's secrets

If it's unclear whether something belongs in `shared/`, the test is: *would this code need to change if a new server's Graph scopes changed?* If yes, it's server-specific, not shared.

## 5. Entra App Registration naming

Display name pattern: **"`<Server Name>` MCP"** in title case, e.g. "Outlook Writeback MCP." (This is one instance of the broader agent-agnostic principle in §9 — the naming rule below is about keeping the *architecture* client-neutral, not about avoiding Claude in docs generally.)

- **Never include "Claude" in the display name.** The tenant already has Anthropic's own official service principals — "M365 MCP Client for Claude" and "M365 MCP Server for Claude" — in Enterprise Applications. A self-hosted app named with "Claude" in it both collides visually with those in the app list and misleadingly implies Anthropic affiliation or endorsement, which matters more here since this repo is open source and others may register their own tenant's copy under a suggested name.
- If it's useful to note which client(s) call a server, put that in the app's **description/publisher-notes field** (free text, not the searchable display name) — e.g. "Self-hosted write-only MCP server; called by Claude Code/Desktop/Cowork over MCP; not an Anthropic product."
- The distinctive, consistent naming (vs. something generic like "Mail Server") is what makes the app instantly recognizable in a tenant's Enterprise Applications list, which can otherwise include many first-party Microsoft and Anthropic entries.

## 6. Secrets handling

- **No production secrets in `local.settings.json`, ever.** `local.settings.json` is the standard Azure Functions local-development convention (values under its `Values` key, read via `IConfiguration` when running a Function locally).
- `local.settings.json` is listed in `.gitignore` from the first commit of every server folder (the Azure Functions project template excludes it by default). A `local.settings.json.example` with placeholder keys ships alongside it so contributors know what local variables are needed.
- Production secrets live in **Azure Key Vault**, one secret set per server (never shared across servers — see §3).
- Each Function App's Application Settings hold **Key Vault references** (`@Microsoft.KeyVault(SecretUri=...)`), not raw values.
- Each Function App reaches Key Vault via its own **Managed Identity** — no stored bootstrap credential is needed to fetch the actual secrets.
- Logging never includes secret values, email/event body content, or other mailbox content — tool invocation metadata only (name, timestamp, success/failure), per each server's own PRD.

## 7. Cost tracking across resource groups

Since each server gets its own resource group (§3), apply a common tag to every resource in every server's resource group:

```
project: sapidus-writeback-mcp
```

This keeps the family visible as one rollup in Azure Cost Management even though the resource groups themselves stay isolated. Each server's own PRD still sets its own budget alert threshold appropriate to its expected volume.

## 8. Adding a new server: checklist

1. Create `sapidus-writeback-mcp/<server-name>/` with its own `docs/active/` spec, scoped to the specific Graph write gap it fills (mirror the shape of an existing server's docs, e.g. `outlook-writeback/docs/`).
2. Register the new server with release-please so it gets its own independent release PR: add a `<server-name>` package entry to `release-please-config.json` (mirror the `outlook-writeback` entry), and seed `.release-please-manifest.json` and `<server-name>/version.txt` with its starting version.
3. Add that server's own `CLAUDE.md` (build/test commands, runtime specifics) and a one-line pointer to it in root `CLAUDE.md`'s `## Servers` section — nothing more detailed than that goes in the root file.
4. Register a new single-tenant Entra App Registration, display name `"<Server Name> MCP"`, requesting only the Graph scopes that server needs.
5. Create `<server-name>-rg` resource group and `<server-name>-func` Function App.
6. Create new Key Vault secrets for this server's Graph credentials — do not reuse another server's Key Vault entries.
7. Tag all new resources `project: sapidus-writeback-mcp`.
8. Only pull code into `shared/` if it's genuinely capability-agnostic transport/auth code per §4 — default to keeping new logic in the server's own folder.
9. Once the server has a boundary-7b Connector app (multi-client OAuth via Easy Auth + Entra ID — see that server's own `DEPLOYMENT.md`), set **"Assignment required" = Yes** on its Enterprise Application object and assign only the intended user(s). Entra allows any tenant user to sign in to an app registration by default once it exists; since Graph calls run under a single seeded server-app refresh token rather than per-caller delegation, an unrestricted Connector app would let any tenant user drive the server with the owner's own Graph access. Confirmed working on Entra ID Free (individual user assignment, not group-based, needs no P1/P2 upgrade).

## 9. MCP protocol and architecture decisions must be agent-agnostic

Per the MCP spec, an MCP server has no concept of which AI agent or model is calling it — it implements the Model Context Protocol (JSON-RPC over stdio/HTTP, `tools/list`, `tools/call`, etc.) for any conforming MCP client. Every server in this repo must preserve that property:

- **Tool schemas, transport choice, and protocol behavior must never assume Claude specifically.** Design and test against the MCP spec, not against one client's quirks. A tool that only works correctly when called by Claude Desktop/Code/Cowork — and would break or misbehave under a different conforming MCP client — is a bug in the server, not an acceptable simplification.
- **Mentioning Claude in docs is fine and expected where it's purely descriptive** — e.g., stating that Claude Desktop/Cowork/Code CLI are a server's current target clients, or giving Claude-specific setup instructions (`claude mcp add ...`). That's documenting who happens to be using the server today, not a design constraint.
- **The line to hold:** anything baked into the *architecture* — Entra app naming (see §5), auth flows, tool contracts, transport behavior — stays agent-agnostic. Anything that's just *describing current usage* — target clients, setup docs, examples — may reference Claude freely.
- If a new capability or design choice would only work for Claude, that's a signal to reconsider the design, not to special-case it — the whole reason this is an MCP server instead of a Claude-specific plugin is so any conforming client can call it later.
