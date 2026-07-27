# Sapidus Writeback MCP

A monorepo of narrow, single-purpose **write-only** [MCP](https://modelcontextprotocol.io) servers, each covering a slice of Microsoft Graph that Anthropic's official Microsoft 365 connector doesn't expose on a Pro plan — write access there lives behind **Organization settings → Connectors**, a Team/Enterprise-only admin surface with no self-service path for Free/Pro/Max users.

Each server is self-hosted on Azure Functions, deployed into your own Azure subscription and your own Entra tenant. Nothing here is a hosted service, and nothing phones home.

## Design premise

Every server in this repo:

- **Writes only.** Read and search stay with the M365 connector and are never duplicated here. That halves the permission surface — a server that can't read your mailbox can't leak it.
- **Requests the minimum Graph scopes for its own job, and nothing else.** `outlook-writeback` asks for `Mail.ReadWrite` and `Calendars.ReadWrite`. It deliberately does *not* request `Mail.Send`, so it structurally cannot send mail on your behalf — only draft it for you to review.
- **Is independently deployable and independently revocable.** Its own Entra app registration, resource group, Function App, and Key Vault secrets. Revoking one server's access never touches another's.

The monorepo is a developer convenience. It is deliberately **not** one deployed identity, one resource group, or one set of Graph scopes — see [`REPO-CONVENTIONS.md`](REPO-CONVENTIONS.md) for the isolation rules every server follows.

## Servers

| Server | Tools | Graph scopes |
|---|---|---|
| [`outlook-writeback`](outlook-writeback/) | `create_draft`, `update_draft`, `create_event`, `update_event`, `delete_event` | `Mail.ReadWrite`, `Calendars.ReadWrite` |

`delete_event` is the only destructive tool, and it's two-step. The first call returns the event's details plus a signed confirmation token and does nothing else; the actual `DELETE` happens only on a second call that echoes that token back. The token is HMAC-signed, bound to the event ID, and expires in five minutes, so a client can't mint its own and skip the round-trip.

Draft and event IDs are expected to come from the M365 connector's read/search tools in the same conversation — they share a Graph ID space, so no translation is needed. This server never implements its own read or search.

## Requirements

- An Azure subscription and a Microsoft 365 tenant you administer (registering the Entra app and granting admin consent needs Global Administrator or Privileged Role Administrator).
- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- [Azure Functions Core Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local)
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli)

## Build and test

From the repo root:

```
dotnet build
dotnet test --filter "Category!=E2E"
```

Tests run in three tiers. `Category=Unit` covers payload mapping and the auth machinery with no I/O. `Category=Integration` exercises real request-building and response-deserialization through the Graph SDK against a stubbed HTTP handler — no network, no credentials, about a second end to end. `Category=E2E` talks to real Microsoft Graph against a real mailbox and self-skips unless you've set the environment variables listed in [`outlook-writeback/CLAUDE.md`](outlook-writeback/CLAUDE.md); it can't run in CI.

For a local smoke test against your own Azure dependencies, `cd outlook-writeback && func start` prints the discovered MCP tool list on startup — a much faster loop than deploy-and-poll.

## Deploying

See [`outlook-writeback/DEPLOYMENT.md`](outlook-writeback/DEPLOYMENT.md). It's written as a full from-scratch runbook: Entra app registration, resource provisioning, Key Vault wiring, the one-time refresh-token bootstrap, multi-client OAuth via Easy Auth, and an optional custom domain. Every resource name in it is a placeholder for you to substitute.

Two things worth knowing before you start:

- Function App, storage account, and Key Vault names are **globally unique** across Azure. Pick your own suffix; the convention defaults in the docs will often be taken.
- There's a known Claude-side bug affecting OAuth token refresh for connectors backed by an external IdP ([anthropics/claude-ai-mcp#228](https://github.com/anthropics/claude-ai-mcp/issues/228)). Entra ID is the cited case. In practice the connector reports "Connected" but forwards an expired token, and you reconnect manually. DEPLOYMENT.md covers the mitigation.

## Agent-agnostic by design

Per the MCP spec, a server has no concept of which client is calling it. The servers here preserve that: tool schemas, transport choice, and protocol behavior never assume a specific client, and a tool that only worked under one client would be a bug. Claude appears in these docs descriptively — it's what's being used today, not a design constraint.

## License

[Apache License 2.0](LICENSE).
