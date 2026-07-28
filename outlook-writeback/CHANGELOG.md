# Changelog

## [0.2.0](https://github.com/keithmendozasr/sapidus-writeback-mcp/compare/v0.1.0...v0.2.0) (2026-07-28)


### Features

* **bootstrap:** Add OutlookWriteback.Bootstrap for one-time token seeding ([4dbfd7c](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/4dbfd7c84e0b94a64dcd6fee2557505db60dbb00))
* **functions:** Add the create_draft MCP tool ([78aff3c](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/78aff3c3b1cf064d269bb7049071722486c5351e))
* **functions:** Add the create_event MCP tool ([34fe503](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/34fe5037d07cd9c30d77f0b6f6ce45d3df59d675))
* **functions:** Add the delete_event MCP tool ([143b779](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/143b779ede73f11e4f0e4d6833bcdfacc359462f))
* **functions:** Add the update_draft MCP tool ([65f4e64](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/65f4e64e0b78ff190c4f59406dd01df2fd675a77))
* **functions:** Add the update_event MCP tool ([0df46ff](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/0df46ffd4ac2ac9bcb6e1691358cf50647d2865e))
* **functions:** Expose isHtml on create_draft and update_draft ([09d9569](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/09d9569ce2297d920499630660d27afb8f28fa9e))
* **functions:** Include the event's date/time in the delete_event preview ([84f4760](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/84f4760fc21c572d6d9c246a839f1e4e2fc80a2f))
* **functions:** Scaffold the OutlookWriteback Functions project ([a54b45c](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/a54b45c7a694b5c02af7f2fba6b50114b9c5efc1))
* **functions:** Tell the calling model to prompt a reconnect on auth failures ([ecf87d4](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/ecf87d40c6429d064b2c5ea054c54227ea30143c))
* **functions:** Wire KeyVaultRefreshTokenStore into the DI container ([91b023e](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/91b023e57ad8bc686d2f60fdae74feafdaa88e2f))
* **functions:** Wire the delete-confirmation services into DI ([dfd729d](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/dfd729d0053112f88fb5e78e69cada34f5b29a0e))
* **graph:** Add attendee support to create_event ([a694e41](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/a694e410d88cc1be49b74f9c334af6cafe61181c))
* **graph:** Add BuildUpdateDraftMessage payload mapping ([8fba94a](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/8fba94ae3cdd07c0ebcdea5c2b81a687fed64845))
* **graph:** Add BuildUpdateEvent payload mapping ([be63a88](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/be63a88a2ef63bb77e1e0d00ef848981b032d264))
* **graph:** Add DeleteConfirmationTokenService ([9b0ed90](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/9b0ed90ddbccb8628a642bb9be9b49ba30be7d39))
* **graph:** Add DeleteEventAsync ([d08172a](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/d08172a53151a24c2464bd708aa4a358eafb7227))
* **graph:** Add EventDeletionService for the two-step delete ([8487253](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/8487253d1fe827cced09dbc682cf0037c171f044))
* **graph:** Add GraphTokenEndpointClient ([997fc71](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/997fc712ee5591a071117457b423faf8477b413c))
* **graph:** Add IRefreshTokenStore for non-interactive auth ([b0d6998](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/b0d6998c34e9bbf82c185e354ec116d513acca47))
* **graph:** Add optional HTML body support to the draft methods ([58eaf4b](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/58eaf4bbea62ea763efb7768d13d245e195d4510))
* **graph:** Add OutlookGraphClient for draft and event writes ([815eefc](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/815eefc51f06504c58dfee2d052f1411a410e2d8))
* **graph:** Add SilentGraphCredential for non-interactive auth ([ba235cc](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/ba235cc5a6da08cb82c8e3c1271b7917cb2006de))
* **graph:** Add the CreateWithSilentRefreshAuth factory ([3f44fab](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/3f44fabacd2af4c68393409f1b17a06286c8f7e7))
* **graph:** Add UpdateDraftAsync ([b273e11](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/b273e11341c91fb89c40209e960505a3bbfd07d5))
* **graph:** Add UpdateEventAsync ([4b2c7f3](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/4b2c7f3f52b720592cf41cdba224a409d61dcabd))


### Bug Fixes

* **bootstrap:** Don't hardcode a domain in the sign-in console message ([73ab3cd](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/73ab3cd114bf54f5954833b395f6829ba6c48845))
* **functions:** Relax the MCP extension system-key gate to Anonymous ([313660f](https://github.com/keithmendozasr/sapidus-writeback-mcp/commit/313660f54d15cf5d9602220e9f5871be5f91d449))
