# Publishing on LobeHub — manifest & score runbook

How **Network Sketcher Local MCP** is updated on the
[LobeHub MCP Marketplace](https://lobehub.com/mcp/cisco-open-network-sketcher).

Unlike Glama, LobeHub does **not** require a hosted Docker release. Quality score
improves when the **owner publishes** an updated manifest (`lhm.plugin.json`) and
the listing shows **Validated** with current tools / prompts / resources.

## In the repo

| File | Purpose |
| --- | --- |
| [`network-sketcher_local_mcp/lhm.plugin.json`](../network-sketcher_local_mcp/lhm.plugin.json) | LobeHub marketplace manifest (tools, prompts, resources, i18n). |
| [`network-sketcher_local_mcp/server.json`](../network-sketcher_local_mcp/server.json) | Official MCP Registry manifest. |

Official MCP Registry (active):  
https://registry.modelcontextprotocol.io/v0/servers/io.github.cisco-open%2Fnetwork-sketcher

## One-time setup (browser required)

```bash
npx -y @lobehub/market-cli login
npx -y @lobehub/market-cli github connect
npx -y @lobehub/market-cli plugin claim cisco-open-network-sketcher
```

> **Org repo:** `cisco-open/network-sketcher` must be claimable by a GitHub user
> with appropriate `cisco-open` org access. If `plugin claim` fails, ask an org
> admin to run the publish steps.

## Publish a new version

```bash
cd network-sketcher/network-sketcher_local_mcp
npx -y @lobehub/market-cli plugin publish --dir "$(pwd)"
```

Bump `version` in `lhm.plugin.json` for each release (e.g. `3.1.2m`).

## After publish

1. Open the [Score tab](https://lobehub.com/mcp/cisco-open-network-sketcher?activeTab=score).
2. Click **Refresh metadata** if shown.
3. Confirm **Unvalidated** is gone and version matches `lhm.plugin.json`.
4. Optional: encourage ratings via README `market-cli mcp comment` snippet.

## Verify

```bash
npx -y @lobehub/market-cli plugin list --output json
```

Full CLI reference: https://market.lobehub.com/s/publish-mcp
