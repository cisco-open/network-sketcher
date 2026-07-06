# Publishing on Glama — build spec & release runbook

How **Network Sketcher Local MCP** is listed and released on the
[Glama MCP directory](https://glama.ai/mcp/servers/cisco-open/network-sketcher).

A **Glama release is not a GitHub release.** Glama builds a container, boots the MCP
server over stdio, runs introspection (`initialize` + `tools/list`), security checks,
and assigns a quality grade. The score badge on awesome-mcp-servers PRs requires this
release to be published.

## In the repo

| File | Purpose |
| --- | --- |
| [`glama.json`](../glama.json) | Maintainer claim metadata (`yuhsukeogawa`). |
| [`Dockerfile`](../Dockerfile) | Runnable stdio image for local/container use. Glama admin can also build from source via the form below. |
| [`network-sketcher_local_mcp/server.json`](../network-sketcher_local_mcp/server.json) | MCP server manifest (tools, install hints). |

## Release steps (Glama admin UI)

1. Open https://glama.ai/mcp/servers/cisco-open/network-sketcher/score and **Claim** the server (if not already claimed).
2. Click **Sync Server** so Glama mirrors the latest `main` commit.
3. Open `…/admin/dockerfile` and configure the build form (Glama clones the repo into `/app` and wraps CMD with `mcp-proxy --`).
4. **Deploy** → wait for the build test to pass (server starts + introspection OK).
5. **Make Release** → enter version (e.g. `3.1.2m`) → publish.
6. Optional: use **Try in Browser** once on the server page to seed tool usage for the quality checklist.

### Form values for Network Sketcher

| Field | Value |
| --- | --- |
| Base image | `debian:trixie-slim` (default) or `python:3.12-slim` |
| Python version | `3.12` (minimum 3.10) |
| **Build steps** | `["python -m pip install -r network-sketcher_local_mcp/requirements_mcp.txt"]` |
| **CMD arguments** | `["python", "network-sketcher_local_mcp/ns_mcp_server.py"]` |
| Environment variables JSON schema | `{"properties":{},"required":[],"type":"object"}` |
| Placeholder parameters | `{}` |
| Pinned commit SHA | empty (after Sync Server) |

The effective container command becomes:

`mcp-proxy -- python network-sketcher_local_mcp/ns_mcp_server.py`

## Local smoke test

```bash
docker build -t network-sketcher-mcp .
docker run --rm -i network-sketcher-mcp
```

The process should stay alive on stdio with logs on stderr. Stop with Ctrl+C.

## After release

- Confirm the [score badge](https://glama.ai/mcp/servers/cisco-open/network-sketcher/badges/score.svg) shows a letter grade (not `?`).
- Reply on [punkpeye/awesome-mcp-servers PR #9021](https://github.com/punkpeye/awesome-mcp-servers/pull/9021) when the grade is visible.
