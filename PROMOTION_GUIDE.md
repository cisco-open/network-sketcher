# Network Sketcher Promotion Guide

This guide collects reusable copy and action items for improving the visibility of Network Sketcher Local MCP across MCP directories, GitHub, demo channels, and community posts.

## Positioning

Primary message:

> Network Sketcher is a local-first MCP server that lets AI agents design Cisco-style networks and generate L1/L2/L3 topology diagrams, device tables, and AI-ready context files.

Short variants:

- AI-native network diagramming MCP server
- Local MCP server for Cisco-style network design automation
- Generate L1/L2/L3 topology diagrams from natural-language workflows
- Network design assistant for Cursor, Claude Code, and other MCP clients

Recommended keywords:

- `mcp-server`
- `model-context-protocol`
- `network-automation`
- `network-diagram`
- `cisco`
- `topology`
- `l1-l2-l3`
- `cursor`
- `claude-code`
- `ai-agent`
- `svg`
- `powerpoint`

## Directory Listing Audit

Checked from public search results on 2026-07-01.

- Remote OpenClaw: listed as `Network Sketcher (Local MCP)` in the AI & ML category. Reported rank: 31. Reported stars: 366.
- LobeHub MCP Marketplace: already listed at `https://lobehub.com/mcp/cisco-open-network-sketcher`.
- MCP.Directory: no direct server listing found for `cisco-open/network-sketcher`; related Cisco/network diagram skill results appear instead.
- MCP Toplist: no direct listing found for `cisco-open/network-sketcher` in public search results.
- Glama: no direct Network Sketcher listing found; related networking/Cisco simulation MCPs appear.
- Smithery: no direct Network Sketcher server listing found; related network diagram skills appear.
- mcp.so: no direct Network Sketcher listing found in public search results.

Recommended follow-up:

- Submit or refresh Network Sketcher where the directory accepts GitHub repository metadata.
- Prefer categories such as `Developer Tools`, `Infrastructure`, `Networking`, `AI & ML`, `Visualization`, and `Automation`.
- Use the same short description across directories so search engines learn a consistent summary.

## Directory Submission Copy

Title:

```text
Network Sketcher (Local MCP)
```

One-line description:

```text
Local-first MCP server that lets AI agents design Cisco-style networks and generate L1/L2/L3 topology diagrams, device tables, and AI-ready context files.
```

Long description:

```text
Network Sketcher Local MCP exposes a network design and diagramming engine through the Model Context Protocol. It lets MCP clients such as Cursor and Claude Code create and edit network master files, generate L1 physical, L2 VLAN/broadcast-domain, and L3 IP topology diagrams, export combined HTML viewers and device tables, and produce AI Context files for design review or continued editing. The server runs locally over stdio and keeps user master files on the local machine.
```

Suggested categories:

- Developer Tools
- Networking
- Infrastructure
- Visualization
- AI Agents
- Automation

Suggested tags:

```text
mcp-server, model-context-protocol, network-automation, network-diagram, cisco, topology, svg, powerpoint, cursor, claude-code
```

Repository:

```text
https://github.com/cisco-open/network-sketcher
```

## Awesome MCP PR Drafts

### punkpeye/awesome-mcp-servers

Suggested section:

```text
Developer Tools / Infrastructure / Visualization
```

Suggested entry:

```markdown
- [Network Sketcher](https://github.com/cisco-open/network-sketcher) 🐍 🏠 🪟 🍎 🐧 - Local-first MCP server for AI-driven network design. Lets Cursor, Claude Code, and other MCP clients create Cisco-style network designs, generate L1/L2/L3 topology diagrams, device tables, and AI Context files.
```

Suggested PR title:

```text
Add Network Sketcher Local MCP
```

Suggested PR body:

```markdown
## Summary

Adds Network Sketcher, a local-first MCP server for AI-driven network design and diagram generation.

## Why it fits

Network Sketcher exposes network design automation through MCP. It lets LLM clients such as Cursor and Claude Code create Cisco-style network master files, generate L1/L2/L3 topology diagrams, export device tables, and produce AI Context files for review or follow-up editing.

## Repository

https://github.com/cisco-open/network-sketcher
```

### mcpHQ/awesome-mcp-servers

Suggested category:

```text
Cloud and Infrastructure
```

Suggested entry:

```json
{
  "name": "Network Sketcher",
  "url": "https://github.com/cisco-open/network-sketcher",
  "description": "Local-first MCP server for AI-driven network design, L1/L2/L3 topology diagrams, device tables, and AI Context generation.",
  "language": "Python",
  "categories": ["Cloud and Infrastructure", "Developer Tools", "AI Agents and Memory"]
}
```

Adjust the exact JSON fields to match the target repository's contribution schema.

## One-Minute Demo Storyboard

Goal: show the value in under 60 seconds without requiring viewers to understand the whole product.

Scene 1: Problem, 0-5 seconds

- Show a blank Network Sketcher workspace or empty master.
- Caption: `Designing network diagrams manually is slow.`

Scene 2: Agent prompt, 5-15 seconds

- In Cursor or Claude Code, ask:

```text
Create a 5-site WAN design with HQ, two data centers, two branches, Internet and WAN waypoints, edge routers, L2 segments, and IP addressing. Then generate L1/L2/L3 diagrams and a device table.
```

Scene 3: Tool calls, 15-30 seconds

- Show the agent calling Local MCP tools:
  - `get_workspace_info`
  - `create_empty_master`
  - `run_commands`
  - `build_default_outputs`

Scene 4: Output, 30-50 seconds

- Open the generated combined HTML viewer.
- Switch tabs:
  - L1 physical topology
  - L2 VLAN/broadcast-domain diagram
  - L3 IP topology diagram
- Open the device table preview.

Scene 5: Close, 50-60 seconds

- Caption:

```text
Network Sketcher Local MCP: AI-native network design, diagrams, and device tables from one conversation.
```

Reusable demo prompt:

```text
Using Network Sketcher Local MCP, create a small 5-site enterprise WAN with HQ, DC-1, DC-2, Branch-1, and Branch-2. Include Internet and WAN waypoints, edge routers in each area, simple L2 segments, and representative IP addressing. Build the default outputs and summarize what was created.
```

Recommended assets:

- `demo/5site-wan/README.md`
- `demo/5site-wan/prompt.txt`
- `demo/5site-wan/screenshots/l1.png`
- `demo/5site-wan/screenshots/l2.png`
- `demo/5site-wan/screenshots/l3.png`
- `demo/5site-wan/screenshots/device-table.png`
- 60-second GIF or MP4 linked from the top of `README.md`

## Launch Post Drafts

### X / Twitter

Short:

```text
Network Sketcher Local MCP lets AI agents design Cisco-style networks and generate L1/L2/L3 topology diagrams, device tables, and AI Context files locally.

Works with Cursor / Claude Code via MCP.

Repo: https://github.com/cisco-open/network-sketcher
```

Thread:

```text
1/ I’ve been working on Network Sketcher Local MCP: a local-first MCP server for AI-driven network design.

2/ Ask an agent to create a WAN or campus LAN design, and it can build the master file, add devices and links, then generate L1/L2/L3 topology diagrams.

3/ Outputs include SVG / PowerPoint diagrams, a combined L1/L2/L3 HTML viewer, device tables, and an AI Context file for further review.

4/ It runs locally over stdio with Cursor / Claude Code. Master files stay on your machine.

5/ Repo: https://github.com/cisco-open/network-sketcher
```

### LinkedIn

```text
Network Sketcher Local MCP is now available for AI-driven network design workflows.

It exposes the Network Sketcher engine as a local Model Context Protocol server, so tools like Cursor and Claude Code can create and update network designs through tool calls.

What it can generate:
- L1 physical topology diagrams
- L2 VLAN / broadcast-domain diagrams
- L3 IP topology diagrams
- Combined L1/L2/L3 HTML viewers
- Device tables
- AI Context files for review or follow-up editing

The goal is to make network design artifacts easier to create, inspect, and iterate from a single AI-assisted workflow.

GitHub: https://github.com/cisco-open/network-sketcher
```

### Reddit

Suggested communities:

- `r/networking`
- `r/networkautomation`
- `r/ClaudeAI`
- `r/Cursor`
- `r/LocalLLaMA`

Post draft:

```text
I built a local MCP server for AI-driven network diagram generation

Network Sketcher Local MCP lets an AI agent create network master data and generate L1/L2/L3 topology diagrams, device tables, and an AI Context file from a conversation.

It runs locally over stdio and is designed for clients such as Cursor and Claude Code. The current focus is Cisco-style network design workflows: WAN/campus layouts, areas, devices, links, VLAN/broadcast-domain views, L3 IP diagrams, and generated HTML/SVG/PowerPoint outputs.

Repository:
https://github.com/cisco-open/network-sketcher

I’m interested in feedback from network engineers on what design workflows would be most useful to automate next.
```

### Hacker News Show HN

```text
Show HN: Network Sketcher Local MCP – AI-driven L1/L2/L3 network diagrams

Network Sketcher Local MCP is a local-first Model Context Protocol server that lets AI agents create and update network design data, then generate L1/L2/L3 topology diagrams, device tables, and AI Context files.

It is designed for Cursor, Claude Code, and other MCP clients. The server runs locally over stdio and produces SVG / PowerPoint diagrams plus a combined HTML viewer.

GitHub: https://github.com/cisco-open/network-sketcher
```

## GitHub Repository Checklist

- Add repository topics:
  - `mcp-server`
  - `model-context-protocol`
  - `network-automation`
  - `network-diagram`
  - `cisco`
  - `topology`
  - `cursor`
  - `claude-code`
  - `ai-agent`
  - `svg`
  - `powerpoint`
- Keep README first-view focused on Local MCP, use cases, and output examples.
- Keep the LobeHub badge near the top.
- Add a Remote OpenClaw ranking note only if periodically maintained.
- Add a short call to action asking users to star the repository if the tool helps their workflow.

## Success Metrics

- GitHub stars: 366 to 400, then 500
- Remote OpenClaw rank: 31 to top 30
- Number of MCP directories with an accurate Network Sketcher listing
- Number of marketplace ratings or comments
- Demo video views and README click-throughs
- Issues or discussions from new users trying Local MCP
