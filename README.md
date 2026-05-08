<p align="center">
  <img src="https://github.com/user-attachments/assets/cc82082d-c4a5-4f13-90f5-adaf162202b2" alt="image" />
</p>

# Network Sketcher

**Network Sketcher generates network configuration diagrams in PowerPoint and manages configuration information in Excel. With AI (LLM) integration, it supports network design creation and updates via an MCP server for LLM clients (Local MCP), a web browser (Online), or a desktop GUI/CLI (Offline).**

Network Sketcher provides three editions:

- [**Network Sketcher Local MCP**](#network-sketcher-local-mcp) — **AI-native MCP server for LLM clients (Cursor, Claude Code, etc.)**. The most direct AI integration: the LLM calls Network Sketcher tools without a browser or copy-paste.
- [**Network Sketcher Online**](#network-sketcher-online) — Browser-based web service.
- [**Network Sketcher Offline**](#network-sketcher-offline) — Desktop GUI + CLI. Runs independently with the `network-sketcher_offline/` folder alone.

You can use any combination.

| | **Local MCP (AI-native)** | Online (Web Service) | Offline (GUI + CLI) |
| --- | --- | --- | --- |
| Interface | **LLM client (Cursor, Claude Code, etc.)** | Web browser | Desktop GUI / Command-line |
| Key dependencies | **Python + MCP SDK** | Python + Flask | Python + tkinter |
| Multi-user | Single user | Multiple users via browser | Single user |
| Client requires | **Python + MCP client** | Web browser only | Python runtime environment |
| AI-native design | **Yes (most direct)** | Yes | No |
| Master format | **`.nsm` only (`.xlsx` via import/export)** | `.nsm` internally; `.xlsx` at boundary | `.xlsx` / `.nsm` both |
| Internal data storage | No | No | No |
| External communication | stdio to LLM client (local) | [HTTPS](#external-communication) | No |
| Tested platforms | Windows (Mac OS, Linux compatible by design) | Windows (Mac OS, Linux untested) | Windows, Mac OS, Linux |
| Folder | **`network-sketcher_local_mcp/`** | `network-sketcher_online/` | `network-sketcher_offline/` |

```
network-sketcher/
├── network-sketcher_local_mcp/  # Local MCP edition — MCP server for LLM clients (AI-native)
├── network-sketcher_online/     # Online edition — Web service (browser-based)
├── network-sketcher_offline/    # Offline edition — GUI + CLI (standalone desktop app)
├── README.md
├── LICENSE
└── ...
```

<br>
<br>

<p align="center">━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━</p>

<br>

# Network Sketcher Local MCP

> **AI-native edition:** Network Sketcher Local MCP exposes the engine as a **Model Context Protocol (MCP) server**, enabling LLM clients such as Cursor and Claude Code to drive network design directly. This is the most direct AI integration of the three editions — no browser, no copy-paste.

<img width="1456" height="782" alt="image" src="https://github.com/user-attachments/assets/6f3db3ba-4f92-401d-8f75-4c09cc0f185d" />

<img width="1211" height="653" alt="image" src="https://github.com/user-attachments/assets/5900104a-3a14-4de9-ba30-38af86f1da20" />


## What is Network Sketcher Local MCP?

Network Sketcher Local MCP is the third edition of Network Sketcher. It wraps `network-sketcher_online/ns_engine/` as a library and exposes the Network Sketcher CLI through the Model Context Protocol so that LLM clients can drive network design via Tool calls.

> **Positioning:** The Online edition is "browser + human + LLM (copy-paste)"; the Offline edition is "desktop GUI / CLI"; this edition is **"AI-native — the LLM executes the CLI directly"**.

## 5min Demo video of the Local MCP edition. 

The AI ​​agent Cursor autonomously creates a network using Network Sketcher's Local MCP functionality. It also simultaneously references best practices from other MCPs.

https://github.com/user-attachments/assets/274d5b66-5f4a-407a-bfb5-f71026971fc4




## Local MCP Features

- **No browser, no copy-paste.** The LLM calls `add device ...` and similar commands as Tool invocations
- Reuses `network-sketcher_online/ns_engine/` **as a library** (no code duplication)
- **No changes** are made to the existing `_online` / `_offline` folders
- stdio transport (designed for local operation)

## Limitations (Local MCP)

- Single-user edition designed to run on a local PC
- Only stdio transport is supported (HTTP/SSE not supported)
- Diagram generation for large networks may take some time
- LLM clients cannot directly view binary output (PPTX / SVG); if visual feedback is needed, the user should open the generated SVG directly
- **Verification status:** End-to-end verified in Cursor (the primary tested host). Claude Code support follows the published MCP specification but is design-validated only at this release; please report any issues you encounter.

## Requirement (Local MCP)

- Python 3.10 or later (required by the MCP SDK; the engine itself supports 3.9+)
- The full Network Sketcher repository (the `network-sketcher_online/` folder must be present)
- **Recommended LLM: Claude Opus 4.7 or later.** The Local MCP edition relies heavily on multi-step tool calling, schema interpretation, and adherence to the layout / workflow rules embedded in the server instructions and AI Context (e.g., RULE 0 / 0.5 layout, RULE 3.5 multi-transport WAN waypoint design, mandatory `get_workspace_info` to `get_ai_context` bootstrap). Weaker or older models may struggle with these workflows.

## Installation (Local MCP)

```bash
git clone https://github.com/cisco-open/network-sketcher/
cd network-sketcher/network-sketcher_local_mcp
python -m pip install -r requirements_mcp.txt
```

## Startup Check (Local MCP)

```bash
python ns_mcp_server.py
```

The server waits indefinitely on stdio — this is normal behaviour (Ctrl+C to exit). Logs are written to stderr.

## Configuration (`mcp_config.json`)

| Key | Description |
| --- | --- |
| `working_directory` | **Optional.** If set to a non-empty path, that path is used as the initial workspace on startup. When left empty (recommended), the server starts with no active workspace; the AI agent must call `suggest_workspace()` → `set_workspace(path)` before other Tools can be used. Leave empty for OS-agnostic portability. |
| `log_level` | `DEBUG` / `INFO` / `WARNING` / `ERROR` |
| `ai_context_show_commands` | List of show commands executed by the `get_network_state` Tool |
| `command_timeout_seconds` | (Reserved; not used in the current version) |

## Connection Examples

Cursor and Claude Code use different configuration mechanisms, so the setup steps are split below. Pick the one that matches your client.

### For Cursor

Add the following to the Cursor MCP configuration file (`File > Preferences > Cursor Settings > MCP` → `mcp.json`):

```json
{
  "mcpServers": {
    "network-sketcher": {
      "command": "python",
      "args": [
        "/path/to/network-sketcher/network-sketcher_local_mcp/ns_mcp_server.py"
      ]
    }
  }
}
```

Replace `/path/to/network-sketcher/` with the actual path where you cloned the repository.
On Windows, you can use either forward slashes (`/`) or escaped backslashes (`\\`).

### For Claude Code

Register the MCP server with the `claude` CLI. The `--` (double dash) separator is required so that the script path is passed to `python` rather than parsed as a flag of `claude mcp add`:

```bash
# Local scope (default; current project only, stored in ~/.claude.json)
claude mcp add network-sketcher -- python "/path/to/network-sketcher/network-sketcher_local_mcp/ns_mcp_server.py"

# User scope (available across all your projects)
claude mcp add --scope user network-sketcher -- python "/path/to/network-sketcher/network-sketcher_local_mcp/ns_mcp_server.py"

# Project scope (shared with team via .mcp.json in project root)
claude mcp add --scope project network-sketcher -- python "/path/to/network-sketcher/network-sketcher_local_mcp/ns_mcp_server.py"
```

Replace `/path/to/network-sketcher/` with the actual path where you cloned the repository. See the [Claude Code MCP installation scopes documentation](https://docs.claude.com/en/docs/claude-code/mcp#mcp-installation-scopes) for details on each scope and when to use which.

## Master File Format: `.nsm` Only

This edition handles only `.nsm` (ZIP + Apache Parquet) master files. `.xlsx` access via openpyxl is slow for large networks, so it is used only at the import/export boundary.

- To use an existing `.xlsx` master → convert with `import_master`
- To open in Excel or the Offline edition → write back with `export_master_xlsx`
- To create a new master → `create_empty_master` generates `.nsm` directly

## Exposed Tools

| Tool | Role |
| --- | --- |
| `suggest_workspace` | Returns candidate workspace directories suited to the current OS (Windows / macOS / Linux) |
| `set_workspace` | Sets the given path as the working directory for this session (restricted to paths under home; auto-creates if absent) |
| `get_workspace_info` | Returns the working directory and the list of master / output files it contains (returns `workspace_active=false` if no workspace is set) |
| `create_empty_master` | Creates an empty `[MASTER]<name>.nsm` (converts xlsx → nsm internally and discards the xlsx) |
| `import_master` | Converts an existing `.xlsx` to `.nsm` inside the working directory |
| `export_master_xlsx` | Writes a `.xlsx` from a `.nsm` (for editing in Excel or the Offline edition) |
| `get_network_state` | Runs the main `show` commands and returns aggregated results (lightweight) |
| `get_ai_context` | Generates `[AI_Context]<name>.txt` and returns its full content (full context) |
| `run_commands` | Executes multiple `add` / `rename` / `delete` / `show` commands in a single call |
| `export_diagram` | Generates an L1 / L2 / L3 diagram in SVG or PPTX format |

## Exposed Prompts

| Prompt | Role |
| --- | --- |
| `start_ns_session(master)` | Session start template. Forces the AI through the `get_workspace_info` → `get_ai_context` → summary report workflow. |

- In **Cursor**, launch with the `/start_ns_session` slash command.
- In **Claude Code**, launch with `/mcp__network-sketcher__start_ns_session` (per Claude Code's MCP prompt naming convention `/mcp__<server>__<prompt>`).

## Exposed Resources

| URI | Content |
| --- | --- |
| `nsm://workspace` | Current working directory state (JSON) |
| `nsm://commands` | Full CLI command reference (`nsm_extensions_cmd_list.txt`) |

## AI Context Guidance Mechanism

Three layers of guidance prevent the LLM from issuing commands without first understanding the current state:

1. **Server instructions (automatic):** `instructions` are sent to the MCP host at FastMCP initialisation. Cursor and Claude Code include them in the system prompt, so the AI is automatically prompted to call `get_workspace_info` → `get_ai_context` in every new session. The instructions also **mandate** that the AI call any other registered MCP server whose capability is relevant to the current sub-task before issuing Network Sketcher mutations (see Additional Policy below).
2. **`/start_ns_session` prompt (explicit launch):** When triggered by the slash command, a step-by-step message is inserted that walks the AI through the workflow. The first step requires the AI to enumerate other registered MCP servers and present a relevance mapping to the user.
3. **Tool docstring PREREQUISITE (fail-safe):** The `run_commands` and `export_diagram` docstrings explicitly state that `get_ai_context` must be called first — a final safety net for when the AI reads the docstring.

### Additional Policy: Coordinated Use of Other MCP Servers (Mandatory)

The server instructions and the `/start_ns_session` prompt include the following policy covering **all MCP servers** registered with the host — not just documentation/RAG servers, but any server regardless of vendor or category (configuration, monitoring, issue tracking, chat, repository, etc.).

- **Enforcement level:** "MANDATORY when relevant." Before starting a sub-task, the LLM must ask itself "Is any registered MCP server's capability relevant here?" If yes, it **must** call that server **before** issuing Network Sketcher mutations, and **must** cite the returned sources in its final answer.
- **Relevance judgment:** Based on each server's description (`serverUseInstructions` / tool descriptors), not its name alone.
- **Typical relevance mappings (non-exhaustive):**
  - Documentation / RAG servers → model selection, best practices, EoS/EoL, config examples, troubleshooting
  - Topology / config / monitoring servers → grounding decisions in live device state
  - Issue tracking / chat / repository servers → when user requests reference existing tickets, code, or history
- **Guardrails (prevent abuse):**
  - Do not call a server whose capability is clearly unrelated to the task
  - Authenticate MCP servers one at a time (no parallel authentication)
  - Keep to roughly **10 external MCP calls or fewer** per user turn
  - On failure, report to the user and continue with remaining sources rather than retrying
- **Scope:** This policy applies only to the Local MCP edition (Online / Offline editions are not accessed via an MCP host and are therefore out of scope). It is not included in the shared AI Context artifact (`[AI_Context]<name>.txt`).

## Workspace Selection (Cross-Platform)

This edition **does not hardcode OS-specific paths**. The AI agent detects the OS and proposes a suitable directory for the host.

### Default Behaviour (Recommended)

1. Set `working_directory` to empty (`""`) in `mcp_config.json`
2. At session start, the AI calls `suggest_workspace()` to retrieve OS-specific candidates:
   - Windows: `~/Documents/ns_workspace`, `~/Desktop/ns_workspace`, `~/ns_workspace`
   - macOS: same as Windows
   - Linux: `$XDG_DATA_HOME/ns_workspace` (or `~/.local/share/ns_workspace`), `~/Documents/ns_workspace`, `~/ns_workspace`
3. The AI proposes one candidate to the user and requests confirmation
4. After user approval, the AI calls `set_workspace(path)`
5. All Tools for the remainder of the session operate in that workspace

### Advanced Usage (Fixed Workspace)

If you want to share the same path across multiple hosts or always use a fixed directory, set an absolute path (or `~/...`) in `mcp_config.json` under `working_directory`. The workspace is then active immediately on startup, and there is no need to call `suggest_workspace` / `set_workspace`.

### Security Boundary

`set_workspace(path)` accepts only paths that satisfy all of the following:
- Resolves to a location under `Path.home()` (`~/` / `$HOME`)
- Is writable
- Is created automatically if it does not yet exist

Paths outside the home directory are rejected to prevent unintended access to the host system.

## Typical Usage Flow (Local MCP)

1. (First time only) Set `working_directory` to empty in `mcp_config.json`, or set a fixed path
2. In Cursor, run `/start_ns_session master='[MASTER]office.nsm'` (or instruct the LLM directly)
3. The LLM calls `suggest_workspace` → `set_workspace` to establish the workspace (if not already set)
4. The LLM calls `get_workspace_info` → `get_ai_context` to understand the current state and reports a summary
   - If no master exists, create one with `create_empty_master(filename='[MASTER]office.nsm')`, or convert an existing `.xlsx` with `import_master`
5. Give the LLM instructions such as "Add Core-SW and bind VLAN10 to the Vlan 10 SVI"
6. The LLM issues multiple commands via `run_commands`
7. Generate L1 / L2 / L3 diagrams with `export_diagram`
8. To open in Excel, call `export_master_xlsx('[MASTER]office.nsm')`

## Migrating Existing `.xlsx` Masters (Local MCP)

`.xlsx` masters created before setting up Local MCP can be imported with the following steps:

```text
# 1. Establish the workspace (if not already set)
suggest_workspace()
set_workspace('~/Documents/ns_workspace')

# 2. Import the existing master using its absolute path
import_master(
    xlsx_path = '/path/to/[MASTER]office100_v2.xlsx',   # absolute path for your OS
    target_name = '[MASTER]office100_v2.nsm'             # optional
)

# 3. The original .xlsx can be deleted manually (it cannot be used by run_commands)
```

## Safety Mechanisms (Local MCP)

- **Workspace boundary:** `set_workspace(path)` accepts only paths under `Path.home()` to prevent unintended access to the host system
- **Master path validation:** Access to files outside the working directory is rejected; `.nsm` extension is mandatory
- **`run_commands` allowlist:** Only `add` / `rename` / `delete` / `show` are permitted; `export` must go through dedicated Tools
- **Automatic `--accept-security-risk`:** Prevents `get_ai_context` from blocking on `input()`
- **Automatic `--master`:** Appended internally so the LLM cannot accidentally target an external master
- **`import_master` / `export_master_xlsx` output location:** Always fixed to the working directory
- **No overwrite of existing files:** `create_empty_master`, `import_master`, and `export_master_xlsx` each stop with an error if the output path already exists

## Architecture (Local MCP)

```text
+-----------------------+   stdio (JSON-RPC)   +----------------------+
| Cursor / Claude Code  | <------------------> | ns_mcp_server.py     |
| (MCP host)            |                      | (FastMCP)            |
+-----------------------+                      +----------+-----------+
                                                          | sys.path insert
                                                          v
                                          +----------------------------+
                                          | network-sketcher_online/   |
                                          |   ns_engine/               |
                                          |     nsm_adapter.run_cli()  |
                                          |     nsm_cli.ns_cli_run()   |
                                          |     ...                    |
                                          +-------------+--------------+
                                                        | file I/O
                                                        v
                                          +----------------------------+
                                          | <workspace>/               |
                                          | (set by set_workspace,     |
                                          |  always under user home)   |
                                          |   [MASTER]*.nsm            |
                                          |   [AI_Context]*.txt        |
                                          |   [Lx_DIAGRAM]*.svg/pptx   |
                                          +----------------------------+
```

## Troubleshooting (Local MCP)

| Symptom | Cause / Resolution |
| --- | --- |
| `[FATAL] ns_engine directory not found` at startup | The `network-sketcher_online/` folder is missing. Place it in the same parent directory as this folder. |
| `[FATAL] The "mcp" package is not installed` | Run `python -m pip install -r requirements_mcp.txt` |
| Server does not appear in Cursor's MCP panel | Verify the path in `mcp.json` is correct and that the Python executable (or virtual environment Python) is specified in `command` |
| `get_ai_context` returns `[ERROR] AI Context file was not generated` | Confirm the master filename starts with `[MASTER]` and ends with `.nsm` |
| `Invalid master filename ... Must start with '[MASTER]' and end with .nsm` | This edition is `.nsm` only. If you have an `.xlsx`, convert it with `import_master`. |
| `[ERROR] No workspace is active for this session` | Call `suggest_workspace()` to see candidates, then `set_workspace(path)` to activate one. |
| `[ERROR] For safety, the workspace must be under your home directory` | Only paths under your home directory are accepted. Use a symbolic link if you need to reference a path outside your home. |
| `run_commands` returns `[ERROR] Verb 'export' is not allowed` | `run_commands` does not accept `export`. Use `export_diagram` or `get_ai_context` instead. |

### Claude Code-specific environment variables

The following environment variables are recognised by Claude Code (not used by Cursor) and may help when running Local MCP under Claude Code on slower machines or with very large networks:

| Variable | Purpose |
| --- | --- |
| `MCP_TIMEOUT=30000` | Extends Claude Code's MCP server startup timeout to 30 seconds. The default is around 10 seconds and may be exceeded by the first import of `pandas` / `pyarrow` / `networkx` on slow disks. |
| `MAX_MCP_OUTPUT_TOKENS=50000` | Raises Claude Code's per-tool output warning threshold. The default is 10,000 tokens; `get_ai_context` on large masters can legitimately exceed it. |

Set them when launching Claude Code, e.g.:

```bash
MCP_TIMEOUT=30000 MAX_MCP_OUTPUT_TOKENS=50000 claude
```

These knobs are documented in the [Claude Code MCP docs](https://docs.claude.com/en/docs/claude-code/mcp).

<br>
<br>

<p align="center">━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━</p>

<br>

# Network Sketcher Online

> **AI-native software:** Network Sketcher Online is designed around AI (LLM) interaction — generate AI context, send it to an LLM, and paste the resulting commands back to update your network design, all within the browser.


<img width="1559" height="879" alt="image" src="https://github.com/user-attachments/assets/0b87904c-f5b3-4585-a7bc-5b3b1e62c610" />


### Demo Video (Ver 3.0.1b)

A demo video of approximately 4 minutes, starting with the installation of Network Sketcher. This demo video demonstrates creating a network configuration using LLM from URL information and performing additional editing. No sound, no captions.


https://github.com/user-attachments/assets/2acaea3b-32f2-4ff0-90ad-a3dc810293d2



## What is Network Sketcher Online?

Network Sketcher Online is a browser-based web service. It wraps the Network Sketcher CLI and provides an intuitive web UI for diagram generation and AI-driven network design — no python on PCs required.

## Online Features

Network Sketcher Online supports two output modes:

- **SVG Mode** (default): Diagrams are rendered as SVG in the browser. All diagrams (L1/L2/L3, all areas and per-area) are generated in parallel and displayed as thumbnails without needing to download individual files. Master files are stored in the high-performance `.nsm` format internally. SVG mode is approximately 30x faster than PPTX mode. For compatibility with the Offline edition, master files can also be downloaded in `.xlsx` format.
- **PPTX Mode**: Diagrams are generated as PowerPoint (.pptx) files, and device files are generated as Excel (.xlsx) files. This is the original output mode and produces the same files as the Offline edition.

- **All uploaded and generated files are automatically deleted from Network Sketcher Online after the session ends — no data is retained on the server.**
- Upload master files via drag-and-drop in a web browser
- Generate L1/L2/L3 diagrams, device files, and AI context files with selectable outputs
- In-browser preview for PowerPoint (.pptx) and Excel (.xlsx) files without requiring Office software
- Copy AI context to clipboard and open LLM with one click
- Describe desired changes in a prompt field; the AI context + prompt is copied together for LLM interaction
- Update master files by pasting CLI commands from LLM output (Update Master)
- Create new master files from scratch via the web UI
- Parallel processing for faster diagram generation
- Attribute-based device coloring for L1 and L3 diagrams
- Session persistence across browser refreshes with automatic cleanup
- HTTPS enforced with auto-generated self-signed certificates
- All settings managed via `ns_web_config.json`

## Limitations (Online)
- Designed for use on internal networks only. Not intended for deployment on the public internet.
- IPv4 only. IPv6 is not supported.
- Excel (.xlsx) and PowerPoint (.pptx) files cannot be synced back to the master file. All editing is performed through CLI commands generated by an LLM.
- In-browser preview for PowerPoint (.pptx) may not render correctly when the browser zoom or display scaling is set to a high value (e.g., 200%).

## Requirement (Online)
- __Tested on Windows only.__ It may work on Mac OS and Linux, but these platforms have not been verified.
- Other requirements are the same as [Network Sketcher Offline](#requirement-offline).

## Installation (Online)
```bash
git clone https://github.com/cisco-open/network-sketcher/
cd network-sketcher/network-sketcher_online
python3 -m pip install -r requirements_online.txt
cd ..
python3 start_ns_online.py
```

Open the URL shown at startup (default: `https://localhost:5443`) in your browser.

- To serve on a specific network interface, edit the `host` and `port` settings in `ns_web_config.json` before starting the server.
- If no SSL certificate exists, a self-signed certificate is auto-generated on first startup.
- If you change the host IP address or other settings in `ns_web_config.json`, manually delete the SSL certificate files in the `Certs/` folder and restart the server. A new certificate matching the updated settings will be auto-generated.
- If the `fqdn` setting in `ns_web_config.json` is configured, the auto-generated SSL certificate's Common Name (CN) will use the specified FQDN.

## User Guide (Online)
Click the <kbd>?</kbd> icons on the web page to view contextual help for each section and feature.
For features not covered in this User Guide, please refer to the [User Guide (Offline)](#user-guide-offline).



https://github.com/user-attachments/assets/d376af22-e100-4647-83aa-e28cd7efefd8




https://github.com/user-attachments/assets/1e506529-1824-43a4-880c-196986d6dae8


[NS_Online_User_Guide_301_en.pdf](https://github.com/user-attachments/files/26778493/NS_Online_User_Guide_301_en.pdf)



[NS_Online_User_Guide_301_jp.pdf](https://github.com/user-attachments/files/26567922/NS_Online_User_Guide_301_jp.pdf)


### Server Management Scripts

| Script | Description |
| --- | --- |
| `python3 start_ns_online.py` | Start `ns_web_start.py` as a background process. Any already-running instance (including those started manually) is stopped first. Output is logged to `logs/server.log`. |
| `python3 stop_ns_online.py` | Stop all running `ns_web_start.py` processes, including those started outside of this script. |

Both scripts work on Windows, Mac OS, and Linux.

### External Communication

| Source | Target | Protocol |
| --- | --- | --- |
| Client PC | NS Online | HTTPS |
| Client PC | Configured LLM | HTTPS |

### Third-Party Libraries (Online)

Network Sketcher Online includes the following third-party JavaScript libraries for in-browser file preview. These are bundled in `network-sketcher_online/static/` and require no additional installation.

| Library | Version | License | Purpose |
| --- | --- | --- | --- |
| [PptxViewJS](https://github.com/gptsci/pptxviewjs) | 1.1.0 | MIT | PowerPoint (.pptx) in-browser preview |
| [Chart.js](https://www.chartjs.org/) | 4.4.8 | MIT | Chart rendering (PptxViewJS dependency) |
| [JSZip](https://stuk.github.io/jszip/) | 3.10.1 | MIT or GPLv3 | ZIP / Office file parsing |
| [SheetJS (xlsx)](https://sheetjs.com/) | 1.15.0 | Apache-2.0 | Excel (.xlsx) in-browser preview |

## Performance Measurement Summary (Online)

| Ver: 3.0.1 (SVG Mode)                                                          | 64 NW devices<br>112 Connections | 256 NW devices<br>480 Connections | 1024 NW devices<br>1984 Connections | 4096 NW devices<br>8064 Connections |
|--------------------------------------------------------------------------------|------------------:|-------------------:|--------------------:|--------------------:|
| Master file creation *1                                                        | 4s                | 4s                 | 26s                 | 8m 12s              |
| Creation of all configuration diagrams, device tables, and AI Context files    | 3s                | 7s                 | 58s                 | 15m 42s             |

---
*1 Reflect only L1 information in the no_data master file. Connect adjacent devices. Measure command execution time.<br>
Test environment: Intel Core Ultra 7 (1.70 GHz), 32.0 GB RAM, Windows 11 Enterprise <br>

<br>
<br>

<p align="center">━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━</p>

<br>

# Network Sketcher Offline

## Overview

https://github.com/user-attachments/assets/9ff207f8-c6b3-4584-b166-98ae4e4c8297

NoLang (no-lang.com)
Otologic (https://otologic.jp) CC BY 4.0

## Demo video of basic usage
https://github.com/cisco-open/network-sketcher/assets/13013736/b76ec8fa-44ad-4d02-a7c2-579f67ad24a9

## AI(LLM) utilization demo video
[Full_Version_link(Youtube)](https://www.youtube.com/watch?v=g5N3yg0jMSg)

https://github.com/user-attachments/assets/5874411a-0e6d-485d-9f85-4cdc85f3ca07




## Concept
**Network Sketcher generates network configuration diagrams in PowerPoint and manages configuration information in Excel. Additionally, exporting a AI ​​context can be used to generate config files using LLM.**
* Automatic generation of each configuration document by metadatization of network configuration information
* Automated synchronization between documents
* Minimize maintenance and training load by automatic generation of common formats
* Facilitate automatic analysis, AI utilization, and inter-system collaboration by metadatization of configuration information.
* Template support for equipment configuration
![image](https://github.com/user-attachments/assets/9f497061-08ee-4c78-9040-d5b37d2f3e69)

![image](https://github.com/cisco-open/network-sketcher/assets/13013736/240ddee0-823d-472f-87d4-8ae7eb1fff7d)


## New Features (Offline)
- Ver 2.6.1<br>
[Network Sketcher Ver 2.6.1 supported the creation of a network configuration with LLM from scratch](https://github.com/cisco-open/network-sketcher/wiki/1%E2%80%905.-Examples-of-General%E2%80%90Purpose-AI-(LLM)-usage-(config-creation,-config-reflection,-analysis,-etc.)_en)

<img width="1423" height="806" alt="image" src="https://github.com/user-attachments/assets/761072de-d64b-4772-bdc7-6224f53fddd8" />


- Ver 2.6.0<br>

[Network Sketcher Ver 2.6.0 now supports master file conversion from Visio, draw.io, NetBox, and CML data to Network Sketcher.](https://github.com/cisco-open/network-sketcher/wiki/1%E2%80%904.-Convert-data-from-other-systems-into-master-files-(Visio,-draw.io,-NetBox-and-CML)_en)

<img alt="image" src="https://github.com/user-attachments/assets/436a1462-bdf7-49cf-bc4f-235be6cb7d42" />
Although Network Sketcher now supports multiple formats, it is not intended to replace the main drawing tool, but rather aims for mutually beneficial development.

    
- Ver 2.5.0
  - [Communication flow management functionality has been added.](https://github.com/cisco-open/network-sketcher/wiki/9%E2%80%901.Exporting-Flow-files)
![image](https://github.com/user-attachments/assets/8683c172-505e-4af8-a87a-dc1a1a86a121)

## Limitations (Offline)
- IPv4 only. IPv6 is not supported.
- A DEVICE file contains multiple sheets, but only one sheet should be updated at a time. Simultaneous synchronization of multiple sheet updates is not supported.
- Do not use Network Sketcher on master files in your One Drive folder.
- Deleting Layer 1 links using the GUI cannot identify individual interfaces and will delete more Layer 2 data than intended. Use the CLI command (delete l1_link) to delete Layer 1 links.
 
## Requirement (Offline)
- __Network Sketcher supports cross-platform. Works with Windows, Mac OS, and Linux.__
  - MAC OS may not display well in Dark mode.
- __Python ver 3.x__
- __Software that can edit .pptx and .xlsx files__
  - Microsoft Powerpoint and Excel are the best
  - Google Slides and Spreadsheets import/export functionality is available. Excel functions display will show an error, but it works fine.
  - Libre Office and Softmaker office cannot be used.

## Installation (Offline)
```bash
git clone https://github.com/cisco-open/network-sketcher/
cd network-sketcher/network-sketcher_offline
python3 -m pip install -r requirements_offline.txt
python3 network_sketcher.py
```
or
```bash
#Download via browser
https://github.com/cisco-open/network-sketcher/archive/refs/heads/main.zip

#Unzip the ZIP file and execute the following in the prompt of the folder
cd network-sketcher_offline
python3 -m pip install -r requirements_offline.txt
python3 network_sketcher.py
```

### Installation Supplement (Offline)
 * Alternative to "python -m pip install -r requirements_offline.txt"
```bash
python3 -m pip install tkinterdnd2
python3 -m pip install "openpyxl>=3.1.3,<=3.1.5"
python3 -m pip install python-pptx
python3 -m pip install ipaddress
python3 -m pip install numpy
python3 -m pip install pyyaml
python3 -m pip install ciscoconfparse
python3 -m pip install networkx
python3 -m pip install svg.path
```

* Mac OS requires the following additional installation.
```bash
brew install tcl-tk
brew install tkdnd
```
* Ubuntu requires the following additional installation.<br>
  GUI drag and drop doesn't work on Ubuntu, you need to compile tkdnd from source or use "Browse" and "Submit".
```bash
sudo apt-get install python3-tk
```

## User Guide (Offline)
| Language  | Link |
| ------------- | ------------- |
| English  | [Link](https://github.com/cisco-open/network-sketcher/wiki/User_Guide(Offline_Edition)%5BEN%5D) |
| Japanese  | [Link](https://github.com/cisco-open/network-sketcher/wiki/User_Guide(Offline_Edition)%5BJP%5D) |
<br>
 
## How to create the exe file for Windows using pyinstaller
 ```bash
pyinstaller.exe [file path]/network-sketcher_offline/network_sketcher.py --onefile --collect-data tkinterdnd2 --additional-hooks-dir  [file path] --clean --add-data "./ns_extensions_cmd_list.txt;." --add-data "./ns_logo.png;."
 ```
<br>

## Performance Measurement Summary (Offline)

| Ver: 2.6.1b                                         | 64 NW devices<br>112 Connections| 256 NW devices<br>480 Connections | 1024 NW devices<br>1984 Connections|
|----------------------------------------------------|-----------:|------------:|-------------:|
| Master file creation *1                            | 51s      | 2m45s      | 25m45s          |
| Layer 1 diagram generation (All Areas with tags)   | 6s         | 29s         | 6m30s         |
| Layer 2 diagram generation                         | 13s        | 51s       | 6m53s          |
| Layer 3 diagram generation (All Areas)             | 10s        | 56s       | 14m23s         |
| Device file export                                 | 19s        | 1m4s         | 5m14s          |

---
*1 Reflect only L1 information in the no_data master file. Connect adjacent devices. Measure command execution time.<br>
Test environment: Intel Core Ultra 7 (1.70 GHz), 32.0 GB RAM, Windows 11 Enterprise <br>

<br>

<p align="center">━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━</p>

<br>

# Appendix

## Feature Support Matrix

| Feature Item | **Local MCP (LLM-driven CLI)** | Online Edition | Offline Edition (GUI) | Offline Edition (CLI) |
| --- | --- | --- | --- | --- |
| Create master file from PowerPoint rough sketch | ❌ | ❌ | ✅ | ❌ |
| Convert master files from Visio, Draw.io, NetBox, CML | ❌ | ❌ | ✅ | ❌ |
| Area placement | ✅ (user-specified) | ✅ (user-specified) | ✅ (automatic) | ✅ (user-specified) |
| Create / delete / modify areas | ✅ | ✅ | ✅ | ✅ |
| Place / create / delete / modify devices | ✅ | ✅ | ✅ | ✅ |
| Place / create / delete / modify waypoints | ✅ | ✅ | ✅ | ✅ |
| Add Layer 1 connections | ✅ | ✅ | ✅ | ✅ |
| Delete Layer 1 connections | ✅ | ✅ | ⚠️ (port cannot be specified) | ✅ |
| Change Layer 1 port names | ✅ | ✅ | ✅ | ✅ |
| Change Layer 1 connection details (e.g., duplex) | ✅ | ✅ | ✅ | ✅ |
| Change Layer 2 segments (VLAN) | ✅ | ✅ | ✅ | ✅ |
| Add / delete virtual ports (SVI, loopback, port-channel) | ✅ | ✅ | ✅ | ✅ |
| Change IP addresses / Layer 3 instances (VRF) | ✅ | ✅ | ✅ | ✅ |
| Change attributes | ✅ | ✅ | ✅ | ✅ |
| Add / delete VPNs | ❌ | ❌ | ✅ | ❌ |
| Flow management | ❌ | ❌ | ✅ | ❌ |
| Export various reports | ❌ | ❌ | ✅ | ❌ |
| Export empty master files (no data) | ✅ | ✅ | ❌ | ✅ |
| Export AI context files | ✅ | ✅ | ✅ | ✅ |
| Export device files | ❌ | ✅ | ✅ | ✅ |
| Generate L1/L2/L3 topology diagrams | ✅ | ✅ | ✅ | ✅ |

## SAMPLE
### - Supports various connections
<img alt="image" src="https://github.com/user-attachments/assets/752a5a6d-fcb8-4bf2-a709-91c4c8f862c5" />

Download : [Sample.figure5.zip](https://github.com/user-attachments/files/24335488/Sample.figure5.zip)

### - Wi-Fi office
Created by using AI context and giving AI (LLM) multiple command generation instructions.
<img  alt="image" src="https://github.com/user-attachments/assets/36ca6a28-e0a6-4e3a-94e1-241bd74a86f0" />

Download : [Sample Office.zip](https://github.com/user-attachments/files/24340917/Sample.Office.zip)

---
<sub>Otologic (https://otologic.jp) CC BY 4.0</sub>

# Author
 
* Yusuke Ogawa - Architect, Cisco | CCIE#17583
 
# License
SPDX-License-Identifier: Apache-2.0

Copyright 2023  Cisco Systems, Inc. and its affiliates

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
