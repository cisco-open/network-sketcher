<p align="center">
  <img src="https://github.com/user-attachments/assets/cc82082d-c4a5-4f13-90f5-adaf162202b2" alt="image" />
</p>

<p align="center">
  <a href="https://lobehub.com/mcp/cisco-open-network-sketcher">
    <img src="https://lobehub.com/badge/mcp/cisco-open-network-sketcher" alt="LobeHub MCP Badge" />
  </a>
</p>

# Network Sketcher

**Network Sketcher generates network configuration diagrams in PowerPoint and manages configuration information in Excel. With AI (LLM) integration, it supports network design creation and updates via an MCP server for LLM clients (Local MCP), a web browser (Online), or a desktop GUI/CLI (Offline).**

<img width="1848" height="1028" alt="image" src="https://github.com/user-attachments/assets/26068524-6293-4f7f-ab0c-f6b7e2c8b842" />

<img width="1852" height="1039" alt="image" src="https://github.com/user-attachments/assets/b3501923-195e-45bc-9120-f6b78396e300" />



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
| Master format | **`.nsm` only** | `.xlsx` / `.nsm` both | `.xlsx` only |
| Internal data storage | No | No | No |
| External communication | stdio to LLM client (local) | [HTTPS](https://github.com/cisco-open/network-sketcher/wiki/User_Guide(Online_Edition)%5BEN%5D#external-communication) | No |
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

## Find on LobeHub MCP Marketplace

Network Sketcher Local MCP is listed on the [LobeHub MCP Plugins Marketplace](https://lobehub.com/mcp/cisco-open-network-sketcher).

If you find the server useful and you trust the official `@lobehub/` npm scope, please leave a rating from your MCP host CLI (the version range below is intentionally pinned to mitigate supply-chain risk):

```bash
npx @lobehub/market-cli@^0.0 mcp comment cisco-open-network-sketcher \
  -c "Used to drive Cisco network design from Cursor / Claude Code." --rating 5
```

## User Guide (Local MCP)
| Language  | Link |
| ------------- | ------------- |
| English  | [Link](https://github.com/cisco-open/network-sketcher/wiki/User_Guide(Local_MCP_Edition)%5BEN%5D) |
| Japanese  | [Link](https://github.com/cisco-open/network-sketcher/wiki/User_Guide(Local_MCP_Edition)%5BJP%5D) |

<br>
<br>

<p align="center">━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━</p>

<br>

# Network Sketcher Online

> **AI-native software:** Network Sketcher Online is designed around AI (LLM) interaction — generate AI context, send it to an LLM, and paste the resulting commands back to update your network design, all within the browser.


<img width="1620" height="898" alt="image" src="https://github.com/user-attachments/assets/cd645a8b-9661-4f74-8bf3-cd12b3395c82" />


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
- **Multi-format diagram download**: SVG / SVG (for Visio) / draw.io / **draw.io (stencil)** — the last variant auto-applies Cisco `mxgraph.cisco.*` stencils
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
| Language  | Link |
| ------------- | ------------- |
| English  | [Link](https://github.com/cisco-open/network-sketcher/wiki/User_Guide(Online_Edition)%5BEN%5D) |
| Japanese  | [Link](https://github.com/cisco-open/network-sketcher/wiki/User_Guide(Online_Edition)%5BJP%5D) |

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
| Place / create / delete / modify areas, devices, waypoints | ✅ | ✅ | ✅ (areas auto-placed in GUI) | ✅ |
| Add / delete / modify Layer 1 connections (port names, duplex, etc.) | ✅ | ✅ | ⚠️ (port cannot be specified on delete) | ✅ |
| Change Layer 2 segments (VLAN) / add / delete virtual ports (SVI, loopback, port-channel) | ✅ | ✅ | ✅ | ✅ |
| Change IP addresses / Layer 3 instances (VRF) | ✅ | ✅ | ✅ | ✅ |
| Change attributes | ✅ | ✅ | ✅ | ✅ |
| Add / delete VPNs | ❌ | ❌ | ✅ | ❌ |
| Flow management | ❌ | ❌ | ✅ | ❌ |
| Export various reports | ✅ (IP Address only) | ✅ (IP Address only) | ✅ | ❌ |
| Export empty master files (no data) | ✅ | ✅ | ❌ | ✅ |
| Export AI context files | ✅ | ✅ | ✅ | ✅ |
| Export device files | ✅ | ✅ | ✅ | ✅ |
| Generate L1/L2/L3 topology diagrams | ✅ | ✅ | ✅ | ✅ |
| Export diagrams as SVG (Visio-compatible) / draw.io (with Cisco stencils) | ❌ | ✅ | ❌ | ❌ |

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
