'''
SPDX-License-Identifier: Apache-2.0

Copyright 2026 Cisco Systems, Inc. and its affiliates

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
'''

"""
Network Sketcher Local MCP Edition.

This MCP server wraps `network-sketcher_online/ns_engine/` and exposes
the Network Sketcher CLI through the Model Context Protocol so that
LLM clients (Cursor, Claude Code, etc.) can drive network design
operations directly without browser-based copy/paste.

Usage:
    python ns_mcp_server.py

The server speaks MCP over stdio and is intended to be launched by an
MCP host. Logs go to stderr so they appear in the host's log panel.
"""

import asyncio
import json
import logging
import os
import shlex
import shutil
import sys
import tempfile
from pathlib import Path
from typing import List, Optional

# ---------------------------------------------------------------------------
# Path setup: import ns_engine from the Online edition without copying it
# ---------------------------------------------------------------------------

_MCP_DIR = Path(__file__).resolve().parent
_PROJECT_DIR = _MCP_DIR.parent
_ONLINE_DIR = _PROJECT_DIR / 'network-sketcher_online'
_ENGINE_DIR = _ONLINE_DIR / 'ns_engine'

if not _ENGINE_DIR.is_dir():
    sys.stderr.write(
        f'[FATAL] ns_engine directory not found at {_ENGINE_DIR}\n'
        f'Network Sketcher Local MCP Edition requires the Online edition '
        f'to be present at {_ONLINE_DIR}\n'
    )
    sys.exit(1)

sys.path.insert(0, str(_ENGINE_DIR))
sys.path.insert(0, str(_ONLINE_DIR))

try:
    from ns_engine.nsm_adapter import bootstrap, run_cli, RunResult  # noqa: E402
    bootstrap()
except ImportError as e:
    sys.stderr.write(
        '[FATAL] Failed to import ns_engine dependencies.\n'
        f'Underlying error: {e}\n'
        'Run: python -m pip install -r requirements_mcp.txt\n'
    )
    sys.exit(1)
except Exception as e:
    sys.stderr.write(
        f'[FATAL] ns_engine bootstrap failed: {e}\n'
        f'Engine directory: {_ENGINE_DIR}\n'
    )
    sys.exit(1)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

_DEFAULT_AI_CONTEXT_SHOW_COMMANDS: List[str] = [
    'show area',
    'show area_device',
    'show area_location',
    'show attribute',
    'show device',
    'show device_interface',
    'show device_location',
    'show l1_interface',
    'show l1_link',
    'show l2_broadcast_domain',
    'show l2_interface',
    'show l3_broadcast_domain',
    'show l3_interface',
    'show waypoint',
    'show waypoint_interface',
]


def _load_config() -> dict:
    """Load mcp_config.json. Each top-level key has a 'value' field."""
    cfg_path = _MCP_DIR / 'mcp_config.json'
    out: dict = {}
    if not cfg_path.is_file():
        return out
    try:
        with cfg_path.open('r', encoding='utf-8') as f:
            raw = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        sys.stderr.write(f'[WARNING] Failed to load {cfg_path}: {e}\n')
        return out
    for k, v in raw.items():
        if k.startswith('_'):
            continue
        if isinstance(v, dict) and 'value' in v:
            out[k] = v['value']
    return out


_CFG = _load_config()


def _resolve_initial_working_directory() -> Optional[Path]:
    """Resolve the initial workspace from mcp_config.json (if provided).

    If `working_directory.value` is non-empty in the config, that path is
    used as the session's initial workspace (auto-created if missing).
    If empty or absent, returns None; the AI agent must then call
    suggest_workspace() and set_workspace() before using any other tool.
    This makes the server cross-platform (no hard-coded OS paths) and
    lets the agent propose an appropriate location per host.
    """
    configured = (_CFG.get('working_directory') or '').strip()
    if not configured:
        return None
    wd = Path(configured).expanduser().resolve()
    wd.mkdir(parents=True, exist_ok=True)
    return wd


_WORK_DIR: Optional[Path] = _resolve_initial_working_directory()

# ---------------------------------------------------------------------------
# Logging (stderr only; stdout is reserved for MCP JSON-RPC traffic)
# ---------------------------------------------------------------------------

_LOG_LEVEL_NAME = str(_CFG.get('log_level', 'INFO')).upper()
_LOG_LEVEL = getattr(logging, _LOG_LEVEL_NAME, logging.INFO)

logging.basicConfig(
    stream=sys.stderr,
    level=_LOG_LEVEL,
    format='%(asctime)s [%(levelname)s] ns_mcp: %(message)s',
)
logger = logging.getLogger('ns_mcp')
if _WORK_DIR is None:
    logger.info('Working directory: <not set; agent must call set_workspace()>')
else:
    logger.info('Working directory: %s', _WORK_DIR)
logger.info('Engine directory:  %s', _ENGINE_DIR)

# ---------------------------------------------------------------------------
# Helpers around run_cli
# ---------------------------------------------------------------------------

# Patterns the engine prints on failure. nsm_cli.py calls bare exit() which
# leaves returncode=0 even for errors, so we must inspect stdout/stderr too.
_ERROR_PATTERNS = (
    '[ERROR]',
    '[WARNING] No ',
    'Input must start with',
)

# Mutating verbs for which we resolve the master path inside the working dir
_MUTATING_VERBS = ('add', 'rename', 'delete')


def _require_workspace() -> Optional[str]:
    """Return an error message if no workspace is active, else None.

    Every tool that touches the filesystem MUST call this and short-circuit
    if a non-None error is returned, so that the agent receives a clear
    instruction to set up the workspace first.
    """
    if _WORK_DIR is None:
        return (
            "[ERROR] No workspace is active for this session. "
            "Call suggest_workspace() to see OS-appropriate candidates, "
            "then call set_workspace(path) with the chosen path. "
            "After confirmation you can use the other tools."
        )
    return None


def _is_safe_filename(name: str) -> bool:
    """Reject path traversal attempts and non-.nsm masters.

    The Local MCP edition uses .nsm (ZIP+Parquet) exclusively for master
    files because openpyxl-based .xlsx I/O is significantly slower than
    the Parquet path. Use the import_master/export_master_xlsx tools to
    convert between formats at the boundary.
    """
    if not name:
        return False
    if '/' in name or '\\' in name or '..' in name:
        return False
    if not (name.startswith('[MASTER]') and name.lower().endswith('.nsm')):
        return False
    return True


def _resolve_master(master: str) -> Path:
    """Resolve a .nsm master filename against the active workspace.

    Accepts either a bare filename like '[MASTER]office.nsm' (joined to
    the workspace) or an absolute path. Absolute paths are checked to
    ensure they reside inside the workspace, to prevent the LLM from
    accessing arbitrary files on the host. Only .nsm files are accepted;
    use import_master to convert .xlsx. Raises ValueError if no workspace
    is active.
    """
    if _WORK_DIR is None:
        raise ValueError(
            "No workspace is active. Call suggest_workspace() then "
            "set_workspace(path) before using master-file tools."
        )
    p = Path(master)
    if p.is_absolute():
        try:
            p.resolve().relative_to(_WORK_DIR)
        except ValueError:
            raise ValueError(
                f"Master path '{master}' is outside the active workspace "
                f'({_WORK_DIR}). Place the file inside the workspace '
                f'or pass just the filename.'
            )
        if not p.name.lower().endswith('.nsm'):
            raise ValueError(
                f"Master path '{master}' is not a .nsm file. The Local MCP "
                f"edition uses .nsm exclusively. Use import_master to "
                f"convert an existing .xlsx into .nsm first."
            )
        return p
    if not _is_safe_filename(master):
        raise ValueError(
            f"Invalid master filename '{master}'. Must start with '[MASTER]' "
            f"and end with .nsm, with no path separators. Use import_master "
            f"to convert an existing .xlsx into .nsm first."
        )
    return _WORK_DIR / master


def _classify_result(result: RunResult) -> tuple[bool, str]:
    """Heuristically classify a CLI result as success or failure.

    Returns (ok, combined_message).
    """
    combined = (result.stdout or '') + (result.stderr or '')
    if result.returncode != 0:
        return False, combined
    for pat in _ERROR_PATTERNS:
        if pat in combined:
            return False, combined
    return True, combined


_WORKER_PATH = _MCP_DIR / '_ns_cli_worker.py'


async def _run_batch(commands: List[List[str]]) -> List[RunResult]:
    """Run one or more CLI commands in a subprocess worker.

    By executing in a separate process, the worker's sys.stdout redirect
    (done internally by run_cli / nsm_adapter) is completely isolated from
    the MCP server's stdio transport.  All commands are batched into a single
    subprocess to amortise Python startup overhead.
    """
    payload = json.dumps({
        'engine_dir': str(_ENGINE_DIR),
        'online_dir': str(_ONLINE_DIR),
        'commands': commands,
    }, ensure_ascii=False).encode('utf-8')

    logger.debug('CLI batch (%d cmd(s)): %s', len(commands),
                 '; '.join(' '.join(c) for c in commands))
    try:
        proc = await asyncio.create_subprocess_exec(
            sys.executable, str(_WORKER_PATH),
            stdin=asyncio.subprocess.PIPE,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        stdout_bytes, stderr_bytes = await proc.communicate(input=payload)
    except OSError as e:
        err = f'[WORKER LAUNCH ERROR] {e}'
        logger.error(err)
        return [RunResult(returncode=1, stderr=err)] * len(commands)

    stderr_text = stderr_bytes.decode('utf-8', errors='replace').strip()
    if stderr_text:
        logger.debug('Worker stderr: %s', stderr_text)

    try:
        data = json.loads(stdout_bytes.decode('utf-8', errors='replace'))
        results = [RunResult(**r) for r in data]
    except (json.JSONDecodeError, KeyError, TypeError) as e:
        err = f'[WORKER PARSE ERROR] {e}\nstdout={stdout_bytes!r}\nstderr={stderr_text}'
        logger.error(err)
        return [RunResult(returncode=1, stderr=err)] * len(commands)

    return results


async def _run(cli_args: List[str]) -> RunResult:
    """Convenience wrapper for a single CLI command."""
    results = await _run_batch([cli_args])
    return results[0]

# ---------------------------------------------------------------------------
# MCP server
# ---------------------------------------------------------------------------

try:
    from mcp.server.fastmcp import FastMCP
except ImportError as e:
    sys.stderr.write(
        '[FATAL] The "mcp" package is not installed.\n'
        'Run: python -m pip install -r requirements_mcp.txt\n'
        f'Underlying error: {e}\n'
    )
    sys.exit(1)

_SERVER_INSTRUCTIONS = (
    "Network Sketcher Local MCP server.\n\n"
    "WORKSPACE SETUP (do this once per session, before anything else): "
    "If get_workspace_info() returns a 'workspace_active: false' state, "
    "you MUST call suggest_workspace() to retrieve OS-appropriate "
    "candidate directories, propose ONE candidate to the user with a "
    "short rationale, ASK for confirmation, then call set_workspace(path). "
    "Only after a workspace is active can you use any master-file tool.\n\n"
    "WORKFLOW REQUIREMENT: When the user asks you to inspect, modify, or "
    "design a Network Sketcher master, your FIRST action MUST be to call "
    "get_workspace_info() to discover the available .nsm master files, "
    "then call get_ai_context(master) to load the current network state "
    "and the full CLI command reference. Only AFTER reading that context "
    "should you plan and issue commands via run_commands() or other "
    "mutating tools. This prevents wasted edits based on stale "
    "assumptions and ensures you use the correct CLI syntax.\n\n"
    "FILE FORMAT: This edition operates on .nsm (ZIP+Parquet) master "
    "files exclusively for performance. If the workspace contains a "
    "legacy .xlsx master, call import_master(xlsx_path) once to convert "
    "it. Use export_master_xlsx(master) to round-trip back to .xlsx for "
    "sharing.\n\n"
    "LAYOUT REQUIREMENT: When constructing add_device_location and "
    "add_area_location commands, follow RULE 0 (TOP-TO-BOTTOM hierarchy) "
    "and RULE 0.5 (L1 LINK CROSSING AVOIDANCE) defined in the AI Context "
    "(loaded via get_ai_context). These rules keep the L1 diagram clean "
    "by avoiding lines that cross area boundaries or pass over other "
    "devices, and by minimising L1 link length when multiple placements "
    "are valid.\n\n"
    "EXTERNAL MCP SERVER USAGE (MANDATORY when relevant): The host may "
    "have other MCP servers registered alongside this one. You MUST "
    "proactively use ANY registered MCP server whose advertised "
    "capability is relevant to the current sub-task. This rule is NOT "
    "limited to documentation/knowledge-base servers; treat every "
    "registered MCP server as a candidate, regardless of vendor or "
    "category.\n"
    "  Procedure: At session start AND before each concrete sub-task "
    "(model selection, capacity planning, config validation, "
    "troubleshooting, ticket lookup, etc.), inspect the list of "
    "available MCP servers (via the host's tool list or any meta-tool "
    "such as list_servers) and ask yourself 'is any registered MCP "
    "server's stated capability relevant here?'. If yes, you MUST call "
    "it BEFORE issuing Network Sketcher mutations (run_commands / "
    "export_diagram / etc.), and you MUST cite the returned "
    "sources/results in your final answer so the user can audit the "
    "recommendations.\n"
    "  Examples of relevance (non-exhaustive):\n"
    "    - Documentation / RAG servers (vendor design guides, EoS/EoL "
    "data, configuration examples) -> use during network design, "
    "model selection, troubleshooting, capacity planning, "
    "best-practice lookup, and whenever you would otherwise rely on "
    "training-data knowledge.\n"
    "    - Topology / configuration / monitoring servers -> use to "
    "ground design decisions in real device state when available.\n"
    "    - Issue trackers / chat / repository servers -> use when the "
    "user's request references existing tickets, code, or "
    "conversations.\n"
    "  Rationale: Grounding decisions in authoritative sources via "
    "registered MCP servers is strictly preferred over relying solely "
    "on the LLM's training data.\n"
    "  Guardrails (avoid abuse):\n"
    "    - Do NOT invoke a server whose advertised capability is "
    "clearly unrelated to the task (e.g. a filesystem-only server "
    "during a pure SVI binding fix). Relevance is judged from each "
    "server's description (serverUseInstructions / tool descriptors), "
    "not from its name alone.\n"
    "    - If a server requires authentication (e.g. exposes an "
    "mcp_auth tool), authenticate one server at a time; never "
    "authenticate in parallel.\n"
    "    - Avoid more than ~10 external MCP calls per single user "
    "turn unless the task genuinely requires it.\n"
    "    - If a server fails (network error, auth error), report the "
    "failure to the user and proceed with the remaining sources "
    "rather than retrying indefinitely.\n"
    "  If, after inspection, NO registered MCP server is relevant to "
    "the current task, proceed using Network Sketcher tools and your "
    "general knowledge alone, and briefly note in your reply that no "
    "external sources were applicable."
)

mcp = FastMCP('network-sketcher', instructions=_SERVER_INSTRUCTIONS)

# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------


@mcp.tool()
async def get_workspace_info() -> str:
    """Return the active workspace and the list of .nsm master files in it.

    Always safe to call. If no workspace has been set yet, returns
    workspace_active=false with a reminder to call suggest_workspace().

    Returns:
        JSON string with keys:
            workspace_active: true/false
            working_directory: absolute path or null
            masters: list of .nsm master filenames ('[MASTER]*.nsm')
            xlsx_masters: list of .xlsx files that need conversion via
                          import_master before they can be used
            other_files: list of generated outputs (DIAGRAM / AI_Context, etc.)
            note: human-readable hint when no workspace is active
    """
    if _WORK_DIR is None:
        return json.dumps({
            'workspace_active': False,
            'working_directory': None,
            'masters': [],
            'xlsx_masters': [],
            'other_files': [],
            'note': (
                "No workspace is active. Call suggest_workspace() to see "
                "OS-appropriate candidate directories, then call "
                "set_workspace(path) with the chosen path."
            ),
        }, indent=2, ensure_ascii=False)

    masters: List[str] = []
    xlsx_masters: List[str] = []
    others: List[str] = []
    for entry in sorted(_WORK_DIR.iterdir()):
        if not entry.is_file():
            continue
        name = entry.name
        if name.startswith('[MASTER]'):
            if name.lower().endswith('.nsm'):
                masters.append(name)
            elif name.lower().endswith('.xlsx'):
                xlsx_masters.append(name)
            else:
                others.append(name)
        elif name.startswith('['):
            others.append(name)
    return json.dumps({
        'workspace_active': True,
        'working_directory': str(_WORK_DIR),
        'masters': masters,
        'xlsx_masters': xlsx_masters,
        'other_files': others,
    }, indent=2, ensure_ascii=False)


@mcp.tool()
async def suggest_workspace() -> str:
    """Suggest OS-appropriate workspace directory candidates.

    Always safe to call. Returns a list of candidate paths under the
    user's home directory (Windows / macOS / Linux), with hints about
    which exist, which are writable, and which the agent should prefer.
    The agent must propose ONE candidate to the user, get confirmation,
    then call set_workspace(path). Custom paths under the user home are
    also acceptable; the candidates are only suggestions.

    Returns:
        JSON string with keys:
            os, current_workspace, home, candidates, guidance
    """
    import platform

    os_name = platform.system()
    home = Path.home().resolve()

    raw_candidates: List[Path] = []
    if os_name == 'Windows':
        raw_candidates = [
            home / 'Documents' / 'ns_workspace',
            home / 'Desktop' / 'ns_workspace',
            home / 'ns_workspace',
        ]
    elif os_name == 'Darwin':
        raw_candidates = [
            home / 'Documents' / 'ns_workspace',
            home / 'Desktop' / 'ns_workspace',
            home / 'ns_workspace',
        ]
    else:
        xdg = os.environ.get('XDG_DATA_HOME', '').strip()
        xdg_base = Path(xdg).expanduser().resolve() if xdg else home / '.local' / 'share'
        raw_candidates = [
            xdg_base / 'ns_workspace',
            home / 'Documents' / 'ns_workspace',
            home / 'ns_workspace',
        ]

    seen = set()
    candidates_out = []
    for idx, c in enumerate(raw_candidates):
        try:
            resolved = c.expanduser().resolve()
        except OSError:
            continue
        key = str(resolved)
        if key in seen:
            continue
        seen.add(key)
        parent = resolved.parent
        candidates_out.append({
            'path': str(resolved),
            'exists': resolved.is_dir(),
            'parent_exists': parent.is_dir(),
            'parent_writable': os.access(str(parent), os.W_OK) if parent.is_dir() else False,
            'preferred': idx == 0,
        })

    guidance = (
        "Pick ONE candidate (preferred=true is the recommended default), "
        "or propose a custom path under the user's home directory. "
        "Show the chosen path to the user for confirmation, then call "
        "set_workspace(path)."
    )

    return json.dumps({
        'os': os_name,
        'current_workspace': str(_WORK_DIR) if _WORK_DIR else None,
        'home': str(home),
        'candidates': candidates_out,
        'guidance': guidance,
    }, indent=2, ensure_ascii=False)


@mcp.tool()
async def set_workspace(path: str) -> str:
    """Set the active workspace directory for this session.

    Validates that the path is under the user's home directory (defence
    against arbitrary filesystem access), creates the directory if it
    does not exist, and confirms it is writable. Once set, subsequent
    tool calls operate against this directory until the server restarts
    or set_workspace is called again.

    Args:
        path: Absolute path to the desired workspace. May use '~' for
              home expansion. Must resolve to a location under
              Path.home().

    Returns:
        Confirmation message including the resolved absolute path,
        or an error if validation failed.
    """
    if not path or not str(path).strip():
        return "[ERROR] path argument is empty."
    try:
        target = Path(path).expanduser().resolve()
    except (OSError, ValueError) as e:
        return f"[ERROR] Invalid path '{path}': {e}"

    home = Path.home().resolve()
    try:
        target.relative_to(home)
    except ValueError:
        return (
            f"[ERROR] For safety, the workspace must be under your home "
            f"directory ({home}). Got: {target}. Choose a path beneath "
            f"your home, e.g. {home / 'ns_workspace'}."
        )

    try:
        target.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return f"[ERROR] Could not create workspace '{target}': {e}"

    if not os.access(str(target), os.W_OK):
        return f"[ERROR] Workspace exists but is not writable: {target}"

    global _WORK_DIR
    previous = _WORK_DIR
    _WORK_DIR = target
    logger.info('Workspace set: %s (was: %s)', _WORK_DIR, previous)

    existing = []
    try:
        for entry in sorted(target.iterdir()):
            if entry.is_file() and entry.name.startswith('['):
                existing.append(entry.name)
    except OSError:
        pass

    msg = (
        f"Workspace activated: {target}\n"
        f"Previous workspace: {previous if previous else '(none)'}\n"
        f"Existing NS files in this directory: "
        f"{len(existing)} (showing first 10)\n"
    )
    for name in existing[:10]:
        msg += f"  - {name}\n"
    msg += (
        "\nNext step: call get_workspace_info() to confirm the state, "
        "then get_ai_context(master) before issuing any edits."
    )
    return msg


@mcp.tool()
async def create_empty_master(filename: Optional[str] = None) -> str:
    """Create a new empty Network Sketcher master file (.nsm) in the working dir.

    The engine's `export master_file_nodata` produces an .xlsx, which this
    tool immediately converts to .nsm via xlsx_to_nsm() and discards.
    The xlsx is written to a system temp directory (never to the working
    directory) so that no .xlsx artefact is left behind.

    Args:
        filename: Optional target name. Must start with '[MASTER]' and end
                  with '.nsm'. No path separators allowed. Defaults to
                  '[MASTER]no_data.nsm'.

    Returns:
        Result message indicating success and final file path, or an error.
    """
    err = _require_workspace()
    if err:
        return err

    target_name = filename or '[MASTER]no_data.nsm'
    if not _is_safe_filename(target_name):
        return (
            f"[ERROR] Invalid filename '{target_name}'. Must start with "
            f"'[MASTER]' and end with '.nsm'."
        )

    final_path = _WORK_DIR / target_name
    if final_path.exists():
        return (
            f"[ERROR] File already exists: {final_path}. "
            f"Delete it first or pick a different name."
        )

    # Engine's `export master_file_nodata` writes '[MASTER]no_data.xlsx'
    # next to --master. We stage this in a temp dir so no xlsx is left
    # behind in the working directory.
    tmp_dir = Path(tempfile.mkdtemp(prefix='nsm_create_'))
    try:
        seed = tmp_dir / '[MASTER]_seed.xlsx'
        seed.write_bytes(b'')
        result = await _run([
            'export', 'master_file_nodata',
            '--master', str(seed),
        ])
        ok, msg = _classify_result(result)
        src_xlsx = tmp_dir / '[MASTER]no_data.xlsx'
        if not ok or not src_xlsx.exists():
            return f"[ERROR] Failed to create empty master.\n{msg.strip()}"

        try:
            from ns_engine.nsm_io import xlsx_to_nsm
            xlsx_to_nsm(str(src_xlsx), str(final_path))
        except Exception as e:
            return (
                f"[ERROR] Engine produced xlsx but xlsx_to_nsm failed: {e}\n"
                f"{msg.strip()}"
            )
    finally:
        shutil.rmtree(str(tmp_dir), ignore_errors=True)

    if not final_path.exists():
        return f"[ERROR] Conversion succeeded but {final_path} is missing."

    return (
        f"Created empty .nsm master file: {final_path.name}\n"
        f"Location: {final_path}\n"
        f"Size: {final_path.stat().st_size:,} bytes"
    )


@mcp.tool()
async def get_network_state(master: str) -> str:
    """Run the standard set of `show` commands and return aggregated results.

    This is the lightweight alternative to `get_ai_context`. It produces a
    plain-text summary of the current network state (areas, devices,
    interfaces, L1/L2/L3 topology, attributes) by invoking the show
    commands listed in mcp_config.json -> ai_context_show_commands.

    Args:
        master: Master filename inside the working directory
                (e.g. '[MASTER]office.xlsx'), or an absolute path inside
                the working directory.

    Returns:
        Concatenated stdout from each show command, grouped by command name.
    """
    err = _require_workspace()
    if err:
        return err
    try:
        master_path = _resolve_master(master)
    except ValueError as e:
        return f'[ERROR] {e}'
    if not master_path.is_file():
        return f"[ERROR] Master file not found: {master_path}"

    show_commands = _CFG.get('ai_context_show_commands') or _DEFAULT_AI_CONTEXT_SHOW_COMMANDS

    # Build token lists for all valid show commands
    cmd_batches: List[List[str]] = []
    cmd_labels: List[str] = []
    for cmd in show_commands:
        try:
            tokens = shlex.split(cmd)
        except ValueError:
            tokens = cmd.split()
        if not tokens or tokens[0] != 'show':
            continue
        cmd_batches.append(tokens + ['--master', str(master_path)])
        cmd_labels.append(cmd)

    # Run all show commands in a single subprocess call
    results = await _run_batch(cmd_batches)

    sections: List[str] = []
    for cmd, result in zip(cmd_labels, results):
        body = (result.stdout or '').rstrip()
        if not body and result.stderr:
            body = '[ERROR] ' + result.stderr.rstrip()
        sections.append(f'** {cmd.replace(" ", "_")}\n{body}')

    return '\n\n'.join(sections) + '\n'


@mcp.tool()
async def get_ai_context(master: str) -> str:
    """Generate the full AI Context file for a master and return its contents.

    This invokes `export ai_context_file` (always with
    --accept-security-risk to avoid the interactive prompt) and reads back
    the generated '[AI_Context]<basename>.txt'. The AI Context bundles the
    network state plus the full CLI command reference, suitable for
    feeding to a language model that needs to plan multi-step edits.

    Args:
        master: Master filename inside the working directory
                (e.g. '[MASTER]office.xlsx').

    Returns:
        Full text content of the generated AI Context file, or an error.
    """
    err = _require_workspace()
    if err:
        return err
    try:
        master_path = _resolve_master(master)
    except ValueError as e:
        return f'[ERROR] {e}'
    if not master_path.is_file():
        return f"[ERROR] Master file not found: {master_path}"

    # --accept-security-risk skips the input() prompt that would otherwise
    # block this server (no TTY).
    result = await _run([
        'export', 'ai_context_file',
        '--master', str(master_path),
        '--accept-security-risk',
    ])
    ok, msg = _classify_result(result)

    # Engine names the AI Context file as '[AI_Context]<basename>.txt'
    # where <basename> is the master stem WITHOUT the '[MASTER]' prefix.
    base = master_path.stem  # e.g. '[MASTER]office' or '[MASTER]office.nsm'
    if base.lower().startswith('[master]'):
        base = base[8:]      # strip '[MASTER]' → 'office'
    if base.lower().endswith('.nsm'):
        base = base[:-4]
    ai_path = _WORK_DIR / f'[AI_Context]{base}.txt'
    if not ai_path.is_file():
        return f"[ERROR] AI Context file was not generated.\n{msg.strip()}"

    try:
        text = ai_path.read_text(encoding='utf-8')
    except OSError as e:
        return f'[ERROR] Failed to read {ai_path}: {e}'

    header = f"# AI Context for {master_path.name}\n# File: {ai_path}\n\n"
    return header + text


@mcp.tool()
async def run_commands(master: str, commands: str) -> str:
    """Execute one or more Network Sketcher CLI commands against a .nsm master.

    PREREQUISITE: You MUST have called get_ai_context(master) (or at least
    get_network_state(master)) at least once in this session before using
    this tool. Without that context you do not know the current state nor
    the available CLI syntax, and your edits are likely to be wrong.

    Each non-empty line of `commands` is parsed with shlex and run as a
    single CLI invocation. Allowed verbs: add, rename, delete, show.
    `export` is intentionally excluded; use create_empty_master,
    export_diagram, or get_ai_context instead.

    The `--master` argument is appended automatically; do NOT include it
    in the command lines.

    Args:
        master: Master filename inside the working directory.
        commands: Newline-separated CLI command lines. Example:
                      add device 'SW-3' --area 'DC1'
                      add l1_link 'SW-2' 'GE 0/0' 'SW-3' 'GE 0/1'

    Returns:
        Per-line results: each command's exit summary and stdout, joined.
    """
    err = _require_workspace()
    if err:
        return err
    try:
        master_path = _resolve_master(master)
    except ValueError as e:
        return f'[ERROR] {e}'
    if not master_path.is_file():
        return f"[ERROR] Master file not found: {master_path}"

    allowed = {'add', 'rename', 'delete', 'show'}

    # --- Pre-validate all lines first, reject invalid ones early ---
    valid_batches: List[List[str]] = []  # token lists for valid commands
    valid_lines: List[str] = []          # original text for each valid command
    pre_errors: List[str] = []           # validation error messages

    for raw_line in commands.splitlines():
        line = raw_line.strip()
        if not line or line.startswith('#'):
            continue
        try:
            tokens = shlex.split(line)
        except ValueError as e:
            pre_errors.append(f"[ERROR] Could not parse line: {line}\n  ({e})")
            continue
        if not tokens:
            continue
        verb = tokens[0]
        if verb not in allowed:
            pre_errors.append(
                f"[ERROR] Verb '{verb}' is not allowed via run_commands. "
                f"Allowed: {sorted(allowed)}. Use a dedicated tool for export."
            )
            continue
        if '--master' in tokens:
            pre_errors.append(
                f"[ERROR] Do not include --master in the command line; it is "
                f"added automatically. Offending line: {line}"
            )
            continue
        valid_batches.append(tokens + ['--master', str(master_path)])
        valid_lines.append(line)

    # --- Run all valid commands in a single subprocess call ---
    output_lines: List[str] = list(pre_errors)
    success_count = 0
    failure_count = len(pre_errors)

    if valid_batches:
        results = await _run_batch(valid_batches)
        for line, result in zip(valid_lines, results):
            ok, msg = _classify_result(result)
            status = 'OK' if ok else 'FAIL'
            if ok:
                success_count += 1
            else:
                failure_count += 1
            output_lines.append(f"[{status}] {line}\n{msg.rstrip()}")

    total = success_count + failure_count
    summary = f"# Executed {total} command(s): {success_count} OK, {failure_count} FAIL\n"
    return summary + '\n\n'.join(output_lines)


@mcp.tool()
async def import_master(xlsx_path: str, target_name: Optional[str] = None) -> str:
    """Convert an existing .xlsx master into a .nsm in the working directory.

    The Local MCP edition operates on .nsm files exclusively. Use this tool
    once at the start of a session to bring an existing .xlsx master under
    .nsm management. The original .xlsx is left untouched.

    Args:
        xlsx_path: Absolute path to the source .xlsx file. The file must
                   exist and have an '.xlsx' extension. The path itself
                   may live outside the working directory.
        target_name: Optional output filename. Must start with '[MASTER]'
                     and end with '.nsm'. If omitted, the source basename
                     is reused with the extension swapped to '.nsm'.

    Returns:
        Result message with the produced .nsm path, or an error.
    """
    err = _require_workspace()
    if err:
        return err
    src = Path(xlsx_path)
    if not src.is_absolute():
        return (
            f"[ERROR] xlsx_path must be an absolute path. Got: {xlsx_path}"
        )
    if not src.is_file():
        return f"[ERROR] Source file not found: {src}"
    if src.suffix.lower() != '.xlsx':
        return f"[ERROR] Source file must have .xlsx extension: {src}"

    if target_name is None:
        base = src.name
        if base.lower().endswith('.xlsx'):
            base = base[:-5]
        if not base.startswith('[MASTER]'):
            base = f'[MASTER]{base}'
        target_name = f'{base}.nsm'

    if not _is_safe_filename(target_name):
        return (
            f"[ERROR] Invalid target_name '{target_name}'. Must start with "
            f"'[MASTER]' and end with '.nsm', no path separators."
        )

    final_path = _WORK_DIR / target_name
    if final_path.exists():
        return (
            f"[ERROR] Output file already exists: {final_path}. "
            f"Delete it first or pick a different target_name."
        )

    try:
        from ns_engine.nsm_io import xlsx_to_nsm
        xlsx_to_nsm(str(src), str(final_path))
    except Exception as e:
        return f"[ERROR] xlsx_to_nsm failed: {e}"

    if not final_path.exists():
        return (
            f"[ERROR] Conversion completed without errors but "
            f"{final_path} is missing."
        )

    return (
        f"Imported master: {src} -> {final_path}\n"
        f"Size: {final_path.stat().st_size:,} bytes\n"
        f"You can now use '{final_path.name}' as the master argument for "
        f"run_commands, get_ai_context, export_diagram, etc."
    )


@mcp.tool()
async def export_master_xlsx(master: str) -> str:
    """Convert a .nsm master back to .xlsx for Excel/Offline edition use.

    The output is placed in the working directory next to the source .nsm,
    using the same basename with '.xlsx' extension. The .nsm itself is
    not modified.

    Args:
        master: Master filename inside the working directory
                (e.g. '[MASTER]office.nsm').

    Returns:
        Result message with the produced .xlsx path, or an error.
    """
    err = _require_workspace()
    if err:
        return err
    try:
        master_path = _resolve_master(master)
    except ValueError as e:
        return f'[ERROR] {e}'
    if not master_path.is_file():
        return f"[ERROR] Master file not found: {master_path}"

    xlsx_name = master_path.stem + '.xlsx'
    xlsx_path = _WORK_DIR / xlsx_name
    if xlsx_path.exists():
        return (
            f"[ERROR] Output file already exists: {xlsx_path}. "
            f"Delete it first or rename it."
        )

    try:
        from ns_engine.nsm_io import nsm_to_xlsx
        nsm_to_xlsx(str(master_path), str(xlsx_path))
    except Exception as e:
        return f"[ERROR] nsm_to_xlsx failed: {e}"

    if not xlsx_path.exists():
        return (
            f"[ERROR] Conversion completed without errors but "
            f"{xlsx_path} is missing."
        )

    return (
        f"Exported xlsx: {master_path.name} -> {xlsx_path.name}\n"
        f"Location: {xlsx_path}\n"
        f"Size: {xlsx_path.stat().st_size:,} bytes"
    )


@mcp.tool()
async def export_diagram(master: str, layer: str, format: str = 'svg') -> str:
    """Export an L1, L2, or L3 network diagram for the given .nsm master.

    PREREQUISITE: You MUST have called get_ai_context(master) (or at least
    get_network_state(master)) at least once in this session for this
    master so that you understand which layer is meaningful to export
    given the current network state.

    Args:
        master: Master filename inside the working directory.
        layer:  One of 'l1', 'l2', 'l3' (case-insensitive).
        format: 'svg' (default) or 'pptx'. SVG is faster and renders in the
                browser. PPTX is editable in PowerPoint.

    Returns:
        A summary describing which files were generated, plus the engine
        stdout. Use get_workspace_info afterwards to enumerate the files.
    """
    err = _require_workspace()
    if err:
        return err
    try:
        master_path = _resolve_master(master)
    except ValueError as e:
        return f'[ERROR] {e}'
    if not master_path.is_file():
        return f"[ERROR] Master file not found: {master_path}"

    layer_norm = (layer or '').strip().lower()
    if layer_norm not in {'l1', 'l2', 'l3'}:
        return f"[ERROR] Invalid layer '{layer}'. Must be one of: l1, l2, l3."
    fmt = (format or 'svg').strip().lower()
    if fmt not in {'svg', 'pptx'}:
        return f"[ERROR] Invalid format '{format}'. Must be 'svg' or 'pptx'."

    cli_args = [
        'export', f'{layer_norm}_diagram',
        '--master', str(master_path),
        '--format', fmt,
    ]
    result = await _run(cli_args)
    ok, msg = _classify_result(result)
    status = 'OK' if ok else 'FAIL'
    return (
        f"[{status}] export {layer_norm}_diagram --format {fmt}\n"
        f"Master: {master_path.name}\n"
        f"{msg.strip()}"
    )

# ---------------------------------------------------------------------------
# Prompts
# ---------------------------------------------------------------------------


@mcp.prompt(
    name='start_ns_session',
    description=(
        'Bootstrap a Network Sketcher session by loading the AI context for '
        'the given master before any edits. Use this whenever you start work '
        'on a master file.'
    ),
)
async def start_ns_session(master: str) -> str:
    """Return a templated user message that drives the agent through the
    standard NS session bootstrap (workspace selection -> workspace info ->
    AI context load -> summary -> wait for instructions).
    """
    return (
        f"I want to work on the master file '{master}'.\n\n"
        f"Please follow this workflow strictly before doing anything else:\n"
        f"-1. List the OTHER MCP servers registered in this host "
        f"(excluding network-sketcher itself). For each one, briefly "
        f"state which of my upcoming sub-tasks (e.g. model selection, "
        f"capacity planning, config validation, troubleshooting, "
        f"ticket/code lookup) its advertised capability is likely to "
        f"be relevant to. Per the server instructions, you MUST call "
        f"every relevant server BEFORE running any Network Sketcher "
        f"mutations, and MUST cite the returned sources/results in "
        f"your final summary. If, after inspection, no registered MCP "
        f"server is relevant to this task, state that explicitly and "
        f"continue with the steps below.\n"
        f"0. Call get_workspace_info(). If 'workspace_active' is false, "
        f"call suggest_workspace(), propose ONE candidate to me with a "
        f"short rationale, ASK for my confirmation, then call "
        f"set_workspace(path).\n"
        f"1. Call get_workspace_info() again to confirm the master file "
        f"exists as a .nsm in the workspace. If only a legacy .xlsx is "
        f"present, call import_master() first.\n"
        f"2. Call get_ai_context('{master}') and read the full output, "
        f"including the network state and the CLI command reference.\n"
        f"3. Summarise back to me: areas, devices, links, VLAN/IP plan, "
        f"and any anomalies you notice.\n"
        f"4. Wait for my next instruction before issuing any add/rename/"
        f"delete/export commands."
    )


# ---------------------------------------------------------------------------
# Resources
# ---------------------------------------------------------------------------


@mcp.resource('nsm://workspace')
async def res_workspace() -> str:
    """Workspace overview: working directory and file inventory (JSON)."""
    return await get_workspace_info()


@mcp.resource('nsm://commands')
async def res_command_reference() -> str:
    """The full Network Sketcher CLI command reference, as bundled
    with the engine (nsm_extensions_cmd_list.txt).
    """
    cmd_list_path = _ENGINE_DIR / 'nsm_extensions_cmd_list.txt'
    if not cmd_list_path.is_file():
        return f'[ERROR] Command reference not found at {cmd_list_path}'
    try:
        return cmd_list_path.read_text(encoding='utf-8')
    except OSError as e:
        return f'[ERROR] Failed to read command reference: {e}'

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    logger.info('Starting Network Sketcher Local MCP server (stdio)')
    mcp.run()
