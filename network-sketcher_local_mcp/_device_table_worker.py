"""
Network Sketcher Device Table Worker - executed as a subprocess by
ns_mcp_server.py's export_device_table_html tool.

Why a separate worker process?
    ns_engine.nsm_adapter.run_cli temporarily redirects sys.stdout /
    sys.stderr to capture CLI output. If that happens inside the long-running
    MCP server process, the redirect interferes with the FastMCP stdio
    transport (JSON-RPC over stdout) and can wedge the server.

    By running the device-table build inside a one-shot subprocess we
    completely isolate any stdio redirection from the MCP transport. The
    parent reads only the final JSON result from this worker's stdout.

Protocol:
    stdin (JSON):
        {
            "engine_dir":  "<abs path to ns_engine>",
            "online_dir":  "<abs path to network-sketcher_online>",
            "master_path": "<abs path to [MASTER]*.nsm>"
        }

    stdout (JSON, on success):
        {
            "ok": true,
            "basename":    "5sites_wan_2dc",
            "output_path": "...\\[DEVICE_TABLE]5sites_wan_2dc.html",
            "size_bytes":  51205,
            "row_summary": "L1 Table=106, L2 Table=155, L3 Table=115, Attribute=38"
        }

    stdout (JSON, on failure):
        {
            "ok": false,
            "error":     "<short message>",
            "traceback": "<python traceback>"
        }

Any Python traceback escapes to stderr so the parent can log it if the
JSON channel is somehow broken.
"""

import json
import sys
import traceback
from pathlib import Path


def _emit(obj: dict) -> None:
    raw = json.dumps(obj, ensure_ascii=False).encode('utf-8')
    sys.stdout.buffer.write(raw)
    sys.stdout.buffer.flush()


def main() -> None:
    try:
        payload = json.loads(sys.stdin.buffer.read())
        engine_dir = payload['engine_dir']
        online_dir = payload['online_dir']
        master_path = Path(payload['master_path'])
    except (json.JSONDecodeError, KeyError, TypeError) as e:
        _emit({
            'ok': False,
            'error': f'invalid payload: {type(e).__name__}: {e}',
            'traceback': traceback.format_exc(),
        })
        return

    sys.path.insert(0, engine_dir)
    sys.path.insert(0, online_dir)

    try:
        from ns_engine.nsm_adapter import bootstrap  # type: ignore
        bootstrap()
        from ns_engine.nsm_device_table_html import (  # type: ignore
            build_device_tabs_data,
            render_device_table_html,
        )
    except ImportError as e:
        _emit({
            'ok': False,
            'error': f'ns_engine import failed: {e}',
            'traceback': traceback.format_exc(),
        })
        return

    if not master_path.is_file():
        _emit({
            'ok': False,
            'error': f'Master file not found: {master_path}',
        })
        return

    try:
        tabs_data, basename = build_device_tabs_data(str(master_path))
    except Exception as e:
        _emit({
            'ok': False,
            'error': f"build_device_tabs_data failed: {e}",
            'traceback': traceback.format_exc(),
        })
        return

    if tabs_data is None or basename is None:
        _emit({
            'ok': False,
            'error': (
                f"Could not load tab data from '{master_path.name}'. "
                f"Verify the .nsm file is valid and not corrupted."
            ),
        })
        return

    try:
        html = render_device_table_html(tabs_data, basename)
    except Exception as e:
        _emit({
            'ok': False,
            'error': f'render_device_table_html failed: {e}',
            'traceback': traceback.format_exc(),
        })
        return

    out_path = master_path.parent / f'[DEVICE_TABLE]{basename}.html'
    try:
        out_path.write_text(html, encoding='utf-8')
    except OSError as e:
        _emit({
            'ok': False,
            'error': f"Failed to write '{out_path.name}': {e}",
        })
        return

    def _tab_row_count(t: dict) -> int:
        sub_tables = t.get('tables') or []
        if sub_tables:
            # Multi-table tab (e.g. Placement): sum all sub-table rows.
            return sum(len(s.get('rows') or []) for s in sub_tables)
        return len(t.get('rows') or [])

    row_summary = ', '.join(
        f"{t['label']}={_tab_row_count(t)}" for t in tabs_data
    )
    _emit({
        'ok': True,
        'basename': basename,
        'output_path': str(out_path),
        'size_bytes': out_path.stat().st_size,
        'row_summary': row_summary,
    })


if __name__ == '__main__':
    main()
