"""Device Table HTML rendering (shared between Online edition and Local MCP).

This module extracts two responsibilities from `ns_web_start.py`:

  1. `build_device_tabs_data(master_path)` -- read L1 / L2 / L3 / Attribute
     data from an `.nsm` (or `.xlsx`) master and return a tabs-data structure
     consumable by both the Online edition's `/device_preview/<job_id>` route
     and the Local MCP `export_device_table_html` tool.

  2. `render_device_table_html(tabs_data, master_basename)` -- emit a
     self-contained HTML page (vanilla JS + inline CSS, no external CDN)
     with sticky-header tables, tab switching, and per-tab CSV/HTML
     download buttons.

The module has no Flask / job-id / upload-directory dependency, so it can
be imported safely from both edition entrypoints.
"""

from __future__ import annotations

import ast
import json
import logging
import os
import sys
from typing import List, Optional, Tuple

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Data-layer helpers
# ---------------------------------------------------------------------------

def _master_basename(master_path: str) -> str:
    """Return a display-friendly basename, stripping the leading [MASTER] tag
    and the .nsm extension.

    Device Table only accepts .nsm masters (see build_device_tabs_data); the
    .xlsx fallback that used to live here is intentionally removed so that
    callers cannot accidentally render a basename for an Excel master.
    """
    name = os.path.basename(str(master_path))
    if name.lower().endswith('.nsm'):
        name = name[: -len('.nsm')]
    return name.replace('[MASTER]', '').strip()


def _ensure_nsm_def_importable() -> Optional[object]:
    """Make sure `nsm_def` is importable from the ns_engine directory.

    The Online edition's BASE_DIR / ns_engine is already on sys.path under
    normal startup, but the Local MCP server imports this module directly,
    so we re-add the engine directory defensively.
    """
    engine_dir = os.path.dirname(os.path.abspath(__file__))
    if engine_dir not in sys.path:
        sys.path.insert(0, engine_dir)
    try:
        import nsm_def  # type: ignore
        return nsm_def
    except ImportError as exc:
        logger.warning('nsm_def import failed: %s', exc)
        return None


def _fetch_attribute_output(master_path: str) -> Optional[str]:
    """Run `show attribute --master <path> --one_msg` in-process via
    `ns_engine.nsm_adapter.run_cli`. Returns stdout text on success, None on
    failure. Wrapping the CLI keeps the in-edition automatic ATTRIBUTE
    repair logic intact (see nsm_cli show attribute)."""
    try:
        from nsm_adapter import run_cli  # type: ignore
    except ImportError:
        # When invoked from outside ns_engine without bootstrap, fall back to
        # the package-qualified import path.
        try:
            from ns_engine.nsm_adapter import run_cli  # type: ignore
        except ImportError as exc:
            logger.warning('nsm_adapter.run_cli import failed: %s', exc)
            return None
    try:
        result = run_cli(['show', 'attribute', '--master',
                          str(master_path), '--one_msg'])
    except Exception as exc:  # pragma: no cover - defensive
        logger.warning('show attribute failed: %s', exc)
        return None
    if not result or result.returncode != 0:
        return None
    return (result.stdout or '').strip()


def build_device_tabs_data(master_path: str
                           ) -> Tuple[Optional[List[dict]], Optional[str]]:
    """Build the 6-tab data structure from a `.nsm` master file.

    Args:
        master_path: absolute path to the ``[MASTER]*.nsm`` file. Device Table
            is strictly an ``.nsm``-only feature; ``.xlsx`` masters are
            rejected with ``ValueError`` so that the underlying
            ``nsm_def.convert_master_to_array`` cannot silently fall back to
            its Excel reader for this path.

    Returns:
        (tabs_data, master_basename) where ``tabs_data`` is
        ``[{id, label, headers, rows, row_colors?}, ...]`` covering the
        l1 / l2 / l3 / attribute / IP Address_Summary / IP Address_List tabs
        in that order; or ``(None, None)`` if the master file is missing or
        unreadable.

    Raises:
        ValueError: if ``master_path`` does not end with ``.nsm``.
    """
    master_path = str(master_path)
    if not master_path.lower().endswith('.nsm'):
        raise ValueError(
            'Device Table only accepts .nsm master files; got: '
            + master_path
        )
    if not os.path.isfile(master_path):
        logger.warning('build_device_tabs_data: master not found: %s',
                       master_path)
        return None, None

    nsm_def = _ensure_nsm_def_importable()
    if nsm_def is None:
        return None, None

    master_basename = _master_basename(master_path)

    def _read_section(ws_name: str, section: str):
        try:
            return nsm_def.convert_master_to_array(ws_name, master_path,
                                                   section)
        except Exception as exc:
            logger.warning('convert_master_to_array %s/%s failed: %s',
                           ws_name, section, exc)
            return []

    # --- POSITION_SHAPE: build device -> area lookup (used by L1 + Attribute) -
    ps_raw = _read_section('Master_Data', '<<POSITION_SHAPE>>')
    device_area_map: dict = {}
    cur_folder: Optional[str] = None
    for item in ps_raw:
        if not isinstance(item, list) or len(item) < 2:
            continue
        row = item[1]
        if not isinstance(row, list):
            continue
        if row and row[0] and row[0] not in (
                '', '<END>', '<<POSITION_SHAPE>>', '_AIR_'):
            cur_folder = row[0]
        if cur_folder:
            area_label = '_N/A_' if '_wp_' in cur_folder else cur_folder
            for val in row:
                if (val and val not in ('', '<END>', '_AIR_',
                                        '<<POSITION_SHAPE>>', cur_folder)
                        and not str(val).startswith('_AIR_')):
                    device_area_map[str(val)] = area_label
        if isinstance(row, list) and len(row) == 1 and row[0] == '<END>':
            cur_folder = None

    # --- L2: Master_Data_L2 / <<L2_TABLE>> -----------------------------------
    # Cols: 0=Area, 1=Device Name, 2=Port Mode(formula/empty), 3=Port Name,
    #       4=Virtual Port Mode(formula/empty), 5=Virtual Port Name,
    #       6=Connected L2 Segment, 7=L2 (L3 Virtual Port)
    l2_raw = _read_section('Master_Data_L2', '<<L2_TABLE>>')
    l2_data = [r[1] for r in l2_raw if isinstance(r, list) and r[0] > 2]
    USE_L2 = [0, 1, 3, 5, 6, 7]
    l2_headers = ['Area', 'Device Name', 'Port Name', 'Virtual Port Name',
                  'Connected L2 Segment', 'L2 (L3 Virtual Port)']
    l2_rows = [
        [str(row[i]) if i < len(row) and row[i] not in (None, '') else ''
         for i in USE_L2]
        for row in l2_data
    ]

    # --- L3: Master_Data_L3 / <<L3_TABLE>> -----------------------------------
    l3_raw = _read_section('Master_Data_L3', '<<L3_TABLE>>')
    l3_hdr_row = next((r[1] for r in l3_raw
                       if isinstance(r, list) and r[0] == 2), [])
    l3_headers = [h for h in l3_hdr_row if h is not None]
    l3_data = [r[1] for r in l3_raw if isinstance(r, list) and r[0] > 2]
    l3_rows = [
        [str(c) if c is not None else '' for c in row[:len(l3_headers)]]
        for row in l3_data
    ]

    # --- L1: Excel-compatible flat format (1 link -> 2 rows) -----------------
    # POSITION_LINE raw col indices:
    #   0=From_Name, 1=To_Name, 2=From_Tag_raw, 3=To_Tag_raw,
    #   12=From_Port_prefix, 13=From_Speed, 14=From_Duplex, 15=From_Port_Type,
    #   16=To_Port_prefix,  17=To_Speed,   18=To_Duplex,   19=To_Port_Type
    pl_raw = _read_section('Master_Data', '<<POSITION_LINE>>')
    pl_data = [r[1] for r in pl_raw if isinstance(r, list) and r[0] > 2]
    l1_headers = [
        'Area', 'Device Name', 'Port Name', 'Abbreviation(Diagram)',
        'Speed', 'Duplex', 'Port Type',
        '[src] Device Name', '[src] Port Name',
        '[dst] Device Name', '[dst] Port Name',
    ]

    def _make_port(raw_tag: str, prefix: str) -> Tuple[str, str]:
        if ' ' in raw_tag:
            parts = raw_tag.split(' ')
            return (prefix + ' ' + parts[-1]).strip(), parts[0]
        return (prefix or raw_tag).strip(), raw_tag

    l1_rows: List[list] = []
    for row in pl_data:
        if len(row) < 20:
            continue
        from_dev = row[0] or ''
        to_dev = row[1] or ''
        from_full, from_abbr = _make_port(row[2] or '', row[12] or '')
        to_full, to_abbr = _make_port(row[3] or '', row[16] or '')
        from_raw = row[2] or ''
        to_raw = row[3] or ''
        l1_rows.append([
            device_area_map.get(from_dev, ''), from_dev,
            from_full, from_abbr,
            row[13] or '', row[14] or '', row[15] or '',
            from_dev, from_raw, to_dev, to_raw,
        ])
        l1_rows.append([
            device_area_map.get(to_dev, ''), to_dev,
            to_full, to_abbr,
            row[17] or '', row[18] or '', row[19] or '',
            from_dev, from_raw, to_dev, to_raw,
        ])

    # Sort: Area asc -> Device Name asc -> numeric port index -> raw port name
    l1_rows.sort(key=lambda x: (
        x[0], x[1],
        nsm_def.get_if_value(x[2]),
        x[2],
    ))

    # --- Attribute: show attribute --one_msg via in-process CLI --------------
    attr_headers = ['Area', 'Device Name']
    attr_rows: List[list] = []
    attr_row_colors: List[list] = []  # parallel to attr_rows
    stdout = _fetch_attribute_output(master_path)
    if stdout:
        try:
            raw_attr = ast.literal_eval(stdout)
            if raw_attr and isinstance(raw_attr[0], list):
                attr_headers = ['Area'] + raw_attr[0]
                for row in raw_attr[1:]:
                    vals: list = []
                    cols: list = []
                    for i, cell in enumerate(row):
                        if i == 0:
                            dev = str(cell) if cell is not None else ''
                            area = device_area_map.get(dev, '')
                            vals = [area, dev]
                            cols = [None, None]
                        else:
                            try:
                                cell_list = ast.literal_eval(str(cell))
                                text = cell_list[0] if cell_list else ''
                                text = ('' if text in ('<EMPTY>', None)
                                        else str(text))
                                rgb = (cell_list[1]
                                       if len(cell_list) > 1 else None)
                                if (rgb and isinstance(rgb, list)
                                        and len(rgb) == 3
                                        and tuple(rgb) != (255, 255, 255)):
                                    color = (f'rgb({rgb[0]},{rgb[1]},'
                                             f'{rgb[2]})')
                                else:
                                    color = None
                            except Exception:
                                text = str(cell) if cell is not None else ''
                                color = None
                            vals.append(text)
                            cols.append(color)
                    attr_rows.append(vals)
                    attr_row_colors.append(cols)
        except Exception as exc:
            logger.warning('Attribute parse failed: %s', exc)

    # Sort Attribute rows by (Area, Device Name) ascending.
    if attr_rows:
        combined = list(zip(attr_rows, attr_row_colors))
        combined.sort(
            key=lambda x: (x[0][0] if x[0] else '',
                           x[0][1] if len(x[0]) > 1 else ''),
            reverse=False,
        )
        attr_rows, attr_row_colors = map(list, zip(*combined))

    tabs_data = [
        {
            'id': 'l1',
            'label': 'L1 Table',
            'headers': l1_headers,
            'rows': l1_rows,
        },
        {
            'id': 'l2',
            'label': 'L2 Table',
            'headers': l2_headers,
            'rows': l2_rows,
        },
        {
            'id': 'l3',
            'label': 'L3 Table',
            'headers': l3_headers,
            'rows': l3_rows,
        },
        {
            'id': 'attribute',
            'label': 'Attribute',
            'headers': attr_headers,
            'rows': attr_rows,
            'row_colors': attr_row_colors,
        },
    ]

    # --- IP Address_Summary / IP Address_List tabs --------------------------
    # Reuse the already-loaded device_area_map / l3_raw (no extra master I/O).
    # The IP report tabs builder is a pure function that never reads from disk
    # and never imports openpyxl, keeping the Device Table path .nsm-only by
    # construction.
    try:
        try:
            from nsm_ip_report_data import build_ip_report_tabs  # type: ignore
        except ImportError:
            from ns_engine.nsm_ip_report_data import (  # type: ignore
                build_ip_report_tabs,
            )
        area_list = sorted(
            {a for a in device_area_map.values() if a and a != '_N/A_'}
        )
        ip_tabs = build_ip_report_tabs(device_area_map, area_list, l3_raw)
        tabs_data.extend(ip_tabs)
    except Exception as exc:  # pragma: no cover - defensive
        logger.warning('IP report tabs build failed: %s', exc)

    return tabs_data, master_basename


# ---------------------------------------------------------------------------
# Presentation layer
# ---------------------------------------------------------------------------

def _tabs_to_json(tabs_data: List[dict]) -> str:
    """Serialise tabs_data to a JSON literal suitable for embedding directly
    inside a <script> block.

    Uses json.dumps which produces valid JS and properly escapes
    HTML/script-sensitive characters (we additionally escape the
    `</script>` sequence to be safe inside an inline script tag).
    """
    normalised: List[dict] = []
    for tab in tabs_data:
        out_rows = []
        for row in tab.get('rows') or []:
            out_rows.append([
                ('' if v is None else str(v))
                for v in row
            ])
        item = {
            'id': str(tab.get('id', '')),
            'label': str(tab.get('label', '')),
            'headers': [str(h) for h in (tab.get('headers') or [])],
            'rows': out_rows,
        }
        if tab.get('row_colors'):
            colors_out = []
            for crow in tab['row_colors']:
                colors_out.append([
                    (None if c is None else str(c))
                    for c in crow
                ])
            item['row_colors'] = colors_out
        normalised.append(item)

    payload = json.dumps(normalised, ensure_ascii=False)
    # Defence in depth for inline-script embedding.
    return (payload
            .replace('</', '<\\/')
            .replace('\u2028', '\\u2028')
            .replace('\u2029', '\\u2029'))


def render_device_table_html(tabs_data: List[dict],
                              master_basename: str) -> str:
    """Render a self-contained Device Preview HTML page.

    The output matches the Online edition's `/device_preview/<job_id>`
    layout (toolbar, sheet tabs, sticky-header table, CSV/HTML download
    buttons) and works as either:

      - a Flask response body (Online), or
      - a standalone HTML file opened directly in a browser (Local MCP).
    """
    safe_title = f'Device Preview - {master_basename}'
    tabs_js = _tabs_to_json(tabs_data or [])
    master_base_js = json.dumps(str(master_basename), ensure_ascii=False)

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Preview - {safe_title}</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        background: #f0f2f5; color: #333; height: 100vh; display: flex; flex-direction: column; }}
.toolbar {{ display: flex; align-items: center; gap: 12px; padding: 10px 20px;
            background: #16213e; color: #fff; flex-shrink: 0; }}
.toolbar h1 {{ font-size: 14px; font-weight: 500; opacity: 0.9; white-space: nowrap;
               overflow: hidden; text-overflow: ellipsis; max-width: 50vw; }}
.toolbar .spacer {{ flex: 1; }}
.toolbar .row-count {{ font-size: 12px; opacity: 0.6; white-space: nowrap; }}
.toolbar button {{ padding: 6px 16px; border: 1px solid rgba(255,255,255,0.4);
                   background: transparent; color: #fff; border-radius: 6px; font-size: 13px;
                   cursor: pointer; transition: all 0.2s; white-space: nowrap; }}
.toolbar button:hover {{ background: rgba(255,255,255,0.15); border-color: rgba(255,255,255,0.7); }}
.sheet-tabs {{ display: flex; gap: 0; background: #dee2e6; border-bottom: 2px solid #4A8FE7;
               padding: 0 16px; flex-shrink: 0; overflow-x: auto; }}
.sheet-tab {{ padding: 8px 20px; font-size: 13px; cursor: pointer; border: none;
              background: #dee2e6; color: #555; border-radius: 6px 6px 0 0;
              transition: all 0.2s; white-space: nowrap; }}
.sheet-tab:hover {{ background: #e9ecef; }}
.sheet-tab.active {{ background: #fff; color: #4A8FE7; font-weight: 600;
                     border-top: 2px solid #4A8FE7; }}
#content {{ flex: 1; display: flex; flex-direction: column; overflow: hidden; min-height: 0; }}
.table-wrap {{ flex: 1; overflow: scroll; padding: 16px; min-height: 0; }}
.table-wrap::-webkit-scrollbar {{ width: 12px; height: 12px; }}
.table-wrap::-webkit-scrollbar-track {{ background: #e9ecef; border-radius: 6px; }}
.table-wrap::-webkit-scrollbar-thumb {{ background: #a0aec0; border-radius: 6px; border: 2px solid #e9ecef; }}
.table-wrap::-webkit-scrollbar-thumb:hover {{ background: #718096; }}
.table-wrap::-webkit-scrollbar-corner {{ background: #e9ecef; }}
.table-wrap table {{ border-collapse: collapse; width: auto; min-width: 100%;
                     background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.1);
                     border-radius: 4px; overflow: hidden; font-size: 13px; }}
.table-wrap th, .table-wrap td {{ border: 1px solid #e0e0e0; padding: 6px 12px;
                                  text-align: left; white-space: nowrap; }}
.table-wrap tr.header-row th {{
    background: #4A8FE7; color: #fff; font-weight: 600; position: sticky; top: 0; z-index: 1; }}
.table-wrap tr:nth-child(even):not(.header-row) td {{ background: #f8f9fa; }}
.table-wrap tr:hover:not(.header-row) td {{ background: #e8f0fe; }}
</style>
</head>
<body>
<div class="toolbar">
    <h1>{safe_title}</h1>
    <span class="spacer"></span>
    <span class="row-count" id="rowCount"></span>
    <button id="btnDlCsv" title="Download current tab as CSV">&#8681; Download CSV</button>
    <button id="btnDlHtml" title="Download all tabs as a single self-contained HTML">&#8681; Download HTML</button>
</div>
<div class="sheet-tabs" id="sheetTabs"></div>
<div id="content">
    <div class="table-wrap" id="tableWrap"></div>
</div>
<script>
(function() {{
    var TABS = {tabs_js};
    var currentTab = 0;
    var masterBase = {master_base_js};

    function escHtml(s) {{
        return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }}

    function buildTable(tab) {{
        var h = '<table><thead><tr class="header-row">';
        for (var i = 0; i < tab.headers.length; i++) {{
            h += '<th>' + escHtml(tab.headers[i]) + '</th>';
        }}
        h += '</tr></thead><tbody>';
        for (var r = 0; r < tab.rows.length; r++) {{
            h += '<tr>';
            var row = tab.rows[r];
            var colors = tab.row_colors ? tab.row_colors[r] : null;
            var colCount = Math.max(tab.headers.length, row.length);
            for (var c = 0; c < colCount; c++) {{
                var v = c < row.length ? row[c] : '';
                var bg = colors && c < colors.length && colors[c] ? colors[c] : null;
                var style = bg ? ' style="background-color:' + bg + '"' : '';
                h += '<td' + style + '>' + (v !== '' ? escHtml(v) : '') + '</td>';
            }}
            h += '</tr>';
        }}
        h += '</tbody></table>';
        return h;
    }}

    function showTab(idx) {{
        currentTab = idx;
        document.getElementById('tableWrap').innerHTML = buildTable(TABS[idx]);
        document.getElementById('rowCount').textContent = TABS[idx].rows.length + ' rows';
        var btns = document.querySelectorAll('.sheet-tab');
        for (var i = 0; i < btns.length; i++) {{
            btns[i].classList.toggle('active', i === idx);
        }}
    }}

    function initTabs() {{
        var container = document.getElementById('sheetTabs');
        for (var i = 0; i < TABS.length; i++) {{
            (function(idx) {{
                var b = document.createElement('button');
                b.className = 'sheet-tab';
                b.textContent = TABS[idx].label;
                b.onclick = function() {{ showTab(idx); }};
                container.appendChild(b);
            }})(i);
        }}
    }}

    function buildCsv(tab) {{
        var lines = [];
        var h = tab.headers.map(function(c) {{ return '"' + String(c).replace(/"/g,'""') + '"'; }}).join(',');
        lines.push(h);
        for (var r = 0; r < tab.rows.length; r++) {{
            var row = tab.rows[r];
            var cells = [];
            for (var c = 0; c < tab.headers.length; c++) {{
                var v = c < row.length ? row[c] : '';
                cells.push('"' + String(v).replace(/"/g,'""') + '"');
            }}
            lines.push(cells.join(','));
        }}
        return lines.join('\\r\\n');
    }}

    function buildFullHtml() {{
        // Snapshot the live document so the saved file contains all 4 tabs
        // and matches the bytes shipped by Local MCP's export_device_table_html.
        //
        // Reset every element that initTabs() / showTab() populated so the
        // saved file matches a "first load" state. Otherwise:
        //   - sheetTabs would already contain the 4 dynamically created
        //     buttons; on reload the IIFE re-runs initTabs() and appends
        //     ANOTHER 4, giving 8 visible tabs.
        //   - rowCount would have the current "<N> rows" text baked in.
        //   - tableWrap would carry the active-tab table markup.
        var wrap = document.getElementById('tableWrap');
        var tabs = document.getElementById('sheetTabs');
        var rowCount = document.getElementById('rowCount');
        var savedWrap = wrap.innerHTML;
        var savedTabs = tabs.innerHTML;
        var savedRowCount = rowCount.textContent;
        wrap.innerHTML = '';
        tabs.innerHTML = '';
        rowCount.textContent = '';
        var snapshot = '<!DOCTYPE html>\\n' + document.documentElement.outerHTML;
        wrap.innerHTML = savedWrap;
        tabs.innerHTML = savedTabs;
        rowCount.textContent = savedRowCount;
        return snapshot;
    }}

    function download(content, filename, mime) {{
        var bom = mime === 'text/csv' ? '\\uFEFF' : '';
        var blob = new Blob([bom + content], {{type: mime}});
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(a.href);
    }}

    document.getElementById('btnDlCsv').onclick = function() {{
        var tab = TABS[currentTab];
        var safeName = tab.label.replace(/[^a-zA-Z0-9_\\-]/g, '_');
        download(buildCsv(tab), masterBase + '_' + safeName + '.csv', 'text/csv');
    }};

    document.getElementById('btnDlHtml').onclick = function() {{
        // Always download the full 4-tab self-contained HTML so the saved
        // artifact matches Local MCP's [DEVICE_TABLE]<basename>.html output.
        download(buildFullHtml(), '[DEVICE_TABLE]' + masterBase + '.html', 'text/html');
    }};

    initTabs();
    // When viewed as a local file (Local MCP-generated artifact, or an
    // Online-downloaded copy opened from disk), hide the Download HTML
    // button: the file is already on disk and re-downloading would only
    // create nested copies. Online served via http(s):// keeps the button.
    if (window.location.protocol === 'file:') {{
        var dlBtn = document.getElementById('btnDlHtml');
        if (dlBtn) {{ dlBtn.style.display = 'none'; }}
    }}
    // Open the tab specified by the URL hash (e.g. #l1, #l2, #l3, #attribute)
    var hashTabId = window.location.hash.replace('#', '');
    var initIdx = 0;
    if (hashTabId) {{
        for (var hi = 0; hi < TABS.length; hi++) {{
            if (TABS[hi].id === hashTabId) {{ initIdx = hi; break; }}
        }}
    }}
    showTab(initIdx);
}})();
</script>
</body>
</html>'''
