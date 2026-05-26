"""Renderer for the combined L1/L2/L3 diagram HTML viewer.

Produces a self-contained HTML page with three tabs (L1 / L2 / L3) that
share the look-and-feel of ``nsm_device_table_html.render_device_table_html``
(.sheet-tabs CSS, sticky toolbar, hash-based initial tab).

Two modes are supported via the ``mode`` parameter:

  - ``mode='standalone'`` (default): All three SVGs are embedded inline as
    ``<svg>`` markup. The page works opened directly from disk (file://)
    with no server interaction. Toolbar download buttons are hidden in
    this mode because the user already has the file locally and the
    server-side conversion endpoints (Visio / draw.io) are not reachable.

  - ``mode='live'``: The page is intended to be served by the Online
    edition's ``/diagram_preview/<job_id>`` route. Each tab fetches its
    SVG from ``/svg_raw/{job}/{filename}`` on demand. The toolbar shows
    four download buttons:
      * "↓html(L1,L2,L3)" -> /diagram_preview_html/<job>?scope=...
      * "↓svg(for visio)" -> /download_visio/<job>/<active filename>
      * "draw.io"         -> /download_drawio/<job>/<active filename>
      * "draw.io(stencil)"-> /download_drawio_stencil/<job>/<active filename>

The two modes share the bulk of the template; differences are gated
inside the rendered ``<script>`` block via JS conditions.

Used by:
  - ``nsm_cli.py`` ``export combined_diagram`` command (standalone mode)
  - ``ns_web_start.py`` ``/diagram_preview/<job_id>`` route (live mode)
"""
from __future__ import annotations

import json
import re
from typing import Dict, Optional


# Layers rendered in this fixed order. The order also drives tab display order.
_LAYERS = ('l1', 'l2', 'l3')
_LAYER_LABELS = {'l1': 'L1 Diagram', 'l2': 'L2 Diagram', 'l3': 'L3 Diagram'}


def _strip_xml_prolog(svg_text: str) -> str:
    """Drop ``<?xml ...?>`` prolog and standalone DOCTYPE so the SVG can be
    embedded directly inside ``<body>`` without confusing the host page's
    HTML parser. The root ``<svg>`` element is preserved verbatim, including
    its xmlns attributes, so the inlined fragment renders identically to
    the source file.
    """
    if not svg_text:
        return ''
    # Strip leading BOM if present.
    if svg_text.startswith('\ufeff'):
        svg_text = svg_text[1:]
    # Remove any leading <?xml ...?> processing instruction.
    svg_text = re.sub(r'^\s*<\?xml[^?]*\?>\s*', '', svg_text, count=1)
    # Remove any leading <!DOCTYPE ...> declaration.
    svg_text = re.sub(r'^\s*<!DOCTYPE[^>]*>\s*', '', svg_text, count=1)
    return svg_text


def render_l1l2l3_html(
    layer_svgs: Dict[str, Optional[str]],
    master_basename: str,
    scope_label: str,
    *,
    mode: str = 'standalone',
    job_id: Optional[str] = None,
    layer_filenames: Optional[Dict[str, Optional[str]]] = None,
) -> str:
    """Render a self-contained or live L1/L2/L3 tabbed HTML page.

    Args:
        layer_svgs: Mapping ``{'l1': '<svg>...</svg>' or None, 'l2': ..., 'l3': ...}``.
                    ``None`` (or missing key) marks a layer as not generated yet;
                    its tab is rendered as disabled and shows a placeholder.
                    In ``mode='live'`` this argument is allowed to be empty
                    ``{}`` since SVGs are fetched on demand.
        master_basename: Base name of the master file (no extension), used
                         for the page title and download filenames.
        scope_label: Human-readable scope label, e.g. ``"All Areas"`` or an
                     area name. Shown in the toolbar.
        mode: ``'standalone'`` or ``'live'``. Default ``'standalone'``.
        job_id: Required for ``mode='live'``. Used to build /svg_raw and
                /download_visio / /download_drawio / /download_drawio_stencil
                URLs.
        layer_filenames: Required for ``mode='live'``. Mapping
                         ``{'l1': '[L1_DIAGRAM]...svg', ...}`` so the
                         in-page JS can hit the right endpoints per tab.
                         Layers whose value is ``None`` render the tab as
                         disabled.

    Returns:
        Complete HTML document as ``str``.
    """
    if mode not in ('standalone', 'live'):
        raise ValueError(f"mode must be 'standalone' or 'live', got {mode!r}")
    if mode == 'live':
        if not job_id:
            raise ValueError("job_id is required for mode='live'")
        if layer_filenames is None:
            raise ValueError("layer_filenames is required for mode='live'")

    layer_svgs = dict(layer_svgs or {})
    layer_filenames = dict(layer_filenames or {})

    # Per-layer availability flags. In standalone mode, a layer is available
    # iff the inlined SVG body is present; in live mode it is available iff
    # a filename was provided (the actual fetch can still 404 -- see the
    # in-page JS fallback).
    if mode == 'standalone':
        availability = {layer: bool(layer_svgs.get(layer)) for layer in _LAYERS}
    else:
        availability = {layer: bool(layer_filenames.get(layer)) for layer in _LAYERS}

    safe_title = f'L1L2L3 Diagram - {master_basename} ({scope_label})'
    safe_title_js = json.dumps(safe_title, ensure_ascii=False)
    master_base_js = json.dumps(str(master_basename), ensure_ascii=False)
    scope_label_js = json.dumps(str(scope_label), ensure_ascii=False)
    job_id_js = json.dumps(str(job_id) if job_id else '', ensure_ascii=False)
    mode_js = json.dumps(mode, ensure_ascii=False)
    availability_js = json.dumps(availability, ensure_ascii=False)
    layer_filenames_js = json.dumps(
        {k: (v or '') for k, v in layer_filenames.items()}, ensure_ascii=False
    )

    # Build the inline SVG sections (standalone mode only). Each layer's
    # SVG is wrapped in a hidden <div> so showTab() can simply toggle the
    # ``active`` class and the active layer becomes visible. We always emit
    # the three wrappers (even in live mode) so the JS layout stays identical;
    # in live mode they are populated dynamically via fetch.
    layer_sections = []
    for layer in _LAYERS:
        body = ''
        if mode == 'standalone' and availability[layer]:
            body = _strip_xml_prolog(layer_svgs[layer] or '')
        layer_sections.append(
            f'<div class="layer-content" data-layer="{layer}" id="layer-{layer}">{body}</div>'
        )
    layers_html = '\n'.join(layer_sections)

    # Toolbar layout (mirrors _render_svg_viewer):
    #   <zoom group>  [sep]  <download group (live only)>
    # The zoom group is always emitted so the user can pan/zoom the
    # standalone HTML viewer too. The download group is gated on live mode
    # because the conversion endpoints (Visio / draw.io / draw.io-stencil)
    # are only reachable from the Online edition.
    zoom_buttons = (
        '<button id="btnZoomOut" title="Zoom Out">-</button>\n'
        '<span class="zoom-display" id="zoomLevel">100%</span>\n'
        '<button id="btnZoomIn" title="Zoom In">+</button>\n'
        '<button id="btnFit" title="Fit to Window">Fit</button>'
    )
    if mode == 'live':
        download_buttons = (
            '<span class="sep"></span>\n'
            '<button id="btnDlAll" title="Download all 3 layers as a single HTML">'
            '&#8681; html(L1,L2,L3)</button>\n'
            '<button id="btnDlVisio" title="Download active layer for Visio">'
            '&#8681; svg(for visio)</button>\n'
            '<button id="btnDlDrawio" title="Download active layer as draw.io">'
            '&#8681; draw.io</button>\n'
            '<button id="btnDlDrawioStencil" '
            'title="Download active layer as draw.io with Cisco stencils">'
            '&#8681; draw.io(stencil)</button>'
        )
    else:
        download_buttons = ''
    toolbar_buttons = zoom_buttons + ('\n' + download_buttons if download_buttons else '')

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{safe_title}</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        background: #f0f2f5; color: #333; height: 100vh; display: flex; flex-direction: column; }}
.toolbar {{ display: flex; align-items: center; gap: 12px; padding: 10px 20px;
            background: #16213e; color: #fff; flex-shrink: 0; }}
.toolbar h1 {{ font-size: 14px; font-weight: 500; opacity: 0.9; white-space: nowrap;
               overflow: hidden; text-overflow: ellipsis; max-width: 50vw; }}
.toolbar .spacer {{ flex: 1; }}
.toolbar button {{ padding: 6px 16px; border: 1px solid rgba(255,255,255,0.4);
                   background: transparent; color: #fff; border-radius: 6px; font-size: 13px;
                   cursor: pointer; transition: all 0.2s; white-space: nowrap; }}
.toolbar button:hover {{ background: rgba(255,255,255,0.15); border-color: rgba(255,255,255,0.7); }}
.toolbar button:disabled {{ opacity: 0.4; cursor: not-allowed; }}
.toolbar .sep {{ width: 1px; height: 20px; background: rgba(255,255,255,0.15); }}
.toolbar .zoom-display {{ font-size: 12px; opacity: 0.6; min-width: 45px; text-align: center; }}
.sheet-tabs {{ display: flex; gap: 0; background: #dee2e6; border-bottom: 2px solid #4A8FE7;
               padding: 0 16px; flex-shrink: 0; overflow-x: auto; }}
.sheet-tab {{ padding: 8px 20px; font-size: 13px; cursor: pointer; border: none;
              background: #dee2e6; color: #555; border-radius: 6px 6px 0 0;
              transition: all 0.2s; white-space: nowrap; }}
.sheet-tab:hover {{ background: #e9ecef; }}
.sheet-tab.active {{ background: #fff; color: #4A8FE7; font-weight: 600;
                     border-top: 2px solid #4A8FE7; }}
.sheet-tab.disabled {{ opacity: 0.45; cursor: not-allowed; }}
#content {{ flex: 1; display: flex; flex-direction: column; overflow: hidden; min-height: 0;
            background: #fff; }}
/* Pan/zoom viewer (mirrors _render_svg_viewer in ns_web_start.py).
   #viewer is the fixed window; .layer-content is positioned absolutely and
   transformed via JS so wheel/drag manipulate the active layer's geometry
   directly, matching the original single-SVG viewer's UX. */
.viewer {{ flex: 1; overflow: hidden; cursor: grab; position: relative; background: #f0f2f5;
           min-height: 0; }}
.viewer.dragging {{ cursor: grabbing; }}
.layer-content {{ position: absolute; left: 0; top: 0;
                  user-select: none; -webkit-user-select: none;
                  display: none; }}
.layer-content.active {{ display: block; }}
.layer-content svg {{ display: block; overflow: visible; }}
.layer-placeholder {{ position: absolute; left: 50%; top: 50%; transform: translate(-50%, -50%);
                      display: flex; align-items: center; justify-content: center;
                      color: #888; font-size: 14px; flex-direction: column;
                      gap: 12px; padding: 24px; text-align: center; }}
.layer-placeholder .icon {{ font-size: 48px; opacity: 0.3; }}
.layer-loading {{ position: absolute; left: 50%; top: 50%; transform: translate(-50%, -50%);
                  color: #666; font-size: 14px; }}
</style>
</head>
<body>
<div class="toolbar">
    <h1 id="titleText">{safe_title}</h1>
    <span class="spacer"></span>
    {toolbar_buttons}
</div>
<div class="sheet-tabs" id="sheetTabs"></div>
<div id="content">
    <div class="viewer" id="viewer">
        {layers_html}
    </div>
</div>
<script>
(function() {{
    var MODE = {mode_js};                  // 'standalone' or 'live'
    var JOB_ID = {job_id_js};
    var MASTER_BASE = {master_base_js};
    var SCOPE_LABEL = {scope_label_js};
    var AVAILABILITY = {availability_js};   // {{l1: bool, l2: bool, l3: bool}}
    var LAYER_FILES = {layer_filenames_js}; // {{l1: filename, ...}} -- live mode only
    var LAYER_LABELS = {{l1: 'L1 Diagram', l2: 'L2 Diagram', l3: 'L3 Diagram'}};
    var LAYERS = ['l1', 'l2', 'l3'];

    var currentLayer = null;
    var fetchedLive = {{}}; // memoize fetched SVG bodies in live mode

    // Per-layer pan/zoom state. Each tab keeps its own scale + center so
    // switching tabs does not reset what the user was looking at on the
    // previous layer. ``fitted`` flips to true after the first
    // fitToWindow() so re-activating the tab restores the user's last
    // pan/zoom rather than re-fitting.
    var perLayerState = {{
        l1: {{scale: 1, cx: 0, cy: 0, naturalW: 800, naturalH: 600, fitted: false}},
        l2: {{scale: 1, cx: 0, cy: 0, naturalW: 800, naturalH: 600, fitted: false}},
        l3: {{scale: 1, cx: 0, cy: 0, naturalW: 800, naturalH: 600, fitted: false}}
    }};

    // Drag state shared across layers; mousedown captures cxStart/cyStart
    // from the active layer's perLayerState entry.
    var dragging = false;
    var dragStartX = 0, dragStartY = 0;
    var cxStart = 0, cyStart = 0;

    var viewerEl = document.getElementById('viewer');
    var zoomEl = document.getElementById('zoomLevel');

    function escHtml(s) {{
        return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }}

    // --- Pan / zoom helpers (mirror _render_svg_viewer) -------------------

    function activeLayerEl() {{
        return currentLayer ? document.getElementById('layer-' + currentLayer) : null;
    }}

    // Read the active layer SVG's intrinsic size (viewBox preferred, then
    // width/height attributes, then a sensible 800x600 default) into
    // perLayerState so updateTransform()/fitToWindow() can scale it.
    function measureSvgIntoState(layer) {{
        var host = document.getElementById('layer-' + layer);
        if (!host) return;
        var svgEl = host.querySelector('svg');
        if (!svgEl) return;
        var st = perLayerState[layer];
        var vb = svgEl.viewBox && svgEl.viewBox.baseVal;
        if (vb && vb.width && vb.height) {{
            st.naturalW = vb.width;
            st.naturalH = vb.height;
        }} else {{
            st.naturalW = parseFloat(svgEl.getAttribute('width')) || 800;
            st.naturalH = parseFloat(svgEl.getAttribute('height')) || 600;
        }}
    }}

    // Apply the active layer's scale + center to its absolutely-positioned
    // wrapper. ``cx/cy`` are coordinates in the SVG's own space that are
    // pinned to the viewer center; this matches the original viewer's
    // wheel-to-cursor zoom anchoring.
    function updateTransform() {{
        if (!currentLayer) return;
        var host = activeLayerEl();
        if (!host) return;
        var st = perLayerState[currentLayer];
        var vw = viewerEl.clientWidth, vh = viewerEl.clientHeight;
        var iw = st.naturalW * st.scale;
        var ih = st.naturalH * st.scale;
        var left = vw / 2 - st.cx * st.scale;
        var top = vh / 2 - st.cy * st.scale;
        host.style.left = left + 'px';
        host.style.top = top + 'px';
        host.style.width = iw + 'px';
        host.style.height = ih + 'px';
        var svgEl = host.querySelector('svg');
        if (svgEl) {{
            svgEl.style.width = iw + 'px';
            svgEl.style.height = ih + 'px';
        }}
        if (zoomEl) zoomEl.textContent = Math.round(st.scale * 100) + '%';
    }}

    // Reset the active layer to fit-to-window, anchored at the SVG center.
    // Caps scale at 1.0 (no upscaling small SVGs) and applies a 0.95
    // breathing-room factor like the original viewer.
    function fitToWindow() {{
        if (!currentLayer) return;
        var st = perLayerState[currentLayer];
        var vw = viewerEl.clientWidth, vh = viewerEl.clientHeight;
        if (vw > 0 && vh > 0 && st.naturalW > 0 && st.naturalH > 0) {{
            st.scale = Math.min(vw / st.naturalW, vh / st.naturalH, 1) * 0.95;
        }}
        st.cx = st.naturalW / 2;
        st.cy = st.naturalH / 2;
        st.fitted = true;
        updateTransform();
    }}

    // Centered zoom (used by the +/- buttons). Scale is clamped to the
    // same [0.02, 80] range as the original viewer so very large diagrams
    // stay zoomable and very small ones can still be inspected closely.
    function zoomCenter(factor) {{
        if (!currentLayer) return;
        var st = perLayerState[currentLayer];
        var newScale = Math.max(0.02, Math.min(80, st.scale * factor));
        st.scale = newScale;
        updateTransform();
    }}

    // Cursor-anchored zoom (used by the wheel handler). The point under
    // the cursor stays put while the rest of the SVG scales around it.
    function zoomAtPoint(px, py, factor) {{
        if (!currentLayer) return;
        var st = perLayerState[currentLayer];
        var newScale = Math.max(0.02, Math.min(80, st.scale * factor));
        var vw = viewerEl.clientWidth, vh = viewerEl.clientHeight;
        var dx = px - vw / 2, dy = py - vh / 2;
        st.cx += dx * (1 / st.scale - 1 / newScale);
        st.cy += dy * (1 / st.scale - 1 / newScale);
        st.scale = newScale;
        updateTransform();
    }}

    function renderPlaceholder(layer, message) {{
        return '<div class="layer-placeholder">'
            + '<div class="icon">&#9888;</div>'
            + '<div>' + escHtml(message) + '</div>'
            + '</div>';
    }}

    // Apply pan/zoom for the active layer: measure the SVG, then either
    // fit-to-window (first activation) or restore the user's previous
    // pan/zoom (re-activation). No-op if the active host has no <svg>
    // (placeholder/loading state).
    function applyLayerView(layer) {{
        var host = document.getElementById('layer-' + layer);
        if (!host) return;
        var svgEl = host.querySelector('svg');
        if (!svgEl) return;
        measureSvgIntoState(layer);
        var st = perLayerState[layer];
        if (!st.fitted) {{
            fitToWindow();
        }} else {{
            updateTransform();
        }}
    }}

    function showTab(layer) {{
        if (LAYERS.indexOf(layer) < 0) return;
        currentLayer = layer;
        // Toggle active class on tab buttons.
        var btns = document.querySelectorAll('.sheet-tab');
        for (var i = 0; i < btns.length; i++) {{
            btns[i].classList.toggle('active', btns[i].dataset.layer === layer);
        }}
        // Toggle active class on layer content wrappers.
        var contents = document.querySelectorAll('.layer-content');
        for (var j = 0; j < contents.length; j++) {{
            contents[j].classList.toggle('active', contents[j].dataset.layer === layer);
        }}
        // In live mode, fetch the SVG on first activation. ensureLiveSvg
        // calls applyLayerView() once the body is in the DOM.
        if (MODE === 'live') {{
            ensureLiveSvg(layer);
        }} else if (!AVAILABILITY[layer]) {{
            // Standalone mode and the layer wasn't bundled: show placeholder.
            var host = document.getElementById('layer-' + layer);
            if (host && !host.dataset.placeheld) {{
                host.innerHTML = renderPlaceholder(layer,
                    LAYER_LABELS[layer] + ' diagram is not generated yet. '
                    + "Click 'Generate Selected' (Online) or run "
                    + "'export l" + layer.charAt(1) + "_diagram' (CLI) to build it.");
                host.dataset.placeheld = '1';
            }}
        }} else {{
            // Standalone mode with bundled SVG: apply pan/zoom directly
            // since the markup is already in the DOM.
            applyLayerView(layer);
        }}
        if (history && history.replaceState) {{
            try {{ history.replaceState(null, '', '#' + layer); }} catch (e) {{}}
        }}
    }}

    function ensureLiveSvg(layer) {{
        var host = document.getElementById('layer-' + layer);
        if (!host) return;
        if (fetchedLive[layer]) {{
            // Already in DOM (or placeholder). Re-apply pan/zoom for this
            // tab activation so the view restores correctly.
            applyLayerView(layer);
            return;
        }}
        var fname = LAYER_FILES[layer];
        if (!fname) {{
            host.innerHTML = renderPlaceholder(layer,
                LAYER_LABELS[layer] + ' diagram is not generated yet. '
                + "Click 'Generate Selected' to build this layer.");
            fetchedLive[layer] = true;
            return;
        }}
        host.innerHTML = '<div class="layer-loading">Loading ' + LAYER_LABELS[layer] + '...</div>';
        var url = '/svg_raw/' + encodeURIComponent(JOB_ID) + '/' + encodeURIComponent(fname);
        fetch(url).then(function(r) {{
            if (!r.ok) throw new Error('HTTP ' + r.status);
            return r.text();
        }}).then(function(svg) {{
            // Strip an XML prolog if present so the SVG plays nicely with HTML
            // parsing. The server responds with 'image/svg+xml' but the bytes
            // may begin with <?xml ...?>.
            svg = svg.replace(/^\\s*<\\?xml[^?]*\\?>\\s*/, '')
                     .replace(/^\\s*<!DOCTYPE[^>]*>\\s*/, '');
            host.innerHTML = svg;
            fetchedLive[layer] = true;
            // Apply pan/zoom now that the SVG is present in the DOM. Only
            // apply if this is still the active layer to avoid scaling a
            // background tab and thereby clobbering state for the visible
            // one.
            if (currentLayer === layer) applyLayerView(layer);
        }}).catch(function(err) {{
            host.innerHTML = renderPlaceholder(layer,
                'Failed to load ' + LAYER_LABELS[layer]
                + ' diagram (' + err.message + '). The file may not be generated yet.');
            fetchedLive[layer] = true;
        }});
    }}

    function initTabs() {{
        var container = document.getElementById('sheetTabs');
        for (var i = 0; i < LAYERS.length; i++) {{
            (function(layer) {{
                var b = document.createElement('button');
                b.className = 'sheet-tab';
                b.dataset.layer = layer;
                b.textContent = LAYER_LABELS[layer];
                if (!AVAILABILITY[layer] && MODE === 'standalone') {{
                    b.classList.add('disabled');
                }}
                b.onclick = function() {{ showTab(layer); }};
                container.appendChild(b);
            }})(LAYERS[i]);
        }}
    }}

    // ----- Live-mode download buttons (no-ops in standalone mode) -----
    function triggerBlobDownload(blob, filename) {{
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }}

    function activeFilename() {{
        return LAYER_FILES[currentLayer] || '';
    }}

    function downloadConverted(prefix, suffix) {{
        var fname = activeFilename();
        if (!fname) {{
            alert(LAYER_LABELS[currentLayer] + ' diagram is not generated yet.');
            return;
        }}
        var url = '/' + prefix + '/' + encodeURIComponent(JOB_ID) + '/' + encodeURIComponent(fname);
        fetch(url).then(function(r) {{
            if (!r.ok) throw new Error('HTTP ' + r.status);
            return r.blob();
        }}).then(function(blob) {{
            triggerBlobDownload(blob, fname.replace(/\\.svg$/i, suffix));
        }}).catch(function(err) {{
            alert('Download failed: ' + err.message);
        }});
    }}

    if (MODE === 'live') {{
        var btnDlAll = document.getElementById('btnDlAll');
        if (btnDlAll) {{
            btnDlAll.onclick = function() {{
                var url = '/diagram_preview_html/' + encodeURIComponent(JOB_ID)
                    + '?scope=' + encodeURIComponent(SCOPE_LABEL_TO_PARAM());
                window.location.href = url;
            }};
        }}
        var btnDlVisio = document.getElementById('btnDlVisio');
        if (btnDlVisio) {{
            btnDlVisio.onclick = function() {{ downloadConverted('download_visio', '_visio.svg'); }};
        }}
        var btnDlDrawio = document.getElementById('btnDlDrawio');
        if (btnDlDrawio) {{
            btnDlDrawio.onclick = function() {{ downloadConverted('download_drawio', '.drawio'); }};
        }}
        var btnDlStencil = document.getElementById('btnDlDrawioStencil');
        if (btnDlStencil) {{
            btnDlStencil.onclick = function() {{ downloadConverted('download_drawio_stencil', '_stencil.drawio'); }};
        }}
    }}

    // The /diagram_preview_html route accepts ?scope=all or ?scope=area:<name>.
    // We derive that param from SCOPE_LABEL: 'All Areas' => 'all', else 'area:<label>'.
    function SCOPE_LABEL_TO_PARAM() {{
        if (SCOPE_LABEL === 'All Areas') return 'all';
        return 'area:' + SCOPE_LABEL;
    }}

    // --- Pan/zoom event wiring (mirrors _render_svg_viewer) ---------------
    function setupPanZoom() {{
        if (!viewerEl) return;
        // Wheel zoom anchored at the cursor. preventDefault stops the page
        // from scrolling under the viewer when the wheel is used.
        viewerEl.addEventListener('wheel', function(e) {{
            e.preventDefault();
            var rect = viewerEl.getBoundingClientRect();
            var px = e.clientX - rect.left;
            var py = e.clientY - rect.top;
            var factor = (e.deltaY < 0) ? 1.15 : (1 / 1.15);
            zoomAtPoint(px, py, factor);
        }}, {{passive: false}});

        // Left-button drag panning. Right/middle buttons are ignored so
        // browser default behavior (context menu, autoscroll) still works.
        viewerEl.addEventListener('mousedown', function(e) {{
            if (e.button !== 0) return;
            if (!currentLayer) return;
            dragging = true;
            viewerEl.classList.add('dragging');
            dragStartX = e.clientX;
            dragStartY = e.clientY;
            var st = perLayerState[currentLayer];
            cxStart = st.cx;
            cyStart = st.cy;
            e.preventDefault();
        }});

        // mousemove/mouseup are bound on window so the drag continues even
        // if the cursor leaves the viewer.
        window.addEventListener('mousemove', function(e) {{
            if (!dragging || !currentLayer) return;
            var st = perLayerState[currentLayer];
            var dx = e.clientX - dragStartX;
            var dy = e.clientY - dragStartY;
            st.cx = cxStart - dx / st.scale;
            st.cy = cyStart - dy / st.scale;
            updateTransform();
        }});
        window.addEventListener('mouseup', function(e) {{
            if (!dragging) return;
            dragging = false;
            viewerEl.classList.remove('dragging');
        }});
        // Recompute layout on resize so the active layer stays centered.
        window.addEventListener('resize', function() {{
            updateTransform();
        }});

        // Toolbar zoom buttons (always present; available in both modes).
        var btnZoomIn = document.getElementById('btnZoomIn');
        if (btnZoomIn) btnZoomIn.onclick = function() {{ zoomCenter(1.25); }};
        var btnZoomOut = document.getElementById('btnZoomOut');
        if (btnZoomOut) btnZoomOut.onclick = function() {{ zoomCenter(1 / 1.25); }};
        var btnFit = document.getElementById('btnFit');
        if (btnFit) btnFit.onclick = function() {{ fitToWindow(); }};
    }}

    initTabs();
    setupPanZoom();
    // Initial tab from URL hash (#l1 / #l2 / #l3); fallback to first available.
    var hashLayer = (window.location.hash || '').replace('#', '');
    var initLayer = null;
    if (LAYERS.indexOf(hashLayer) >= 0) {{
        initLayer = hashLayer;
    }} else {{
        for (var i = 0; i < LAYERS.length; i++) {{
            if (AVAILABILITY[LAYERS[i]] || MODE === 'live') {{
                initLayer = LAYERS[i];
                break;
            }}
        }}
        if (!initLayer) initLayer = LAYERS[0];
    }}
    showTab(initLayer);
}})();
</script>
</body>
</html>'''
