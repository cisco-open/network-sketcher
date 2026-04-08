'''
SPDX-License-Identifier: Apache-2.0

Copyright 2023 Cisco Systems, Inc. and its affiliates

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
SVG rendering engine for Network Sketcher L1 diagrams.

Mirrors the drawing logic of nsm_ddx_figure.py but outputs SVG.
Uses ThreadPoolExecutor for parallel string generation and
NumPy/CuPy for vectorised coordinate math.
"""

import os
import math
from concurrent.futures import ThreadPoolExecutor

import nsm_svg_compute as sc

DPI = 96
PT_TO_PX = DPI / 72.0
FONT_FAMILY = 'Calibri, Segoe UI, Arial, sans-serif'


def _in(inches):
    """Convert inches to SVG pixels."""
    return round(inches * DPI, 2)


def _pt(points):
    """Convert points to SVG pixels."""
    return round(points * PT_TO_PX, 2)


def _rgb(r, g, b):
    return f'rgb({int(r)},{int(g)},{int(b)})'


def _escape(text):
    """Escape text for SVG XML."""
    if text is None:
        return ''
    s = str(text)
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def _render_device_chunk(args):
    """Worker function for parallel device rendering.

    Called from ThreadPoolExecutor with (devices, style_index, attribute_colors).
    """
    devices, style_index, attribute_colors = args
    parts = []
    for dev in devices:
        name = dev['name']
        display_name = dev['display_name']
        x, y, w, h = dev['x'], dev['y'], dev['w'], dev['h']
        roundness = dev.get('roundness', 0.0)

        is_air = '_AIR_' in str(name)

        rx = round(roundness * min(w, h) * 0.5, 2)
        fill_color = sc.resolve_device_color(name, style_index, attribute_colors)

        px, py, pw, ph = _in(x), _in(y), _in(w), _in(h)

        if is_air:
            parts.append(
                f'<rect x="{px}" y="{py}" width="{pw}" height="{ph}" '
                f'rx="{_in(rx)}" ry="{_in(rx)}" '
                f'class="device-air"/>\n')
        else:
            fill = _rgb(*fill_color) if fill_color else 'none'
            parts.append(
                f'<rect x="{px}" y="{py}" width="{pw}" height="{ph}" '
                f'rx="{_in(rx)}" ry="{_in(rx)}" '
                f'fill="{fill}" class="device"/>\n')
            tx = px + pw / 2
            ty = py + ph / 2
            parts.append(
                f'<text x="{round(tx, 2)}" y="{round(ty, 2)}" '
                f'class="device-text">{_escape(display_name)}</text>\n')
    return ''.join(parts)


def _svg_header(slide_w, slide_h, shape_font_size, folder_font_size):
    w = _in(slide_w)
    h = _in(slide_h)
    return (
        f'<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{w}" height="{h}" viewBox="0 0 {w} {h}">\n'
        f'<defs>\n'
        f'<style>\n'
        f'  .bg {{ fill: white; }}\n'
        f'  .root {{ fill: white; stroke: black; stroke-width: {_pt(1.5)}px; }}\n'
        f'  .folder {{ fill: none; stroke: black; stroke-width: {_pt(1.5)}px; }}\n'
        f'  .device {{ stroke: black; stroke-width: {_pt(1.0)}px; }}\n'
        f'  .device-air {{ fill: white; stroke: white; stroke-width: {_pt(1.0)}px; }}\n'
        f'  .line {{ stroke: black; stroke-width: {_pt(0.5)}px; fill: none; }}\n'
        f'  .channel {{ fill: none; stroke: black; stroke-width: {_pt(0.3)}px; }}\n'
        f'  .tag {{ fill: white; stroke: black; stroke-width: {_pt(0.1)}px; }}\n'
        f'  .bullet {{ fill: black; stroke: none; }}\n'
        f'  .segment-line {{ stroke: black; stroke-width: {_pt(0.8)}px; fill: none; }}\n'
        f'  text {{ font-family: {FONT_FAMILY}; }}\n'
        f'  .title-text {{ font-size: {_pt(18)}px; fill: black; }}\n'
        f'  .folder-text {{ font-size: {_pt(folder_font_size)}px; fill: black; text-anchor: middle; }}\n'
        f'  .device-text {{ font-size: {_pt(shape_font_size)}px; fill: black; '
        f'text-anchor: middle; dominant-baseline: central; }}\n'
        f'  .tag-text {{ font-size: {_pt(6)}px; fill: black; text-anchor: middle; '
        f'dominant-baseline: central; }}\n'
        f'  .segment-text {{ font-size: {_pt(8)}px; fill: black; text-anchor: end; '
        f'dominant-baseline: auto; }}\n'
        f'</style>\n'
        f'</defs>\n'
        f'<rect width="{w}" height="{h}" class="bg"/>\n'
    )


def _render_title(text, left, top):
    return (f'<text x="{_in(left)}" y="{_in(top)}" '
            f'class="title-text">{_escape(text)}</text>\n')


def _render_root_folder(left, top, width, height):
    return (f'<rect x="{_in(left)}" y="{_in(top)}" '
            f'width="{_in(width)}" height="{_in(height)}" class="root"/>\n')


def _render_sub_folder(left, top, width, height,
                       label, text_pos, font_size):
    rx = _in(0.1 * min(width, height) * 0.5)
    parts = [f'<rect x="{_in(left)}" y="{_in(top)}" '
             f'width="{_in(width)}" height="{_in(height)}" '
             f'rx="{rx}" ry="{rx}" class="folder"/>\n']

    if label and text_pos:
        tx = _in(left + width * 0.5)
        if text_pos == 'UP':
            ty = _in(top + 0.15)
            parts.append(f'<text x="{tx}" y="{ty}" '
                         f'class="folder-text" dominant-baseline="hanging">'
                         f'{_escape(label)}</text>\n')
        elif text_pos == 'DOWN':
            ty = _in(top + height - 0.05)
            parts.append(f'<text x="{tx}" y="{ty}" '
                         f'class="folder-text" dominant-baseline="auto">'
                         f'{_escape(label)}</text>\n')
    return ''.join(parts)


def _compute_devices_in_folder(coord_list, device_grid, f_left, f_top, f_width, f_height,
                               style_index, attribute_colors):
    """Compute device positions within a folder. Returns list of device dicts."""
    devices = []

    total_row_height = 0.0
    row_heights = []
    for row in device_grid:
        max_h = 0.0
        for dev_name in row:
            _, h, _ = sc.get_shape_dims(dev_name, style_index)
            if h > max_h:
                max_h = h
        row_heights.append(max_h)
        total_row_height += max_h

    num_rows = len(device_grid)
    v_margin = (f_height - total_row_height) / (num_rows + 1) if num_rows > 0 else 0.0

    cur_top = f_top
    for r_idx, row in enumerate(device_grid):
        cur_top += v_margin
        row_h = row_heights[r_idx]

        total_row_width = 0.0
        for dev_name in row:
            w, _, _ = sc.get_shape_dims(dev_name, style_index)
            total_row_width += w

        num_cols = len(row)
        h_margin = (f_width - total_row_width) / (num_cols + 1) if num_cols > 0 else 0.0

        cur_left = f_left
        for dev_name in row:
            cur_left += h_margin
            w, h, roundness = sc.get_shape_dims(dev_name, style_index)

            display_name = dev_name
            tag_name = dev_name
            if '<' in dev_name:
                display_name = dev_name.split('<')[0]
                tag_name = '<' + dev_name.split('<')[-1].split('>')[0] + '>'

            is_segment = dev_name.startswith('<SEGMENT>')
            if is_segment:
                cur_left += w
                continue

            is_air = '_AIR_' in dev_name
            if dev_name and not is_segment:
                dev_dict = {
                    'name': tag_name,
                    'display_name': display_name,
                    'x': cur_left, 'y': cur_top,
                    'w': w, 'h': h,
                    'roundness': roundness,
                }
                devices.append(dev_dict)

                if not is_air:
                    x_l = cur_left
                    x_m = cur_left + w * 0.5
                    x_r = cur_left + w
                    y_t = cur_top
                    y_m = cur_top + h * 0.5
                    y_d = cur_top + h
                    coord_list.append([tag_name, x_l, x_m, x_r, y_t, y_m, y_d])

            cur_left += w
        cur_top += row_h

    return devices


def _render_lines_vectorised(coord_list, line_records, tag_type, tag_overrides, font_size):
    """Render connection lines, channel marks, and tags using vectorised math.

    NumPy/CuPy batch-computes all angles and tag offsets before generating SVG.
    """
    import numpy as np

    if not line_records:
        return ''

    per_char_width = 0.045
    font_size_height = 0.014
    tag_font_size = 6
    size_bullet = 0.04

    coord_map = {}
    for entry in coord_list:
        coord_map[entry[0]] = entry

    n = len(line_records)
    x1_arr = np.empty(n, dtype=np.float64)
    y1_arr = np.empty(n, dtype=np.float64)
    x2_arr = np.empty(n, dtype=np.float64)
    y2_arr = np.empty(n, dtype=np.float64)
    valid = np.zeros(n, dtype=bool)

    for i, rec in enumerate(line_records):
        from_entry = coord_map.get(rec['from_name'])
        to_entry = coord_map.get(rec['to_name'])
        if not from_entry or not to_entry:
            continue

        from_cx, to_cx = 2, 2
        from_cy, to_cy = 5, 5

        if from_entry[6] >= to_entry[4]:
            from_cy = 4
            to_cy = 6
        else:
            from_cy = 6
            to_cy = 4

        if rec['from_side'] == 'RIGHT':
            from_cx, from_cy = 3, 5
        elif rec['from_side'] == 'LEFT':
            from_cx, from_cy = 1, 5
        if rec['to_side'] == 'RIGHT':
            to_cx, to_cy = 3, 5
        elif rec['to_side'] == 'LEFT':
            to_cx, to_cy = 1, 5

        x1 = float(from_entry[from_cx])
        y1 = float(from_entry[from_cy])
        x2 = float(to_entry[to_cx])
        y2 = float(to_entry[to_cy])

        if rec['from_offset_x'] is not None:
            x1 += rec['from_offset_x']
        if rec['from_offset_y'] is not None:
            y1 += rec['from_offset_y']

        to_ox = rec['to_offset_x']
        to_oy = rec['to_offset_y']
        if to_ox is not None:
            if str(to_ox) == '<FROM_X>':
                x2 = x1
            elif isinstance(to_ox, (int, float)):
                x2 += to_ox
        if to_oy is not None:
            if str(to_oy) == '<FROM_Y>':
                y2 = y1
            elif isinstance(to_oy, (int, float)):
                y2 += to_oy

        x1_arr[i] = x1
        y1_arr[i] = y1
        x2_arr[i] = x2
        y2_arr[i] = y2
        valid[i] = True

    angles_deg = sc.batch_compute_angles(
        x1_arr[valid], y1_arr[valid], x2_arr[valid], y2_arr[valid])

    angle_lookup = np.zeros(n, dtype=np.float64)
    angle_lookup[valid] = angles_deg

    parts = []
    for i, rec in enumerate(line_records):
        if not valid[i]:
            continue

        x1, y1, x2, y2 = x1_arr[i], y1_arr[i], x2_arr[i], y2_arr[i]
        deg = angle_lookup[i]

        if rec['visible'] != 'NO':
            parts.append(f'<line x1="{_in(x1)}" y1="{_in(y1)}" '
                         f'x2="{_in(x2)}" y2="{_in(y2)}" class="line"/>\n')

        if rec['channel_height'] is not None:
            ch = rec['channel_height']
            cx = x1 + (x2 - x1) * 0.5
            cy = y1 + (y2 - y1) * 0.5
            ow = 0.1
            parts.append(
                f'<ellipse cx="{_in(cx)}" cy="{_in(cy)}" '
                f'rx="{_in(ow * 0.5)}" ry="{_in(ch * 0.5)}" '
                f'transform="rotate({round(deg, 2)} {_in(cx)} {_in(cy)})" '
                f'class="channel"/>\n')

        for side in ('From', 'To'):
            tag_text = rec['from_tag'] if side == 'From' else rec['to_tag']
            if not tag_text or tag_text == 'None':
                continue

            target_name = rec['from_name'] if side == 'From' else rec['to_name']
            tx = x1 if side == 'From' else x2
            ty = y1 if side == 'From' else y2

            tw = sc.get_east_asian_width_count(tag_text) * per_char_width
            th = tag_font_size * font_size_height

            override = tag_overrides.get(target_name, {})
            cur_tag_type = override.get('type', tag_type)
            tag_deg = 0

            if cur_tag_type == 'SHAPE':
                tl = tx - tw * 0.5
                tt = ty - th * 0.5
                if override.get('offset_x') is not None:
                    tl += override['offset_x']
                if override.get('offset_y') is not None:
                    tt += override['offset_y']
            elif cur_tag_type == 'LINE':
                tl = tx - tw * 0.5
                tt = ty - th * 0.5
                offset_inch = 0.02
                line_offset = len(tag_text) * offset_inch

                if override.get('line_offset') is not None:
                    cos_v = math.cos(math.radians(deg))
                    sin_v = math.sin(math.radians(deg))
                    if side == 'From':
                        tl += cos_v * line_offset
                        tt += sin_v * line_offset
                    else:
                        tl -= cos_v * line_offset
                        tt -= sin_v * line_offset

                if override.get('rotation') == 'YES':
                    tag_deg = deg
                    if (x2 - x1) < 0:
                        tag_deg += 180
            else:
                tl = tx - tw * 0.5
                tt = ty - th * 0.5

            if tag_text == '<BULLET>':
                bx = tx
                by = ty
                parts.append(
                    f'<rect x="{_in(bx - size_bullet * 0.5)}" '
                    f'y="{_in(by - size_bullet * 0.5)}" '
                    f'width="{_in(size_bullet)}" height="{_in(size_bullet)}" '
                    f'rx="{_in(size_bullet * 0.5)}" ry="{_in(size_bullet * 0.5)}" '
                    f'class="bullet"/>\n')
            else:
                rx = _in(min(tw, th) * 0.5 * 0.99)
                transform = ''
                if tag_deg != 0:
                    cx_t = _in(tl + tw * 0.5)
                    cy_t = _in(tt + th * 0.5)
                    transform = f' transform="rotate({round(tag_deg, 2)} {cx_t} {cy_t})"'
                parts.append(
                    f'<rect x="{_in(tl)}" y="{_in(tt)}" '
                    f'width="{_in(tw)}" height="{_in(th)}" '
                    f'rx="{rx}" ry="{rx}" class="tag"{transform}/>\n')
                parts.append(
                    f'<text x="{_in(tl + tw * 0.5)}" y="{_in(tt + th * 0.5)}" '
                    f'class="tag-text"{transform}>{_escape(tag_text)}</text>\n')

    return ''.join(parts)


class ns_ddx_svg_run:
    """SVG renderer that mirrors nsm_ddx_figure.ns_ddx_figure_run for L1 diagrams."""

    def __init__(self):
        """Called with self being the L1 create object that has all layout data set."""

        master_file = getattr(self, 'full_filepath', None)
        if master_file is None:
            return

        import nsm_def

        coord_list = []
        self.coord_list = coord_list

        outline_margin_root = 0.2
        outline_margin_sub = 0.1
        folder_font_size = 10
        shape_font_size = 6

        root_left = float(self.root_left)
        root_top = float(self.root_top)
        root_width = float(self.root_width)
        root_height = float(self.root_hight)

        slide_w = root_width + root_left * 2 + outline_margin_root * 2 + 1.0
        slide_h = root_height + root_top * 1.5 + outline_margin_root * 2 + 1.0

        rf_left = root_left
        rf_top = root_top
        rf_width = root_width + outline_margin_root * 2
        rf_height = root_height + outline_margin_root * 2

        content_left = rf_left + outline_margin_root
        content_top = rf_top + outline_margin_root
        content_width = rf_width - outline_margin_root * 2
        content_height = rf_height - outline_margin_root * 2

        def _arr_to_sd(arr):
            if arr and isinstance(arr[0], str) and arr[0] == '_NOT_FOUND_':
                return sc.SectionData([])
            filtered = []
            for item in arr:
                if isinstance(item, list) and len(item) == 2 and isinstance(item[1], list):
                    row_vals = item[1]
                    if row_vals and str(row_vals[0]).startswith('<<'):
                        continue
                    filtered.append(row_vals)
            return sc.SectionData(filtered)

        bulk = getattr(self, '_preloaded_bulk', None)
        if bulk:
            def _get(tag):
                return _arr_to_sd(bulk.get(tag, ['_NOT_FOUND_', 1]))
        else:
            sections = [
                '<<POSITION_FOLDER>>', '<<POSITION_SHAPE>>', '<<STYLE_SHAPE>>',
                '<<STYLE_FOLDER>>', '<<POSITION_LINE>>', '<<POSITION_TAG>>',
            ]
            bulk = nsm_def.convert_master_to_arrays_bulk('Master_Data', master_file, sections)

            def _get(tag):
                return _arr_to_sd(bulk.get(tag, ['_NOT_FOUND_', 1]))

        sd_folder = _get('<<POSITION_FOLDER>>')
        sd_shape = _get('<<POSITION_SHAPE>>')
        sd_style_shape = _get('<<STYLE_SHAPE>>')
        sd_style_folder = _get('<<STYLE_FOLDER>>')
        sd_line = _get('<<POSITION_LINE>>')
        sd_tag = _get('<<POSITION_TAG>>')

        style_index, default_style = sc.build_style_shape_index(sd_style_shape)
        folder_style, default_folder_style = sc.build_style_folder_index(sd_style_folder)
        attribute_colors = getattr(self, 'attribute_tuple1_1', {})

        title_text = '[L1] All Areas'
        root_tuple = getattr(self, 'root_folder_tuple', {})
        if root_tuple and root_tuple.get((2, 2)) == 'Summary Diagram':
            title_text = 'Summary Diagram'
        click = getattr(self, 'click_value', '')
        if click == 'VPN-1-1':
            title_text = '[VPNs on L1] All Areas'
        elif click in ('2-4-1', '2-4-2'):
            title_text = '[L1] ' + str(getattr(self, 'tmp_folder_name', ''))

        row_weights, per_row_col_weights, cell_name_rows = sc.compute_folder_grid(sd_folder)
        row_sum = sum(row_weights) if row_weights else 1.0

        svg_parts = []
        svg_parts.append(_svg_header(slide_w, slide_h, shape_font_size, folder_font_size))
        svg_parts.append(_render_title(title_text, rf_left, 0.5))
        svg_parts.append(_render_root_folder(rf_left, rf_top, rf_width, rf_height))

        all_devices = []
        sub_y = content_top

        for row_idx, rw in enumerate(row_weights):
            sub_x = content_left
            folder_h = (rw / row_sum) * content_height

            row_col_weights = per_row_col_weights[row_idx] if row_idx < len(per_row_col_weights) else [1.0]
            col_sum = sum(row_col_weights) if row_col_weights else 1.0

            for col_idx, cw in enumerate(row_col_weights):
                folder_w = (cw / col_sum) * content_width

                folder_name = None
                if row_idx < len(cell_name_rows) and col_idx < len(cell_name_rows[row_idx]):
                    folder_name = cell_name_rows[row_idx][col_idx]

                display_folder = folder_name
                tag_name = folder_name
                if folder_name and '<' in folder_name:
                    display_folder = folder_name.split('<')[0]
                    tag_name = '<' + folder_name.split('<')[-1].split('>')[0] + '>'

                fs = folder_style.get(tag_name or '', folder_style.get(folder_name or '', default_folder_style))
                visible = fs[0]
                text_pos = fs[1]
                margin_top_adj = sc._safe_float(fs[2], None) if fs[2] != '<AUTO>' else None
                margin_bottom_adj = sc._safe_float(fs[3], None) if fs[3] != '<AUTO>' else None

                if folder_name is None:
                    if folder_style.get('<EMPTY>', ('NO',))[0] == 'NO':
                        sub_x += folder_w
                        continue

                if visible != 'NO':
                    svg_parts.append(_render_sub_folder(
                        sub_x, sub_y, folder_w, folder_h,
                        display_folder, text_pos, folder_font_size))

                if tag_name:
                    x_l, x_m, x_r = sub_x, sub_x + folder_w * 0.5, sub_x + folder_w
                    y_t, y_m, y_d = sub_y, sub_y + folder_h * 0.5, sub_y + folder_h
                    coord_list.append([tag_name, x_l, x_m, x_r, y_t, y_m, y_d])

                shape_fl = sub_x + outline_margin_sub
                shape_ft = sub_y + outline_margin_sub
                shape_fw = folder_w - outline_margin_sub * 2
                shape_fh = folder_h - outline_margin_sub * 2

                if margin_top_adj is not None:
                    shape_ft += margin_top_adj
                    shape_fh -= margin_top_adj
                if margin_bottom_adj is not None:
                    shape_fh -= margin_bottom_adj

                lookup_tag = folder_name
                if lookup_tag and '<' in lookup_tag:
                    lookup_tag = '<' + lookup_tag.split('<')[-1].split('>')[0] + '>'

                device_grid = sc.compute_shape_grid(sd_shape, lookup_tag or '')
                if device_grid:
                    devs = _compute_devices_in_folder(
                        coord_list, device_grid, shape_fl, shape_ft, shape_fw, shape_fh,
                        style_index, attribute_colors)
                    all_devices.extend(devs)

                sub_x += folder_w
            sub_y += folder_h

        num_devices = len(all_devices)
        if num_devices > 200:
            svg_parts.append(_render_devices_parallel(all_devices, style_index, attribute_colors))
        else:
            svg_parts.append(_render_device_chunk((all_devices, style_index, attribute_colors)))

        line_records = sc.compute_line_data(sd_line)
        tag_type, tag_overrides = sc.compute_tag_config(sd_tag)
        svg_parts.append(
            _render_lines_vectorised(coord_list, line_records, tag_type, tag_overrides, shape_font_size))

        svg_parts.append('</svg>\n')
        self._svg_content = ''.join(svg_parts)

        output_file = getattr(self, 'output_svg_file', None)
        if output_file:
            _save_svg(self._svg_content, output_file)


def _render_devices_parallel(all_devices, style_index, attribute_colors):
    """Render devices using subprocess worker for true multiprocessing.

    Falls back to in-process ThreadPoolExecutor if subprocess fails.
    """
    import json
    import subprocess
    import sys

    worker_path = os.path.join(os.path.dirname(__file__), 'nsm_svg_worker.py')
    num_workers = min(max(1, os.cpu_count() or 1), 8)

    payload = json.dumps({
        'devices': all_devices,
        'style_index': {k: list(v) for k, v in style_index.items()},
        'attribute_colors': {k: list(v) if isinstance(v, tuple) else v
                             for k, v in (attribute_colors or {}).items()},
        'num_workers': num_workers,
    })

    try:
        startupinfo = None
        if sys.platform == 'win32':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
        proc = subprocess.run(
            [sys.executable, worker_path],
            input=payload.encode('utf-8'),
            capture_output=True,
            timeout=120,
            startupinfo=startupinfo,
        )
        if proc.returncode == 0:
            result = json.loads(proc.stdout)
            return result['svg']
    except (subprocess.TimeoutExpired, json.JSONDecodeError, KeyError, OSError):
        pass

    num_workers = min(num_workers, 8)
    chunk_size = max(1, len(all_devices) // num_workers)
    chunks = [all_devices[i:i + chunk_size] for i in range(0, len(all_devices), chunk_size)]
    chunk_args = [(chunk, style_index, attribute_colors) for chunk in chunks]
    with ThreadPoolExecutor(max_workers=num_workers) as pool:
        results = list(pool.map(_render_device_chunk, chunk_args))
    return ''.join(results)


def _save_svg(content, output_path):
    """Write SVG content to file."""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f'SVG saved: {output_path}')
