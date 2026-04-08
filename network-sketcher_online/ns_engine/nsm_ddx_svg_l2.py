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
L2 SVG rendering engine for Network Sketcher L2 diagrams.

Renders L2 area diagrams including:
- Folder/device topology (shared with L1)
- L2 segments inside each device
- Physical IF / Virtual port tags
- L2/L3 tag colour coding
- Device-internal connection lines
- Inter-device L2 connection lines

Architecture mirrors nsm_ddx_figure.py extended class but outputs SVG.
"""

import nsm_svg_compute as sc

DPI = 96
PT_TO_PX = DPI / 72.0
FONT_FAMILY = 'Calibri, Segoe UI, Arial, sans-serif'


def _in(inches):
    return round(inches * DPI, 2)


def _pt(points):
    return round(points * PT_TO_PX, 2)


def _rgb(r, g, b):
    return f'rgb({int(r)},{int(g)},{int(b)})'


def _escape(text):
    if text is None:
        return ''
    s = str(text)
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


# ---------- L2 shape style definitions (from nsm_ddx_figure.py extended.add_shape) ----------

L2_STYLES = {
    'L2_SEGMENT':      {'fill': (247,245,249), 'stroke': (112,48,160), 'sw': 0.5, 'rx': 0.15, 'tc': (112,48,160)},
    'L2_SEGMENT_GRAY': {'fill': (255,255,255), 'stroke': (127,127,127), 'sw': 0.5, 'rx': 0.15, 'tc': (127,127,127)},
    'L2_TAG':          {'fill': (255,255,255), 'stroke': (0,112,192),   'sw': 0.5, 'rx': 0.50, 'tc': (0,112,192)},
    'L3_TAG':          {'fill': (255,255,255), 'stroke': (192,0,0),     'sw': 0.5, 'rx': 0.50, 'tc': (192,0,0)},
    'TAG_NORMAL':      {'fill': (255,255,255), 'stroke': (0,0,0),       'sw': 0.5, 'rx': 0.50, 'tc': (0,0,0)},
    'GRAY_TAG':        {'fill': (255,255,255), 'stroke': (127,127,127), 'sw': 0.5, 'rx': 0.50, 'tc': (127,127,127)},
    'DEVICE_FRAME':    {'fill': (250,251,247), 'stroke': (0,0,0),       'sw': 1.0, 'rx': 0.0,  'tc': (0,0,0)},
    'WAY_POINT':       {'fill': (237,242,249), 'stroke': (0,0,0),       'sw': 1.0, 'rx': 0.2,  'tc': (0,0,0)},
    'DEVICE_NORMAL':   {'fill': (235,241,222), 'stroke': (0,0,0),       'sw': 1.0, 'rx': 0.0,  'tc': (0,0,0)},
    'L2SEG_TEXT':      {'fill': None,          'stroke': None,          'sw': 0,   'rx': 0.0,  'tc': (163,101,209)},
    'FOLDER_NORMAL':   {'fill': None,          'stroke': (205,205,205), 'sw': 1.0, 'rx': 0.015,'tc': (0,0,0)},
    'OUTLINE_NORMAL':  {'fill': (255,255,255), 'stroke': (0,0,0),       'sw': 1.0, 'rx': 0.0,  'tc': (0,0,0)},
    'L3_SEGMENT_GRAY': {'fill': (249,249,249), 'stroke': (0,0,0),       'sw': 0.75,'rx': 0.30, 'tc': (0,0,0)},
    'L3_SEGMENT_VPN':  {'fill': (248,243,251), 'stroke': (112,48,160),  'sw': 0.75,'rx': 0.30, 'tc': (0,0,0)},
    'L3_INSTANCE':     {'fill': (230,224,236), 'stroke': (0,0,0),       'sw': 1.0, 'rx': 0.20, 'tc': (0,0,0)},
    'IP_ADDRESS_TAG':  {'fill': None,          'stroke': None,          'sw': 0.75,'rx': 0.0,  'tc': (0,0,0)},
}

L2_LINE_STYLES = {
    'NORMAL':          {'stroke': (0,112,192),  'sw': 0.75},
    'L3_SEGMENT':      {'stroke': (0,0,0),      'sw': 2.5,  'marker_start': 'diamond', 'marker_end': 'diamond'},
    'L3_SEGMENT-L3IF': {'stroke': (0,0,0),      'sw': 0.7,  'marker_end': 'diamond'},
    'L3_SEGMENT-VPN':  {'stroke': (112,48,160), 'sw': 0.7,  'marker_end': 'diamond'},
    'L3_INSTANCE':     {'stroke': (96,74,123),  'sw': 0.7},
}


def _svg_l2_header(slide_w, slide_h, shape_font_size):
    w = _in(slide_w)
    h = _in(slide_h)
    return (
        f'<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{w}" height="{h}" viewBox="0 0 {w} {h}">\n'
        f'<defs>\n'
        f'<marker id="diamond" viewBox="0 0 10 10" refX="5" refY="5" '
        f'markerWidth="8" markerHeight="8" orient="auto-start-reverse">'
        f'<path d="M5 0 L10 5 L5 10 L0 5 Z" fill="black"/></marker>\n'
        f'<style>\n'
        f'  text {{ font-family: {FONT_FAMILY}; }}\n'
        f'</style>\n'
        f'</defs>\n'
        f'<rect width="{w}" height="{h}" fill="white"/>\n'
    )


def _render_l2_shape(shape_type, x, y, w, h, text, rotation=0, font_size=6):
    """Render a single L2 shape (rect + text) as SVG."""
    style = L2_STYLES.get(shape_type, L2_STYLES['DEVICE_FRAME'])
    parts = []

    fill_str = _rgb(*style['fill']) if style['fill'] else 'none'
    stroke_str = _rgb(*style['stroke']) if style['stroke'] else 'none'
    sw = _pt(style['sw']) if style['sw'] else 0
    rx = _in(style['rx'] * min(w, h) * 0.5) if style['rx'] else 0

    transform = ''
    if rotation != 0:
        cx_r = _in(x + w / 2)
        cy_r = _in(y + h / 2)
        transform = f' transform="rotate({round(rotation, 2)} {cx_r} {cy_r})"'

    if style['fill'] is not None or style['stroke'] is not None:
        parts.append(
            f'<rect x="{_in(x)}" y="{_in(y)}" width="{_in(w)}" height="{_in(h)}" '
            f'rx="{rx}" ry="{rx}" fill="{fill_str}" stroke="{stroke_str}" '
            f'stroke-width="{sw}"{transform}/>\n')

    if text and str(text).strip():
        tc = style['tc']
        tc_str = _rgb(*tc) if tc else 'black'
        fs = _pt(font_size)
        tx = _in(x + w / 2)
        ty = _in(y + h / 2)
        baseline = 'central'

        anchor = 'middle'
        if shape_type in ('DEVICE_FRAME', 'WAY_POINT'):
            anchor = 'start'
            tx = _in(x + 0.1)
            ty = _in(y + 0.05)
            baseline = 'hanging'
        elif shape_type in ('L2SEG_TEXT', 'IP_ADDRESS_TAG', 'DEVICE_L3_INSTANCE'):
            anchor = 'start'
            tx = _in(x + 0.05)

        parts.append(
            f'<text x="{tx}" y="{ty}" font-size="{fs}px" '
            f'fill="{tc_str}" text-anchor="{anchor}" dominant-baseline="{baseline}"'
            f'{transform}>{_escape(text)}</text>\n')

    return ''.join(parts)


def _render_l2_line(line_type, x1, y1, x2, y2):
    """Render a single L2 line as SVG."""
    style = L2_LINE_STYLES.get(line_type, L2_LINE_STYLES['NORMAL'])
    stroke_str = _rgb(*style['stroke'])
    sw = _pt(style['sw'])

    markers = ''
    if style.get('marker_start'):
        markers += ' marker-start="url(#diamond)"'
    if style.get('marker_end'):
        markers += ' marker-end="url(#diamond)"'

    return (f'<line x1="{_in(x1)}" y1="{_in(y1)}" x2="{_in(x2)}" y2="{_in(y2)}" '
            f'stroke="{stroke_str}" stroke-width="{sw}"{markers}/>\n')


# ---------- L2 Data preparation (ported from l2_device_materials STEP1.1) ----------

_text_wh_cache = {}

def _text_wh(font_size, text):
    """Compute (width, height) in inches for text at given font_size (points). Cached."""
    key = (font_size, str(text))
    result = _text_wh_cache.get(key)
    if result is None:
        import nsm_def
        result = nsm_def.get_description_width_hight(font_size, str(text))
        _text_wh_cache[key] = result
    return result


def _prepare_l2_data(ctx):
    """Prepare L2 shared data from context object.

    ctx must have: l2_table_array, position_line_tuple, position_shape_array,
                   position_folder_array
    Returns dict with processed L2 data needed for rendering.
    """
    import nsm_def

    new_l2_table_array = []
    for tmp in ctx.l2_table_array:
        if tmp[0] != 1 and tmp[0] != 2:
            row = list(tmp[1])
            row.extend([''] * 8)
            del row[8:]
            new_l2_table_array.append([tmp[0], row])

    update_l2_table_array = []
    for tmp_tmp in new_l2_table_array:
        offset_excel = 2
        row = tmp_tmp[1]

        if row[offset_excel + 3] == "":
            if row[offset_excel + 4] == "":
                row[offset_excel] = '' if row[offset_excel + 1] == "" else 'Routed (L3)'
            else:
                row[offset_excel] = '' if row[offset_excel + 1] == "" else 'Switch (L2)'
        else:
            row[offset_excel] = '' if row[offset_excel + 1] == "" else 'Switch (L2)'

        offset_excel = 4
        if row[offset_excel + 1] == "":
            row[offset_excel] = '' if row[offset_excel + 1] == "" else 'Routed (L3)'
        else:
            if row[offset_excel + 2] == "":
                row[offset_excel] = 'Loopback (L3)' if row[offset_excel - 1] == "" else 'Routed (L3)'
            else:
                row[offset_excel] = 'Routed (L3)' if row[offset_excel - 1] == "" else 'Switch (L2)'

        update_l2_table_array.append(row)

    device_l2name_array = []
    unique_l2name_array = []
    for tmp in new_l2_table_array:
        if tmp[1][6] != '':
            tmp_l2seg = []
            for seg in tmp[1][6].split(','):
                seg = seg.replace(' ', '').strip()
                tmp_l2seg.append(seg)
                if seg not in unique_l2name_array:
                    unique_l2name_array.append(seg)
            device_l2name_array.append([tmp[1][1], tmp_l2seg])
    unique_l2name_array.sort()

    device_list_array = []
    wp_list_array_local = []
    for tmp in new_l2_table_array:
        name = tmp[1][1]
        if name not in device_list_array and name not in wp_list_array_local:
            if tmp[1][0] == 'N/A':
                wp_list_array_local.append(name)
            else:
                device_list_array.append(name)

    shape_list_array = list(device_list_array) + list(wp_list_array_local)

    # Build device-indexed L2 table for O(1) lookup
    l2_rows_by_device = {}
    for row in update_l2_table_array:
        dev = row[1]
        if dev not in l2_rows_by_device:
            l2_rows_by_device[dev] = []
        l2_rows_by_device[dev].append(row)

    shape_if_array = []
    for shape_name in shape_list_array:
        ifs = []
        for tmp in new_l2_table_array:
            if shape_name == tmp[1][1] and tmp[1][3] != '':
                ifs.append(tmp[1][3])
        shape_if_array.append([shape_name, ifs])

    modify_position_shape_array = []
    tmp_folder_name = ''
    for tmp in ctx.position_shape_array:
        if tmp[0] != 1 and tmp[1][0] != '<END>':
            if tmp[1][0] != '':
                tmp_folder_name = tmp[1][0]
            else:
                tmp[1][0] = tmp_folder_name
            modify_position_shape_array.append(tmp)

    modify_position_folder_array = []
    for tmp in ctx.position_folder_array:
        if tmp[0] != 1 and tmp[1][0] != '<SET_WIDTH>':
            row_copy = list(tmp[1])
            row_copy[0] = ''
            modify_position_folder_array.append([tmp[0], row_copy])

    direction_if_array = _compute_if_directions(
        shape_if_array, wp_list_array_local,
        ctx.position_line_tuple,
        modify_position_shape_array,
        modify_position_folder_array)

    new_direction_if_array = _sort_if_directions(direction_if_array)

    # Build direction index for O(1) lookup by device name
    direction_by_device = {}
    for entry in new_direction_if_array:
        direction_by_device[entry[0]] = entry

    # Build device L2 segment name index
    l2name_by_device = {}
    for entry in device_l2name_array:
        dev = entry[0]
        if dev not in l2name_by_device:
            l2name_by_device[dev] = []
        for seg in entry[1]:
            if seg not in l2name_by_device[dev]:
                l2name_by_device[dev].append(seg)

    return {
        'new_l2_table_array': new_l2_table_array,
        'update_l2_table_array': update_l2_table_array,
        'device_l2name_array': device_l2name_array,
        'unique_l2name_array': unique_l2name_array,
        'device_list_array': device_list_array,
        'wp_list_array': wp_list_array_local,
        'shape_if_array': shape_if_array,
        'new_direction_if_array': new_direction_if_array,
        'modify_position_shape_array': modify_position_shape_array,
        'l2_rows_by_device': l2_rows_by_device,
        'direction_by_device': direction_by_device,
        'l2name_by_device': l2name_by_device,
    }


def _compute_if_directions(shape_if_array, wp_list_array,
                           position_line_tuple,
                           modify_position_shape_array,
                           modify_position_folder_array):
    """Determine UP/DOWN/RIGHT/LEFT direction for each IF on each device.

    Uses pre-built index for O(1) device lookup instead of O(P) full scan.
    """
    # Build device-name -> [(row_idx, col_idx)] index for O(1) lookup
    device_rows_index = {}
    for key in position_line_tuple:
        row_idx, col_idx = key
        if row_idx in (1, 2) or col_idx not in (1, 2):
            continue
        dev = position_line_tuple[key]
        if dev not in device_rows_index:
            device_rows_index[dev] = []
        device_rows_index[dev].append((row_idx, col_idx))

    # Pre-build shape/folder position indexes for _determine_up_down
    shape_pos_index = {}
    for arr in modify_position_shape_array:
        for idx, item in enumerate(arr[1]):
            if idx != 0:
                shape_pos_index[item] = (arr[1][0], arr[0])

    folder_pos_index = {}
    for arr in modify_position_folder_array:
        for item in arr[1]:
            if item:
                folder_pos_index[item] = arr[0]

    direction_if_array = []
    for tmp_shape_if in shape_if_array:
        device_name = tmp_shape_if[0]
        tmp_dir = [device_name, [], [], [], []]

        device_entries = device_rows_index.get(device_name, [])

        for if_name in tmp_shape_if[1]:
            for row_idx, col_idx in device_entries:
                offset_col = 0 if col_idx == 1 else 1

                tmp_tag = position_line_tuple.get((row_idx, 3 + offset_col), '')
                if not tmp_tag or ' ' not in str(tmp_tag):
                    continue

                prefix = position_line_tuple.get((row_idx, 13 + offset_col * 4), '')
                space_idx = str(tmp_tag).find(' ')
                modify_if_name = str(prefix) + ' ' + str(tmp_tag)[space_idx + 1:]

                if if_name != modify_if_name:
                    continue

                side = position_line_tuple.get((row_idx, 5 + offset_col), '')
                tag_offset_key_lr = (row_idx, 8 + offset_col * 2)
                tag_offset_key_ud = (row_idx, 7 + offset_col * 2)

                tag_offset_lr = position_line_tuple.get(tag_offset_key_lr, 0.0)
                if tag_offset_lr == '':
                    tag_offset_lr = 0.0
                tag_offset_ud = position_line_tuple.get(tag_offset_key_ud, 0.0)
                if tag_offset_ud == '':
                    tag_offset_ud = 0.0

                if side == 'RIGHT':
                    tmp_dir[3].append([modify_if_name, float(tag_offset_lr)])
                elif side == 'LEFT':
                    tmp_dir[4].append([modify_if_name, float(tag_offset_lr)])
                else:
                    if offset_col == 0:
                        opposite = position_line_tuple.get((row_idx, col_idx + 1), '')
                    else:
                        opposite = position_line_tuple.get((row_idx, col_idx - 1), '')

                    is_up = _determine_up_down_fast(
                        device_name, opposite, wp_list_array,
                        shape_pos_index, folder_pos_index)
                    if is_up:
                        tmp_dir[1].append([modify_if_name, float(tag_offset_ud)])
                    else:
                        tmp_dir[2].append([modify_if_name, float(tag_offset_ud)])

        direction_if_array.append(tmp_dir)
    return direction_if_array


def _determine_up_down(device_name, opposite_name, wp_list_array,
                       modify_position_shape_array, modify_position_folder_array):
    """Return True if opposite is above (device should connect UP). Legacy version."""
    if device_name in wp_list_array or opposite_name in wp_list_array:
        origin_folder = ''
        opposite_folder = ''
        origin_num = 0
        opposite_num = 0
        for arr in modify_position_shape_array:
            for idx, item in enumerate(arr[1]):
                if idx != 0:
                    if device_name == item:
                        origin_folder = arr[1][0]
                    if opposite_name == item:
                        opposite_folder = arr[1][0]
        for arr in modify_position_folder_array:
            if origin_folder in arr[1]:
                origin_num = arr[0]
            if opposite_folder in arr[1]:
                opposite_num = arr[0]
        return origin_num > opposite_num
    else:
        origin_num = 0
        opposite_num = 0
        for arr in modify_position_shape_array:
            if device_name in arr[1]:
                origin_num = arr[0]
            if opposite_name in arr[1]:
                opposite_num = arr[0]
        return origin_num > opposite_num


def _determine_up_down_fast(device_name, opposite_name, wp_list_array,
                            shape_pos_index, folder_pos_index):
    """Return True if opposite is above. O(1) indexed version."""
    if device_name in wp_list_array or opposite_name in wp_list_array:
        origin_info = shape_pos_index.get(device_name, ('', 0))
        opposite_info = shape_pos_index.get(opposite_name, ('', 0))
        origin_num = folder_pos_index.get(origin_info[0], 0)
        opposite_num = folder_pos_index.get(opposite_info[0], 0)
        return origin_num > opposite_num
    else:
        origin_info = shape_pos_index.get(device_name, ('', 0))
        opposite_info = shape_pos_index.get(opposite_name, ('', 0))
        return origin_info[1] > opposite_info[1]


def _sort_if_directions(direction_if_array):
    """Sort IFs within each direction by offset."""
    result = []
    for entry in direction_if_array:
        sorted_entry = [entry[0]]
        for i in range(1, 5):
            if not entry[i]:
                sorted_entry.append([])
            elif len(entry[i]) == 1:
                sorted_entry.append([entry[i][0][0]])
            else:
                sorted_data = sorted(entry[i], key=lambda x: x[1])
                sorted_entry.append([item[0] for item in sorted_data])
        result.append(sorted_entry)
    return result


# ---------- L2 Device Material Rendering ----------

def _render_l2_device_materials(target_device_name, l2_data, ctx, offset_x, offset_y, font_size=6):
    """Render L2 materials for a single device.

    Ported from nsm_ddx_figure.py add_l2_material / l2_device_materials.

    offset_x, offset_y: offset to convert device-local coords to absolute.
    Returns: (list_of_svg_strings, [left, top, width, height], tag_size_array)
    """
    import nsm_def

    parts = []

    update_l2_table_array = l2_data['update_l2_table_array']
    new_l2_table_array = l2_data['new_l2_table_array']
    new_direction_if_array = l2_data['new_direction_if_array']
    wp_list_array = l2_data['wp_list_array']
    l2_rows_by_device = l2_data.get('l2_rows_by_device', {})

    shape_width_min = 0.5
    shape_hight_min = 0.1
    shape_interval_width_ratio = 0.75
    shape_interval_hight_ratio = 0.75
    between_tag = 0.2
    l2seg_size_margin = 0.7
    l2seg_size_margin_left_right_add = 0.4

    target_device_l2_array = list(l2_data.get('l2name_by_device', {}).get(target_device_name, []))
    target_device_l2_array.sort()

    flag_l2_segment_empty = len(target_device_l2_array) == 0
    if flag_l2_segment_empty:
        target_device_l2_array.append('_DummyL2Segment_')

    # --- Compute L2 segment positions ---
    count = 0
    pre_shape_width = 0
    pre_offset_left = 0
    l2seg_size_array = []
    seg_offset_left = 0.0
    seg_offset_top = 0.0

    for seg_name in target_device_l2_array:
        wh = _text_wh(font_size, seg_name)
        seg_w = max(wh[0], shape_width_min)
        seg_h = wh[1]

        if flag_l2_segment_empty:
            seg_w = 0.01
            seg_h = 0.01

        if count > 0:
            seg_offset_left -= pre_shape_width
            seg_offset_left += pre_shape_width * shape_interval_width_ratio
            seg_offset_top += seg_h * shape_interval_hight_ratio

            right_edge = pre_offset_left + pre_shape_width + seg_w * shape_interval_width_ratio
            if right_edge > (seg_offset_left + seg_w):
                seg_offset_left += (right_edge - (seg_offset_left + seg_w))

        pre_offset_left = seg_offset_left
        pre_shape_width = seg_w

        shape_type = 'L2_SEGMENT'
        if not flag_l2_segment_empty:
            for row in l2_rows_by_device.get(target_device_name, []):
                if row[3] == '' and row[5] == '':
                    segs = row[6].replace(' ', '').split(',')
                    if seg_name in segs:
                        shape_type = 'L2_SEGMENT_GRAY'
                        break

            parts.append(_render_l2_shape(
                shape_type,
                offset_x + seg_offset_left, offset_y + seg_offset_top,
                seg_w, seg_h, seg_name, font_size=font_size))

        l2seg_size_array.append([
            seg_offset_left, seg_offset_top, seg_w, seg_h, seg_name])

        seg_offset_left += seg_w
        seg_offset_top += seg_h
        count += 1

    if not l2seg_size_array:
        l2seg_size_array = [[0.0, 0.0, 0.01, 0.01, '_DummyL2Segment_']]

    # --- Compute device frame size ---
    dev_left = l2seg_size_array[0][0]
    dev_top = l2seg_size_array[0][1]
    dev_right = l2seg_size_array[-1][0] + l2seg_size_array[-1][2]
    dev_bottom = l2seg_size_array[-1][1] + l2seg_size_array[-1][3]
    dev_w = dev_right - dev_left
    dev_h = dev_bottom - dev_top

    device_size = [
        dev_left - l2seg_size_margin,
        dev_top - l2seg_size_margin,
        dev_w + l2seg_size_margin * 2,
        dev_h + l2seg_size_margin * 2
    ]

    # --- Get current device IF directions (O(1) lookup) ---
    current_dir = l2_data.get('direction_by_device', {}).get(
        target_device_name, [target_device_name, [], [], [], []])

    # Deduplicate IF names
    for i in range(1, 5):
        if current_dir[i]:
            current_dir[i] = list(dict.fromkeys(current_dir[i]))

    # --- Get virtual ports (using device-indexed L2 table) ---
    target_vport_array = []
    target_vport_if_array = []
    for row in l2_rows_by_device.get(target_device_name, []):
        if row[5] != '':
            if row[5] not in target_vport_array:
                target_vport_array.append(row[5])
                target_vport_if_array.append([row[5], [row[3]]])
            else:
                for vp in target_vport_if_array:
                    if vp[0] == row[5]:
                        vp[1].append(row[3])

    # --- Check IF/Vport existence per direction ---
    has_if = [False, False, False, False]
    has_vport = [False, False, False, False]
    vport_count = [0, 0, 0, 0]
    vport_l3_count = [0, 0, 0, 0]

    dir_map = {0: 1, 1: 2, 2: 3, 3: 4}
    for dir_idx in range(4):
        arr_idx = dir_map[dir_idx]
        if current_dir[arr_idx]:
            has_if[dir_idx] = True
            for if_name in current_dir[arr_idx]:
                for vp in target_vport_if_array:
                    if if_name in vp[1]:
                        has_vport[dir_idx] = True
                        vport_count[dir_idx] += 1
                        for row in l2_rows_by_device.get(target_device_name, []):
                            if (row[5] == vp[0] and row[4] != 'Switch (L2)'):
                                vport_l3_count[dir_idx] += 1

    other_if_array = []
    for vp in target_vport_if_array:
        if vp[1] == ['']:
            other_if_array.append(vp[0])

    # --- Adjust device size for vports ---
    if has_vport[0] or other_if_array:
        device_size[1] -= l2seg_size_margin
        device_size[3] += l2seg_size_margin
    if has_vport[1]:
        device_size[3] += l2seg_size_margin
    if has_vport[2]:
        device_size[2] += (l2seg_size_margin + l2seg_size_margin_left_right_add)
    if has_vport[3]:
        device_size[0] -= (l2seg_size_margin + l2seg_size_margin_left_right_add)
        device_size[2] += (l2seg_size_margin + l2seg_size_margin_left_right_add)

    # --- Adjust for IF count (vertical space) ---
    need_left = len(current_dir[4]) * (shape_hight_min + between_tag) + l2seg_size_array[0][1] + l2seg_size_array[0][3]
    need_right = l2seg_size_array[-1][1] + l2seg_size_array[-1][3] + vport_l3_count[2] * (shape_hight_min + between_tag) - (l2seg_size_margin * 0.75)
    keep_down = device_size[1] + device_size[3] - l2seg_size_margin - (shape_hight_min * 1.5)

    extra_down = max(need_left, need_right) - keep_down
    if extra_down > 0:
        device_size[3] += extra_down

    need_right_up = len(current_dir[3]) * (shape_hight_min + between_tag) + l2seg_size_margin
    if has_vport[0]:
        need_right_up += l2seg_size_margin
    need_left_up = (l2seg_size_array[-1][1] - l2seg_size_array[0][1]) + vport_l3_count[3] * (shape_hight_min + between_tag) + (l2seg_size_margin * 0.75)
    keep_up = l2seg_size_array[-1][1] - device_size[1]

    extra_up = max(need_right_up, need_left_up) - keep_up
    if extra_up > 0:
        device_size[1] -= extra_up
        device_size[3] += extra_up

    # --- Render physical IF and vport tags ---
    # Use device-filtered L2 table for O(avg_if) instead of O(L)
    device_l2_rows = l2_rows_by_device.get(target_device_name, [])
    tag_size_array = []
    used_vport_names = []

    _render_if_tags_direction(
        parts, tag_size_array, used_vport_names,
        'UP', current_dir[1], has_if[0], has_vport[0],
        target_device_name, device_l2_rows, target_vport_if_array,
        other_if_array, l2seg_size_array, device_size,
        l2seg_size_margin, between_tag, font_size,
        offset_x, offset_y, ctx)

    # Expand device width based on UP physical IF tag span (PPT line 3839-3843)
    up_phys_ifs = set(current_dir[1]) if has_if[0] else set()
    up_phys_tags = [t for t in tag_size_array if t[6] == 'UP' and t[5] in up_phys_ifs]
    if up_phys_tags:
        rightmost = max(up_phys_tags, key=lambda t: t[0] + t[2])
        l2seg_rightside = l2seg_size_array[-1][0] + l2seg_size_array[-1][2] + l2seg_size_margin
        tag_rightside = rightmost[0] + rightmost[2] + l2seg_size_margin
        if tag_rightside > l2seg_rightside:
            device_size[2] += (tag_rightside - l2seg_rightside)

    # Expand device width based on UP vport tag span (PPT line 3916-3918)
    up_vport_tags = [t for t in tag_size_array if t[6] == 'UP' and t[5] not in up_phys_ifs]
    if up_vport_tags:
        rightmost_vp = max(up_vport_tags, key=lambda t: t[0] + t[2])
        vp_rightside = rightmost_vp[0] + rightmost_vp[2] + l2seg_size_margin
        current_right = device_size[0] + device_size[2]
        if vp_rightside > current_right:
            device_size[2] += (vp_rightside - current_right)

    # Pre-calculate DOWN tag offset before rendering (PPT line 3920-3941)
    down_tag_offset = 0.0
    if has_if[1]:
        import nsm_def as _nsm_def_down
        down_tag_sum = l2seg_size_array[0][0]
        for if_name in current_dir[2]:
            tag_name = _nsm_def_down.get_tag_name_from_full_name(
                target_device_name, if_name, ctx.position_line_tuple)
            if tag_name == '_NO_MATCH_':
                arr = _nsm_def_down.adjust_portname(if_name)
                tag_name = str(arr[0]) + ' ' + str(arr[2])
            tw = _text_wh(font_size, tag_name)[0]
            down_tag_sum += tw + between_tag

        if down_tag_sum > l2seg_size_array[-1][0]:
            down_tag_offset = down_tag_sum - l2seg_size_array[-1][0]
            device_size[0] -= down_tag_offset
            device_size[2] += down_tag_offset

    _render_if_tags_direction(
        parts, tag_size_array, used_vport_names,
        'DOWN', current_dir[2], has_if[1], has_vport[1],
        target_device_name, device_l2_rows, target_vport_if_array,
        other_if_array, l2seg_size_array, device_size,
        l2seg_size_margin, between_tag, font_size,
        offset_x, offset_y, ctx,
        down_tag_offset=down_tag_offset)

    _render_if_tags_direction(
        parts, tag_size_array, used_vport_names,
        'RIGHT', current_dir[3], has_if[2], has_vport[2],
        target_device_name, device_l2_rows, target_vport_if_array,
        other_if_array, l2seg_size_array, device_size,
        l2seg_size_margin, between_tag, font_size,
        offset_x, offset_y, ctx)

    _render_if_tags_direction(
        parts, tag_size_array, used_vport_names,
        'LEFT', current_dir[4], has_if[3], has_vport[3],
        target_device_name, device_l2_rows, target_vport_if_array,
        other_if_array, l2seg_size_array, device_size,
        l2seg_size_margin, between_tag, font_size,
        offset_x, offset_y, ctx)

    # --- Render internal connection lines ---
    _render_l2_internal_lines(
        parts, tag_size_array, used_vport_names,
        current_dir, target_device_name,
        device_l2_rows, other_if_array,
        l2seg_size_array, device_size,
        has_if, has_vport,
        l2seg_size_margin, l2seg_size_margin_left_right_add,
        offset_x, offset_y)

    # --- Ensure device frame is wide enough for device name text ---
    import nsm_def as _nsm_def_name
    name_w = _nsm_def_name.get_description_width_hight(16, target_device_name)[0] + 0.2
    if device_size[2] < name_w:
        device_size[2] = name_w

    # --- Render device frame (background) ---
    frame_type = 'WAY_POINT' if target_device_name in wp_list_array else 'DEVICE_FRAME'
    frame_font = 16
    frame_svg = _render_l2_shape(
        frame_type,
        offset_x + device_size[0], offset_y + device_size[1],
        device_size[2], device_size[3],
        target_device_name, font_size=frame_font)
    parts.insert(0, frame_svg)

    abs_device_size = [
        offset_x + device_size[0],
        offset_y + device_size[1],
        device_size[2],
        device_size[3]
    ]

    abs_tag_array = []
    for t in tag_size_array:
        abs_tag_array.append([
            offset_x + t[0], offset_y + t[1],
            t[2], t[3], t[4], t[5], t[6]])

    return parts, abs_device_size, abs_tag_array


def _render_if_tags_direction(parts, tag_size_array, used_vport_names,
                              direction, if_list, has_if_flag, has_vport_flag,
                              device_name, update_l2_table_array,
                              target_vport_if_array, other_if_array,
                              l2seg_size_array, device_size,
                              l2seg_size_margin, between_tag, font_size,
                              offset_x, offset_y, ctx,
                              down_tag_offset=0.0):
    """Render physical IF tags and virtual port tags for a direction."""
    import nsm_def

    render_vport_only = False
    if not has_if_flag and direction in ('UP', 'DOWN'):
        if direction == 'UP' and (has_vport_flag or other_if_array):
            render_vport_only = True
        else:
            return

    if not has_if_flag and not render_vport_only:
        return

    # --- Render physical IF tags ---
    shape_hight_min = 0.1
    iter_list = list(reversed(if_list)) if direction == 'RIGHT' else if_list
    right_offset_h = 0.0

    for if_name in iter_list:
        tag_type = 'GRAY_TAG'
        for row in update_l2_table_array:
            if row[1] == device_name and row[3] == if_name:
                if 'L2' in str(row[2]):
                    tag_type = 'L2_TAG'
                    break
                elif 'L3' in str(row[2]):
                    tag_type = 'L3_TAG'
                    break
                elif 'Routed (L3)' == row[2]:
                    tag_type = 'L3_TAG'
                    break

        tag_name = nsm_def.get_tag_name_from_full_name(
            device_name, if_name, ctx.position_line_tuple)
        if tag_name == '_NO_MATCH_':
            arr = nsm_def.adjust_portname(if_name)
            tag_name = str(arr[0]) + ' ' + str(arr[2])

        tw = _text_wh(font_size, tag_name)[0]
        th = _text_wh(font_size, tag_name)[1]

        if direction == 'UP':
            tag_left = l2seg_size_array[0][0] + l2seg_size_array[0][2] * 0.5
            if tag_size_array:
                last = [t for t in tag_size_array if t[6] == 'UP']
                if last:
                    tag_left = last[-1][0] + last[-1][2] + between_tag
            tag_top = device_size[1] - th * 0.5
        elif direction == 'DOWN':
            tag_left = l2seg_size_array[0][0] - down_tag_offset
            if tag_size_array:
                last = [t for t in tag_size_array if t[6] == 'DOWN']
                if last:
                    tag_left = last[-1][0] + last[-1][2] + between_tag
            tag_top = device_size[1] + device_size[3] - th * 0.5
        elif direction == 'RIGHT':
            tag_left = device_size[0] + device_size[2] - tw * 0.5
            tag_top = l2seg_size_array[-1][1] - l2seg_size_array[0][3] * 2 + right_offset_h
            right_offset_h -= (shape_hight_min + between_tag)
        elif direction == 'LEFT':
            tag_left = device_size[0] - tw * 0.5
            offset_h = 0.0
            last = [t for t in tag_size_array if t[6] == 'LEFT']
            if last:
                offset_h = len(last) * (shape_hight_min + between_tag)
            tag_top = l2seg_size_array[0][1] + l2seg_size_array[0][3] * 2 + offset_h
        else:
            tag_left = 0
            tag_top = 0

        parts.append(_render_l2_shape(
            tag_type,
            offset_x + tag_left, offset_y + tag_top,
            tw, th, tag_name, font_size=font_size))

        tag_size_array.append([tag_left, tag_top, tw, th, tag_name, if_name, direction])

    # --- Render virtual port tags ---
    _render_vport_tags(
        parts, tag_size_array, used_vport_names, direction,
        if_list, has_vport_flag, other_if_array,
        device_name, update_l2_table_array, target_vport_if_array,
        l2seg_size_array, device_size, l2seg_size_margin,
        between_tag, font_size, offset_x, offset_y)


def _render_l2seg_text(parts, tag_left, tag_top, tag_hight, l2seg_str,
                       font_size, offset_x, offset_y, direction):
    """Render L2 segment name labels under a virtual port tag."""
    import math
    if not l2seg_str:
        return
    segs = l2seg_str.replace(' ', '').split(',')

    if direction in ('UP', 'DOWN'):
        off_h = 0.0
        off_l = 0.05
        for seg in segs:
            off_h += tag_hight
            tw = _text_wh(font_size, seg)[0]
            parts.append(_render_l2_shape(
                'L2SEG_TEXT',
                offset_x + tag_left + off_l, offset_y + tag_top + off_h,
                tw, tag_hight, seg, font_size=font_size))
    else:
        half_num = math.floor(len(segs) * 0.5)
        upper = ''
        lower = ''
        for i, seg in enumerate(segs):
            if i < half_num:
                upper += seg + ' '
            else:
                lower += seg + ' '
        if not upper and lower:
            upper = lower
            lower = ''
        off_l = 0.05
        off_h = tag_hight
        if upper:
            tw = _text_wh(font_size, upper)[0]
            parts.append(_render_l2_shape(
                'L2SEG_TEXT',
                offset_x + tag_left + off_l, offset_y + tag_top + off_h,
                tw, tag_hight, upper, font_size=font_size))
        if lower:
            off_h += _text_wh(font_size, lower)[1]
            tw = _text_wh(font_size, lower)[0]
            parts.append(_render_l2_shape(
                'L2SEG_TEXT',
                offset_x + tag_left + off_l, offset_y + tag_top + off_h,
                tw, tag_hight, lower, font_size=font_size))


def _render_vport_tags(parts, tag_size_array, used_vport_names, direction,
                       if_list, has_vport_flag, other_if_array,
                       device_name, update_l2_table_array, target_vport_if_array,
                       l2seg_size_array, device_size, l2seg_size_margin,
                       between_tag, font_size, offset_x, offset_y):
    """Render virtual port tags and L2SEG_TEXT for all four directions."""
    import nsm_def

    l2seg_size_margin_left_right_add = 0.4
    shape_hight_min = 0.1

    if direction == 'UP' and (has_vport_flag or other_if_array):
        # Find the last physical IF tag position (not vport)
        up_phys_ifs = set(if_list)
        phys_tags = [t for t in tag_size_array if t[6] == 'UP' and t[5] in up_phys_ifs]
        last_phys_left = l2seg_size_array[0][0] + l2seg_size_array[0][2] * 0.5
        last_phys_width = 0.0
        if phys_tags:
            last_phys = phys_tags[-1]
            last_phys_left = last_phys[0]
            last_phys_width = last_phys[2]

        offset_vport_l3 = 0.0

        for row in update_l2_table_array:
            if (row[1] == device_name and row[5] != '' and
                    row[5] not in used_vport_names and
                    (row[3] in if_list or row[3] == '')):
                vp_tag_type = _classify_vport_type(row[4])

                vp_arr = nsm_def.adjust_portname(row[5])
                vp_name = str(vp_arr[0]) + ' ' + str(vp_arr[2])
                vp_w = _text_wh(font_size, vp_name)[0]
                vp_h = _text_wh(font_size, vp_name)[1]
                vp_top = device_size[1] + l2seg_size_margin - vp_h * 0.5

                if vp_tag_type == 'L2_TAG':
                    phy_tag = _find_tag(tag_size_array, row[3])
                    if phy_tag:
                        vp_left = phy_tag[0] + (phy_tag[2] - vp_w) * 0.5
                elif vp_tag_type in ('L3_TAG', 'GRAY_TAG'):
                    vp_left = last_phys_left + last_phys_width + offset_vport_l3 + between_tag * 0.5
                    offset_vport_l3 += vp_w + between_tag * 0.5
                else:
                    vp_left = last_phys_left + last_phys_width + offset_vport_l3 + between_tag * 0.5

                # other_if_array (loopback etc)
                for other_if in other_if_array:
                    if other_if == row[5]:
                        vp_left = last_phys_left + last_phys_width + offset_vport_l3 + between_tag * 0.5
                        offset_vport_l3 += vp_w + between_tag * 0.5

                parts.append(_render_l2_shape(
                    vp_tag_type, offset_x + vp_left, offset_y + vp_top,
                    vp_w, vp_h, vp_name, font_size=font_size))
                tag_size_array.append([vp_left, vp_top, vp_w, vp_h, vp_name, row[5], 'UP'])
                used_vport_names.append(row[5])

                if row[5] != '' and row[6] == '' and row[7] != '':
                    _render_l2seg_text(parts, vp_left, vp_top, vp_h,
                                       row[7], font_size, offset_x, offset_y, 'UP')

    elif direction == 'DOWN' and has_vport_flag:
        last_tags = [t for t in tag_size_array if t[6] == 'DOWN']
        vp_left = l2seg_size_array[0][0]
        if last_tags:
            vp_left = last_tags[0][0] - between_tag * 0.5

        for row in reversed(update_l2_table_array):
            if (row[1] == device_name and row[5] != '' and
                    row[5] not in used_vport_names and
                    row[3] in if_list):
                vp_tag_type = _classify_vport_type(row[4])

                vp_arr = nsm_def.adjust_portname(row[5])
                vp_name = str(vp_arr[0]) + ' ' + str(vp_arr[2])
                vp_w = _text_wh(font_size, vp_name)[0]
                vp_h = _text_wh(font_size, vp_name)[1]
                vp_top = device_size[1] + device_size[3] - vp_h * 0.5 - l2seg_size_margin

                if vp_tag_type == 'L2_TAG':
                    phy_tag = _find_tag(tag_size_array, row[3])
                    if phy_tag:
                        vp_left = phy_tag[0] + (phy_tag[2] - vp_w) * 0.5
                else:
                    vp_left -= vp_w + between_tag * 0.5
                    device_size[0] -= vp_w + between_tag * 0.5
                    device_size[2] += vp_w + between_tag * 0.5

                parts.append(_render_l2_shape(
                    vp_tag_type, offset_x + vp_left, offset_y + vp_top,
                    vp_w, vp_h, vp_name, font_size=font_size))
                tag_size_array.append([vp_left, vp_top, vp_w, vp_h, vp_name, row[5], 'DOWN'])
                used_vport_names.append(row[5])

                if row[5] != '' and row[6] == '' and row[7] != '':
                    _render_l2seg_text(parts, vp_left, vp_top, vp_h,
                                       row[7], font_size, offset_x, offset_y, 'DOWN')

    elif direction == 'RIGHT' and has_vport_flag:
        offset_hight = 0.0
        for row in update_l2_table_array:
            if (row[1] == device_name and row[5] != '' and
                    row[5] not in used_vport_names and
                    row[3] in if_list):
                for vp_entry in target_vport_if_array:
                    if (row[3] in vp_entry[1] and
                            vp_entry[0] not in used_vport_names):
                        used_vport_names.append(vp_entry[0])

                        vp_arr = nsm_def.adjust_portname(vp_entry[0])
                        vp_name = str(vp_arr[0]) + ' ' + str(vp_arr[2])
                        vp_w = _text_wh(font_size, vp_name)[0]
                        vp_h = _text_wh(font_size, vp_name)[1]

                        vp_left = (device_size[0] + device_size[2] -
                                   vp_w * 0.5 - l2seg_size_margin -
                                   l2seg_size_margin_left_right_add)
                        vp_top = (l2seg_size_array[-1][1] +
                                  l2seg_size_array[-1][3] * 2 + offset_hight)

                        if 'Routed (L3)' == row[4]:
                            vp_tag_type = 'L3_TAG'
                        elif 'Switch (L2)' == row[4]:
                            vp_tag_type = 'L2_TAG'
                            phy_tag = _find_tag(tag_size_array, row[3])
                            if phy_tag:
                                vp_top = phy_tag[1]
                        elif 'Loopback (L3)' == row[4]:
                            vp_tag_type = 'GRAY_TAG'
                        else:
                            vp_tag_type = 'GRAY_TAG'

                        parts.append(_render_l2_shape(
                            vp_tag_type, offset_x + vp_left, offset_y + vp_top,
                            vp_w, vp_h, vp_name, font_size=font_size))
                        tag_size_array.append([vp_left, vp_top, vp_w, vp_h,
                                               vp_name, vp_entry[0], 'RIGHT'])

                        if vp_tag_type != 'L2_TAG':
                            offset_hight += vp_h + between_tag
                            if row[5] != '' and row[6] == '' and row[7] != '':
                                _render_l2seg_text(parts, vp_left, vp_top, vp_h,
                                                   row[7], font_size, offset_x,
                                                   offset_y, 'RIGHT')
                        break

    elif direction == 'LEFT' and has_vport_flag:
        vport_l3_count = 0
        for row in update_l2_table_array:
            if (row[1] == device_name and row[5] != '' and
                    row[3] in if_list and row[4] != 'Switch (L2)'):
                vport_l3_count += 1

        offset_hight = 0.0
        for row in update_l2_table_array:
            if (row[1] == device_name and row[5] != '' and
                    row[5] not in used_vport_names and
                    row[3] in if_list):
                for vp_entry in target_vport_if_array:
                    if (row[3] in vp_entry[1] and
                            vp_entry[0] not in used_vport_names):
                        used_vport_names.append(vp_entry[0])

                        vp_arr = nsm_def.adjust_portname(vp_entry[0])
                        vp_name = str(vp_arr[0]) + ' ' + str(vp_arr[2])
                        vp_w = _text_wh(font_size, vp_name)[0]
                        vp_h = _text_wh(font_size, vp_name)[1]

                        vp_left = (device_size[0] - vp_w * 0.5 +
                                   l2seg_size_margin +
                                   l2seg_size_margin_left_right_add)
                        vp_top = (l2seg_size_array[0][1] -
                                  vport_l3_count *
                                  (between_tag + l2seg_size_array[0][3]) +
                                  offset_hight)

                        if 'Routed (L3)' == row[4]:
                            vp_tag_type = 'L3_TAG'
                        elif 'Switch (L2)' == row[4]:
                            vp_tag_type = 'L2_TAG'
                            phy_tag = _find_tag(tag_size_array, row[3])
                            if phy_tag:
                                vp_top = phy_tag[1]
                        elif 'Loopback (L3)' == row[4]:
                            vp_tag_type = 'GRAY_TAG'
                        else:
                            vp_tag_type = 'GRAY_TAG'

                        parts.append(_render_l2_shape(
                            vp_tag_type, offset_x + vp_left, offset_y + vp_top,
                            vp_w, vp_h, vp_name, font_size=font_size))
                        tag_size_array.append([vp_left, vp_top, vp_w, vp_h,
                                               vp_name, vp_entry[0], 'LEFT'])

                        if vp_tag_type != 'L2_TAG':
                            offset_hight += vp_h + between_tag
                            if row[5] != '' and row[6] == '' and row[7] != '':
                                _render_l2seg_text(parts, vp_left, vp_top, vp_h,
                                                   row[7], font_size, offset_x,
                                                   offset_y, 'LEFT')
                        break


def _classify_vport_type(l2_type_str):
    """Classify virtual port tag type from L2 table type string."""
    if 'L2' in str(l2_type_str):
        return 'L2_TAG'
    elif 'Routed (L3)' in str(l2_type_str):
        return 'L3_TAG'
    return 'GRAY_TAG'


def _render_l2_internal_lines(parts, tag_size_array, used_vport_names,
                              current_dir, device_name,
                              update_l2_table_array, other_if_array,
                              l2seg_size_array, device_size,
                              has_if, has_vport,
                              l2seg_size_margin, l2seg_size_margin_left_right_add,
                              offset_x, offset_y):
    """Render internal connection lines (PPT write lines equivalent).

    Handles direction-aware connection points:
    - UP: from bottom of phys-IF, to top of vport
    - DOWN: from top of phys-IF, to bottom of vport
    - RIGHT: from left of phys-IF, to right of vport
    - LEFT: from right of phys-IF, to left of vport
    """
    used_vport_for_lines = []

    def _get_if_direction(if_name):
        if if_name in current_dir[1]:
            return 'UP'
        if if_name in current_dir[2]:
            return 'DOWN'
        if if_name in current_dir[3]:
            return 'RIGHT'
        if if_name in current_dir[4]:
            return 'LEFT'
        return 'UP'

    for row in update_l2_table_array:
        if row[1] != device_name:
            continue

        # Physical IF to Virtual port
        if row[3] != '' and row[5] != '':
            from_tag = _find_tag(tag_size_array, row[3])
            to_tag = _find_tag(tag_size_array, row[5])
            if from_tag and to_tag:
                direction = _get_if_direction(row[3])
                if direction == 'UP':
                    x1 = from_tag[0] + from_tag[2] * 0.5
                    y1 = from_tag[1] + from_tag[3]
                    x2 = to_tag[0] + to_tag[2] * 0.5
                    y2 = to_tag[1]
                elif direction == 'DOWN':
                    x1 = from_tag[0] + from_tag[2] * 0.5
                    y1 = from_tag[1]
                    x2 = to_tag[0] + to_tag[2] * 0.5
                    y2 = to_tag[1] + to_tag[3]
                elif direction == 'RIGHT':
                    x1 = from_tag[0]
                    y1 = from_tag[1] + from_tag[3] * 0.5
                    x2 = to_tag[0] + to_tag[2]
                    y2 = to_tag[1] + to_tag[3] * 0.5
                elif direction == 'LEFT':
                    x1 = from_tag[0] + from_tag[2]
                    y1 = from_tag[1] + from_tag[3] * 0.5
                    x2 = to_tag[0]
                    y2 = to_tag[1] + to_tag[3] * 0.5
                else:
                    x1 = from_tag[0] + from_tag[2] * 0.5
                    y1 = from_tag[1] + from_tag[3]
                    x2 = to_tag[0] + to_tag[2] * 0.5
                    y2 = to_tag[1]

                parts.append(_render_l2_line(
                    'NORMAL',
                    offset_x + x1, offset_y + y1,
                    offset_x + x2, offset_y + y2))

        # Virtual port to L2 Segment
        if row[6] != '' and row[5] != '' and row[5] not in used_vport_for_lines:
            used_vport_for_lines.append(row[5])
            vp_tag = _find_tag(tag_size_array, row[5])

            if row[5] in other_if_array and vp_tag:
                x1 = vp_tag[0] + vp_tag[2] * 0.5
                y1 = vp_tag[1] + vp_tag[3]
                segs = str(row[6]).replace(' ', '').split(',')
                for seg_name in segs:
                    for seg in l2seg_size_array:
                        if seg[4] == seg_name:
                            x2 = seg[0] + seg[2] * 0.9
                            y2 = seg[1]
                            parts.append(_render_l2_line(
                                'NORMAL',
                                offset_x + x1, offset_y + y1,
                                offset_x + x2, offset_y + y2))

            elif vp_tag:
                direction = _get_if_direction(row[3])
                segs = str(row[6]).replace(' ', '').split(',')

                if direction == 'UP':
                    x1 = vp_tag[0] + vp_tag[2] * 0.5
                    y1 = vp_tag[1] + vp_tag[3]
                    for seg_name in segs:
                        for seg in l2seg_size_array:
                            if seg[4] == seg_name:
                                parts.append(_render_l2_line(
                                    'NORMAL',
                                    offset_x + x1, offset_y + y1,
                                    offset_x + seg[0] + seg[2] * 0.9,
                                    offset_y + seg[1]))
                elif direction == 'DOWN':
                    x1 = vp_tag[0] + vp_tag[2] * 0.5
                    y1 = vp_tag[1]
                    for seg_name in segs:
                        for seg in l2seg_size_array:
                            if seg[4] == seg_name:
                                parts.append(_render_l2_line(
                                    'NORMAL',
                                    offset_x + x1, offset_y + y1,
                                    offset_x + seg[0] + seg[2] * 0.1,
                                    offset_y + seg[1] + seg[3]))
                elif direction == 'RIGHT':
                    x1 = vp_tag[0]
                    y1 = vp_tag[1] + vp_tag[3] * 0.5
                    for seg_name in segs:
                        for seg in l2seg_size_array:
                            if seg[4] == seg_name:
                                parts.append(_render_l2_line(
                                    'NORMAL',
                                    offset_x + x1, offset_y + y1,
                                    offset_x + seg[0] + seg[2],
                                    offset_y + seg[1] + seg[3] * 0.5))
                elif direction == 'LEFT':
                    x1 = vp_tag[0] + vp_tag[2]
                    y1 = vp_tag[1] + vp_tag[3] * 0.5
                    for seg_name in segs:
                        for seg in l2seg_size_array:
                            if seg[4] == seg_name:
                                parts.append(_render_l2_line(
                                    'NORMAL',
                                    offset_x + x1, offset_y + y1,
                                    offset_x + seg[0],
                                    offset_y + seg[1] + seg[3] * 0.5))

        # Physical IF directly to L2 Segment
        if row[6] != '' and row[5] == '' and row[3] != '':
            if_tag = _find_tag(tag_size_array, row[3])
            if if_tag:
                direction = _get_if_direction(row[3])
                segs = str(row[6]).replace(' ', '').split(',')

                if direction == 'UP':
                    x1 = if_tag[0] + if_tag[2] * 0.5
                    y1 = if_tag[1] + if_tag[3]
                    if has_vport[0]:
                        y_mid = device_size[1] + l2seg_size_margin
                        parts.append(_render_l2_line(
                            'NORMAL',
                            offset_x + x1, offset_y + y1,
                            offset_x + x1, offset_y + y_mid))
                        y1 = y_mid
                    for seg_name in segs:
                        for seg in l2seg_size_array:
                            if seg[4] == seg_name:
                                parts.append(_render_l2_line(
                                    'NORMAL',
                                    offset_x + x1, offset_y + y1,
                                    offset_x + seg[0] + seg[2] * 0.9,
                                    offset_y + seg[1]))

                elif direction == 'DOWN':
                    x1 = if_tag[0] + if_tag[2] * 0.5
                    y1 = if_tag[1]
                    if has_vport[1]:
                        y_mid = device_size[1] + device_size[3] - 0.1 - l2seg_size_margin
                        parts.append(_render_l2_line(
                            'NORMAL',
                            offset_x + x1, offset_y + y1,
                            offset_x + x1, offset_y + y_mid))
                        y1 = y_mid
                    for seg_name in segs:
                        for seg in l2seg_size_array:
                            if seg[4] == seg_name:
                                parts.append(_render_l2_line(
                                    'NORMAL',
                                    offset_x + x1, offset_y + y1,
                                    offset_x + seg[0] + seg[2] * 0.1,
                                    offset_y + seg[1] + seg[3]))

                elif direction == 'RIGHT':
                    x1 = if_tag[0]
                    y1 = if_tag[1] + if_tag[3] * 0.5
                    if has_vport[2]:
                        x_mid = if_tag[0] - l2seg_size_margin - l2seg_size_margin_left_right_add
                        parts.append(_render_l2_line(
                            'NORMAL',
                            offset_x + x1, offset_y + y1,
                            offset_x + x_mid, offset_y + y1))
                        x1 = x_mid
                    for seg_name in segs:
                        for seg in l2seg_size_array:
                            if seg[4] == seg_name:
                                parts.append(_render_l2_line(
                                    'NORMAL',
                                    offset_x + x1, offset_y + y1,
                                    offset_x + seg[0] + seg[2],
                                    offset_y + seg[1] + seg[3] * 0.5))

                elif direction == 'LEFT':
                    x1 = if_tag[0] + if_tag[2]
                    y1 = if_tag[1] + if_tag[3] * 0.5
                    if has_vport[3]:
                        x_mid = if_tag[0] + if_tag[2] * 0.5 + l2seg_size_margin + l2seg_size_margin_left_right_add + 0.3
                        parts.append(_render_l2_line(
                            'NORMAL',
                            offset_x + x1, offset_y + y1,
                            offset_x + x_mid, offset_y + y1))
                        x1 = x_mid
                    for seg_name in segs:
                        for seg in l2seg_size_array:
                            if seg[4] == seg_name:
                                parts.append(_render_l2_line(
                                    'NORMAL',
                                    offset_x + x1, offset_y + y1,
                                    offset_x + seg[0],
                                    offset_y + seg[1] + seg[3] * 0.5))


def _find_tag(tag_size_array, name):
    """Find tag entry by name (5th element)."""
    for t in tag_size_array:
        if t[5] == name:
            return t
    return None


# ---------- Inter-device L2 connection lines ----------

def _render_l2_inter_device_lines(all_tag_size_array, shapes_size_array,
                                  new_direction_if_array, position_line_tuple):
    """Render inter-device L2 connection lines (PPT add_l2_line equivalent)."""
    import nsm_def

    parts = []
    check_devices = [s[0] for s in shapes_size_array]

    new_all_tags = []
    for tag_entry in all_tag_size_array:
        for shape_entry in shapes_size_array:
            if tag_entry[0] == shape_entry[0]:
                new_all_tags.append(tag_entry)
                break

    for plt_key in position_line_tuple:
        if plt_key[0] in (1, 2) or plt_key[1] != 1:
            continue

        row_idx = plt_key[0]
        from_device = str(position_line_tuple.get((row_idx, 1), ''))
        to_device = str(position_line_tuple.get((row_idx, 2), ''))

        if from_device not in check_devices or to_device not in check_devices:
            continue

        from_tag_raw = position_line_tuple.get((row_idx, 3), '')
        to_tag_raw = position_line_tuple.get((row_idx, 4), '')
        from_prefix = str(position_line_tuple.get((row_idx, 13), ''))
        to_prefix = str(position_line_tuple.get((row_idx, 17), ''))

        if not from_tag_raw or not to_tag_raw:
            continue

        from_split = nsm_def.split_portname(str(from_tag_raw))
        to_split = nsm_def.split_portname(str(to_tag_raw))
        from_fullname = from_prefix + ' ' + str(from_split[1])
        to_fullname = to_prefix + ' ' + str(to_split[1])

        visible = str(position_line_tuple.get((row_idx, 12), ''))

        fx, fy, tx, ty = None, None, None, None
        from_tag_wh = [0, 0]
        to_tag_wh = [0, 0]

        for tag_entry in new_all_tags:
            if tag_entry[0] == from_device:
                for t in tag_entry[1]:
                    if t[5] == from_fullname:
                        fx = t[0] + t[2] * 0.5
                        fy = t[1] + t[3] * 0.5
                        from_tag_wh = [t[2], t[3]]
                        break

        for tag_entry in new_all_tags:
            if tag_entry[0] == to_device:
                for t in tag_entry[1]:
                    if t[5] == to_fullname:
                        tx = t[0] + t[2] * 0.5
                        ty = t[1] + t[3] * 0.5
                        to_tag_wh = [t[2], t[3]]
                        break

        if fx is None or tx is None:
            continue

        for dir_entry in new_direction_if_array:
            if dir_entry[0] == from_device:
                if from_fullname in dir_entry[1]:
                    fy -= from_tag_wh[1] * 0.5
                elif from_fullname in dir_entry[2]:
                    fy += from_tag_wh[1] * 0.5
                elif from_fullname in dir_entry[3]:
                    fx += from_tag_wh[0] * 0.5
                elif from_fullname in dir_entry[4]:
                    fx -= from_tag_wh[0] * 0.5

            if dir_entry[0] == to_device:
                if to_fullname in dir_entry[1]:
                    ty -= to_tag_wh[1] * 0.5
                elif to_fullname in dir_entry[2]:
                    ty += to_tag_wh[1] * 0.5
                elif to_fullname in dir_entry[3]:
                    tx += to_tag_wh[0] * 0.5
                elif to_fullname in dir_entry[4]:
                    tx -= to_tag_wh[0] * 0.5

        if visible == 'NO':
            continue

        stroke_str = _rgb(0, 0, 0)
        sw = _pt(0.5)
        parts.append(
            f'<line x1="{_in(fx)}" y1="{_in(fy)}" '
            f'x2="{_in(tx)}" y2="{_in(ty)}" '
            f'stroke="{stroke_str}" stroke-width="{sw}"/>\n')

    return ''.join(parts)


# ---------- Main L2 SVG renderer class ----------

class ns_ddx_svg_l2_run:
    """L2 SVG renderer (PPT-compatible flow).

    Mirrors PPT's L2-3-2 mode:
    1. Layout devices in folders using L2-computed sizes
    2. Skip device shape drawing (record positions only)
    3. Draw L2 materials (segments, IFs, vports, internal lines) per device
    4. Draw inter-device L2 connection lines
    """

    def __init__(self):
        import nsm_def

        master_file = getattr(self, 'full_filepath', None)
        if master_file is None:
            return

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

        svg_min_w, svg_min_h = 13.4, 7.5
        slide_w = max(svg_min_w, root_width + root_left * 2 + outline_margin_root * 2)
        slide_h = max(svg_min_h, root_height + root_top * 1.5 + outline_margin_root * 2)

        rf_left = root_left
        rf_top = root_top
        rf_width = root_width + outline_margin_root * 2
        rf_height = root_height + outline_margin_root * 2

        content_left = rf_left + outline_margin_root
        content_top = rf_top + outline_margin_root
        content_width = rf_width - outline_margin_root * 2
        content_height = rf_height - outline_margin_root * 2

        bulk = getattr(self, '_preloaded_bulk', None)
        if bulk:
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

            def _get(tag):
                return _arr_to_sd(bulk.get(tag, ['_NOT_FOUND_', 1]))
        else:
            sections = [
                '<<POSITION_FOLDER>>', '<<POSITION_SHAPE>>', '<<STYLE_SHAPE>>',
                '<<STYLE_FOLDER>>', '<<POSITION_LINE>>', '<<POSITION_TAG>>',
            ]
            bulk_loaded = nsm_def.convert_master_to_arrays_bulk('Master_Data', master_file, sections)

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

            def _get(tag):
                return _arr_to_sd(bulk_loaded.get(tag, ['_NOT_FOUND_', 1]))

        sd_folder = _get('<<POSITION_FOLDER>>')
        sd_shape = _get('<<POSITION_SHAPE>>')
        sd_style_shape = _get('<<STYLE_SHAPE>>')
        sd_style_folder = _get('<<STYLE_FOLDER>>')
        sd_line = _get('<<POSITION_LINE>>')
        sd_tag = _get('<<POSITION_TAG>>')

        style_index, default_style = sc.build_style_shape_index(sd_style_shape)
        folder_style, default_folder_style = sc.build_style_folder_index(sd_style_folder)
        attribute_colors = getattr(self, 'attribute_tuple1_1', {})

        target_area = getattr(self, 'target_area', '')
        title_text = f'[L2] {target_area}' if target_area else '[L2] All Areas'

        row_weights, per_row_col_weights, cell_name_rows = sc.compute_folder_grid(sd_folder)
        row_sum = sum(row_weights) if row_weights else 1.0

        svg_parts = []
        svg_parts.append(_svg_l2_header(slide_w, slide_h, shape_font_size))

        svg_parts.append(
            f'<text x="{_in(rf_left)}" y="{_in(0.5)}" '
            f'font-size="{_pt(18)}px" fill="black">{_escape(title_text)}</text>\n')

        # ---------- Phase 1: Layout (compute device positions, draw folders) ----------
        from nsm_ddx_svg import _compute_devices_in_folder, _render_sub_folder

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

        # ---------- Phase 2: Outline rect (drawn first as background) ----------
        svg_parts.append(
            f'<rect x="{_in(rf_left)}" y="{_in(rf_top)}" '
            f'width="{_in(rf_width)}" height="{_in(rf_height)}" '
            f'fill="white" stroke="black" stroke-width="{_pt(1.5)}"/>\n')

        # ---------- Phase 3: L2 material rendering (PPT add_l2_material equivalent) ----------
        has_l2_data = (hasattr(self, 'l2_table_array') and
                       hasattr(self, 'position_line_tuple') and
                       hasattr(self, 'position_shape_array') and
                       hasattr(self, 'position_folder_array'))

        device_size_dict = getattr(self, 'device_size_dict', {})

        shapes_size_array = []
        all_tag_size_array = []

        if has_l2_data and device_size_dict:
            l2_data = _prepare_l2_data(self)

            for dev in all_devices:
                dev_name = dev['name']
                if '_AIR_' in str(dev_name):
                    continue
                if dev_name not in device_size_dict:
                    continue

                pos = [dev['x'], dev['y'], dev['w'], dev['h']]
                size = device_size_dict[dev_name]
                offset_x = pos[0] - size[0]
                offset_y = pos[1] - size[1]

                dev_parts, dev_size, dev_tag_array = _render_l2_device_materials(
                    dev_name, l2_data, self,
                    offset_x, offset_y,
                    font_size=shape_font_size)
                svg_parts.extend(dev_parts)
                shapes_size_array.append([dev_name, pos])
                all_tag_size_array.append([dev_name, dev_tag_array])

        # ---------- Phase 4: Inter-device L2 connection lines ----------
        if has_l2_data and all_tag_size_array:
            new_dir = l2_data.get('new_direction_if_array', [])
            svg_parts.append(
                _render_l2_inter_device_lines(
                    all_tag_size_array, shapes_size_array,
                    new_dir, self.position_line_tuple))

        svg_parts.append('</svg>\n')

        import io
        buf = io.StringIO()
        for part in svg_parts:
            buf.write(part)
        self._svg_content = buf.getvalue()

        output_file = getattr(self, 'output_svg_file', None)
        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(self._svg_content)
            try:
                print(f'L2 SVG saved: {output_file}')
            except UnicodeEncodeError:
                pass
