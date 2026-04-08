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
SVG coordinate computation engine.

Uses NumPy for vectorised coordinate calculations.
Provides fast section data access for rendering L1 diagrams as SVG.
"""

import unicodedata
import math

import numpy as np


class SectionData:
    """Fast section data access backed by lists-of-lists from Parquet.

    Replaces slow openpyxl cell-by-cell access with indexed lookups.
    Data is loaded from load_section() or convert_master_to_array().
    """

    def __init__(self, rows):
        """rows: list of [row_idx, [cell_values, ...]] from convert_master_to_array,
           or list of lists (raw cell values per row)."""
        self._rows = []
        self._max_cols = 0
        if not rows:
            return
        if isinstance(rows[0], list) and len(rows[0]) == 2 and isinstance(rows[0][1], list):
            for _, cell_values in rows:
                self._rows.append(cell_values)
                if len(cell_values) > self._max_cols:
                    self._max_cols = len(cell_values)
        else:
            for r in rows:
                if isinstance(r, list):
                    self._rows.append(r)
                    if len(r) > self._max_cols:
                        self._max_cols = len(r)

    @property
    def num_rows(self):
        return len(self._rows)

    def cell(self, row, col):
        """Access cell value (0-based row, 0-based col). Returns None if out of bounds."""
        if 0 <= row < len(self._rows) and 0 <= col < len(self._rows[row]):
            v = self._rows[row][col]
            if v == '' or v is None:
                return None
            return v
        return None

    def find_section_start(self, section_tag):
        """Find the row index (0-based) of a section tag like '<<POSITION_FOLDER>>'."""
        for i, row in enumerate(self._rows):
            if row and str(row[0]) == section_tag:
                return i
        return -1

    def iter_rows(self, start_row=0):
        """Iterate rows starting from start_row, yielding (row_index, row_data)."""
        for i in range(start_row, len(self._rows)):
            yield i, self._rows[i]


def build_style_shape_index(section_data):
    """Build {shape_name: (width, height, roundness, color_name)} from <<STYLE_SHAPE>> data.

    Returns (style_dict, default_style) where default_style is the <DEFAULT> entry.
    """
    style = {}
    default_style = (1.0, 0.5, 0.0, None)
    empty_style = (0.3, 0.3, 0.0, None)

    for i, row in enumerate(section_data._rows):
        if not row:
            continue
        name = str(row[0]) if row[0] is not None else ''
        if name.startswith('<<'):
            continue
        if name == '':
            break

        width = _safe_float(row[1] if len(row) > 1 else None, 1.0)
        height = _safe_float(row[2] if len(row) > 2 else None, 0.5)
        roundness = _safe_float(row[3] if len(row) > 3 else None, 0.0)
        color_name = row[4] if len(row) > 4 and row[4] else None

        entry = (width, height, roundness, color_name)
        if name == '<DEFAULT>':
            default_style = entry
        elif name == '<EMPTY>':
            empty_style = entry
        else:
            style[name] = entry

    style['<DEFAULT>'] = default_style
    style['<EMPTY>'] = empty_style
    return style, default_style


def build_style_folder_index(section_data):
    """Build {folder_name: (visible, text_pos, margin_top, margin_bottom)} from <<STYLE_FOLDER>>.

    visible: 'YES'/'NO', text_pos: 'UP'/'DOWN'/None
    """
    style = {}
    default_style = ('YES', None, '<AUTO>', '<AUTO>')
    empty_style = ('NO', None, '<AUTO>', '<AUTO>')

    for _, row in enumerate(section_data._rows):
        if not row:
            continue
        name = str(row[0]) if row[0] is not None else ''
        if name.startswith('<<'):
            continue
        if name == '':
            break

        visible = str(row[1]) if len(row) > 1 and row[1] else 'YES'
        text_pos = str(row[2]) if len(row) > 2 and row[2] else None
        margin_top = row[3] if len(row) > 3 and row[3] else '<AUTO>'
        margin_bottom = row[4] if len(row) > 4 and row[4] else '<AUTO>'

        entry = (visible, text_pos, margin_top, margin_bottom)
        if name == '<DEFAULT>':
            default_style = entry
        elif name == '<EMPTY>':
            empty_style = entry
        else:
            style[name] = entry

    style['<DEFAULT>'] = default_style
    style['<EMPTY>'] = empty_style
    return style, default_style


COLOR_MAP = {
    'ORANGE': (253, 234, 218),
    'BLUE': (220, 230, 242),
    'GREEN': (235, 241, 222),
    'GRAY': (242, 242, 242),
}


def resolve_device_color(shape_name, style_index, attribute_colors):
    """Resolve final RGB color for a device shape.

    Priority: attribute_colors > style_index color > transparent.
    Returns (r, g, b) or None for transparent.
    """
    if '_AIR_' in str(shape_name):
        return (255, 255, 255)

    tag_stripped = shape_name
    if '<' in str(shape_name):
        tag_stripped = '<' + str(shape_name).split('<')[-1].split('>')[0] + '>'

    color = None
    entry = style_index.get(tag_stripped) or style_index.get(shape_name)
    if entry and entry[3]:
        color = COLOR_MAP.get(entry[3])

    if attribute_colors and tag_stripped in attribute_colors:
        rgb = attribute_colors[tag_stripped]
        if isinstance(rgb, (list, tuple)) and len(rgb) == 3:
            color = (int(rgb[0]), int(rgb[1]), int(rgb[2]))

    return color


def get_shape_dims(shape_name, style_index):
    """Get (width, height, roundness) for a shape name from the style index."""
    if shape_name is None:
        shape_name = '<EMPTY>'
    name = str(shape_name)
    if '<' in name and not name.startswith('<'):
        tag = '<' + name.split('<')[-1].split('>')[0] + '>'
        entry = style_index.get(tag)
        if entry:
            return entry[0], entry[1], entry[2]

    entry = style_index.get(name)
    if entry:
        return entry[0], entry[1], entry[2]

    default = style_index.get('<DEFAULT>', (1.0, 0.5, 0.0))
    return default[0], default[1], default[2]


def get_east_asian_width_count(text):
    """Count display width accounting for East Asian wide characters."""
    count = 0
    for c in str(text):
        if unicodedata.east_asian_width(c) in 'FWA':
            count += 2
        else:
            count += 1
    return count


def compute_folder_grid(section_data):
    """Parse <<POSITION_FOLDER>> section into grid structure.

    Each data row has format: [row_weight, folder_name_or_empty, folder_name_or_empty, ...]
    Empty cells ('') represent empty slots in the grid.
    Each row may have a preceding <SET_WIDTH> row with per-row column weights.

    Returns:
        row_weights: list of float (row height proportions)
        col_weights: list of list of float (per-row column weights)
        cell_names: list of list of str/None (folder names in grid)
    """
    row_weights = []
    col_weights_rows = []
    cell_name_rows = []
    pending_set_width = None

    for _, row in enumerate(section_data._rows):
        if not row:
            continue
        first = str(row[0]) if row[0] is not None else ''
        if first.startswith('<<POSITION_FOLDER>>'):
            continue
        if first == '' or first == 'None':
            break
        if first == '<SET_WIDTH>':
            weights = []
            for j in range(1, len(row)):
                v = row[j]
                if v is None or str(v) == '' or str(v) == 'None':
                    break
                weights.append(_safe_float(v, 1.0))
            pending_set_width = weights
        else:
            rw = _safe_float(first, 1.0)
            row_weights.append(rw)
            names = []
            for j in range(1, len(row)):
                v = row[j]
                if v is None or str(v) == 'None':
                    continue
                if str(v) == '':
                    names.append(None)
                else:
                    names.append(str(v))
            while names and names[-1] is None:
                names.pop()
            cell_name_rows.append(names)
            col_weights_rows.append(pending_set_width)
            pending_set_width = None

    default_col_count = max((len(r) for r in cell_name_rows), default=1)
    resolved_weights = []
    for cw in col_weights_rows:
        if cw:
            resolved_weights.append(cw)
        else:
            resolved_weights.append([1.0] * default_col_count)

    return row_weights, resolved_weights, cell_name_rows


def compute_shape_grid(section_data, folder_tag):
    """Parse <<POSITION_SHAPE>> to get device grid for a specific folder.

    The folder name and the first device row share the same row:
      ['Office_LAN', '_AIR_', 'Core-SW', ..., '<END>']
      ['',           'Access-Stack1', ..., '<END>']

    Returns list of lists: [[device_name, ...], [device_name, ...], ...]
    """
    found = False
    rows = []

    for _, row in enumerate(section_data._rows):
        if not row:
            continue
        first = str(row[0]) if row[0] is not None else ''
        if first.startswith('<<'):
            if found:
                break
            continue
        if not found:
            if first == folder_tag:
                found = True
            else:
                continue
        if first == '<END>':
            break
        device_row = []
        for j in range(1, len(row)):
            v = row[j]
            if v is None or str(v) == '' or str(v) == 'None':
                break
            if str(v) == '<END>':
                break
            device_row.append(str(v))
        if device_row:
            rows.append(device_row)

    return rows


def compute_line_data(section_data):
    """Parse <<POSITION_LINE>> section into structured line records.

    Returns list of dicts with keys:
        from_name, to_name, from_tag, to_tag, from_side, to_side,
        from_offset_x, from_offset_y, to_offset_x, to_offset_y,
        channel_height, visible
    """
    lines = []
    for _, row in enumerate(section_data._rows):
        if not row:
            continue
        first = str(row[0]) if row[0] is not None else ''
        if first.startswith('<<') or first == '' or first == 'None':
            continue

        rec = {
            'from_name': str(row[0]) if row[0] else None,
            'to_name': str(row[1]) if len(row) > 1 and row[1] else None,
            'from_tag': str(row[2]) if len(row) > 2 and row[2] else None,
            'to_tag': str(row[3]) if len(row) > 3 and row[3] else None,
            'from_side': str(row[4]) if len(row) > 4 and row[4] else None,
            'to_side': str(row[5]) if len(row) > 5 and row[5] else None,
            'from_offset_x': _safe_float(row[6] if len(row) > 6 else None, None),
            'from_offset_y': _safe_float(row[7] if len(row) > 7 else None, None),
            'to_offset_x': row[8] if len(row) > 8 else None,
            'to_offset_y': row[9] if len(row) > 9 else None,
            'channel_height': _safe_float(row[10] if len(row) > 10 else None, None),
            'visible': str(row[11]) if len(row) > 11 and row[11] else None,
        }
        if rec['to_offset_x'] is not None and str(rec['to_offset_x']) != '<FROM_X>':
            rec['to_offset_x'] = _safe_float(rec['to_offset_x'], None)
        if rec['to_offset_y'] is not None and str(rec['to_offset_y']) != '<FROM_Y>':
            rec['to_offset_y'] = _safe_float(rec['to_offset_y'], None)

        if rec['from_name'] and rec['to_name']:
            lines.append(rec)

    return lines


def compute_tag_config(section_data):
    """Parse <<POSITION_TAG>> section.

    The _load() filter already removes the '<<POSITION_TAG>>' header row.
    Remaining rows: '<DEFAULT>' row sets defaults, then per-device overrides.

    Returns (default_tag_type, tag_overrides) where tag_overrides is
    {device_name: {type, offset_x, offset_y, line_offset, rotation}}.
    """
    default_type = 'SHAPE'
    overrides = {}
    current_type = default_type

    for _, row in enumerate(section_data._rows):
        if not row:
            continue
        first = str(row[0]) if row[0] is not None else ''
        if first.startswith('<<'):
            continue
        if first == '' or first == 'None':
            break

        type_val = str(row[1]) if len(row) > 1 and row[1] and str(row[1]).strip() else None
        if type_val:
            current_type = type_val

        if first == '<DEFAULT>':
            if type_val:
                default_type = type_val
                current_type = default_type
            continue

        entry = {
            'type': current_type,
            'offset_x': _safe_float(row[2] if len(row) > 2 else None, None),
            'offset_y': _safe_float(row[3] if len(row) > 3 else None, None),
            'line_offset': _safe_float(row[4] if len(row) > 4 else None, None),
            'rotation': str(row[5]) if len(row) > 5 and row[5] and str(row[5]).strip() else None,
        }
        overrides[first] = entry

    return default_type, overrides


def batch_compute_angles(x1, y1, x2, y2):
    """Vectorised angle computation for line endpoints.

    All inputs are NumPy arrays. Returns angles in degrees.
    """
    dx = np.asarray(x2 - x1, dtype=np.float64)
    dy = np.asarray(y2 - y1, dtype=np.float64)
    return np.degrees(np.arctan2(dy, dx))


def _safe_float(v, default):
    """Convert value to float, returning default on failure."""
    if v is None or str(v).strip() == '' or str(v) == 'None':
        return default
    try:
        return float(v)
    except (ValueError, TypeError):
        return default
