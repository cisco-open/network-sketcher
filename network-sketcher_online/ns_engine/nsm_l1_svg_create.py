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
L1 SVG diagram creation module.

Prepares layout data from the master file and delegates rendering
to nsm_ddx_svg.  Mirrors the data-preparation logic of
nsm_l1_diagram_create.py for the "All Areas + Tags" (click_value 2-4-4)
export path, but does NOT use openpyxl or create staging xlsx files.

Uses bulk loading to read all sections in a single file open.
"""

import os
import re

import nsm_def
import nsm_ddx_svg


_SVG_SECTIONS = [
    '<<POSITION_FOLDER>>', '<<POSITION_SHAPE>>', '<<POSITION_LINE>>',
    '<<STYLE_SHAPE>>', '<<STYLE_FOLDER>>', '<<POSITION_TAG>>',
    '<<ROOT_FOLDER>>',
]


def _safe_area_name(name):
    """Sanitize an area name for use in a filename."""
    safe = re.sub(r'[\\/*?:"<>|]', '-', str(name))
    safe = safe.strip('. ')
    return safe or 'Area'


class nsm_l1_svg_create:
    def __init__(self):
        ws_name = 'Master_Data'
        ppt_meta_file = str(self.inFileTxt_2_1.get())

        if getattr(self, 'click_value_dummy', '') == '12-3':
            ppt_meta_file = str(self.inFileTxt_12_2.get())

        self.full_filepath = ppt_meta_file

        bulk = nsm_def.convert_master_to_arrays_bulk(ws_name, ppt_meta_file, _SVG_SECTIONS)

        self.position_folder_array = bulk['<<POSITION_FOLDER>>']
        self.position_shape_array = bulk['<<POSITION_SHAPE>>']
        self.position_line_array = bulk['<<POSITION_LINE>>']
        self.position_style_shape_array = bulk['<<STYLE_SHAPE>>']
        self.position_tag_array = bulk['<<POSITION_TAG>>']
        self.root_folder_array = bulk['<<ROOT_FOLDER>>']

        self._preloaded_bulk = bulk

        self.position_folder_tuple = nsm_def.convert_array_to_tuple(self.position_folder_array)
        self.position_shape_tuple = nsm_def.convert_array_to_tuple(self.position_shape_array)
        self.position_line_tuple = nsm_def.convert_array_to_tuple(self.position_line_array)
        self.position_style_shape_tuple = nsm_def.convert_array_to_tuple(self.position_style_shape_array)

        try:
            if (hasattr(self, 'inFileTxt_L2_3_1') and
                    hasattr(self, 'comboATTR_1_1') and
                    hasattr(self.inFileTxt_L2_3_1, 'get') and
                    hasattr(self.comboATTR_1_1, 'get')):
                file_path = self.inFileTxt_L2_3_1.get()
                combo_value = self.comboATTR_1_1.get()
                if file_path and combo_value and file_path.strip() and combo_value.strip():
                    self.attribute_tuple1_1 = nsm_def.get_global_attribute_tuple(file_path, combo_value)
        except Exception:
            pass

        click = getattr(self, 'click_value', '2-4-4')

        if click in ('2-4-4', '2-4-3'):
            min_tag_inches = 0.3
            master_folder_size_array = nsm_def.get_folder_width_size(
                self.position_folder_tuple,
                self.position_style_shape_tuple,
                self.position_shape_tuple,
                min_tag_inches)

            self.root_left = 0.28
            self.root_top = 1.42
            self.root_width = float(master_folder_size_array[0]) + 1.0
            self.root_hight = float(master_folder_size_array[3]) + 1.0

            nsm_ddx_svg.ns_ddx_svg_run.__init__(self)

        elif click in ('2-4-1', '2-4-2'):
            _render_l1_per_area_svg(self, ppt_meta_file, ws_name, bulk, click)


def _render_l1_per_area_svg(ctx, ppt_meta_file, ws_name, orig_bulk, click):
    """Render L1 per-area SVG: each area as a panel stacked vertically."""
    folder_wp_name_array = nsm_def.get_folder_wp_array_from_master(ws_name, ppt_meta_file)

    if not folder_wp_name_array or not folder_wp_name_array[0]:
        return

    # Build wp_with_folder_tuple (WP device → WP folder name)
    wp_with_folder_tuple = {}
    for tmp_wp_folder_name in folder_wp_name_array[1]:
        current_row = 1
        flag_start = False
        flag_end = False
        start_row = end_row = 1
        while not flag_end:
            v = ctx.position_shape_tuple.get((current_row, 1), '')
            if str(v) == tmp_wp_folder_name:
                start_row = current_row
                flag_start = True
            if flag_start and str(v) == '<END>':
                flag_end = True
                end_row = current_row - 1
            current_row += 1
        for i in range(start_row, end_row + 1):
            col = 2
            while True:
                val = ctx.position_shape_tuple.get((i, col), '')
                if str(val) == '<END>' or val == '' or val is None:
                    break
                wp_with_folder_tuple[val] = tmp_wp_folder_name
                col += 1

    min_tag_inches = 0.3
    all_slide_max_width = 10.0
    all_slide_max_hight = 5.0

    # Pass 1: collect per-area extract data and find max slide dimensions
    per_area_data = []
    for tmp_folder_name in folder_wp_name_array[0]:
        # Find shapes in this folder
        current_row = 1
        flag_start = False
        flag_end = False
        start_row = end_row = 1
        while not flag_end:
            v = ctx.position_shape_tuple.get((current_row, 1), '')
            if str(v) == tmp_folder_name:
                start_row = current_row
                flag_start = True
            if flag_start and str(v) == '<END>':
                flag_end = True
                end_row = current_row - 1
            current_row += 1

        tmp_folder_array = []
        for i in range(start_row, end_row + 1):
            col = 2
            while True:
                val = ctx.position_shape_tuple.get((i, col), '')
                if str(val) == '<END>' or val == '' or val is None:
                    break
                tmp_folder_array.append(val)
                col += 1

        # Find connected WP folders
        connected_wp_folder_array = []
        for shape_name in tmp_folder_array:
            for k in ctx.position_line_tuple:
                if k[0] == 1:
                    continue
                v1 = ctx.position_line_tuple.get((k[0], 1), '')
                v2 = ctx.position_line_tuple.get((k[0], 2), '')
                if shape_name == str(v1) and str(v2) in wp_with_folder_tuple:
                    wf = str(wp_with_folder_tuple[str(v2)])
                    if wf not in connected_wp_folder_array:
                        connected_wp_folder_array.append(wf)
                if shape_name == str(v2) and str(v1) in wp_with_folder_tuple:
                    wf = str(wp_with_folder_tuple[str(v1)])
                    if wf not in connected_wp_folder_array:
                        connected_wp_folder_array.append(wf)

        # Build extract_folder_tuple for this area
        extract_folder_tuple = {}
        for k, v in ctx.position_folder_tuple.items():
            if v == tmp_folder_name or v in connected_wp_folder_array:
                extract_folder_tuple[k] = v
                extract_folder_tuple[(k[0] - 1, k[1])] = ctx.position_folder_tuple.get((k[0] - 1, k[1]), '')
                extract_folder_tuple[(k[0], 1)] = ctx.position_folder_tuple.get((k[0], 1), '')
                extract_folder_tuple[(k[0] - 1, 1)] = ctx.position_folder_tuple.get((k[0] - 1, 1), '')

        convert_array = nsm_def.convert_tuple_to_array(extract_folder_tuple)
        if not convert_array:
            continue

        offset_row = convert_array[0][0] * -1 + 2
        current_y_grid_array = []
        if convert_array[0][1][0] == '<SET_WIDTH>':
            for arr in convert_array:
                current_y_grid_array.append([arr[0] + offset_row, arr[1]])
        else:
            for arr in convert_array:
                current_y_grid_array.append([arr[0] + offset_row - 1, arr[1]])

        convert_tuple = nsm_def.convert_array_to_tuple(current_y_grid_array)

        master_folder_size_array = nsm_def.get_folder_width_size(
            convert_tuple, ctx.position_style_shape_tuple,
            ctx.position_shape_tuple, min_tag_inches)

        master_root_folder_tuple = nsm_def.get_root_folder_tuple(
            ctx, master_folder_size_array, tmp_folder_name)

        try:
            rw = float(master_root_folder_tuple[2, 7])
            rh = float(master_root_folder_tuple[2, 8])
        except Exception:
            rw = float(master_folder_size_array[0]) + 1.0
            rh = float(master_folder_size_array[3]) + 1.0

        if rw > all_slide_max_width:
            all_slide_max_width = rw
        if rh + 1.0 > all_slide_max_hight:
            all_slide_max_hight = rh + 1.0

        per_area_data.append((tmp_folder_name, current_y_grid_array, convert_tuple))

    if not per_area_data:
        return

    # Pass 2: render each area at max dimensions, save one SVG per area
    save_svg_file = getattr(ctx, 'output_svg_file', None)
    orig_pf_array = ctx.position_folder_array
    orig_pf_tuple = ctx.position_folder_tuple
    orig_bulk_saved = getattr(ctx, '_preloaded_bulk', None)

    ctx.output_svg_file = None  # suppress file writes during per-area rendering

    # Derive per-area file path template: /dir/[L1_DIAGRAM]PerAreaTag_<base>_<area>.svg
    base_no_ext = os.path.splitext(save_svg_file)[0] if save_svg_file else None

    saved_svg_paths = []
    for tmp_folder_name, current_y_grid_array, convert_tuple in per_area_data:
        ctx.tmp_folder_name = tmp_folder_name
        ctx.position_folder_array = current_y_grid_array
        ctx.position_folder_tuple = convert_tuple

        area_bulk = dict(orig_bulk_saved) if orig_bulk_saved is not None else {}
        area_bulk['<<POSITION_FOLDER>>'] = current_y_grid_array
        if click == '2-4-1':
            area_bulk['<<POSITION_TAG>>'] = []
        ctx._preloaded_bulk = area_bulk

        ctx.root_width = all_slide_max_width
        ctx.root_hight = all_slide_max_hight + 1.0
        ctx.root_left = 0.28
        ctx.root_top = 1.42

        nsm_ddx_svg.ns_ddx_svg_run.__init__(ctx)

        svg_content = getattr(ctx, '_svg_content', '')
        if svg_content and base_no_ext:
            safe_area = _safe_area_name(tmp_folder_name)
            area_svg_path = base_no_ext + '_' + safe_area + '.svg'
            with open(area_svg_path, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            saved_svg_paths.append(area_svg_path)

    # Restore original state
    ctx.position_folder_array = orig_pf_array
    ctx.position_folder_tuple = orig_pf_tuple
    ctx._preloaded_bulk = orig_bulk_saved
    ctx.output_svg_file = save_svg_file

    # Store list of saved per-area SVG paths for caller use
    ctx._per_area_svg_files = saved_svg_paths


if __name__ == '__main__':
    nsm_l1_svg_create()
