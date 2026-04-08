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
L2 SVG diagram creation module.

Prepares layout data from the master file and delegates rendering
to nsm_ddx_svg_l2.  Mirrors the PPT flow in nsm_l2_diagram_create:
  1. Compute L2 device sizes (RETURN_DEVICE_SIZE equivalent)
  2. Override STYLE_SHAPE with L2-computed sizes
  3. Extract target folder + connected WP folders
  4. Recalculate root dimensions
  5. Render via nsm_ddx_svg_l2
"""

import copy
import os
import nsm_def


_L2_SECTIONS = [
    '<<POSITION_FOLDER>>', '<<POSITION_SHAPE>>', '<<POSITION_LINE>>',
    '<<STYLE_SHAPE>>', '<<STYLE_FOLDER>>', '<<POSITION_TAG>>',
    '<<ROOT_FOLDER>>',
]

# Per-process worker context for ProcessPoolExecutor
_SVG_WORKER_L2DATA = None


def _svg_size_worker_init(l2_data_serialized):
    """Pool initializer: restore l2_data once per worker process."""
    global _SVG_WORKER_L2DATA
    _SVG_WORKER_L2DATA = l2_data_serialized


def _svg_size_worker(device_name):
    """Worker function for parallel _compute_device_l2_size."""
    global _SVG_WORKER_L2DATA
    try:
        return (device_name, _compute_device_l2_size(_SVG_WORKER_L2DATA, device_name), None)
    except Exception as e:
        return (device_name, [0.0, 0.0, 1.0, 1.0], str(e))


def _compute_device_l2_size(l2_data, device_name):
    """Compute L2 bounding box for a device (RETURN_DEVICE_SIZE equivalent).

    Returns [left, top, width, height] in relative coordinates.
    """
    update_l2_table_array = l2_data['update_l2_table_array']
    device_l2name_array = l2_data['device_l2name_array']
    new_l2_table_array = l2_data['new_l2_table_array']
    new_direction_if_array = l2_data['new_direction_if_array']
    wp_list_array = l2_data['wp_list_array']

    shape_width_min = 0.5
    shape_hight_min = 0.1
    shape_interval_width_ratio = 0.75
    shape_interval_hight_ratio = 0.75
    between_tag = 0.2
    l2seg_size_margin = 0.7
    l2seg_size_margin_left_right_add = 0.4
    font_size = 6

    target_device_l2_array = list(l2_data.get('l2name_by_device', {}).get(device_name, []))
    target_device_l2_array.sort()

    flag_l2_segment_empty = len(target_device_l2_array) == 0
    if flag_l2_segment_empty:
        target_device_l2_array.append('_DummyL2Segment_')

    count = 0
    pre_shape_width = 0
    pre_offset_left = 0
    l2seg_size_array = []
    seg_offset_left = 0.0
    seg_offset_top = 0.0

    for seg_name in target_device_l2_array:
        wh = nsm_def.get_description_width_hight(font_size, str(seg_name))
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

        l2seg_size_array.append([seg_offset_left, seg_offset_top, seg_w, seg_h, seg_name])
        seg_offset_left += seg_w
        seg_offset_top += seg_h
        count += 1

    if not l2seg_size_array:
        return [0.0, 0.0, 1.0, 1.0]

    dev_left = l2seg_size_array[0][0]
    dev_top = l2seg_size_array[0][1]
    dev_right = l2seg_size_array[-1][0] + l2seg_size_array[-1][2]
    dev_bottom = l2seg_size_array[-1][1] + l2seg_size_array[-1][3]

    device_size = [
        dev_left - l2seg_size_margin,
        dev_top - l2seg_size_margin,
        (dev_right - dev_left) + l2seg_size_margin * 2,
        (dev_bottom - dev_top) + l2seg_size_margin * 2
    ]

    current_dir = l2_data.get('direction_by_device', {}).get(
        device_name, [device_name, [], [], [], []])
    for i in range(1, 5):
        if current_dir[i]:
            current_dir[i] = list(dict.fromkeys(current_dir[i]))

    target_vport_if_array = []
    for tmp in new_l2_table_array:
        row = tmp[1]
        if row[1] == device_name and row[5] != '':
            found = False
            for vp in target_vport_if_array:
                if vp[0] == row[5]:
                    vp[1].append(row[3])
                    found = True
            if not found:
                target_vport_if_array.append([row[5], [row[3]]])

    has_vport = [False, False, False, False]
    vport_l3_count = [0, 0, 0, 0]
    dir_map = {0: 1, 1: 2, 2: 3, 3: 4}
    for dir_idx in range(4):
        arr_idx = dir_map[dir_idx]
        if current_dir[arr_idx]:
            for if_name in current_dir[arr_idx]:
                for vp in target_vport_if_array:
                    if if_name in vp[1]:
                        has_vport[dir_idx] = True
                        for row in update_l2_table_array:
                            if (row[1] == device_name and row[5] == vp[0] and
                                    row[4] != 'Switch (L2)'):
                                vport_l3_count[dir_idx] += 1

    other_if_count = sum(1 for vp in target_vport_if_array if vp[1] == [''])

    if has_vport[0] or other_if_count > 0:
        device_size[1] -= l2seg_size_margin
        device_size[3] += l2seg_size_margin
    if has_vport[1]:
        device_size[3] += l2seg_size_margin
    if has_vport[2]:
        device_size[2] += (l2seg_size_margin + l2seg_size_margin_left_right_add)
    if has_vport[3]:
        device_size[0] -= (l2seg_size_margin + l2seg_size_margin_left_right_add)
        device_size[2] += (l2seg_size_margin + l2seg_size_margin_left_right_add)

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

    # --- Width expansion based on UP IF tag span (PPT line 3838-3843) ---
    has_if_up = len(current_dir[1]) > 0
    if has_if_up:
        up_tag_sum = l2seg_size_array[0][0] + l2seg_size_array[0][2] * 0.5
        last_tag_left = up_tag_sum
        last_tag_width = 0.0
        for if_name in current_dir[1]:
            tag_name = nsm_def.get_tag_name_from_full_name(
                device_name, if_name,
                l2_data.get('_position_line_tuple', {}))
            if tag_name == '_NO_MATCH_':
                arr = nsm_def.adjust_portname(if_name)
                tag_name = str(arr[0]) + ' ' + str(arr[2])
            tw = nsm_def.get_description_width_hight(font_size, tag_name)[0]
            last_tag_left = up_tag_sum
            last_tag_width = tw
            up_tag_sum += tw + between_tag

        l2seg_rightside = l2seg_size_array[-1][0] + l2seg_size_array[-1][2] + l2seg_size_margin
        if_tag_rightside = last_tag_left + last_tag_width + l2seg_size_margin
        if if_tag_rightside > l2seg_rightside:
            device_size[2] += (if_tag_rightside - l2seg_rightside)

        # UP vport L3/GRAY tags extend width further (PPT line 3916-3918)
        vp_right_edge = last_tag_left + last_tag_width
        offset_vport_l3 = 0.0
        up_if_set = set(current_dir[1])
        for row in l2_data.get('l2_rows_by_device', {}).get(device_name, []):
            if (row[5] != '' and (row[3] in up_if_set or row[3] == '')):
                vp_type = row[4]
                if 'L2' not in str(vp_type):
                    arr = nsm_def.adjust_portname(row[5])
                    vp_name = str(arr[0]) + ' ' + str(arr[2])
                    vp_w = nsm_def.get_description_width_hight(font_size, vp_name)[0]
                    offset_vport_l3 += vp_w + between_tag * 0.5
        vp_total_right = vp_right_edge + offset_vport_l3 + l2seg_size_margin
        if vp_total_right > (device_size[0] + device_size[2]):
            device_size[2] += (vp_total_right - (device_size[0] + device_size[2]))

    # --- Width expansion based on DOWN IF tag span (PPT line 3934-3939) ---
    has_if_down = len(current_dir[2]) > 0
    if has_if_down:
        down_tag_sum = l2seg_size_array[0][0]
        for if_name in current_dir[2]:
            tag_name = nsm_def.get_tag_name_from_full_name(
                device_name, if_name,
                l2_data.get('_position_line_tuple', {}))
            if tag_name == '_NO_MATCH_':
                arr = nsm_def.adjust_portname(if_name)
                tag_name = str(arr[0]) + ' ' + str(arr[2])
            tw = nsm_def.get_description_width_hight(font_size, tag_name)[0]
            down_tag_sum += tw + between_tag

        if down_tag_sum > l2seg_size_array[-1][0]:
            tag_offset = down_tag_sum - l2seg_size_array[-1][0]
            device_size[0] -= tag_offset
            device_size[2] += tag_offset

    # --- Width expansion for DOWN L3/GRAY vport tags (PPT line 4005-4010) ---
    used_down_vp = []
    down_if_set = set(current_dir[2]) if has_if_down else set()
    if has_if_down:
        for row in reversed(l2_data.get('l2_rows_by_device', {}).get(device_name, [])):
            if (row[5] != '' and
                    row[5] not in used_down_vp and
                    row[3] in down_if_set):
                if 'L2' not in str(row[4]):
                    arr = nsm_def.adjust_portname(row[5])
                    vp_name = str(arr[0]) + ' ' + str(arr[2])
                    vp_w = nsm_def.get_description_width_hight(font_size, vp_name)[0]
                    device_size[0] -= vp_w + between_tag * 0.5
                    device_size[2] += vp_w + between_tag * 0.5
                used_down_vp.append(row[5])

    # --- Ensure device frame is wide enough for device name text (16pt) ---
    name_w = nsm_def.get_description_width_hight(16, device_name)[0] + 0.2
    if device_size[2] < name_w:
        device_size[2] = name_w

    return device_size


class nsm_l2_svg_create:
    def __init__(self):
        ws_name = 'Master_Data'
        ws_l2_name = 'Master_Data_L2'
        ppt_meta_file = str(self.inFileTxt_L2_3_1.get())

        self.full_filepath = ppt_meta_file

        bulk = nsm_def.convert_master_to_arrays_bulk(ws_name, ppt_meta_file, _L2_SECTIONS)
        l2_bulk = nsm_def.convert_master_to_arrays_bulk(ws_l2_name, ppt_meta_file, ['<<L2_TABLE>>'])

        self.position_folder_array = bulk['<<POSITION_FOLDER>>']
        self.position_shape_array = bulk['<<POSITION_SHAPE>>']
        self.position_line_array = bulk['<<POSITION_LINE>>']
        self.position_style_shape_array = bulk['<<STYLE_SHAPE>>']
        self.position_tag_array = bulk['<<POSITION_TAG>>']
        self.root_folder_array = bulk['<<ROOT_FOLDER>>']
        self.l2_table_array = l2_bulk['<<L2_TABLE>>']

        self.position_folder_tuple = nsm_def.convert_array_to_tuple(self.position_folder_array)
        self.position_shape_tuple = nsm_def.convert_array_to_tuple(self.position_shape_array)
        self.position_line_tuple = nsm_def.convert_array_to_tuple(self.position_line_array)
        self.position_style_shape_tuple = nsm_def.convert_array_to_tuple(self.position_style_shape_array)
        self.position_tag_tuple = nsm_def.convert_array_to_tuple(self.position_tag_array)
        self.root_folder_tuple = nsm_def.convert_array_to_tuple(self.root_folder_array)

        self.folder_wp_name_array = nsm_def.get_folder_wp_array_from_master(ws_name, ppt_meta_file)

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

        target_area = str(self.comboL2_3_6.get())

        new_l2_table_array = []
        for item in self.l2_table_array:
            if item[0] != 1 and item[0] != 2:
                item[1].extend(['', '', '', '', '', '', '', ''])
                del item[1][8:]
                new_l2_table_array.append(item)

        device_list_array = []
        wp_list_array = []
        device_set = set()
        wp_set = set()

        for item in new_l2_table_array:
            device_name = item[1][1]
            area_name = item[1][0]
            if area_name == 'N/A':
                if device_name not in wp_set:
                    wp_set.add(device_name)
                    wp_list_array.append(device_name)
            else:
                if target_area == area_name:
                    if device_name not in device_set:
                        device_set.add(device_name)
                        device_list_array.append(device_name)

        self.device_list_array = device_list_array
        self.wp_list_array = wp_list_array
        self.target_area = target_area

        click = getattr(self, 'click_value', 'L2-3-2')

        if click == 'L2-3-2':
            # ===== STEP2: Compute L2 device sizes (PPT equivalent) =====
            import nsm_ddx_svg_l2
            l2_data = nsm_ddx_svg_l2._prepare_l2_data(self)
            l2_data['_position_line_tuple'] = self.position_line_tuple

            all_dev_names = device_list_array + wp_list_array
            device_size_dict = {}
            all_device_l2_size_array = []

            for dev_name in all_dev_names:
                size = _compute_device_l2_size(l2_data, dev_name)
                device_size_dict[dev_name] = size
                all_device_l2_size_array.append([dev_name, size])

            self.device_size_dict = device_size_dict
            self.all_device_l2_size_array = all_device_l2_size_array

            # ===== Override STYLE_SHAPE with L2-computed sizes =====
            l2_style_shape_array = []
            for item in self.position_style_shape_array:
                if item[0] in [1, 2, 3]:
                    l2_style_shape_array.append(item)
                else:
                    row = list(item[1])
                    dev_name = row[0]
                    if dev_name in device_size_dict:
                        size = device_size_dict[dev_name]
                        row[1] = size[2]
                        row[2] = size[3]
                    l2_style_shape_array.append([item[0], row])

            l2_style_shape_tuple = nsm_def.convert_array_to_tuple(l2_style_shape_array)

            # ===== Extract target folder + connected WP folders =====
            wp_with_folder_tuple = {}
            if hasattr(self, 'folder_wp_name_array') and len(self.folder_wp_name_array) > 1:
                for wp_folder_name in self.folder_wp_name_array[1]:
                    current_row = 1
                    flag_start = False
                    flag_end = False
                    start_row = 1
                    end_row = 1
                    while not flag_end:
                        key = (current_row, 1)
                        val = self.position_shape_tuple.get(key, '')
                        if str(val) == wp_folder_name:
                            start_row = current_row
                            flag_start = True
                        if flag_start and str(val) == '<END>':
                            flag_end = True
                            end_row = current_row - 1
                        current_row += 1
                        if current_row > 50000:
                            break

                    for i in range(start_row, end_row + 1):
                        col = 2
                        while True:
                            key2 = (i, col)
                            val2 = self.position_shape_tuple.get(key2, '')
                            if str(val2) == '<END>' or val2 == '' or val2 is None:
                                break
                            wp_with_folder_tuple[str(val2)] = wp_folder_name
                            col += 1

            connected_wp_folder_set = set()
            tmp_folder_array = []
            current_row = 1
            flag_start = False
            flag_end = False
            start_row = 1
            end_row = 1
            while not flag_end:
                key = (current_row, 1)
                val = self.position_shape_tuple.get(key, '')
                if str(val) == target_area:
                    start_row = current_row
                    flag_start = True
                if flag_start and str(val) == '<END>':
                    flag_end = True
                    end_row = current_row - 1
                current_row += 1
                if current_row > 50000:
                    break

            for i in range(start_row, end_row + 1):
                col = 2
                while True:
                    key2 = (i, col)
                    val2 = self.position_shape_tuple.get(key2, '')
                    if str(val2) == '<END>' or val2 == '' or val2 is None:
                        break
                    tmp_folder_array.append(str(val2))
                    col += 1

            for shape_name in tmp_folder_array:
                for plt_key in self.position_line_tuple:
                    if plt_key[0] == 1:
                        continue
                    val1 = self.position_line_tuple.get((plt_key[0], 1), '')
                    val2 = self.position_line_tuple.get((plt_key[0], 2), '')
                    if shape_name == str(val1):
                        if str(val2) in wp_with_folder_tuple:
                            connected_wp_folder_set.add(wp_with_folder_tuple[str(val2)])
                    if shape_name == str(val2):
                        if str(val1) in wp_with_folder_tuple:
                            connected_wp_folder_set.add(wp_with_folder_tuple[str(val1)])

            extract_folder_tuple = {}
            for key in self.position_folder_tuple:
                val = self.position_folder_tuple[key]
                if val == target_area or val in connected_wp_folder_set:
                    extract_folder_tuple[key] = val
                    extract_folder_tuple[(key[0] - 1, key[1])] = self.position_folder_tuple.get((key[0] - 1, key[1]), '')
                    extract_folder_tuple[(key[0], 1)] = self.position_folder_tuple.get((key[0], 1), '')
                    extract_folder_tuple[(key[0] - 1, 1)] = self.position_folder_tuple.get((key[0] - 1, 1), '')

            if extract_folder_tuple:
                convert_array = nsm_def.convert_tuple_to_array(extract_folder_tuple)
                offset_row = convert_array[0][0] * -1 + 2

                current_y_grid_array = []
                flag_first = True
                if convert_array[0][1][0] == '<SET_WIDTH>':
                    for arr in convert_array:
                        current_y_grid_array.append([arr[0] + offset_row, arr[1]])
                else:
                    for arr in convert_array:
                        current_y_grid_array.append([arr[0] + offset_row - 1, arr[1]])

                extract_folder_array = current_y_grid_array
                extract_folder_tuple_new = nsm_def.convert_array_to_tuple(extract_folder_array)
            else:
                extract_folder_tuple_new = self.position_folder_tuple
                extract_folder_array = self.position_folder_array

            # ===== Recalculate root dimensions with L2 sizes =====
            master_folder_size_array = nsm_def.get_folder_width_size(
                extract_folder_tuple_new, l2_style_shape_tuple,
                self.position_shape_tuple, 0.8)

            # ===== Update folder grid with measured sizes (PPT line 856-888) =====
            update_y_grid = copy.deepcopy(current_y_grid_array)

            for grid_key in extract_folder_tuple_new:
                folder_val = extract_folder_tuple_new[grid_key]
                if not isinstance(folder_val, str):
                    continue
                for measured in master_folder_size_array[2]:
                    if (folder_val == measured[1][0][0] and
                            measured[1][0][0] != 10 and
                            isinstance(measured[1][0][0], str)):
                        row_idx = grid_key[0]
                        col_idx = grid_key[1]
                        if measured[1][0][1] != 0:
                            for i, grid_row in enumerate(current_y_grid_array):
                                if grid_row[0] == row_idx - 1:
                                    if col_idx - 1 < len(update_y_grid[i][1]):
                                        update_y_grid[i][1][col_idx - 1] = measured[1][0][1]
                                    break

            for i, grid_row in enumerate(current_y_grid_array):
                if isinstance(grid_row[1][0], (int, float)):
                    max_h = grid_row[1][0]
                    for col_entry in grid_row[1:]:
                        if isinstance(col_entry, list):
                            for sub in col_entry:
                                for measured in master_folder_size_array[2]:
                                    if measured[1][0][0] == sub:
                                        if max_h < measured[1][0][2]:
                                            max_h = measured[1][0][2]
                        elif isinstance(col_entry, str):
                            for measured in master_folder_size_array[2]:
                                if measured[1][0][0] == col_entry:
                                    if max_h < measured[1][0][2]:
                                        max_h = measured[1][0][2]
                    update_y_grid[i][1][0] = max_h

            extract_folder_array = update_y_grid
            extract_folder_tuple_new = nsm_def.convert_array_to_tuple(update_y_grid)

            # ===== Set root dimensions =====
            try:
                master_root = nsm_def.get_root_folder_tuple(self, master_folder_size_array, target_area)
                self.root_width = master_root[2, 7]
                self.root_hight = master_root[2, 8]
            except Exception:
                self.root_width = float(master_folder_size_array[0]) + 1.0
                self.root_hight = float(master_folder_size_array[3]) + 1.0

            self.root_left = 0.28
            self.root_top = 1.42

            # Pass L2-overridden bulk to renderer
            l2_bulk_override = copy.deepcopy(bulk)
            l2_bulk_override['<<STYLE_SHAPE>>'] = l2_style_shape_array
            l2_bulk_override['<<POSITION_FOLDER>>'] = extract_folder_array
            self._preloaded_bulk = l2_bulk_override

            nsm_ddx_svg_l2.ns_ddx_svg_l2_run.__init__(self)


if __name__ == '__main__':
    nsm_l2_svg_create()
