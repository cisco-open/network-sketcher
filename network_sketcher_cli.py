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

import tkinter as tk ,tkinter.ttk as ttk,tkinter.filedialog, tkinter.messagebox
import sys, os, subprocess ,webbrowser ,openpyxl

class ns_cli_run():
    def __init__(self,argv_array):
        # GET file path of master
        master_file_path = get_next_arg(argv_array, '--master')
        if master_file_path == None:
            print('[ERROR] Master file must be selected --master <filepath>.')
            exit()

        # run show command
        if 'show' in argv_array:
            print_type(self,argv_array, ns_cli_run.cli_show(self, master_file_path,argv_array))
            exit()
        elif 'add' in argv_array:
            print_type(self,argv_array, ns_cli_run.cli_add(self, master_file_path,argv_array))
            exit()
        elif 'rename' in argv_array:
            print_type(self,argv_array, ns_cli_run.cli_rename(self, master_file_path,argv_array))
            exit()
        elif 'delete' in argv_array:
            print_type(self, argv_array, ns_cli_run.cli_delete(self, master_file_path, argv_array))
            exit()
        else:
            print('[ERROR] Supported commands are as follows', 'add', 'delete', 'rename', 'show', sep='\n')

    def cli_rename(self, master_file_path, argv_array): # add at ver 2.5.4
        import ns_def
        next_arg = get_next_arg(argv_array, 'rename')
        rename_command_list = [ \
            'rename area', \
            'rename device', \
            'rename l3_instance', \
            'rename port', \
            ]

        if next_arg == None or '--' in next_arg or str('rename ' + next_arg) not in rename_command_list:
            print('[ERROR] Supported commands are as follows')
            for tmp_add_command_list in rename_command_list:
                print(tmp_add_command_list)
            exit()

        if next_arg == 'area':
            idx = argv_array.index('area')
            updated_name_array = [str(argv_array[idx + 1]), str(argv_array[idx + 2]).replace(' ', '').replace('　', '')]

            if isinstance(updated_name_array[1], str) and updated_name_array[1].strip(' \u3000') == '':
                return_text = '[ERROR] This area name is invalid.  --- ' + ' ' + updated_name_array[1]
                return ([return_text])
            if  '--' in str(argv_array[idx + 2]):
                return_text = '[ERROR] This area name is invalid.  --- ' + ' ' + updated_name_array[1]
                return ([return_text])

            '''rename area name in Master_Data.'''
            ws_name = 'Master_Data'
            ppt_meta_file = master_file_path

            position_folder_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_FOLDER>>')
            style_folder_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<STYLE_FOLDER>>')
            position_shape_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_SHAPE>>')

            # Replace updated_name_array[0] with updated_name_array[1] (no def).
            src = updated_name_array[0]
            dst = updated_name_array[1]
            sentinels = {"<SET_WIDTH>", "<DEFAULT>", "<END>","N/A","<EMPTY>"}

            # position_folder_array and style_folder_array: target all elements in the inner list
            # (exclude sentinels; exclude items whose leading number is 1).
            for arr in (position_folder_array, style_folder_array):
                for item in arr:
                    if not isinstance(item, list) or len(item) < 2:
                        continue
                    idx, values = item[0], item[1]
                    if idx == 1 or not isinstance(values, list):
                        continue
                    for j, val in enumerate(values):
                        if isinstance(val, str) and val == src and val not in sentinels:
                            values[j] = dst

            # position_shape_array: only the first item of the second element (the inner list) is targeted
            # (exclude sentinels; exclude items whose leading number is 1).
            for item in position_shape_array:
                if not isinstance(item, list) or len(item) < 2:
                    continue
                idx, values = item[0], item[1]
                if idx == 1 or not isinstance(values, list) or len(values) == 0:
                    continue
                if updated_name_array[1] == item[1][0]:
                    return_text = '[ERROR] This Area name is already exist.  --- ' + updated_name_array[1]
                    return ([return_text])
                first = values[0]
                if isinstance(first, str) and first == src and first not in sentinels:
                    values[0] = dst

            '''rename area name in Master_Data_L2 and Master_Data_L3'''
            ws_name_l2 = 'Master_Data_L2'
            l2_table_array = ns_def.convert_master_to_array(ws_name_l2, ppt_meta_file, '<<L2_TABLE>>')
            ws_name_l3 = 'Master_Data_L3'
            l3_table_array = ns_def.convert_master_to_array(ws_name_l3, ppt_meta_file, '<<L3_TABLE>>')

            # l2_table_array: only replace the first item of the inner list when it matches src and is not a sentinel.
            for item in l2_table_array:
                if not isinstance(item, list) or len(item) < 2:
                    continue
                idx, values = item[0], item[1]
                if idx in (1, 2) or not isinstance(values, list) or len(values) == 0:
                    continue
                first = values[0]
                if isinstance(first, str) and first == src and first not in sentinels:
                    values[0] = dst

            # l3_table_array: only replace the first item of the inner list when it matches src and is not a sentinel.
            for item in l3_table_array:
                if not isinstance(item, list) or len(item) < 2:
                    continue
                idx, values = item[0], item[1]
                if idx in (1, 2) or not isinstance(values, list) or len(values) == 0:
                    continue
                first = values[0]
                if isinstance(first, str) and first == src and first not in sentinels:
                    values[0] = dst

            self.position_folder_tuple = ns_def.convert_array_to_tuple(position_folder_array)
            self.style_folder_tuple = ns_def.convert_array_to_tuple(style_folder_array)
            self.position_shape_tuple = ns_def.convert_array_to_tuple(position_shape_array)
            self.l2_table_tuple = ns_def.convert_array_to_tuple(l2_table_array)
            self.l3_table_tuple = ns_def.convert_array_to_tuple(l3_table_array)

            excel_file_path = ppt_meta_file
            worksheet_name = ws_name
            offset_row = 0
            offset_column = 0

            master_excel_meta = self.position_folder_tuple
            section_write_to = '<<POSITION_FOLDER>>'
            ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

            master_excel_meta = self.style_folder_tuple
            section_write_to = '<<STYLE_FOLDER>>'
            ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

            master_excel_meta = self.position_shape_tuple
            section_write_to = '<<POSITION_SHAPE>>'
            ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

            worksheet_name = ws_name_l2
            master_excel_meta = self.l2_table_tuple
            section_write_to = '<<L2_TABLE>>'
            ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

            worksheet_name = ws_name_l3
            master_excel_meta = self.l3_table_tuple
            section_write_to = '<<L3_TABLE>>'
            ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

            return_text = '--- Area renamed --- ' + str(updated_name_array[0]) + ' -> ' + str(updated_name_array[1])
            return ([return_text])


        if next_arg == 'port':
            import ns_def
            idx = argv_array.index('port')
            updated_name_array = [str(argv_array[idx + 1]), str(argv_array[idx + 2]),str(argv_array[idx + 3])]
            updated_name_array.append(ns_def.adjust_portname(updated_name_array[1]))
            updated_name_array.append(ns_def.adjust_portname(updated_name_array[2]))
            flag_l1_port_name = False
            flag_l2_port_name = False

            if isinstance(updated_name_array[0], str) and updated_name_array[0].strip(' \u3000') == '':
                return_text = '[ERROR] This device name is invalid.  --- ' + ' ' + updated_name_array[0]
                return ([return_text])

            if ns_def.get_if_value(str(updated_name_array[1])) == -1 or ns_def.get_if_value(str(updated_name_array[2])) == -1:
                return_text = '[ERROR] The port name is invalid.  --- ' + ' ' + updated_name_array[1] + '  or  ' + updated_name_array[2]
                return ([return_text])


            '''
            rename L1 port name in Master_Data
            '''
            ws_name = 'Master_Data'
            ppt_meta_file = master_file_path

            position_line_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_LINE>>')

            # position_line_array: The first and second lines are not included. Only the first (0) and second (1) are replaced.
            for row in position_line_array:
                if isinstance(row, list) and len(row) == 2:
                    line_no, fields = row
                    if line_no not in (1, 2) and isinstance(fields, list):
                        #print(fields)

                        if fields[0] == updated_name_array[0]:
                            if updated_name_array[4][1] == fields[12] and updated_name_array[4][2] == str(ns_def.split_portname(fields[2])[1]):
                                return_text = '[ERROR] This port name is already exist.  --- ' + updated_name_array[0]  + ' ' + updated_name_array[2]
                                return ([return_text])

                            #print(updated_name_array[0] , fields[12],ns_def.split_portname(fields[2])[1])
                            if updated_name_array[3][1] == fields[12] and updated_name_array[3][2] == str(ns_def.split_portname(fields[2])[1]):
                                fields[2] = updated_name_array[4][0] + ' ' + updated_name_array[4][2]
                                fields[12] = updated_name_array[4][1]
                                flag_l1_port_name = True

                        if fields[1] == updated_name_array[0]:
                            if updated_name_array[4][1] == fields[16] and updated_name_array[4][2] == str(ns_def.split_portname(fields[3])[1]):
                                return_text = '[ERROR] This port name is already exist.  --- ' + updated_name_array[0] + ' ' + updated_name_array[2]
                                return ([return_text])

                            #print(updated_name_array[0] , fields[16],ns_def.split_portname(fields[3])[1])
                            if updated_name_array[3][1] == fields[16] and updated_name_array[3][2] == str(ns_def.split_portname(fields[3])[1]):
                                fields[3] = updated_name_array[4][0] + ' ' + updated_name_array[4][2]
                                fields[16] = updated_name_array[4][1]
                                flag_l1_port_name = True

            if flag_l1_port_name == True:
                excel_file_path = ppt_meta_file
                worksheet_name = ws_name
                offset_row = 0
                offset_column = 0

                self.position_line_tuple = ns_def.convert_array_to_tuple(position_line_array)
                master_excel_meta = self.position_line_tuple
                section_write_to = '<<POSITION_LINE>>'
                ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

                '''rename portname in Master_Data_L2 and Master_Data_L3.'''
                self.full_filepath = master_file_path
                self.update_port_num_array = [[updated_name_array[0],updated_name_array[1],updated_name_array[2]]]
                import ns_sync_between_layers
                ns_sync_between_layers.l1_device_port_name_sync_with_l2l3_master(self)

                return_text = '--- Physical Port Name renamed --- ' + updated_name_array[0] + ' ' + updated_name_array[1] + ' -> ' + updated_name_array[2]
                return ([return_text])

            '''
            rename L2 port name in Master_Data
            '''
            ws_name = 'Master_Data_L2'
            l2_table_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<L2_TABLE>>')
            #print(l2_table_array)

            # l2_table_array: The first and second lines are not included. Only the first (0) and second (1) are replaced.
            for row in l2_table_array:
                if isinstance(row, list) and len(row) == 2:
                    line_no, fields = row
                    if line_no not in (1, 2) and isinstance(fields, list):
                        #print(fields)

                        if fields[1] == updated_name_array[0]:
                            #print(updated_name_array)
                            if updated_name_array[2] == fields[5]:
                                return_text = '[ERROR] This port name is already exist.  --- ' + updated_name_array[0]  + ' ' + updated_name_array[2]
                                return ([return_text])

                            #print(updated_name_array)
                            if updated_name_array[1] == fields[5]:
                                fields[5] = updated_name_array[2]
                                flag_l2_port_name = True

            if flag_l2_port_name == True:
                excel_file_path = ppt_meta_file
                worksheet_name = ws_name
                offset_row = 0
                offset_column = 0

                self.l2_table_tuple = ns_def.convert_array_to_tuple(l2_table_array)
                master_excel_meta = self.l2_table_tuple
                section_write_to = '<<L2_TABLE>>'
                ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

                '''rename Virtual port name in Master_Data_L3.'''
                ws_name = 'Master_Data_L3'
                l3_table_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<L3_TABLE>>')
                #print(l3_table_array)

                # l2_table_array: The first and second lines are not included. Only the first (0) and second (1) are replaced.
                for row in l3_table_array:
                    if isinstance(row, list) and len(row) == 2:
                        line_no, fields = row
                        if line_no not in (1, 2) and isinstance(fields, list):
                            #print(fields)

                            if fields[1] == updated_name_array[0]:
                                #print(updated_name_array)
                                if updated_name_array[1] == fields[2]:
                                    fields[2] = updated_name_array[2]
                                    break

                self.l3_table_tuple = ns_def.convert_array_to_tuple(l3_table_array)
                master_excel_meta = self.l3_table_tuple
                section_write_to = '<<L3_TABLE>>'
                worksheet_name = 'Master_Data_L3'
                ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

                return_text = '--- Virtual Port Name renamed --- ' + updated_name_array[0] + ' ' + updated_name_array[1] + ' -> ' + updated_name_array[2]
                return ([return_text])

            return_text = '[ERROR]The port name did not exist. --- ' + updated_name_array[0] + ' ' + updated_name_array[1]
            return ([return_text])

        if next_arg == 'device':
            idx = argv_array.index('device')
            updated_name_array = [[str(argv_array[idx + 1]), str(argv_array[idx + 2]).replace(' ', '').replace('　', '')]]

            if isinstance(updated_name_array[0][1], str) and updated_name_array[0][1].strip(' \u3000') == '':
                return_text = '[ERROR] This device name is invalid.  --- ' + ' ' + updated_name_array[0][1]
                return ([return_text])

            self.updated_name_array = updated_name_array

            '''rename device name in Master_Data.'''
            ws_name = 'Master_Data'
            ppt_meta_file = master_file_path

            position_shape_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_SHAPE>>')
            position_line_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_LINE>>')
            position_style_shape_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<STYLE_SHAPE>>')
            position_tag_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_TAG>>')
            attribute_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<ATTRIBUTE>>')

            mapping = {}
            for pair in updated_name_array:
                if isinstance(pair, (list, tuple)) and len(pair) == 2:
                    old, new = pair
                    if isinstance(old, str) and isinstance(new, str) and old != '':
                        mapping[old] = new

            # position_shape_array: The first line is ignored, and all other string elements in the line are replaced.
            for row in position_shape_array:
                if isinstance(row, list) and len(row) == 2:
                    line_no, fields = row
                    if line_no != 1 and isinstance(fields, list):
                        for i in range(len(fields)):
                            v = fields[i]
                            if isinstance(v, str) and v in mapping:
                                fields[i] = mapping[v]

            # position_line_array: The first and second lines are not included. Only the first (0) and second (1) are replaced.
            for row in position_line_array:
                if isinstance(row, list) and len(row) == 2:
                    line_no, fields = row
                    if line_no not in (1, 2) and isinstance(fields, list):
                        if len(fields) > 0 and isinstance(fields[0], str) and fields[0] in mapping:
                            fields[0] = mapping[fields[0]]
                        if len(fields) > 1 and isinstance(fields[1], str) and fields[1] in mapping:
                            fields[1] = mapping[fields[1]]

            # position_style_shape_array: Lines 1, 2, and 3 are not included. For the others, only the first (0) is replaced.
            flag_source_exist = False
            for row in position_style_shape_array:
                if isinstance(row, list) and len(row) == 2:
                    line_no, fields = row
                    if line_no not in (1, 2, 3) and isinstance(fields, list) and len(fields) > 0:
                        if isinstance(fields[0], str) and fields[0] == updated_name_array[0][1]:
                            return_text = '[ERROR] The device name already exists.  --- ' + ' ' + updated_name_array[0][1]
                            return ([return_text])

                        if isinstance(fields[0], str) and fields[0] in mapping:
                            fields[0] = mapping[fields[0]]
                            flag_source_exist = True
            if flag_source_exist == False:
                return_text = '[ERROR] This device name does not exist.  --- ' + ' ' + updated_name_array[0][0]
                return ([return_text])

            # position_tag_array: The first and second lines are not included. For the others, only the first (0) is replaced.
            for row in position_tag_array:
                if isinstance(row, list) and len(row) == 2:
                    line_no, fields = row
                    if line_no not in (1, 2) and isinstance(fields, list) and len(fields) > 0:
                        if isinstance(fields[0], str) and fields[0] in mapping:
                            fields[0] = mapping[fields[0]]

            # attribute_array: The first line is not included. For the others, only the first (0) is replaced.
            for row in attribute_array:
                if isinstance(row, list) and len(row) == 2:
                    line_no, fields = row
                    if line_no != 1 and isinstance(fields, list) and len(fields) > 0:
                        if isinstance(fields[0], str) and fields[0] in mapping:
                            fields[0] = mapping[fields[0]]

            self.position_shape_tuple = ns_def.convert_array_to_tuple(position_shape_array)
            self.position_line_tuple = ns_def.convert_array_to_tuple(position_line_array)
            self.position_style_shape_tuple = ns_def.convert_array_to_tuple(position_style_shape_array)
            self.position_tag_tuple = ns_def.convert_array_to_tuple(position_tag_array)
            self.attribute_tuple = ns_def.convert_array_to_tuple(attribute_array)

            excel_file_path = ppt_meta_file
            worksheet_name = ws_name
            offset_row = 0
            offset_column = 0

            master_excel_meta = self.position_shape_tuple
            section_write_to = '<<POSITION_SHAPE>>'
            ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)
            master_excel_meta = self.position_line_tuple
            section_write_to = '<<POSITION_LINE>>'
            ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)
            master_excel_meta = self.position_style_shape_tuple
            section_write_to = '<<STYLE_SHAPE>>'
            ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)
            master_excel_meta = self.position_tag_tuple
            section_write_to = '<<POSITION_TAG>>'
            ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)
            master_excel_meta = self.attribute_tuple
            section_write_to = '<<ATTRIBUTE>>'
            ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

            '''rename device name in Master_Data_L2 and Master_Data_L3.'''
            self.full_filepath = master_file_path
            import ns_sync_between_layers
            ns_sync_between_layers.l1_sketch_device_name_sync_with_l2l3_master(self)

            return_text = '--- Device Name renamed --- ' + ' ' + updated_name_array[0][0] + ' -> ' + updated_name_array[0][1]
            return ([return_text])

        if next_arg == 'l3_instance':
            import ns_def
            l3_attribute_array = ns_def.convert_master_to_array('Master_Data_L3', master_file_path,'<<L3_TABLE>>')

            TARGET_LEN = 7
            for row in l3_attribute_array:
                if isinstance(row, list) and len(row) >= 2 and isinstance(row[1], list):
                    vals = row[1]
                    if len(vals) < TARGET_LEN:
                        vals += [''] * (TARGET_LEN - len(vals))

            idx = argv_array.index('l3_instance')

            target = idx + 3
            if target >= len(argv_array):
                argv_array.extend([None] * (target - len(argv_array) + 1))
                argv_array[target] = '__default__'
            elif '--' in argv_array[target]:
                argv_array.insert(target, '__default__')

            try:
                # The element right after 'l3_instance' is the hostname
                hostname = argv_array[idx + 1]
                # The next element is the portname
                portname = argv_array[idx + 2]
                # The next element is the l3 instance name
                add_l3instance_name = argv_array[idx + 3]
            except IndexError:
                # If there are not enough elements after 'l3_instance', print an error
                return ([f"[ERROR] hostname or portname or l3_instance is missing"])

            match_found = False
            for row in l3_attribute_array[2:]:
                values = row[1]
                if len(values) > 1 and values[1] == hostname:
                    if (len(values) > 3 and values[2] == portname):
                        match_found = True
                        break
            if not match_found:
                return ([f"Error: No matching entry found for hostname: {hostname} and portname: {portname}"])

            add_l3instance_name = add_l3instance_name.replace(' ', '')

            for entry in l3_attribute_array:
                if entry == row:
                    # Get the 4th element (index 3)
                    entry[1][3] = add_l3instance_name

                    if add_l3instance_name == '__default__':
                        entry[1][3] = ''
                        add_l3instance_name = ''

                    break


            #write to Master file
            excel_maseter_file = master_file_path
            excel_master_ws_name_l3 = 'Master_Data_L3'
            last_l3_table_tuple = {}
            last_l3_table_tuple = ns_def.convert_array_to_tuple(l3_attribute_array)

            # delete L2 Table sheet
            ns_def.remove_excel_sheet(excel_maseter_file, excel_master_ws_name_l3)
            # create L2 Table sheet
            ns_def.create_excel_sheet(excel_maseter_file, excel_master_ws_name_l3)
            # write tuple to excel master data
            ns_def.write_excel_meta(last_l3_table_tuple, excel_maseter_file, excel_master_ws_name_l3, '_template_', 0,0)

            return_text = '--- l3 instance renamed --- ' + ' ' + hostname + ',' + portname + ',' + add_l3instance_name
            return ([return_text])

    def cli_add(self, master_file_path, argv_array): # add at ver 2.5.3
        import ns_def
        next_arg = get_next_arg(argv_array, 'add')
        add_command_list = [ \
            'add ip_address', \
            'add l2_segment', \
            'add portchannel', \
            'add virtual_port', \
            'add vport_l1if_direct_binding', \
            'add vport_l2_direct_binding', \
            ]

        if next_arg == None or '--' in next_arg or str('add ' + next_arg) not in add_command_list:
            print('[ERROR] Supported commands are as follows')
            for tmp_add_command_list in add_command_list:
                print(tmp_add_command_list)
            exit()

        if next_arg == 'ip_address':
            import ns_def
            l3_attribute_array = ns_def.convert_master_to_array('Master_Data_L3', master_file_path,'<<L3_TABLE>>')

            TARGET_LEN = 7
            for row in l3_attribute_array:
                if isinstance(row, list) and len(row) >= 2 and isinstance(row[1], list):
                    vals = row[1]
                    if len(vals) < TARGET_LEN:
                        vals += [''] * (TARGET_LEN - len(vals))

            idx = argv_array.index('ip_address')
            try:
                # The element right after 'l2_segment' is the hostname
                hostname = argv_array[idx + 1]
                # The next element is the portname
                portname = argv_array[idx + 2]
                add_ipaddress_name = argv_array[idx + 3]
            except IndexError:
                return ([f"[ERROR] hostname or portname or ip address is missing"])

            # check IP Addresses
            if ns_def.check_ip_format(add_ipaddress_name) != 'IPv4':
                return ([f"Error: IP Address format is invalid: {add_ipaddress_name}"])

            match_found = False
            for row in l3_attribute_array[2:]:
                values = row[1]
                if len(values) > 1 and values[1] == hostname:
                    if (len(values) > 3 and values[2] == portname):
                        match_found = True
                        break
            if not match_found:
                return ([f"Error: No matching entry found for hostname: {hostname} and portname: {portname}"])

            add_ipaddress_name = add_ipaddress_name.replace(' ', '')

            for entry in l3_attribute_array:
                if entry == row:
                    # Get the 5th element (index 4)
                    current_value = entry[1][4]
                    # Split by comma, remove extra spaces
                    ips = [seg.strip() for seg in current_value.split(',')] if current_value else []
                    # Only add if not already present
                    if add_ipaddress_name not in ips:
                        if current_value.strip() == '':
                            entry[1][4] = add_ipaddress_name
                        else:
                            entry[1][4] = current_value + ',' + add_ipaddress_name
                    else:
                        return ([add_ipaddress_name + ' already exists in the ipaddress. No change made'])
                    break


            #write to Master file
            excel_maseter_file = master_file_path
            excel_master_ws_name_l3 = 'Master_Data_L3'
            last_l3_table_tuple = {}
            last_l3_table_tuple = ns_def.convert_array_to_tuple(l3_attribute_array)

            # delete L2 Table sheet
            ns_def.remove_excel_sheet(excel_maseter_file, excel_master_ws_name_l3)
            # create L2 Table sheet
            ns_def.create_excel_sheet(excel_maseter_file, excel_master_ws_name_l3)
            # write tuple to excel master data
            ns_def.write_excel_meta(last_l3_table_tuple, excel_maseter_file, excel_master_ws_name_l3, '_template_', 0,0)

            return_text = '--- IP Address added --- ' + ' ' + hostname + ',' + portname + ',' + add_ipaddress_name
            return ([return_text])

        if next_arg == 'l2_segment' or 'virtual_port' or 'portchannel' or 'vport_l2_direct_binding' or 'vport_l1if_direct_binding':
            import ns_def
            l2_attribute_array = ns_def.convert_master_to_array('Master_Data_L2', master_file_path,'<<L2_TABLE>>')

            TARGET_LEN = 9
            for row in l2_attribute_array:
                if isinstance(row, list) and len(row) >= 2 and isinstance(row[1], list):
                    vals = row[1]
                    if len(vals) < TARGET_LEN:
                        vals += [''] * (TARGET_LEN - len(vals))

            #print(argv_array)
            # Check if the input data exists in argv_array
            if 'l2_segment' in argv_array:
                idx = argv_array.index('l2_segment')
                try:
                    # The element right after 'l2_segment' is the hostname
                    hostname = argv_array[idx + 1]
                    # The next element is the portname
                    portname = argv_array[idx + 2]
                    add_l2seg_name = argv_array[idx + 3]
                except IndexError:
                    # If there are not enough elements after 'l2_segment', print an error
                    return (['[Error] hostname or portname is missing'])

            elif 'virtual_port' in argv_array:
                idx = argv_array.index('virtual_port')
                try:
                    hostname = argv_array[idx + 1]
                    add_vport_name = argv_array[idx + 2]
                except IndexError:
                    return (['[Error] hostname is missing'])

            elif 'portchannel' in argv_array:
                idx = argv_array.index('portchannel')
                try:
                    # The element right after 'portchannel' is the hostname
                    hostname = argv_array[idx + 1]
                    # The next element is the portname
                    portname = argv_array[idx + 2]
                    add_portchannel_name = argv_array[idx + 3]
                except IndexError:
                    # If there are not enough elements after 'portchannel', print an error
                    return (['[Error] hostname or portname is missing'])

            elif 'vport_l2_direct_binding' in argv_array:
                idx = argv_array.index('vport_l2_direct_binding')
                try:
                    # The element right after 'vport_l2_direct_binding' is the hostname
                    hostname = argv_array[idx + 1]
                    # The next element is the portname
                    portname = argv_array[idx + 2]
                    add_l2seg_name = argv_array[idx + 3]
                except IndexError:
                    # If there are not enough elements after 'vport_l2_direct_binding', print an error
                    return (['[Error] hostname or portname is missing'])

            elif 'vport_l1if_direct_binding' in argv_array:
                idx = argv_array.index('vport_l1if_direct_binding')
                try:
                    # The element right after 'vport_l1if_direct_binding' is the hostname
                    hostname = argv_array[idx + 1]
                    # The next element is the portname
                    portname = argv_array[idx + 2]
                    add_vport_name = argv_array[idx + 3]
                except IndexError:
                    # If there are not enough elements after 'vport_l1if_direct_binding', print an error
                    return (['[Error] hostname or portname is missing'])

            else:
                print("Not found in arguments")

            match_found = False
            row_array = []
            flag_l2_segment_vport = False
            if 'l2_segment' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if len(values) > 3 and values[3] == portname:
                            if values[7] != '':
                                return ([f"[ERROR] Any vport_l2_direct_binding is already configured for hostname: {hostname} and virtual portname: {portname}"])
                            match_found = True
                            break
                        elif len(values) > 5 and values[5] == portname:
                            if values[7] != '':
                                return ([f"[ERROR] Any vport_l2_direct_binding is already configured for hostname: {hostname} and virtual portname: {portname}"])
                            #for virtual port
                            match_found = True
                            flag_l2_segment_vport = True
                            row_array.append(row)
                if flag_l2_segment_vport == True:
                    row = []

            elif 'virtual_port' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        match_found = True
                        break

                #check to L1 interface duplicate
                for row3 in l2_attribute_array[2:]:
                    values2 = row3[1]
                    if len(values2) > 3 and values2[1] == hostname and values2[3] == add_vport_name :
                        return (["[ERROR] There is a Layer 1 interface with the same name as the virtual port."])

            elif 'portchannel' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if (len(values) > 3 and values[3] == portname) or (len(values) > 5 and values[5] == portname):
                            match_found = True
                            break

            elif 'vport_l2_direct_binding' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if len(values) > 5 and values[5] == portname:
                            if values[6] != '':
                                return ([f"[ERROR] Any L2 segment is already configured for hostname: {hostname} and virtual portname: {portname}"])
                            #for virtual port
                            match_found = True
                            flag_l2_segment_vport = True
                            row_array.append(row)
                if flag_l2_segment_vport == True:
                    row = []

            elif 'vport_l1if_direct_binding' in argv_array:
                flag_vport_l1if_direct_binding_exists = False

                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if (len(values) > 3 and values[3] == portname) and (len(values) > 5 and values[5] == add_vport_name):
                            return ([f"[ERROR] {add_vport_name} is already configured for hostname: {hostname} and portname: {portname}"])

                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if (len(values) > 3 and values[3] == portname):
                            match_found = True
                            if (len(values) > 5 and values[5] != ''):
                                flag_vport_l1if_direct_binding_exists = True
                            break

                #check to L1 interface duplicate
                for row3 in l2_attribute_array[2:]:
                    values2 = row3[1]
                    if len(values2) > 3 and values2[1] == hostname and values2[3] == add_vport_name :
                        return (["[ERROR] There is a Layer 1 interface with the same name as the virtual port."])

            if not match_found:
                if 'l2_segment' in argv_array:
                    return ([f"No matching entry found for hostname: {hostname} and portname: {portname}"])
                elif 'virtual_port' in argv_array:
                    return ([f"No matching entry found for hostname: {hostname}"])
                elif 'portchannel' in argv_array:
                    return ([f"No matching entry found for hostname: {hostname} and portname: {portname}"])
                elif 'vport_l2_direct_binding' in argv_array:
                    return ([f"No matching entry found for hostname: {hostname} and virtual portname: {portname}"])
                elif 'vport_l1if_direct_binding' in argv_array:
                    return ([f"No matching entry found for hostname: {hostname} and portname: {portname}"])
                exit()

            if 'l2_segment' in argv_array:
                add_l2seg_name = add_l2seg_name.replace(' ', '')
                for entry in l2_attribute_array:
                    if entry == row:
                        # Get the 7th element (index 6)
                        current_value = entry[1][6]
                        # Split by comma, remove extra spaces
                        segs = [seg.strip() for seg in current_value.split(',')] if current_value else []
                        # Only add if not already present
                        if add_l2seg_name not in segs:
                            if current_value.strip() == '':
                                entry[1][6] = add_l2seg_name
                                entry[1][6] = ','.join(sorted(s.strip() for s in entry[1][6].split(',')))
                            else:
                                entry[1][6] = current_value + ',' + add_l2seg_name
                                entry[1][6] = ','.join(sorted(s.strip() for s in entry[1][6].split(',')))
                        else:
                            #print(f"'{add_l2seg_name}' already exists in the l2 segments. No change made.")
                            return ([add_l2seg_name + ' already exists in the l2 segments. No change made'])
                        break

                    #for virtual port
                    if entry in row_array:
                        # Get the 7th element (index 6)
                        current_value = entry[1][6]
                        # Split by comma, remove extra spaces
                        segs = [seg.strip() for seg in current_value.split(',')] if current_value else []
                        # Only add if not already present
                        if add_l2seg_name not in segs:
                            if current_value.strip() == '':
                                entry[1][6] = add_l2seg_name
                                entry[1][6] = ','.join(sorted(s.strip() for s in entry[1][6].split(',')))
                            else:
                                entry[1][6] = current_value + ',' + add_l2seg_name
                                entry[1][6] = ','.join(sorted(s.strip() for s in entry[1][6].split(',')))
                        else:
                            #print(f"'{add_l2seg_name}' already exists in the l2 segments. No change made.")
                            return ([add_l2seg_name + ' already exists in the l2 segments. No change made'])

            elif 'virtual_port' in argv_array:
                if ns_def.get_if_value(add_vport_name) == -1:
                    return ([f"Invalid virtual port name: {add_vport_name}"])
                # Check if the virtual port name already exists for this hostname
                for entry in l2_attribute_array:
                    if entry[0] > 2 and entry[1][1] == hostname and entry[1][5] == add_vport_name:
                        return ([f"{add_vport_name} already exists as virtual port for hostname: {hostname}"])

                # get insert number
                tmp_if_value_array = []
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if (len(values) > 5 and values[5] != '') and (len(values) > 3 and values[3] == '') :
                            tmp_if_value_array.append(ns_def.get_if_value(values[5]))
                target_if_value = ns_def.get_if_value(add_vport_name)
                from bisect import bisect_left
                insert_idx = bisect_left(tmp_if_value_array, target_if_value)

                # Find the index of the first matching hostname row
                insert_index = -1
                for i, entry in enumerate(l2_attribute_array):
                    if entry[0] > 2 and entry[1][1] == hostname:
                        insert_index = i
                        break

                if insert_index != -1:
                    # Create new entry with the same area and hostname
                    new_entry = [0, [l2_attribute_array[insert_index][1][0], hostname, '', '', '', add_vport_name, '', '']]
                    # Insert the new entry at the found index (before the existing entry)
                    l2_attribute_array.insert(insert_index + insert_idx, new_entry)

                    # Renumber all entries
                    for i, entry in enumerate(l2_attribute_array):
                        entry[0] = i + 1

                return_text = '--- Virtual Port added --- ' + ' ' + hostname + ',' + add_vport_name

            elif 'portchannel' in argv_array:
                if ns_def.get_if_value(add_portchannel_name) == -1:
                    return ([f"Invalid portchannel name: {add_portchannel_name}"])

                for entry in l2_attribute_array:
                    if entry == row:
                        entry[1][5] = add_portchannel_name
                        return_text = '--- portchannel added --- ' + ' ' + hostname + ',' + portname + ',' + add_portchannel_name
                        break

            elif 'vport_l2_direct_binding' in argv_array:
                add_l2seg_name = add_l2seg_name.replace(' ', '')
                for entry in l2_attribute_array:
                    #for virtual port
                    if entry in row_array:
                        # Get the 8th element (index 7)
                        current_value = entry[1][7]
                        # Split by comma, remove extra spaces
                        segs = [seg.strip() for seg in current_value.split(',')] if current_value else []
                        # Only add if not already present
                        if add_l2seg_name not in segs:
                            if current_value.strip() == '':
                                entry[1][7] = add_l2seg_name
                                entry[1][7] = ','.join(sorted(s.strip() for s in entry[1][7].split(',')))
                            else:
                                entry[1][7] = current_value + ',' + add_l2seg_name
                                entry[1][7] = ','.join(sorted(s.strip() for s in entry[1][7].split(',')))
                        else:
                            #print(f"'{add_l2seg_name}' already exists in the l2 segments. No change made.")
                            return ([add_l2seg_name + ' already exists in the vport_l2_direct_binding. No change made'])

            elif 'vport_l1if_direct_binding' in argv_array:
                if ns_def.get_if_value(add_vport_name) == -1:
                    return ([f"Invalid virtual port name: {add_vport_name}"])

                if flag_vport_l1if_direct_binding_exists == False:
                    for entry in l2_attribute_array:
                        if entry == row:
                            entry[1][5] = add_vport_name
                            return_text = '--- vport_l1if_direct_binding added --- ' + ' ' + hostname + ',' + portname + ',' + add_vport_name
                            break
                elif flag_vport_l1if_direct_binding_exists == True:
                    #get insert number
                    tmp_if_value_array = []
                    for row in l2_attribute_array[2:]:
                        values = row[1]
                        if len(values) > 1 and values[1] == hostname:
                            if (len(values) > 3 and values[3] == portname) and (len(values) > 5 and values[5] != ''):
                                tmp_if_value_array.append(ns_def.get_if_value(values[5]))
                    target_if_value = ns_def.get_if_value(add_vport_name)
                    from bisect import bisect_left
                    insert_idx = bisect_left(tmp_if_value_array, target_if_value)

                    # Find the index of the first matching hostname row
                    insert_index = -1
                    for i, entry in enumerate(l2_attribute_array):
                        if entry[0] > 2 and entry[1][1] == hostname:
                            insert_index = i
                            break

                    if insert_index != -1:
                        # Create new entry with the same area and hostname
                        new_entry = [0, [l2_attribute_array[insert_index][1][0], hostname, '', portname, '', add_vport_name, '', '']]
                        # Insert the new entry at the found index (before the existing entry)
                        l2_attribute_array.insert(insert_index + 1 + insert_idx, new_entry)

                        # Renumber all entries
                        for i, entry in enumerate(l2_attribute_array):
                            entry[0] = i + 1
                    return_text = '--- vport_l1if_direct_binding added --- ' + ' ' + hostname + ',' + portname + ',' + add_vport_name

            #write to Master file
            excel_maseter_file = master_file_path
            excel_master_ws_name_l2 = 'Master_Data_L2'
            last_l2_table_tuple = {}
            last_l2_table_tuple = ns_def.convert_array_to_tuple(l2_attribute_array)

            # delete L2 Table sheet
            ns_def.remove_excel_sheet(excel_maseter_file, excel_master_ws_name_l2)
            # create L2 Table sheet
            ns_def.create_excel_sheet(excel_maseter_file, excel_master_ws_name_l2)
            # write tuple to excel master data
            ns_def.write_excel_meta(last_l2_table_tuple, excel_maseter_file, excel_master_ws_name_l2, '_template_', 0,0)

            if 'virtual_port' in argv_array or 'portchannel' in argv_array or 'vport_l1if_direct_binding' in argv_array:
                # sync l2 sheet of Master file to L3 sheet
                dummy_tk = tk.Toplevel()
                self.inFileTxt_L3_1_1 = tk.Entry(dummy_tk )
                self.inFileTxt_L3_1_1 .delete(0, tkinter.END)
                self.inFileTxt_L3_1_1 .insert(tk.END, master_file_path)

                self.outFileTxt_11_2 = tk.Entry(dummy_tk )
                self.outFileTxt_11_2.delete(0, tkinter.END)
                self.outFileTxt_11_2.insert(tk.END, master_file_path)

                self.inFileTxt_L2_1_1 = tk.Entry(dummy_tk )
                self.inFileTxt_L2_1_1.delete(0, tkinter.END)
                self.inFileTxt_L2_1_1.insert(tk.END, master_file_path)

                tmp_delete_excel_name = self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L3_TABLE]')

                import ns_l3_table_from_master
                ns_l3_table_from_master.ns_l3_table_from_master.__init__(self)

                if os.path.isfile(tmp_delete_excel_name) == True:
                    os.remove(tmp_delete_excel_name)

                return ([return_text])

            elif 'vport_l2_direct_binding' in argv_array:
                return_text = '--- vport_l2_direct_binding added --- ' + ' ' + hostname + ',' + portname + ',' + add_l2seg_name
                return ([return_text])

            else:
                return_text = '--- l2 Segment added --- ' + ' ' + hostname + ',' + portname + ',' + add_l2seg_name
                return ([return_text])

    def cli_delete(self, master_file_path, argv_array):
        import ns_def
        next_arg = get_next_arg(argv_array, 'delete')
        delete_command_list = [
            'delete ip_address', \
            'delete l2_segment', \
            'delete portchannel', \
            'delete virtual_port', \
            'delete vport_l1if_direct_binding', \
            'delete vport_l2_direct_binding', \
        ]

        if next_arg is None or '--' in next_arg or str('delete ' + next_arg) not in delete_command_list:
            print(next_arg)
            print('[ERROR] Supported commands are as follows')
            for tmp_delete_command_list in delete_command_list:
                print(tmp_delete_command_list)
            exit()

        if next_arg == 'ip_address':
            import ns_def
            l3_attribute_array = ns_def.convert_master_to_array('Master_Data_L3', master_file_path,'<<L3_TABLE>>')

            TARGET_LEN = 7
            for row in l3_attribute_array:
                if isinstance(row, list) and len(row) >= 2 and isinstance(row[1], list):
                    vals = row[1]
                    if len(vals) < TARGET_LEN:
                        vals += [''] * (TARGET_LEN - len(vals))

            idx = argv_array.index('ip_address')
            try:
                # The element right after 'ip address' is the hostname
                hostname = argv_array[idx + 1]
                # The next element is the portname
                portname = argv_array[idx + 2]
                delete_ipaddress_name = argv_array[idx + 3]
            except IndexError:
                # If there are not enough elements after 'ip address', print an error
                return ([f"[ERROR] hostname or portname or ip address is missing"])

            # check IP Addresses
            if ns_def.check_ip_format(delete_ipaddress_name) != 'IPv4':
                return ([f"Error: IP Address format is invalid: {delete_ipaddress_name}"])

            match_found = False
            target_row = None
            for row in l3_attribute_array[2:]:
                values = row[1]
                if len(values) > 1 and values[1] == hostname:
                    if (len(values) > 3 and values[2] == portname):
                        match_found = True
                        target_row = row
                        break
            if not match_found:
                return ([f"Error: No matching entry found for hostname: {hostname} and portname: {portname}"])

            delete_ipaddress_name = delete_ipaddress_name.replace(' ', '')

            for entry in l3_attribute_array:
                if entry == target_row:
                    current_value = entry[1][4]
                    segs = [seg.strip() for seg in current_value.split(',')] if current_value else []
                    if delete_ipaddress_name in segs:
                        segs.remove(delete_ipaddress_name)
                        entry[1][4] = ','.join(segs)
                    else:
                        return ([f"{delete_ipaddress_name} does not exist in the ip address. No change made"])
                    break

            #write to Master file
            excel_maseter_file = master_file_path
            excel_master_ws_name_l3 = 'Master_Data_L3'
            last_l3_table_tuple = {}
            last_l3_table_tuple = ns_def.convert_array_to_tuple(l3_attribute_array)

            # delete L2 Table sheet
            ns_def.remove_excel_sheet(excel_maseter_file, excel_master_ws_name_l3)
            # create L2 Table sheet
            ns_def.create_excel_sheet(excel_maseter_file, excel_master_ws_name_l3)
            # write tuple to excel master data
            ns_def.write_excel_meta(last_l3_table_tuple, excel_maseter_file, excel_master_ws_name_l3, '_template_', 0,0)

            return_text = '--- IP Address deleted--- ' + ' ' + hostname + ',' + portname + ',' + delete_ipaddress_name

            return ([return_text])

        if next_arg == 'l2_segment' or 'virtual_port' or 'portchannel' or 'vport_l2_direct_binding' or 'vport_l1if_direct_binding':
            import ns_def
            l2_attribute_array = ns_def.convert_master_to_array('Master_Data_L2', master_file_path, '<<L2_TABLE>>')

            TARGET_LEN = 9
            for row in l2_attribute_array:
                if isinstance(row, list) and len(row) >= 2 and isinstance(row[1], list):
                    vals = row[1]
                    if len(vals) < TARGET_LEN:
                        vals += [''] * (TARGET_LEN - len(vals))

            if 'l2_segment' in argv_array:
                idx = argv_array.index('l2_segment')
                try:
                    hostname = argv_array[idx + 1]
                    portname = argv_array[idx + 2]
                    del_l2seg_name = argv_array[idx + 3]
                except IndexError:
                    print("Error: hostname or portname is missing")
                    return
            elif 'virtual_port' in argv_array:
                idx = argv_array.index('virtual_port')
                try:
                    hostname = argv_array[idx + 1]
                    del_vport_name = argv_array[idx + 2]
                except IndexError:
                    print("Error: hostname or virtual_port name is missing")
                    return
            elif 'portchannel' in argv_array:
                idx = argv_array.index('portchannel')
                try:
                    hostname = argv_array[idx + 1]
                    portname = argv_array[idx + 2]
                    #del_portchannel_name = argv_array[idx + 3]
                except IndexError:
                    print("Error: hostname or portname is missing")
                    return
            elif 'vport_l2_direct_binding' in argv_array:
                idx = argv_array.index('vport_l2_direct_binding')
                try:
                    hostname = argv_array[idx + 1]
                    portname = argv_array[idx + 2]
                    del_l2seg_name = argv_array[idx + 3]
                except IndexError:
                    print("Error: hostname or virtual portname is missing")
                    return
            elif 'vport_l1if_direct_binding' in argv_array:
                idx = argv_array.index('vport_l1if_direct_binding')
                try:
                    hostname = argv_array[idx + 1]
                    del_vport_name = argv_array[idx + 2]
                except IndexError:
                    print("Error: hostname or portname is missing")
                    return
            else:
                print("Not found in the argument array")
                return

            match_found = False
            target_row = None
            row_array = []
            flag_l2_segment_vport = False
            if 'l2_segment' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if len(values) > 3 and values[3] == portname:
                            match_found = True
                            target_row = row
                            break
                        elif len(values) > 5 and values[5] == portname:
                            # for virtual port
                            match_found = True
                            flag_l2_segment_vport = True
                            row_array.append(row)
                if flag_l2_segment_vport == True:
                    row = []

            elif 'virtual_port' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname and values[5] == del_vport_name:
                        if values[3] != '':
                            return (['[ERROR] To delete a virtual port bound to a Layer 1 interface, use the delete vport_l1if_direct_binding command.'])

                        match_found = True
                        target_row = row
                        break
            elif 'portchannel' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if (len(values) > 3 and values[3] == portname) or (len(values) > 5 and values[5] == portname):
                            match_found = True
                            target_row = row
                            break
            elif 'vport_l2_direct_binding' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if len(values) > 5 and values[5] == portname:
                            # for virtual port
                            match_found = True
                            flag_l2_segment_vport = True
                            row_array.append(row)
                if flag_l2_segment_vport == True:
                    row = []
            elif 'vport_l1if_direct_binding' in argv_array:
                match_l1_if = ''
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if len(values) > 5 and values[5] == del_vport_name:
                            match_found = True
                            target_row = row
                            #count match l1 interface
                            match_l1_if = values[3]
                            break

                count_match_l1if = 0
                for tmp_l2_attribute_array in l2_attribute_array[2:]:
                    value_l2_attribute_array = tmp_l2_attribute_array[1]
                    if len(value_l2_attribute_array) > 1 and value_l2_attribute_array[1] == hostname:
                        if len(value_l2_attribute_array) > 3 and value_l2_attribute_array[3] == match_l1_if:
                            count_match_l1if += 1
                #print(count_match_l1if)

            if not match_found:
                if 'l2_segment' in argv_array:
                    print(f"No matching entry found for hostname: {hostname} and portname: {portname}")
                elif 'portchannel' in argv_array:
                    print(f"No matching entry found for hostname: {hostname} and portname: {portname}")
                elif 'vport_l2_direct_binding' in argv_array:
                    print(f"No matching entry found for hostname: {hostname} and virtual portname: {portname}")
                else:
                    print(f"No matching entry found for hostname: {hostname} and virtual port: {del_vport_name}")
                exit()

            if 'l2_segment' in argv_array:
                del_l2seg_name = del_l2seg_name.replace(' ', '')
                for entry in l2_attribute_array:
                    if entry == target_row:
                        current_value = entry[1][6]
                        segs = [seg.strip() for seg in current_value.split(',')] if current_value else []
                        if del_l2seg_name in segs:
                            segs.remove(del_l2seg_name)
                            entry[1][6] = ','.join(segs)
                        else:
                            return ([f"{del_l2seg_name} does not exist in the l2 segments. No change made"])
                        break

                    #for virtual port
                    if entry in row_array:
                        # Get the 7th element (index 6)
                        current_value = entry[1][6]
                        # Split by comma, remove extra spaces
                        segs = [seg.strip() for seg in current_value.split(',')] if current_value else []
                        # Only add if not already present
                        #print(del_l2seg_name,segs)
                        if del_l2seg_name in segs:
                            segs.remove(del_l2seg_name)
                            entry[1][6] = ','.join(segs)
                        else:
                            return ([del_l2seg_name + ' does not exist in the l2 segments. No change made'])


            elif 'virtual_port' in argv_array:
                # Remove the entire row for virtual port
                l2_attribute_array.remove(target_row)
                # Renumber all entries
                for i, entry in enumerate(l2_attribute_array):
                    entry[0] = i + 1
                    return_text = '--- Virtual Port deleted--- ' + ' ' + hostname + ',' + del_vport_name

            elif 'portchannel' in argv_array:
                for entry in l2_attribute_array:
                    if entry == target_row:
                        del_portchannel_name = entry[1][5]
                        entry[1][5] = ''
                        entry[1][6] = ''
                        return_text = '--- portchannel deleted --- ' + ' ' + hostname + ',' + portname + ',' + del_portchannel_name
                        break

            elif 'vport_l2_direct_binding' in argv_array:
                del_l2seg_name = del_l2seg_name.replace(' ', '')
                for entry in l2_attribute_array:
                    #for virtual port
                    if entry in row_array:
                        # Get the 7th element (index 6)
                        current_value = entry[1][7]
                        # Split by comma, remove extra spaces
                        segs = [seg.strip() for seg in current_value.split(',')] if current_value else []
                        # Only add if not already present
                        #print(del_l2seg_name,segs)
                        if del_l2seg_name in segs:
                            segs.remove(del_l2seg_name)
                            entry[1][7] = ','.join(segs)
                        else:
                            return ([del_l2seg_name + ' does not exist in the vport_l2_direct_binding. No change made'])

            elif 'vport_l1if_direct_binding' in argv_array:
                if count_match_l1if >= 2:
                    # Remove the entire row for virtual port
                    l2_attribute_array.remove(target_row)
                    # Renumber all entries
                    for i, entry in enumerate(l2_attribute_array):
                        entry[0] = i + 1
                        return_text = '--- vport_l1if_direct_binding deleted--- ' + ' ' + hostname + ',' + del_vport_name
                elif count_match_l1if == 1:
                    for entry in l2_attribute_array:
                        if entry == target_row:
                            del_vport_name = entry[1][5]
                            entry[1][5] = ''
                            entry[1][6] = ''
                            entry[1][7] = ''
                            return_text = '--- vport_l1if_direct_binding deleted --- ' + ' ' + hostname + ',' + del_vport_name
                            break

            #write to Master file
            excel_maseter_file = master_file_path
            excel_master_ws_name_l2 = 'Master_Data_L2'
            last_l2_table_tuple = ns_def.convert_array_to_tuple(l2_attribute_array)
            ns_def.remove_excel_sheet(excel_maseter_file, excel_master_ws_name_l2)
            ns_def.create_excel_sheet(excel_maseter_file, excel_master_ws_name_l2)
            ns_def.write_excel_meta(last_l2_table_tuple, excel_maseter_file, excel_master_ws_name_l2, '_template_', 0, 0)

            if 'virtual_port' in argv_array or 'portchannel' in argv_array or 'vport_l1if_direct_binding' in argv_array:
                #sync l2 sheet of Master file to L3 sheet
                dummy_tk = tk.Toplevel()
                self.inFileTxt_L3_1_1 = tk.Entry(dummy_tk)
                self.inFileTxt_L3_1_1.delete(0, tkinter.END)
                self.inFileTxt_L3_1_1.insert(tk.END, master_file_path)

                self.outFileTxt_11_2 = tk.Entry(dummy_tk)
                self.outFileTxt_11_2.delete(0, tkinter.END)
                self.outFileTxt_11_2.insert(tk.END, master_file_path)

                self.inFileTxt_L2_1_1 = tk.Entry(dummy_tk)
                self.inFileTxt_L2_1_1.delete(0, tkinter.END)
                self.inFileTxt_L2_1_1.insert(tk.END, master_file_path)

                tmp_delete_excel_name = self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L3_TABLE]')

                import ns_l3_table_from_master
                ns_l3_table_from_master.ns_l3_table_from_master.__init__(self)

                if os.path.isfile(tmp_delete_excel_name) == True:
                    os.remove(tmp_delete_excel_name)
            elif 'vport_l2_direct_binding' in argv_array:
                return_text = '--- vport_l2_direct_binding deleted --- ' + ' ' + hostname + ',' + portname + ',' + del_l2seg_name

            else:
                return_text = '--- l2 Segment deleted --- ' + ' ' + hostname + ',' + portname + ',' + del_l2seg_name

            return ([return_text])

    def cli_show(self, master_file_path, argv_array):
        next_arg = get_next_arg(argv_array, 'show')
        show_command_list = [\
            'show area', \
            'show area_device', \
            'show area_location', \
            'show attribute', \
            'show attribute_color', \
            'show device', \
            'show device_interface', \
            'show device_location', \
            'show l1_interface', \
            'show l1_link', \
            'show l2_broadcast_domain', \
            'show l2_interface', \
            'show l3_broadcast_domain', \
            'show l3_interface', \
            'show waypoint', \
            'show waypoint_interface', \
            ]

        if next_arg == None or '--' in next_arg or str('show ' + next_arg) not in show_command_list:
            print(next_arg)
            print('[ERROR] Supported commands are as follows')
            for tmp_show_command_list in show_command_list:
                print(tmp_show_command_list)
            exit()

        if next_arg == 'area' or 'area_location':
            import ns_def
            position_folder_array = ns_def.convert_master_to_array('Master_Data', master_file_path,'<<STYLE_FOLDER>>')

            folder_name_list = []
            area_wp_name_list = []
            area_name_list = []
            for item in position_folder_array:
                if item[0] not in [1, 2, 3]:
                    folder_name_list.append(item[1][0])

            for folder_name in folder_name_list:
                if "_wp_" in folder_name:
                    area_wp_name_list.append(folder_name)
                else:
                    area_name_list.append(folder_name)

            #print(folder_name_list)
            area_name_list = sorted(area_name_list,  reverse=False)
            #print(area_name_list)
            #print(area_wp_name_list)

            if next_arg == 'area':
                return (area_name_list)

        if next_arg == 'area_location':
            import ns_def
            position_folder_array = ns_def.convert_master_to_array('Master_Data', master_file_path,'<<POSITION_FOLDER>>')
            update_position_folder_array =[]

            for tmp_position_folder_array in position_folder_array:
                update_position_folder_array.append(tmp_position_folder_array[1])

            tmp_return = update_position_folder_array

            for sublist in tmp_return:
                for i in range(len(sublist)):
                    if sublist[i] not in area_name_list:
                        sublist[i] = ''

            tmp_return = [sublist for sublist in tmp_return if any(item != '' for item in sublist)]
            tmp_return = [[item for item in sublist if item != ''] for sublist in tmp_return]

            return(tmp_return)

        if next_arg == 'device_interface' or 'waypoint_interface' or 'l1_link' or 'l1_interface':
            import ns_def
            position_line_array = ns_def.convert_master_to_array('Master_Data', master_file_path, '<<POSITION_LINE>>')
            position_line_tuple = ns_def.convert_array_to_tuple(position_line_array)

            line_list = []
            device_interface_list = []
            waypoint_interface_list = []
            interface_detail_list = []
            device_waypoint_array = get_device_waypoint_array(master_file_path)
            for item in position_line_array:

                if item[0] not in [1,2]:
                    # get full if name
                    if_full_name_src = ns_def.get_full_name_from_tag_name(item[1][0], item[1][2],position_line_tuple)
                    if_full_name_tar = ns_def.get_full_name_from_tag_name(item[1][1], item[1][3],position_line_tuple)

                    line_list.append([[item[1][0], if_full_name_src], [item[1][1], if_full_name_tar]])
                    if str(item[1][0]) in device_waypoint_array[0]:
                        device_interface_list.append([item[1][0],if_full_name_src])
                    else:
                        waypoint_interface_list.append([item[1][0],if_full_name_src])

                    if str(item[1][1]) in device_waypoint_array[0]:
                        device_interface_list.append([item[1][1],if_full_name_tar])
                    else:
                        waypoint_interface_list.append([item[1][1],if_full_name_src])

                    interface_detail_list.append([item[1][0], item[1][2], if_full_name_src, item[1][13], item[1][14], item[1][15]])
                    interface_detail_list.append([item[1][1], item[1][3], if_full_name_tar, item[1][17], item[1][18], item[1][19]])

            if next_arg == 'l1_link':
                line_list = sorted(line_list, key=lambda x: (x[0][0],x[0][1]), reverse=False)
                return (line_list)

            elif next_arg == 'device_interface':
                device_interface_list = sorted(device_interface_list, key=lambda x: (x[0][0],x[0][1]), reverse=False)
                #print(interface_list)

                from collections import defaultdict
                grouped_data = defaultdict(list)
                for device, interface in device_interface_list:
                    grouped_data[device].append(interface, )

                result = [[device, interfaces] for device, interfaces in grouped_data.items()]
                return (result)

            elif next_arg == 'waypoint_interface':
                waypoint_interface_list = sorted(waypoint_interface_list, key=lambda x: (x[0][0],x[0][1]), reverse=False)
                #print(interface_list)

                from collections import defaultdict
                grouped_data = defaultdict(list)
                for device, interface in waypoint_interface_list:
                    grouped_data[device].append(interface, )

                result = [[device, interfaces] for device, interfaces in grouped_data.items()]
                return (result)

            elif next_arg == 'l1_interface':
                interface_detail_list = sorted(interface_detail_list, key=lambda x: (x[0], x[1]), reverse=False)
                return (interface_detail_list)

        if next_arg == 'device' or 'waypoint':
            device_waypoint_array = get_device_waypoint_array(master_file_path)
            if next_arg == 'device':
                device_list_array = sorted(device_waypoint_array[0], reverse=False)
                return (device_list_array)

            if next_arg == 'waypoint':
                wp_list_array = sorted(device_waypoint_array[1], reverse=False)
                return (wp_list_array)

        if next_arg == 'area_device':
            import ns_def
            l2_table_array = ns_def.convert_master_to_array('Master_Data_L2', master_file_path,   '<<L2_TABLE>>')
            new_l2_table_array = []
            for tmp_l2_table_array in l2_table_array:
                if tmp_l2_table_array[0] != 1 and tmp_l2_table_array[0] != 2:
                    tmp_l2_table_array[1].extend(['', '', '', '', '', '', '', ''])
                    del tmp_l2_table_array[1][8:]
                    new_l2_table_array.append(tmp_l2_table_array)

            device_list_array = []
            area_device_list_array = []
            wp_list_array = []
            for tmp_new_l2_table_array in new_l2_table_array:
                if tmp_new_l2_table_array[1][1] not in device_list_array and tmp_new_l2_table_array[1][1] not in wp_list_array:
                    if tmp_new_l2_table_array[1][0] == 'N/A':
                        wp_list_array.append(tmp_new_l2_table_array[1][1])
                    else:
                        device_list_array.append(tmp_new_l2_table_array[1][1])
                        area_device_list_array.append([tmp_new_l2_table_array[1][0], tmp_new_l2_table_array[1][1]])
            area_device_list_array = sorted(area_device_list_array, key=lambda x: (x[0],x[1]), reverse=False)

            from collections import defaultdict
            grouped_data = defaultdict(list)
            for area, device in area_device_list_array:
                grouped_data[area].append(device)
            result = [[area, devices] for area, devices in grouped_data.items()]
            return (result)

        if next_arg == 'device_location':
            import ns_def
            position_shape_array = ns_def.convert_master_to_array('Master_Data', master_file_path,'<<POSITION_SHAPE>>')
            update_position_shape_array = []

            current_folder_name = ''
            current_folder_shape = []
            flag_1st = True
            for tmp_position_folder_array in position_shape_array:
                if tmp_position_folder_array[1][0] != '<<POSITION_SHAPE>>':
                    if tmp_position_folder_array[1][0] != '<END>' and tmp_position_folder_array[0] != '':
                        if flag_1st == True and tmp_position_folder_array[1][0] != '<END>' and tmp_position_folder_array[1][0] != '':
                            current_folder_name = tmp_position_folder_array[1][0]

                        if flag_1st == False:
                            if '_wp_' not in str(current_folder_name):
                                update_position_shape_array.append([current_folder_name, current_folder_shape])

                            current_folder_name = tmp_position_folder_array[1][0]
                            current_folder_shape = []
                            flag_1st = True

                    if tmp_position_folder_array[1][0] != '<END>':
                        current_folder_shape.append(tmp_position_folder_array[1][1:][:-1])

                    if tmp_position_folder_array[1][0] == '<END>':
                        flag_1st = False
            else:
                if '_wp_' not in str(current_folder_name):
                    update_position_shape_array.append([current_folder_name, current_folder_shape])

            return (update_position_shape_array)

        if next_arg == 'l3_interface':
            import ns_def
            l3_interface_array = ns_def.convert_master_to_array('Master_Data_L3', master_file_path,'<<L3_TABLE>>')

            update_l3_interface_array = []
            for tmp_l3_interface_array in l3_interface_array:
                if tmp_l3_interface_array[0] != 1 and tmp_l3_interface_array[0] != 2:
                    update_l3_interface_array.append(tmp_l3_interface_array[1][1:])
            padded_data = [sublist + [''] * (4 - len(sublist)) for sublist in update_l3_interface_array]
            padded_data = sorted(padded_data, key=lambda x: (x[0]), reverse=False)
            return (padded_data)

        if next_arg == 'l2_interface':
            import ns_def
            l2_attribute_array = ns_def.convert_master_to_array('Master_Data_L2', master_file_path,'<<L2_TABLE>>')

            update_attribute_array = []
            for tmp_attribute_array in l2_attribute_array:
                if tmp_attribute_array[0] != 1 and tmp_attribute_array[0] != 2:
                    update_attribute_array.append(tmp_attribute_array[1][1:])

            padded_data = [sublist + [''] * (7 - len(sublist)) for sublist in update_attribute_array]
            update_padded_data = []

            for tmp_padded_data in padded_data:
                del tmp_padded_data[3]
                del tmp_padded_data[1]
                update_padded_data.append(tmp_padded_data)

            update_padded_data = sorted(padded_data, key=lambda x: (x[0]), reverse=False)
            return (update_padded_data)

        if next_arg == 'l2_broadcast_domain' or 'l3_broadcast_domain':
            import ns_def
            return_get_l2_broadcast_domains = []
            result_get_l2_broadcast_domains = ns_def.get_l2_broadcast_domains.run(self, master_file_path)  ## '[0] self.update_l2_table_array, [1] device_l2_boradcast_domain_array, [2] device_l2_directly_l3vport_array, [3] device_l2_other_array, [4] marged_l2_broadcast_group_array'
            if next_arg == 'l2_broadcast_domain':
                for tmp4_result_get_l2_broadcast_domains in result_get_l2_broadcast_domains[4]:
                    kari_l2_array = []
                    for tmp1_result_get_l2_broadcast_domains in result_get_l2_broadcast_domains[1]:
                        if tmp1_result_get_l2_broadcast_domains != []:
                            for tmp1_tmp1_result_get_l2_broadcast_domains in tmp1_result_get_l2_broadcast_domains:
                                if tmp1_tmp1_result_get_l2_broadcast_domains[0] in tmp4_result_get_l2_broadcast_domains:
                                    kari_l2_array.append([tmp1_tmp1_result_get_l2_broadcast_domains[1],tmp1_tmp1_result_get_l2_broadcast_domains[2]])
                                    break
                    return_get_l2_broadcast_domains.append([tmp4_result_get_l2_broadcast_domains,kari_l2_array])

                return (return_get_l2_broadcast_domains)

            if next_arg == 'l3_broadcast_domain':
                retrun_target_l2_broadcast_group_array = []
                for tmp_target_l2_broadcast_group_array in self.target_l2_broadcast_group_array:
                    if tmp_target_l2_broadcast_group_array[1] != []:
                        retrun_target_l2_broadcast_group_array.append(tmp_target_l2_broadcast_group_array)

                return (self.target_l2_broadcast_group_array)

        if next_arg == 'attribute' or 'attribute_color':
            import ns_def
            l2_attribute_array = ns_def.convert_master_to_array('Master_Data', master_file_path,'<<ATTRIBUTE>>')

            update_attribute_array = []
            for tmp_attribute_array in l2_attribute_array:
                if tmp_attribute_array[0] != 1:
                    update_attribute_array.append(tmp_attribute_array[1][1:])

            update_padded_data = []

            for tmp_padded_data in update_attribute_array:
                del tmp_padded_data[-1]
                update_padded_data.append(tmp_padded_data)

            update_padded_data = sorted(update_attribute_array, key=lambda x: (x[0]), reverse=False)

            if next_arg == 'attribute_color':
                kari_return_attribute = []
                for row in update_padded_data:
                    transformed_row = [item.replace('<EMPTY>', '') for item in row]
                    kari_return_attribute.append(transformed_row)
                return (kari_return_attribute)

            output_data = []
            for line in update_padded_data:
                if isinstance(line[0], str) and not line[0].startswith('['):
                    # Add header directly to output
                    output_data.append(line)
                else:
                    # Extract the first element from each formatted string and replace '<EMPTY>' with ''
                    extracted_elements = [eval(element)[0] if eval(element)[0] != '<EMPTY>' else '' for element in line]
                    output_data.append(extracted_elements)

            if next_arg == 'attribute':
                return (output_data)


def get_next_arg(argv_array, target):
    try:
        index = argv_array.index(target)
        return argv_array[index + 1]
    except (ValueError, IndexError):
        return None

def get_next_next_arg(argv_array, target):
    try:
        index = argv_array.index(target)
        return argv_array[index + 2]
    except (ValueError, IndexError):
        return None

def print_type(self, argv_array,source):
    if  '--one_msg' in argv_array:
        print(source)
    else:
        for item in source:
            print(item)
    return

def get_device_waypoint_array(master_file_path):
    import ns_def
    l2_table_array = ns_def.convert_master_to_array('Master_Data_L2', master_file_path, '<<L2_TABLE>>')

    new_l2_table_array = []
    for tmp_l2_table_array in l2_table_array:
        if tmp_l2_table_array[0] != 1 and tmp_l2_table_array[0] != 2:
            tmp_l2_table_array[1].extend(['', '', '', '', '', '', '', ''])
            del tmp_l2_table_array[1][8:]
            new_l2_table_array.append(tmp_l2_table_array)

    device_list_array = []
    area_device_list_array = []
    wp_list_array = []
    for tmp_new_l2_table_array in new_l2_table_array:
        if tmp_new_l2_table_array[1][1] not in device_list_array and tmp_new_l2_table_array[1][1] not in wp_list_array:
            if tmp_new_l2_table_array[1][0] == 'N/A':
                wp_list_array.append(tmp_new_l2_table_array[1][1])
            else:
                device_list_array.append(tmp_new_l2_table_array[1][1])
                area_device_list_array.append([tmp_new_l2_table_array[1][0], tmp_new_l2_table_array[1][1]])

    return ([device_list_array,wp_list_array])
