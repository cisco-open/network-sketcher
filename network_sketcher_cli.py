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
        elif 'delete' in argv_array:
            print_type(self, argv_array, ns_cli_run.cli_delete(self, master_file_path, argv_array))
            exit()
        else:
            print('[ERROR] Supported commands are as follows')
            print('add')
            print('delete')
            print('show')

    def cli_add(self, master_file_path, argv_array): # add at ver 2.5.3
        next_arg = get_next_arg(argv_array, 'add')
        add_command_list = [ \
            'add ip_address', \
            'add l2_segment', \
            'add virtual_port', \
            ]

        if next_arg == None or '--' in next_arg or str('add ' + next_arg) not in add_command_list:
            print(next_arg)
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
                # If there are not enough elements after 'l2_segment', print an error
                print("Error: hostname or portname or ipaddress is missing")

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

        if next_arg == 'l2_segment' or 'virtual_port':
            import ns_def
            l2_attribute_array = ns_def.convert_master_to_array('Master_Data_L2', master_file_path,'<<L2_TABLE>>')

            TARGET_LEN = 9
            for row in l2_attribute_array:
                if isinstance(row, list) and len(row) >= 2 and isinstance(row[1], list):
                    vals = row[1]
                    if len(vals) < TARGET_LEN:
                        vals += [''] * (TARGET_LEN - len(vals))

            #print(argv_array)
            # Check if 'l2_segment' exists in argv_array
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
                    print("Error: hostname or portname is missing")

            elif 'virtual_port' in argv_array:
                idx = argv_array.index('virtual_port')
                try:
                    hostname = argv_array[idx + 1]
                    add_vport_name = argv_array[idx + 2]
                except IndexError:
                    print("Error: hostname is missing")
            else:
                print("Not found in arguments")


            match_found = False
            if 'l2_segment' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if (len(values) > 3 and values[3] == portname) or (len(values) > 5 and values[5] == portname):
                            match_found = True
                            break
            elif 'virtual_port' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        match_found = True
                        break

            if not match_found:
                if 'l2_segment' in argv_array:
                    return ([f"No matching entry found for hostname: {hostname} and portname: {portname}"])
                elif 'virtual_port' in argv_array:
                    return ([f"No matching entry found for hostname: {hostname}"])

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
                            else:
                                entry[1][6] = current_value + ',' + add_l2seg_name
                        else:
                            #print(f"'{add_l2seg_name}' already exists in the l2 segments. No change made.")
                            return ([add_l2seg_name + ' already exists in the l2 segments. No change made'])
                        break

            elif 'virtual_port' in argv_array:
                if ns_def.get_if_value(add_vport_name) == -1:
                    return ([f"Invalid virtual port name: {add_vport_name}"])
                # Check if the virtual port name already exists for this hostname
                for entry in l2_attribute_array:
                    if entry[0] > 2 and entry[1][1] == hostname and entry[1][5] == add_vport_name:
                        return ([f"{add_vport_name} already exists as virtual port for hostname: {hostname}"])

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
                    l2_attribute_array.insert(insert_index, new_entry)

                    # Renumber all entries
                    for i, entry in enumerate(l2_attribute_array):
                        entry[0] = i + 1

                return_text = '--- Virtual Port added --- ' + ' ' + hostname + ',' + add_vport_name

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

            if 'virtual_port' in argv_array:
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
            else:
                return_text = '--- l2 Segment added --- ' + ' ' + hostname + ',' + portname + ',' + add_l2seg_name
                return ([return_text])

    def cli_delete(self, master_file_path, argv_array):
        next_arg = get_next_arg(argv_array, 'delete')
        delete_command_list = [
            'delete ip_address', \
            'delete l2_segment', \
            'delete virtual_port', \
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
                # The element right after 'l2_segment' is the hostname
                hostname = argv_array[idx + 1]
                # The next element is the portname
                portname = argv_array[idx + 2]
                delete_ipaddress_name = argv_array[idx + 3]
            except IndexError:
                # If there are not enough elements after 'l2_segment', print an error
                print("Error: hostname or portname or ipaddress is missing")

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

        if next_arg == 'l2_segment' or 'virtual_port':
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
            else:
                if next_arg == 'l2_segment':
                    print("'l2_segment' not found in the argument array")
                else:
                    print("'virtual_port' not found in the argument array")
                return

            match_found = False
            target_row = None
            if 'l2_segment' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname:
                        if (len(values) > 3 and values[3] == portname) or (len(values) > 5 and values[5] == portname):
                            match_found = True
                            target_row = row
                            break
            elif 'virtual_port' in argv_array:
                for row in l2_attribute_array[2:]:
                    values = row[1]
                    if len(values) > 1 and values[1] == hostname and values[5] == del_vport_name:
                        match_found = True
                        target_row = row
                        break

            if not match_found:
                if 'l2_segment' in argv_array:
                    print(f"No matching entry found for hostname: {hostname} and portname: {portname}")
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
            elif 'virtual_port' in argv_array:
                # Remove the entire row for virtual port
                l2_attribute_array.remove(target_row)
                # Renumber all entries
                for i, entry in enumerate(l2_attribute_array):
                    entry[0] = i + 1

            #write to Master file
            excel_maseter_file = master_file_path
            excel_master_ws_name_l2 = 'Master_Data_L2'
            last_l2_table_tuple = ns_def.convert_array_to_tuple(l2_attribute_array)
            ns_def.remove_excel_sheet(excel_maseter_file, excel_master_ws_name_l2)
            ns_def.create_excel_sheet(excel_maseter_file, excel_master_ws_name_l2)
            ns_def.write_excel_meta(last_l2_table_tuple, excel_maseter_file, excel_master_ws_name_l2, '_template_', 0, 0)

            if 'virtual_port' in argv_array:
                return_text = '--- Virtual Port deleted--- ' + ' ' + hostname + ',' + del_vport_name

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
