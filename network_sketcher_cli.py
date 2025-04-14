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
        else:
            print('[ERROR] Supported commands are as follows')
            print('show <sub-command>')

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
