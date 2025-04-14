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

from pptx import *
import sys, os, re
import numpy as np
import math
import openpyxl
import ns_def ,ns_ddx_figure

class  ns_l1_table_from_master():
    def __init__(self):
        '''
        make device table excel file
        '''
        print('--- Device file create ---')
        #parameter
        ws_name = 'Master_Data'
        tmp_ws_name = '_tmp_'
        excel_maseter_file = str(self.inFileTxt_11_1.get())
        ppt_meta_file = excel_maseter_file

        ### click sync excel master ###
        if self.click_value == '12-3':
            excel_maseter_file = self.inFileTxt_12_2.get()
            ppt_meta_file = excel_maseter_file

        ### click create first set ###
        if self.click_value == '1-4':
            excel_maseter_file = self.outFileTxt_1_2.get()
            ppt_meta_file = excel_maseter_file

        #convert from master to array and convert to tuple
        self.position_folder_array = ns_def.convert_master_to_array(ws_name, excel_maseter_file,'<<POSITION_FOLDER>>')
        self.position_shape_array = ns_def.convert_master_to_array(ws_name, excel_maseter_file, '<<POSITION_SHAPE>>')
        self.position_line_array = ns_def.convert_master_to_array(ws_name, excel_maseter_file, '<<POSITION_LINE>>')
        self.position_style_shape_array = ns_def.convert_master_to_array(ws_name, excel_maseter_file, '<<STYLE_SHAPE>>')
        self.position_tag_array = ns_def.convert_master_to_array(ws_name, excel_maseter_file, '<<POSITION_TAG>>')
        self.root_folder_array = ns_def.convert_master_to_array(ws_name, excel_maseter_file, '<<ROOT_FOLDER>>')
        self.position_folder_tuple = ns_def.convert_array_to_tuple(self.position_folder_array)
        self.position_shape_tuple = ns_def.convert_array_to_tuple(self.position_shape_array)
        self.position_line_tuple = ns_def.convert_array_to_tuple(self.position_line_array)
        self.position_style_shape_tuple = ns_def.convert_array_to_tuple(self.position_style_shape_array)
        self.position_tag_tuple = ns_def.convert_array_to_tuple(self.position_tag_array)
        self.root_folder_tuple = ns_def.convert_array_to_tuple(self.root_folder_array)

        #print('---- self.position_folder_tuple ----')
        #print(self.position_folder_tuple)
        #print('---- self.position_shape_tuple ----')
        #print(self.position_shape_tuple)
        #print('---- self.position_line_tuple ----')
        #print(self.position_line_tuple)

        # GET Folder and wp name List
        folder_wp_name_array = ns_def.get_folder_wp_array_from_master(ws_name, excel_maseter_file)
        #print('---- folder_wp_name_array ----')
        #print(folder_wp_name_array)

        '''make device table excel file'''
        master_device_table_tuple = {}
        device_table_array = []
        egt_maker_width_array = ['25', '25', '25', '8', '8', '8', '20', '20', '20', '20', '20']  # for Network Sketcher Ver 2.0
        device_table_array.append([1, ['<RANGE>', '1', '1', '1', '1', '1', '1', '1', '1', '1', '1', '1', '<END>']])
        device_table_array.append([2, ['<HEADER>', 'Area', 'Device Name', 'Port Name', 'Abbreviation(Diagram)', 'Speed', 'Duplex', 'Port Type', '[src] Device Name', '[src] Port Name','[dst] Device Name', '[dst] Port Name', '<END>']])

        start_row = 3
        start_column = 2
        egt_prefix = '>>' # for egt maker value.

        #sort folder name
        #print('---- new_folder_wp_name_array ----')
        new_folder_wp_name_array = sorted(folder_wp_name_array[0], key=str.lower)
        new_folder_wp_name_array.extend(sorted(folder_wp_name_array[1], key=str.lower))
        #print(new_folder_wp_name_array)

        #Run each folder name
        for tmp_new_folder_wp_name_array in new_folder_wp_name_array:
            tmp_current_array = []
            Flag_first_folder = True
            tmp_sort_array = []

            # extend folder name
            tmp_current_array.extend([''])
            if '_wp_' not in tmp_new_folder_wp_name_array:
                tmp_current_array.extend([tmp_new_folder_wp_name_array])
            else:
                tmp_current_array.extend(['N/A'])
            #get shape name in the folder and sort
            shape_folder_tuple = ns_def.get_shape_folder_tuple(self.position_shape_tuple)
            #print('---- shape_folder_tuple ----')
            #print(shape_folder_tuple)
            for tmp_shape_folder_tuple in shape_folder_tuple:
                #print(shape_folder_tuple[tmp_shape_folder_tuple] ,tmp_new_folder_wp_name_array )
                if shape_folder_tuple[tmp_shape_folder_tuple] == tmp_new_folder_wp_name_array:
                    tmp_sort_array.append(tmp_shape_folder_tuple)
            new_tmp_sort_array = sorted(tmp_sort_array, key=str.lower)
            #print('---- new_tmp_sort_array ----')
            #print(new_tmp_sort_array)

            '''get tag of the shape'''
            flag_shape_first = True
            for tmp_new_tmp_sort_array in new_tmp_sort_array:
                for tmp_position_line_tuple in self.position_line_tuple:
                    if tmp_position_line_tuple[0] != 1 and tmp_position_line_tuple[0] != 2:
                        if tmp_position_line_tuple[1] == 1 and self.position_line_tuple[tmp_position_line_tuple] == tmp_new_tmp_sort_array:
                            if flag_shape_first == True and Flag_first_folder == True:
                                tmp_current_array.extend([self.position_line_tuple[tmp_position_line_tuple[0], 1]])
                                flag_shape_first = False
                                Flag_first_folder = False
                            elif flag_shape_first == True and Flag_first_folder == False:
                                tmp_current_array.extend(['', ''])
                                tmp_current_array.extend([self.position_line_tuple[tmp_position_line_tuple[0], 1]])
                                flag_shape_first = False
                            else:
                                tmp_current_array.extend(['','',''])

                            # change part of port number between port name and abbr
                            from_split_port_name = str(self.position_line_tuple[tmp_position_line_tuple[0], 3]).split(' ')
                            if ' ' in str(self.position_line_tuple[tmp_position_line_tuple[0], 3]):
                                from_port_full = self.position_line_tuple[tmp_position_line_tuple[0], 13] + ' ' + from_split_port_name[-1]
                                from_port_abbr = from_split_port_name[0]
                            else:
                                from_port_full = self.position_line_tuple[tmp_position_line_tuple[0],13]
                                from_port_abbr = self.position_line_tuple[tmp_position_line_tuple[0], 3]

                            to_split_port_name = str(self.position_line_tuple[tmp_position_line_tuple[0], 4]).split(' ')
                            if ' ' in str(self.position_line_tuple[tmp_position_line_tuple[0], 4]):
                                to_port_full = self.position_line_tuple[tmp_position_line_tuple[0], 17] + ' ' + to_split_port_name[-1]
                                to_port_abbr = to_split_port_name[0]
                            else:
                                to_port_full = self.position_line_tuple[tmp_position_line_tuple[0], 17]
                                to_port_abbr = self.position_line_tuple[tmp_position_line_tuple[0], 4]
                            tmp_current_array.extend([egt_prefix + from_port_full ,egt_prefix + from_port_abbr,egt_prefix + str(self.position_line_tuple[tmp_position_line_tuple[0],14]),\
                                                      egt_prefix + str(self.position_line_tuple[tmp_position_line_tuple[0],15]), egt_prefix + str(self.position_line_tuple[tmp_position_line_tuple[0],16]), \
                                                      self.position_line_tuple[tmp_position_line_tuple[0], 1], self.position_line_tuple[tmp_position_line_tuple[0], 3], \
                                                      self.position_line_tuple[tmp_position_line_tuple[0],2],self.position_line_tuple[tmp_position_line_tuple[0],4],'<END>'])
                            device_table_array.append([start_row, tmp_current_array])
                            #print(tmp_current_array)
                            tmp_current_array = []
                            start_row += 1

                        if tmp_position_line_tuple[1] == 2 and self.position_line_tuple[tmp_position_line_tuple] == tmp_new_tmp_sort_array:
                            if flag_shape_first == True and Flag_first_folder == True:
                                tmp_current_array.extend([self.position_line_tuple[tmp_position_line_tuple[0], 2]])
                                flag_shape_first = False
                                Flag_first_folder = False
                            elif flag_shape_first == True and Flag_first_folder == False:
                                tmp_current_array.extend(['', ''])
                                tmp_current_array.extend([self.position_line_tuple[tmp_position_line_tuple[0], 2]])
                                flag_shape_first = False
                            else:
                                tmp_current_array.extend(['', '', ''])

                            # split part of port number between port name and abbr
                            from_split_port_name = str(self.position_line_tuple[tmp_position_line_tuple[0], 4]).split(' ')
                            if ' ' in str(self.position_line_tuple[tmp_position_line_tuple[0], 4]):
                                from_port_full = self.position_line_tuple[tmp_position_line_tuple[0], 17] + ' ' + from_split_port_name[-1]
                                from_port_abbr = from_split_port_name[0]
                            else:
                                from_port_full = self.position_line_tuple[tmp_position_line_tuple[0], 17]
                                from_port_abbr = self.position_line_tuple[tmp_position_line_tuple[0], 4]


                            to_split_port_name = str(self.position_line_tuple[tmp_position_line_tuple[0], 3]).split(' ')
                            if ' ' in str(self.position_line_tuple[tmp_position_line_tuple[0], 3]):
                                to_port_full = self.position_line_tuple[tmp_position_line_tuple[0], 13] + ' ' + to_split_port_name[-1]
                                to_port_abbr = to_split_port_name[0]
                            else:
                                to_port_full = self.position_line_tuple[tmp_position_line_tuple[0], 13]
                                to_port_abbr = self.position_line_tuple[tmp_position_line_tuple[0], 3]
                            tmp_current_array.extend([egt_prefix + from_port_full,egt_prefix + from_port_abbr,egt_prefix + str(self.position_line_tuple[tmp_position_line_tuple[0],18]),\
                                                      egt_prefix + str(self.position_line_tuple[tmp_position_line_tuple[0],19]), egt_prefix + str(self.position_line_tuple[tmp_position_line_tuple[0],20]), \
                                                      self.position_line_tuple[tmp_position_line_tuple[0], 2], self.position_line_tuple[tmp_position_line_tuple[0], 4],\
                                                      self.position_line_tuple[tmp_position_line_tuple[0],1],self.position_line_tuple[tmp_position_line_tuple[0],3],'<END>'])
                            device_table_array.append([start_row, tmp_current_array])
                            #print(tmp_current_array)
                            tmp_current_array = []
                            start_row += 1
                flag_shape_first = True

        device_table_array.append([start_row, ['<END>']])

        #print('---- device_table_array ----')
        #print(device_table_array)

        ''' sort Interafce number'''
        current_if_array = []
        tuple_num = 1
        overwrite_if_array = []
        for tmp_device_table_array in device_table_array:
            #print(tmp_device_table_array[1])
            if str(tmp_device_table_array[1][0]) == '':
                if str(tmp_device_table_array[1][2]) != '':
                    if len(current_if_array) != 0:
                        # sort interface number
                        current_if_array = sorted(current_if_array, reverse=False, key=lambda x:(x[11], x[10]))  # sort for tmp_end_top

                        for tmp_current_if_array in current_if_array:
                            overwrite_if_array.append([tuple_num,tmp_current_if_array[:9]])
                            #print(tuple_num, tmp_current_if_array[:7])
                            tuple_num += 1

                    current_if_array = []
                    current_shape_name = tmp_device_table_array[1][2]
                    if_value = ns_def.get_if_value(tmp_device_table_array[1][3])
                    tmp_device_table_array[1].extend([if_value])
                    tmp_device_table_array[1].extend([str(ns_def.split_portname(tmp_device_table_array[1][3])[0])])
                    current_if_array.append(tmp_device_table_array[1][3:])
                else:
                    if_value = ns_def.get_if_value(tmp_device_table_array[1][3])
                    tmp_device_table_array[1].extend([if_value])
                    tmp_device_table_array[1].extend([str(ns_def.split_portname(tmp_device_table_array[1][3])[0])])
                    current_if_array.append(tmp_device_table_array[1][3:])
                    last_append = current_if_array

        ### kyuusai last shape interfaces
        if len(current_if_array) >= 2:  ## Modify ver 1.1
            for tmp_current_if_array in current_if_array:
                last_append = sorted(last_append, reverse=False, key=lambda x:(x[11], x[10]))  # sort for tmp_end_top
                overwrite_if_array.append([tuple_num, tmp_current_if_array[:9]])
                tuple_num += 1

        #print('---- overwrite_if_array ----')
        #print(overwrite_if_array)
        overwrite_if_tuple = ns_def.convert_array_to_tuple(overwrite_if_array)
        #print('---- overwrite_if_tuple ----')
        #print(overwrite_if_tuple)
        ### Convert to tuple
        master_device_table_tuple = ns_def.convert_array_to_tuple(device_table_array)
        #print('---- master_device_table_tuple ----')
        #print(master_device_table_tuple)

        # Create _tmp_ sheet
        ns_def.create_excel_sheet(ppt_meta_file, tmp_ws_name)

        # Write normal tuple to excel
        master_excel_meta = master_device_table_tuple
        excel_file_path = ppt_meta_file
        worksheet_name = tmp_ws_name
        section_write_to = '<<N/A>>'
        offset_row = 0
        offset_column = 0
        ns_def.write_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

        # Write overwrite tuple to excel
        master_excel_meta = overwrite_if_tuple
        excel_file_path = ppt_meta_file
        worksheet_name = tmp_ws_name
        section_write_to = '<<N/A>>'
        offset_row = 2
        offset_column = 3
        ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

        '''Create Device table file'''
        import ns_egt_maker #module import

        #create Device table file
        tmp_tmp_ws_name = '_tmp_tmp_'
        ns_def.create_excel_sheet(ppt_meta_file, tmp_tmp_ws_name)
        input_tree_tuple = {}
        input_tree_tuple[1,1] = 'Dummy' # first excel sheet has bug. That's why the first sheet is dummy sheet. this sheet will be deleted soon.
        input_tree_tuple[2, 1] = 'L1 Table'
        master_excel_meta = input_tree_tuple
        excel_file_path = ppt_meta_file
        worksheet_name = tmp_tmp_ws_name
        section_write_to = '<<N/A>>'
        offset_row = 0
        offset_column = 0
        ns_def.write_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)
        #Add Device table sheet
        input_excel_name = ppt_meta_file
        output_excel_name = self.outFileTxt_11_2.get().replace('[MASTER]', '')
        ### click sync file ###
        if self.click_value == '12-3':
            output_excel_name = self.inFileTxt_12_1.get()

        ### click create first set ###
        if self.click_value == '1-4':
            output_excel_name = str(self.outFileTxt_1_4_1.get()).replace('[L1_DIAGRAM]AllAreasTag_','')

        NEW_OR_ADD = 'NEW'
        ns_egt_maker.create_excel_gui_tree(input_excel_name,output_excel_name,NEW_OR_ADD, egt_maker_width_array)
        ns_def.remove_excel_sheet(output_excel_name, 'Dummy')

        #Add Device table from meta
        #print(output_excel_name)
        self.input_tree_excel = openpyxl.load_workbook(output_excel_name)
        worksheet_name = 'L1 Table'
        start_row = 1
        start_column = 0
        custom_table_name = ppt_meta_file
        self.input_tree_excel = ns_egt_maker.insert_custom_excel_table(self.input_tree_excel, worksheet_name, start_row, start_column, custom_table_name)
        self.input_tree_excel.save(output_excel_name)

        # Remove _tmp_ sheet from excel master
        ns_def.remove_excel_sheet(ppt_meta_file, tmp_ws_name)
        ns_def.remove_excel_sheet(ppt_meta_file, tmp_tmp_ws_name)
