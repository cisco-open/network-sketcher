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

class  ns_l3_table_from_master():
    def __init__(self):
        '''
        make L3 Table excel file
        '''
        #parameter
        ws_name = 'Master_Data'
        ws_name_l2 = 'Master_Data_L2'
        ws_name_l3 = 'Master_Data_L3'
        tmp_ws_name = '_tmp_'
        excel_maseter_file = self.inFileTxt_L3_1_1.get()
        #write_excel_file = self.outFileTxt_11_2.get().replace('[MASTER]', '')

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

        # GET L2 TABLE
        self.l2_table_array = ns_def.convert_master_to_array(ws_name_l2, excel_maseter_file, '<<L2_TABLE>>')
        self.l2_table_tuple = ns_def.convert_array_to_tuple(self.l2_table_array)

        #print('--- self.l2_table_array ---')
        #print(self.l2_table_array)
        #print('--- self.l2_table_tuple ---')
        #print(self.l2_table_tuple)

        ## sort interface and identify L3
        used_l3_if_name = []
        pre_L3_table_array = []
        now_shape_name = ''
        for tmp_l2_table_array in self.l2_table_array:
            if tmp_l2_table_array[0] != 1 and tmp_l2_table_array[0] != 2:
                tmp_l2_table_array[1].extend(['', '', '', '', '', ''])
                del tmp_l2_table_array[1][8:]

                if tmp_l2_table_array[1][1] != now_shape_name:
                    used_l3_if_name = []
                    now_shape_name = tmp_l2_table_array[1][1]

                # identify L3 IF and remove duplex virtual port name and sort
                if tmp_l2_table_array[1][3] != '' and tmp_l2_table_array[1][5] == '' and tmp_l2_table_array[1][6] == '' and tmp_l2_table_array[1][7] == '' and tmp_l2_table_array[1][3] not in used_l3_if_name:
                    if_name_array = ns_def.split_portname(tmp_l2_table_array[1][3])
                    if_number = ns_def.get_if_value(tmp_l2_table_array[1][3])
                    pre_L3_table_array.append(['',tmp_l2_table_array[1][0], tmp_l2_table_array[1][1], tmp_l2_table_array[1][3], '<EMPTY>', '<EMPTY>','<END>',if_name_array[0],if_number])
                    used_l3_if_name.append(tmp_l2_table_array[1][3])

                elif tmp_l2_table_array[1][3] == '' and tmp_l2_table_array[1][5] != '' and tmp_l2_table_array[1][6] != '' and tmp_l2_table_array[1][7] == '' and tmp_l2_table_array[1][5] not in used_l3_if_name:
                    if_name_array = ns_def.split_portname(tmp_l2_table_array[1][5])
                    if_number = ns_def.get_if_value(tmp_l2_table_array[1][5])
                    pre_L3_table_array.append(['',tmp_l2_table_array[1][0], tmp_l2_table_array[1][1], tmp_l2_table_array[1][5], '<EMPTY>', '<EMPTY>','<END>',if_name_array[0],if_number])
                    used_l3_if_name.append(tmp_l2_table_array[1][5])

                elif tmp_l2_table_array[1][3] != '' and tmp_l2_table_array[1][5] != '' and tmp_l2_table_array[1][6] == '' and tmp_l2_table_array[1][7] != '' and tmp_l2_table_array[1][5] not in used_l3_if_name:
                    if_name_array = ns_def.split_portname(tmp_l2_table_array[1][5])
                    if_number = ns_def.get_if_value(tmp_l2_table_array[1][5])
                    pre_L3_table_array.append(['',tmp_l2_table_array[1][0], tmp_l2_table_array[1][1], tmp_l2_table_array[1][5], '<EMPTY>', '<EMPTY>','<END>',if_name_array[0],if_number])
                    used_l3_if_name.append(tmp_l2_table_array[1][5])

                elif tmp_l2_table_array[1][3] != '' and tmp_l2_table_array[1][5] != '' and tmp_l2_table_array[1][6] == '' and tmp_l2_table_array[1][7] == '' and tmp_l2_table_array[1][5] not in used_l3_if_name:
                    if_name_array = ns_def.split_portname(tmp_l2_table_array[1][5])
                    if_number = ns_def.get_if_value(tmp_l2_table_array[1][5])
                    pre_L3_table_array.append(['',tmp_l2_table_array[1][0], tmp_l2_table_array[1][1], tmp_l2_table_array[1][5], '<EMPTY>', '<EMPTY>','<END>',if_name_array[0],if_number])
                    used_l3_if_name.append(tmp_l2_table_array[1][5])

                elif tmp_l2_table_array[1][3] == '' and tmp_l2_table_array[1][5] != '' and tmp_l2_table_array[1][6] == '' and tmp_l2_table_array[1][7] == '' and tmp_l2_table_array[1][5] not in used_l3_if_name:
                    if_name_array = ns_def.split_portname(tmp_l2_table_array[1][5])
                    if_number = ns_def.get_if_value(tmp_l2_table_array[1][5])
                    pre_L3_table_array.append(['',tmp_l2_table_array[1][0], tmp_l2_table_array[1][1], tmp_l2_table_array[1][5], '<EMPTY>', '<EMPTY>','<END>',if_name_array[0],if_number])
                    used_l3_if_name.append(tmp_l2_table_array[1][5])

        pre_L3_table_array = sorted(pre_L3_table_array, reverse=False, key=lambda x: (x[1], x[2], x[7], x[8], x[7], x[4], x[5]))  # sort

        #print('### pre_L3_table_array ###')
        #print(pre_L3_table_array)

        '''
        Create L3 Table excel file
        '''

        master_L3_table_tuple = {}
        L3_table_array = []

        egt_maker_width_array = ['25', '25', '30', '20', '60', '50', '50']
        L3_table_array.append([1, ['<RANGE>', '1', '1', '1', '1', '1', '1', '1', '<END>']])
        L3_table_array.append([2, ['<HEADER>', 'Area', 'Device Name', 'L3 Port Name','L3 Instance Name', 'IP Address / Subnet mask (Comma Separated)', '[VPN] Target Device Name (Comma Separated)', '[VPN] Target L3 Port Name (Comma Separated)', '<END>']])



        start_row = 3
        start_column = 2

        #make L3 table array
        now_shape_name = ''
        for tmp_pre_L3_table_array in pre_L3_table_array:
            L3_table_array.append([start_row, tmp_pre_L3_table_array])
            start_row += 1
        L3_table_array.append([start_row, ['<END>']])

        #print('---- L3_table_array ----')
        #print(L3_table_array)

        update_L3_table_array = []

        current_folder_name = ''
        current_shape_name  = ''
        for tmp_L3_table_array in L3_table_array:
            if tmp_L3_table_array[0] != 1 and tmp_L3_table_array[0] != 2 and tmp_L3_table_array[1][0] != '<END>':
                if tmp_L3_table_array[1][1] == current_folder_name:
                    tmp_L3_table_array[1][1] = ''
                else:
                    current_folder_name = tmp_L3_table_array[1][1]

                if tmp_L3_table_array[1][2] == current_shape_name:
                    tmp_L3_table_array[1][2] = ''
                else:
                    current_shape_name = tmp_L3_table_array[1][2]

            update_L3_table_array.append(tmp_L3_table_array)


        L3_table_tuple = {}
        L3_table_tuple = ns_def.convert_array_to_tuple(update_L3_table_array)

        # Create _tmp_ sheet
        ns_def.create_excel_sheet(excel_maseter_file, tmp_ws_name)

        # Write normal tuple to excel
        master_excel_meta = L3_table_tuple
        excel_file_path = excel_maseter_file
        worksheet_name = tmp_ws_name
        section_write_to = '<<N/A>>'
        offset_row = 0
        offset_column = 0
        ns_def.write_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

        '''Create L3 Table file'''
        import ns_egt_maker #module import

        #create L3 Table file
        tmp_tmp_ws_name = '_tmp_tmp_'
        ns_def.create_excel_sheet(excel_maseter_file, tmp_tmp_ws_name)
        input_tree_tuple = {}
        input_tree_tuple[1,1] = 'Dummy' # first excel sheet has bug. That's why the first sheet is dummy sheet. this sheet will be deleted soon.
        input_tree_tuple[2, 1] = 'L3 Table'
        master_excel_meta = input_tree_tuple
        excel_file_path = excel_maseter_file
        worksheet_name = tmp_tmp_ws_name
        section_write_to = '<<N/A>>'
        offset_row = 0
        offset_column = 0
        ns_def.write_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)
        #Add L3 Table sheet
        input_excel_name = excel_maseter_file
        output_excel_name =  self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L3_TABLE]')

        NEW_OR_ADD = 'NEW'
        ns_egt_maker.create_excel_gui_tree(input_excel_name,output_excel_name,NEW_OR_ADD, egt_maker_width_array)
        ns_def.remove_excel_sheet(output_excel_name, 'Dummy')

        #Add L3 Table from meta
        #print(output_excel_name)
        self.input_tree_excel = openpyxl.load_workbook(output_excel_name)
        worksheet_name = 'L3 Table'
        start_row = 1
        start_column = 0
        custom_table_name = excel_maseter_file
        self.input_tree_excel = ns_egt_maker.insert_custom_excel_table(self.input_tree_excel, worksheet_name, start_row, start_column, custom_table_name)
        self.input_tree_excel.save(output_excel_name)

        # Remove _tmp_ sheet from excel master
        ns_def.remove_excel_sheet(excel_maseter_file, tmp_ws_name)
        ns_def.remove_excel_sheet(excel_maseter_file, tmp_tmp_ws_name)

        '''
        ADD L3 MASTER DATA Sheet to Excel Master file
        '''
        excel_L3_table_array = []
        excel_L3_table_array.append([1, ['<<L3_TABLE>>']])

        tmp_offset_row = 0
        now_folder_name = ''
        now_shape_name = ''
        for tmp_L3_table_array in L3_table_array:
            if tmp_L3_table_array[0] != 1 and tmp_L3_table_array[1][0] != '<END>':
                del tmp_L3_table_array[1][0]
                del tmp_L3_table_array[1][-1]
                if tmp_L3_table_array[1][-4] == '<EMPTY>' or tmp_L3_table_array[1][-5] == '<EMPTY>':
                    del tmp_L3_table_array[1][-2:]
                    if tmp_L3_table_array[1][-1] == '<EMPTY>':
                        tmp_L3_table_array[1][-1] = ''
                    if tmp_L3_table_array[1][-2] == '<EMPTY>':
                        tmp_L3_table_array[1][-2] = ''

                if now_folder_name != tmp_L3_table_array[1][0] and tmp_L3_table_array[1][0] != '':
                    now_folder_name = tmp_L3_table_array[1][0]
                tmp_L3_table_array[1][0] = now_folder_name

                if now_shape_name != tmp_L3_table_array[1][1] and tmp_L3_table_array[1][1] != '':
                    now_shape_name = tmp_L3_table_array[1][1]
                tmp_L3_table_array[1][1] = now_shape_name

                excel_L3_table_array.append(([tmp_L3_table_array[0] + tmp_offset_row,tmp_L3_table_array[1]]))
        #print('--- excel_L3_table_array ---')
        #print(excel_L3_table_array)
        L3_master_data_tuple = {}
        L3_master_data_tuple = ns_def.convert_array_to_tuple(excel_L3_table_array)
        #print('L3_master_data_tuple')
        #print(L3_master_data_tuple)
        # create Master_Data_L3 sheet
        ns_def.create_excel_sheet(excel_maseter_file, ws_name_l3 )

        # write L3_master_data_array tupple to Master data L3 sheet
        offset_row = 0
        offset_column = 0
        write_to_section = '_template_'
        ns_def.write_excel_meta(L3_master_data_tuple, excel_maseter_file, ws_name_l3, write_to_section, offset_row, offset_column)

        ### TEST MODE ###
        #ns_def.remove_excel_sheet(excel_maseter_file, 'Master_Data_L3')


class ns_l3_table_from_master_l3_sheet():
    def __init__(self):
        '''
        make L3 Table sheet from master L3 sheet
        '''
        #print('---ns_L3_table_from_master_L3()---')

        tmp_ws_name = '_tmp_'
        excel_maseter_file = self.inFileTxt_L3_1_1.get()
        write_excel_file = self.outFileTxt_11_2.get().replace('[MASTER]', '')
        excel_master_ws_name_L3 = 'Master_Data_L3'

        master_L3_array = []
        master_L3_array = ns_def.convert_master_to_array('Master_Data_L3', excel_maseter_file, '<<L3_TABLE>>')
        #print('--- master_L3_array ---')
        #print(master_L3_array)

        L3_table_array = []

        egt_maker_width_array = ['25', '25', '30', '20', '60', '50', '50']
        L3_table_array.append([1, ['<RANGE>', '1', '1', '1', '1', '1', '1', '1', '<END>']])
        L3_table_array.append([2, ['<HEADER>', 'Area', 'Device Name', 'L3 Port Name','L3 Instance Name', 'IP Address / Subnet mask (Comma Separated)', '[VPN] Target Device Name (Comma Separated)', '[VPN] Target L3 Port Name (Comma Separated)', '<END>']])


        max_row_num = 0
        for tmp_master_L3_array in master_L3_array:
            if tmp_master_L3_array[0] != 1 and tmp_master_L3_array[0] != 2:

                tmp_master_L3_array[1].extend(['<EMPTY>','<EMPTY>','<EMPTY>','<EMPTY>','<EMPTY>','<EMPTY>','<EMPTY>','<EMPTY>'])
                tmp_master_L3_array[1].insert(0, '')
                del tmp_master_L3_array[1][8:]
                tmp_master_L3_array[1].append('<END>')
                max_row_num = tmp_master_L3_array[0]

                #insert '>>' or <EMPTY> to column when greater than 7
                for i in range(4,8):
                    if tmp_master_L3_array[1][i] != '<EMPTY>' and tmp_master_L3_array[1][i] !=  '':
                        tmp_master_L3_array[1][i] = '>>' + str(tmp_master_L3_array[1][i])
                    elif tmp_master_L3_array[1][i] == '':
                        tmp_master_L3_array[1][i] = '<EMPTY>'

                L3_table_array.append([tmp_master_L3_array[0],tmp_master_L3_array[1]])

        L3_table_array.append([max_row_num + 1 ,['<END>']])

        #print('--- L3_table_array ---')
        #print(L3_table_array)

        # remove duplicated value in row 1 or 2
        new_L3_table_array = []
        tmp_current_value = 'dummy'
        for tmp_L3_table_array in L3_table_array:
            if tmp_L3_table_array[0] != 1 and tmp_L3_table_array[0] != 2 and tmp_L3_table_array[1][0] != '<END>':
                if tmp_L3_table_array[1][2] == tmp_current_value:
                    tmp_L3_table_array[1][2] = ''
                else:
                    tmp_current_value = tmp_L3_table_array[1][2]
                new_L3_table_array.append(tmp_L3_table_array)
            else:
                new_L3_table_array.append(tmp_L3_table_array)
                if tmp_L3_table_array[1][0] != '<END>':
                    tmp_current_value = tmp_L3_table_array[1][2]

        tmp_current_value = '_dummy_'
        new_new_L3_table_array = []
        for tmp_L3_table_array in new_L3_table_array:
            if tmp_L3_table_array[0] != 1 and tmp_L3_table_array[0] != 2 and tmp_L3_table_array[1][0] != '<END>':
                if tmp_L3_table_array[1][1] == tmp_current_value or tmp_L3_table_array[1][2] == '':
                    tmp_L3_table_array[1][1] = ''
                else:
                    tmp_current_value = tmp_L3_table_array[1][1]
                new_new_L3_table_array.append(tmp_L3_table_array)
            else:
                new_new_L3_table_array.append(tmp_L3_table_array)
                if tmp_L3_table_array[1][0] != '<END>':
                    tmp_current_value = tmp_L3_table_array[1][1]



        L3_table_tuple = {}
        L3_table_tuple = ns_def.convert_array_to_tuple(new_new_L3_table_array)

        '''Create _tmp_ sheet for L3_table_tuple'''
        ns_def.create_excel_sheet(write_excel_file, tmp_ws_name)

        # Write normal tuple to excel
        master_excel_meta = L3_table_tuple
        excel_file_path = write_excel_file
        worksheet_name = tmp_ws_name
        section_write_to = '<<N/A>>'
        offset_row = 0
        offset_column = 0
        #print('---master_excel_meta---')
        #print(master_excel_meta)
        ns_def.write_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

        '''Create L3 Table file'''
        import ns_egt_maker  # module import

        # create L3 Table file
        tmp_tmp_ws_name = '_tmp_tmp_'
        ns_def.create_excel_sheet(write_excel_file, tmp_tmp_ws_name)
        input_tree_tuple = {}
        input_tree_tuple[1, 1] = 'Dummy'  # first excel sheet has bug. That's why the first sheet is dummy sheet. this sheet will be deleted soon.
        input_tree_tuple[2, 1] = 'L3 Table'
        master_excel_meta = input_tree_tuple
        excel_file_path = write_excel_file
        worksheet_name = tmp_tmp_ws_name
        section_write_to = '<<N/A>>'
        offset_row = 0
        offset_column = 0
        ns_def.write_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)
        # Add L3 Table sheet
        input_excel_name = write_excel_file
        output_excel_name = write_excel_file

        NEW_OR_ADD = 'ADD'
        ns_egt_maker.create_excel_gui_tree(input_excel_name, output_excel_name, NEW_OR_ADD, egt_maker_width_array)
        ns_def.remove_excel_sheet(output_excel_name, 'Dummy')

        # Add L3 Table from meta
        # print(output_excel_name)
        self.input_tree_excel = openpyxl.load_workbook(output_excel_name)
        worksheet_name = 'L3 Table'
        start_row = 1
        start_column = 0
        custom_table_name = write_excel_file
        self.input_tree_excel = ns_egt_maker.insert_custom_excel_table(self.input_tree_excel, worksheet_name, start_row, start_column, custom_table_name)
        self.input_tree_excel.save(output_excel_name)

        # Remove _tmp_ sheet from excel master
        ns_def.remove_excel_sheet(write_excel_file, tmp_ws_name)
        ns_def.remove_excel_sheet(write_excel_file, tmp_tmp_ws_name)

