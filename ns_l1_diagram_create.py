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

class  ns_l1_diagram_create():
    def __init__(self):
        #parameter
        ws_name = 'Master_Data'
        tmp_ws_name = '_tmp_'
        ppt_meta_file = str(self.inFileTxt_2_1.get())
        self.all_slide_max_width = 10.0
        self.all_slide_max_hight = 5.0

        ### click action
        if self.click_value_dummy == '12-3':
            ppt_meta_file = str(self.inFileTxt_12_2.get())

        #convert from master to array and convert to tuple
        self.position_folder_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file,'<<POSITION_FOLDER>>')
        self.position_shape_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_SHAPE>>')
        self.position_line_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_LINE>>')
        self.position_style_shape_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<STYLE_SHAPE>>')
        self.position_tag_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_TAG>>')
        self.root_folder_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<ROOT_FOLDER>>')
        self.position_folder_tuple = ns_def.convert_array_to_tuple(self.position_folder_array)
        self.position_shape_tuple = ns_def.convert_array_to_tuple(self.position_shape_array)
        self.position_line_tuple = ns_def.convert_array_to_tuple(self.position_line_array)
        self.position_style_shape_tuple = ns_def.convert_array_to_tuple(self.position_style_shape_array)
        self.position_tag_tuple = ns_def.convert_array_to_tuple(self.position_tag_array)
        self.root_folder_tuple = ns_def.convert_array_to_tuple(self.root_folder_array)


        print('---- self.position_folder_tuple ----')
        #print(self.position_folder_array)
        #print(self.position_folder_tuple)
        print('---- self.position_shape_tuple ----')
        #print(self.position_shape_tuple)
        print('---- self.position_line_tuple ----')
        #print(self.position_line_tuple)

        # GET Folder and wp name List
        folder_wp_name_array = ns_def.get_folder_wp_array_from_master(ws_name, ppt_meta_file)
        print('---- folder_wp_name_array ----')
        #print(folder_wp_name_array)

        #GET way point with folder tuple
        wp_with_folder_tuple = {}
        for tmp_wp_folder_name in folder_wp_name_array[1]:
            current_row = 1
            flag_start_row = False
            flag_end_row = False

            while flag_end_row == False:
                if str(self.position_shape_tuple[current_row,1]) == tmp_wp_folder_name:
                    start_row = current_row
                    flag_start_row = True
                if flag_start_row == True and str(self.position_shape_tuple[current_row,1]) == '<END>':
                    flag_end_row = True
                    end_row = current_row - 1
                current_row += 1
            #print(tmp_wp_folder_name,start_row,end_row)

            for i in range(start_row,end_row+1):
                flag_start_column = False
                current_column = 2
                while flag_start_column == False:
                    if str(self.position_shape_tuple[i,current_column]) != '<END>':
                        wp_with_folder_tuple[self.position_shape_tuple[i,current_column]] = tmp_wp_folder_name
                    else:
                        flag_start_column = True
                    current_column += 1

        print('---- wp_with_folder_tuple ----')
        #print(wp_with_folder_tuple)

        '''
        Each self.click_value
        '''
        ''' for per folder button'''
        if self.click_value == '2-4-1' or self.click_value == '2-4-2':
            ''' Create PPT '''
            for tmp_folder_name in folder_wp_name_array[0]:
                current_row = 1
                flag_start_row = False
                flag_end_row = False

                while flag_end_row == False:
                    if str(self.position_shape_tuple[current_row, 1]) == tmp_folder_name:
                        start_row = current_row
                        flag_start_row = True
                    if flag_start_row == True and str(self.position_shape_tuple[current_row, 1]) == '<END>':
                        flag_end_row = True
                        end_row = current_row - 1
                    current_row += 1
                # print(tmp_folder_name,start_row,end_row)
                tmp_folder_array = []
                for i in range(start_row, end_row + 1):
                    flag_start_column = False
                    current_column = 2
                    while flag_start_column == False:
                        if str(self.position_shape_tuple[i, current_column]) != '<END>':
                            tmp_folder_array.append(self.position_shape_tuple[i, current_column])
                            # print(tmp_folder_name,self.position_shape_tuple[i,current_column])
                        else:
                            flag_start_column = True
                        current_column += 1
                #print(tmp_folder_name, tmp_folder_array)

                # GET connected wp up down left right
                connected_wp_folder_array =[]

                for tmp_shpae_name in tmp_folder_array:
                    for tmp_position_line_tuple in self.position_line_tuple:
                        if tmp_position_line_tuple[0] != 1:
                            if tmp_shpae_name == self.position_line_tuple[tmp_position_line_tuple[0],1]:
                                for tmp_wp_with_folder_tuple in wp_with_folder_tuple:
                                    if tmp_wp_with_folder_tuple == self.position_line_tuple[tmp_position_line_tuple[0],2]:
                                        if str(wp_with_folder_tuple[tmp_wp_with_folder_tuple]) not in connected_wp_folder_array:
                                            connected_wp_folder_array.append(wp_with_folder_tuple[tmp_wp_with_folder_tuple])
                            if tmp_shpae_name == self.position_line_tuple[tmp_position_line_tuple[0],2]:
                                for tmp_wp_with_folder_tuple in wp_with_folder_tuple:
                                    if tmp_wp_with_folder_tuple == self.position_line_tuple[tmp_position_line_tuple[0],1]:
                                        if str(wp_with_folder_tuple[tmp_wp_with_folder_tuple]) not in connected_wp_folder_array:
                                            connected_wp_folder_array.append(wp_with_folder_tuple[tmp_wp_with_folder_tuple])
                print('---- connected_wp_folder_array ----')
                #print(tmp_folder_name,connected_wp_folder_array)

                #GET extract_folder_tuple
                extract_folder_tuple = {}
                for tmp_position_folder_tuple in self.position_folder_tuple:
                    if self.position_folder_tuple[tmp_position_folder_tuple] == tmp_folder_name or self.position_folder_tuple[tmp_position_folder_tuple] in connected_wp_folder_array:
                        extract_folder_tuple[tmp_position_folder_tuple] = self.position_folder_tuple[tmp_position_folder_tuple]
                        extract_folder_tuple[tmp_position_folder_tuple[0]-1,tmp_position_folder_tuple[1]] = self.position_folder_tuple[tmp_position_folder_tuple[0]-1,tmp_position_folder_tuple[1]]
                        extract_folder_tuple[tmp_position_folder_tuple[0],1] = self.position_folder_tuple[tmp_position_folder_tuple[0],1]
                        extract_folder_tuple[tmp_position_folder_tuple[0] - 1, 1] = self.position_folder_tuple[tmp_position_folder_tuple[0] - 1, 1]

                print('---- extract_folder_tuple ----')
                #print(extract_folder_tuple)

                # copy Master_Data sheet to _tmp_
                ns_def.copy_excel_sheet(ws_name, ppt_meta_file, tmp_ws_name)

                # clear values a selected section in _tmp_ sheet
                #### clear <<POSITION_FOLDER>>
                clear_section_tuple = self.position_folder_tuple
                ns_def.clear_section_sheet(tmp_ws_name, ppt_meta_file, clear_section_tuple)

                #### make <<POSITION_FOLDER>> tuple
                convert_array = []
                convert_array = ns_def.convert_tuple_to_array(extract_folder_tuple)
                #print(convert_array)

                # adjust for <<POSITION_FOLDER>>
                offset_row = convert_array[0][0] * -1 +2

                current_y_grid_array = []
                flag_first_array = True
                if convert_array[0][1][0] == '<SET_WIDTH>':
                    for tmp_array in convert_array:
                        current_y_grid_array.append([tmp_array[0] + offset_row,tmp_array[1]])
                else:
                    for tmp_array in convert_array:
                        if flag_first_array == True:
                            current_y_grid_array.append([tmp_array[0] + offset_row - 1, tmp_array[1]])
                            flag_first_array = False
                        else:
                            current_y_grid_array.append([tmp_array[0] + offset_row -1,tmp_array[1]])

                print('---- current_y_grid_array ----')
                #print(current_y_grid_array)

                convert_tuple = {}
                convert_tuple = ns_def.convert_array_to_tuple(current_y_grid_array)
                print('---- convert_tuple ----')
                #print(convert_tuple)

                master_excel_meta = convert_tuple
                excel_file_path = ppt_meta_file
                worksheet_name = tmp_ws_name
                section_write_to = '<<POSITION_FOLDER>>'
                offset_row = 0
                offset_column = 0
                ns_def.write_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

                'parameter'
                master_folder_tuple = convert_tuple
                master_style_shape_tuple = self.position_style_shape_tuple
                master_shape_tuple = self.position_shape_tuple
                min_tag_inches = 0.3  # inches,  between side of folder and eghe shape. left and right.

                #### GET best width size ####
                master_folder_size_array = ns_def.get_folder_width_size(master_folder_tuple, master_style_shape_tuple, master_shape_tuple, min_tag_inches)
                print('###master_folder_size_array###  \n'  + str(master_folder_size_array))

                ### get root folder tuple
                master_root_folder_tuple = ns_def.get_root_folder_tuple(self,master_folder_size_array,tmp_folder_name)
                #print('###master_root_folder_tuple###  \n'  + str(master_root_folder_tuple))

                #### clear <<ROOT_FOLDER>>
                clear_section_tuple = master_root_folder_tuple
                clear_section_tuple[1,1] = '<<ROOT_FOLDER>>'
                ns_def.clear_section_sheet(tmp_ws_name, ppt_meta_file, clear_section_tuple)

                ### write root folder
                write_to_section = '<<ROOT_FOLDER>>'
                ns_def.write_excel_meta(master_root_folder_tuple, excel_file_path, worksheet_name, write_to_section, offset_row, offset_column)

                if self.click_value == '2-4-1':
                    #### clear <<POSITION_TAG>>
                    clear_section_tuple = self.position_tag_tuple
                    clear_section_tuple[1,1] = '<<POSITION_TAG>>'
                    ns_def.clear_section_sheet(tmp_ws_name, ppt_meta_file, clear_section_tuple)

                    #### clear tag in <<POSITION_LINE>>
                    clear_section_tuple = self.position_line_tuple
                    clear_section_tuple[1,1] = '<<POSITION_LINE>>'
                    ns_def.clear_tag_in_position_line(tmp_ws_name, ppt_meta_file, clear_section_tuple)

                # adjust for per slide size
                self.root_width = self.all_slide_max_width
                self.root_hight = self.all_slide_max_hight + 1.0  # top side margin + 1.0

                ### Create ppt
                #self.output_diagram_path = self.outFileTxt_2_1.get()
                self.excel_file_path = ppt_meta_file
                self.worksheet_name = tmp_ws_name

                print('---- master_root_folder_tuple ---- ')
                #print(master_root_folder_tuple)

                if self.all_slide_max_width < master_root_folder_tuple[2,7]:
                    self.all_slide_max_width = master_root_folder_tuple[2, 7]
                if self.all_slide_max_hight < master_root_folder_tuple[2,8]  + 1.0:  # top side margin + 1.0
                    self.all_slide_max_hight = master_root_folder_tuple[2, 8]  + 1.0 # top side margin + 1.0

                ns_ddx_figure.ns_ddx_figure_run.__init__(self)
                #exit()

            # Remove _tmp_ sheet from excel master
            ns_def.remove_excel_sheet(ppt_meta_file, tmp_ws_name)

        ''' for Entire NW buttom'''
        if self.click_value == '2-4-3' or self.click_value == '2-4-4':
            # copy Master_Data sheet to _tmp_
            ns_def.copy_excel_sheet(ws_name, ppt_meta_file, tmp_ws_name)

            if self.click_value == '2-4-3':
                #### clear <<POSITION_TAG>>
                clear_section_tuple = self.position_tag_tuple
                clear_section_tuple[1, 1] = '<<POSITION_TAG>>'
                ns_def.clear_section_sheet(tmp_ws_name, ppt_meta_file, clear_section_tuple)

                #### clear tag in <<POSITION_LINE>>
                clear_section_tuple = self.position_line_tuple
                clear_section_tuple[1, 1] = '<<POSITION_LINE>>'
                ns_def.clear_tag_in_position_line(tmp_ws_name, ppt_meta_file, clear_section_tuple)

            self.root_left = self.root_folder_tuple[2,5]
            self.root_top =  self.root_folder_tuple[2,6]
            self.root_width = self.root_folder_tuple[2,7]
            self.root_hight = self.root_folder_tuple[2,8]

            ### Create ppt
            self.excel_file_path = ppt_meta_file
            self.worksheet_name = tmp_ws_name

            ns_ddx_figure.ns_ddx_figure_run.__init__(self)

            # Remove _tmp_ sheet from excel master
            ns_def.remove_excel_sheet(ppt_meta_file, tmp_ws_name)

if __name__ == '__main__':
    ns_l1_diagram_create()
