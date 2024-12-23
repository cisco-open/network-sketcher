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
import ns_def , ns_ddx_figure
import openpyxl
from pptx import Presentation
from pptx.enum.dml import MSO_LINE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor, MSO_THEME_COLOR
from pptx.dml.line import LineFormat
from pptx.shapes.connector import Connector
from pptx.util import Inches, Cm, Pt

class  ns_l2_diagram_create():
    def __init__(self):
        #print('--- ns_l2_diagram_create ---')

        '''
        STEP0 get values of Master Data
        '''
        #parameter
        ws_name = 'Master_Data'
        ws_l2_name = 'Master_Data_L2'
        excel_maseter_file = self.inFileTxt_L2_3_1.get()

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
        self.folder_wp_name_array = ns_def.get_folder_wp_array_from_master(ws_name, excel_maseter_file)
        #print('---- folder_wp_name_array ----')
        #print(self.folder_wp_name_array)

        # GET L2 Table sheet
        self.l2_table_array = ns_def.convert_master_to_array(ws_l2_name, excel_maseter_file, '<<L2_TABLE>>')

        '''
        STEP1 get size of Devices
        '''
        new_l2_table_array = []
        for tmp_l2_table_array in self.l2_table_array:
            if tmp_l2_table_array[0] != 1 and tmp_l2_table_array[0] != 2:
                tmp_l2_table_array[1].extend(['', '', '', '', '', '', '', ''])
                del tmp_l2_table_array[1][8:]
                new_l2_table_array.append(tmp_l2_table_array)

        device_list_array = []
        wp_list_array = []
        for tmp_new_l2_table_array in new_l2_table_array:
            if tmp_new_l2_table_array[1][1] not in device_list_array and tmp_new_l2_table_array[1][1] not in wp_list_array:
                if tmp_new_l2_table_array[1][0] == 'N/A':
                    wp_list_array.append(tmp_new_l2_table_array[1][1])
                else:
                    if str(self.comboL2_3_6.get()) == str(tmp_new_l2_table_array[1][0]):  # add at ver 2.3.0(b)
                        device_list_array.append(tmp_new_l2_table_array[1][1])

        #print('''--- device_list_array ---''')
        #print(device_list_array)
        #print('''--- wp_list_array ---''')
        #print(wp_list_array)

        '''GET all device l2 size'''
        self.ppt_edge_margin = 1.0  # inches
        self.ppt_width = 15.0  # inches
        self.ppt_hight = 10.0  # inches

        self.all_device_l2_size_array = []
        #print('--- l2_device_materials   RETURN_DEVICE_SIZE ---')
        for tmp_device_list_array in device_list_array:
            self.active_ppt = Presentation()  #define target ppt object
            action_type = 'RETURN_DEVICE_SIZE' #'RETURN_DEVICE_SIZE' - > return array[left, top , width, hight] , 'WRITE_DEVICE_L2' -> write device l2 materials
            input_device_name = tmp_device_list_array  #device_name
            write_width_hight_array = []
            device_size_array = ns_ddx_figure.extended.l2_device_materials(self,action_type,input_device_name,write_width_hight_array,wp_list_array)  # offset_left, offset_top , right , left
            self.all_device_l2_size_array.append([input_device_name,device_size_array])

        for tmp_wp_list_array in wp_list_array:
            self.active_ppt = Presentation()  #define target ppt object
            action_type = 'RETURN_DEVICE_SIZE' #'RETURN_DEVICE_SIZE' - > return array[left, top , width, hight] , 'WRITE_DEVICE_L2' -> write device l2 materials
            input_device_name = tmp_wp_list_array  #device_name
            write_width_hight_array = []
            device_size_array = ns_ddx_figure.extended.l2_device_materials(self,action_type,input_device_name,write_width_hight_array,wp_list_array)  # offset_left, offset_top , right , left
            self.all_device_l2_size_array.append([input_device_name,device_size_array])
        #print('--- self.all_device_l2_size_array ---')
        #print(self.all_device_l2_size_array)

        '''
        Create per area l2 ppt
        '''
        if self.click_value == 'L2-3-2':
            action_type = 'CREATE_L2_AREA' # or 'GET_SIZE'
            data_array = []
            ns_l2_diagram_create.l2_area_create(self, action_type, data_array)

        '''
        Create per device l2 ppt
        '''
        if self.click_value == 'L2-3-3':
            #get ppt size
            for tmp_all_device_size_array in self.all_device_l2_size_array:
                if tmp_all_device_size_array[1][2] > self.ppt_width:
                    self.ppt_width = tmp_all_device_size_array[1][2]
                if tmp_all_device_size_array[1][3] > self.ppt_hight:
                    self.ppt_hight = tmp_all_device_size_array[1][3]

            self.active_ppt = Presentation()  # define target ppt object
            #print('--- l2_device_materials   WRITE_DEVICE_L2 ---')
            for tmp_all_device_size_array in self.all_device_l2_size_array:
                input_device_name = tmp_all_device_size_array[0]  # device_name
                device_size_array =  tmp_all_device_size_array[1]
                action_type = 'WRITE_DEVICE_L2' #'RETURN_DEVICE_SIZE' - > return array[width, hight] , 'WRITE_DEVICE_L2' -> write device l2 materials
                write_left_top_array = [3.0 , 3.0 ,device_size_array] # [left , top , [offset_left, offset_top , right , left]]
                ns_ddx_figure.extended.l2_device_materials(self, action_type, input_device_name, write_left_top_array,wp_list_array)

            ### save pptx file
            self.active_ppt.save(self.output_ppt_file)

            return

    def l2_area_create(self, action_type, data_array):
        #parameter
        ws_name = 'Master_Data'
        tmp_ws_name = '_tmp_'
        ppt_meta_file = self.inFileTxt_L2_3_1.get()
        self.all_slide_max_width = 0
        self.all_slide_max_hight = 0

        #### change shape of position_style size to L2 Device
        l2_position_style_shape_array = []
        for tmp_position_style_shape_array  in self.position_style_shape_array:
            if tmp_position_style_shape_array[0] in [1,2,3]:
                l2_position_style_shape_array.append(tmp_position_style_shape_array)
            else:
                for tmp_all_device_l2_size_array in self.all_device_l2_size_array:
                    if tmp_position_style_shape_array[1][0] == tmp_all_device_l2_size_array[0]:
                        tmp_position_style_shape_array[1][1] = tmp_all_device_l2_size_array[1][2]
                        tmp_position_style_shape_array[1][2] = tmp_all_device_l2_size_array[1][3]
                        break

                l2_position_style_shape_array.append(tmp_position_style_shape_array)

        #print('--- l2_position_style_shape_array ---')
        #print(l2_position_style_shape_array)

        l2_position_style_shape_tuple = ns_def.convert_array_to_tuple(l2_position_style_shape_array)

        # GET way point with folder tuple
        wp_with_folder_tuple = {}
        for tmp_wp_folder_name in self.folder_wp_name_array[1]:
            current_row = 1
            flag_start_row = False
            flag_end_row = False

            while flag_end_row == False:
                if str(self.position_shape_tuple[current_row, 1]) == tmp_wp_folder_name:
                    start_row = current_row
                    flag_start_row = True
                if flag_start_row == True and str(self.position_shape_tuple[current_row, 1]) == '<END>':
                    flag_end_row = True
                    end_row = current_row - 1
                current_row += 1
            # print(tmp_wp_folder_name,start_row,end_row)

            for i in range(start_row, end_row + 1):
                flag_start_column = False
                current_column = 2
                while flag_start_column == False:
                    if str(self.position_shape_tuple[i, current_column]) != '<END>':
                        wp_with_folder_tuple[self.position_shape_tuple[i, current_column]] = tmp_wp_folder_name
                    else:
                        flag_start_column = True
                    current_column += 1

        #print('---- wp_with_folder_tuple ----')
        # print(wp_with_folder_tuple)
        # GET connected way point of each folder
        ''' GET Size of PPT slide '''
        #for tmp_folder_name in self.folder_wp_name_array[0]:
        tmp_folder_name = self.comboL2_3_6.get()
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
                    tmp_folder_array.append(self.position_shape_tuple[i, current_column])  # print(tmp_folder_name,self.position_shape_tuple[i,current_column])
                else:
                    flag_start_column = True
                current_column += 1
        # print(tmp_folder_name, tmp_folder_array)

        # GET connected wp up down left right
        connected_wp_folder_array = []

        for tmp_shpae_name in tmp_folder_array:
            for tmp_position_line_tuple in self.position_line_tuple:
                if tmp_position_line_tuple[0] != 1:
                    if tmp_shpae_name == self.position_line_tuple[tmp_position_line_tuple[0], 1]:
                        for tmp_wp_with_folder_tuple in wp_with_folder_tuple:
                            if tmp_wp_with_folder_tuple == self.position_line_tuple[tmp_position_line_tuple[0], 2]:
                                if str(wp_with_folder_tuple[tmp_wp_with_folder_tuple]) not in connected_wp_folder_array:
                                    connected_wp_folder_array.append(wp_with_folder_tuple[tmp_wp_with_folder_tuple])
                    if tmp_shpae_name == self.position_line_tuple[tmp_position_line_tuple[0], 2]:
                        for tmp_wp_with_folder_tuple in wp_with_folder_tuple:
                            if tmp_wp_with_folder_tuple == self.position_line_tuple[tmp_position_line_tuple[0], 1]:
                                if str(wp_with_folder_tuple[tmp_wp_with_folder_tuple]) not in connected_wp_folder_array:
                                    connected_wp_folder_array.append(wp_with_folder_tuple[tmp_wp_with_folder_tuple])
        #print('---- connected_wp_folder_array ----')
        # print(tmp_folder_name,connected_wp_folder_array)

        # GET extract_folder_tuple
        extract_folder_tuple = {}
        for tmp_position_folder_tuple in self.position_folder_tuple:
            if self.position_folder_tuple[tmp_position_folder_tuple] == tmp_folder_name or self.position_folder_tuple[tmp_position_folder_tuple] in connected_wp_folder_array:
                extract_folder_tuple[tmp_position_folder_tuple] = self.position_folder_tuple[tmp_position_folder_tuple]
                extract_folder_tuple[tmp_position_folder_tuple[0] - 1, tmp_position_folder_tuple[1]] = self.position_folder_tuple[tmp_position_folder_tuple[0] - 1, tmp_position_folder_tuple[1]]
                extract_folder_tuple[tmp_position_folder_tuple[0], 1] = self.position_folder_tuple[tmp_position_folder_tuple[0], 1]
                extract_folder_tuple[tmp_position_folder_tuple[0] - 1, 1] = self.position_folder_tuple[tmp_position_folder_tuple[0] - 1, 1]

        #print('---- extract_folder_tuple ----')
        # print(extract_folder_tuple)

        # copy Master_Data sheet to _tmp_
        print(ws_name, ppt_meta_file, tmp_ws_name)
        ns_def.copy_excel_sheet(ws_name, ppt_meta_file, tmp_ws_name)  # uncommented at 2.0

        # clear values a selected section in _tmp_ sheet
        #### clear <<POSITION_FOLDER>>
        clear_section_tuple = self.position_folder_tuple
        ns_def.clear_section_sheet(tmp_ws_name, ppt_meta_file, clear_section_tuple) # uncommented at 2.0

        #### make <<POSITION_FOLDER>> tuple
        convert_array = []
        convert_array = ns_def.convert_tuple_to_array(extract_folder_tuple)
        # print(convert_array)

        # adjust for <<POSITION_FOLDER>>
        offset_row = convert_array[0][0] * -1 + 2

        current_y_grid_array = []
        flag_first_array = True
        if convert_array[0][1][0] == '<SET_WIDTH>':
            for tmp_array in convert_array:
                current_y_grid_array.append([tmp_array[0] + offset_row, tmp_array[1]])
        else:
            for tmp_array in convert_array:
                if flag_first_array == True:
                    current_y_grid_array.append([tmp_array[0] + offset_row - 1, tmp_array[1]])
                    flag_first_array = False
                else:
                    current_y_grid_array.append([tmp_array[0] + offset_row - 1, tmp_array[1]])

        #print('---- current_y_grid_array ----')
        # print(current_y_grid_array)

        convert_tuple = {}
        convert_tuple = ns_def.convert_array_to_tuple(current_y_grid_array)
        #print('---- convert_tuple ----')
        # print(convert_tuple)

        master_excel_meta = convert_tuple
        excel_file_path = ppt_meta_file
        worksheet_name = tmp_ws_name
        section_write_to = '<<POSITION_FOLDER>>'
        offset_row = 0
        offset_column = 0

        if action_type == 'CREATE_L2_AREA':
            ns_def.write_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column) # uncommented at 2.0

        ### change <<STYLE_SHAPE>> for L2 Device size ###
        #### clear <<STYLE_SHAPE>>
        clear_section_tuple = self.position_style_shape_tuple
        if action_type == 'CREATE_L2_AREA':
            ns_def.clear_section_sheet(tmp_ws_name, ppt_meta_file, clear_section_tuple)  # uncommented at 2.0

        #### make <<STYLE_SHAPE>> tuple
        master_excel_meta = l2_position_style_shape_tuple
        excel_file_path = ppt_meta_file
        worksheet_name = tmp_ws_name
        section_write_to = '<<STYLE_SHAPE>>'
        offset_row = 0
        offset_column = 0
        ns_def.write_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)

        '''parameter'''
        master_folder_tuple = convert_tuple
        master_style_shape_tuple = l2_position_style_shape_tuple
        master_shape_tuple = self.position_shape_tuple
        min_tag_inches = 0.8  # inches,  between side of folder and edge shape. left and right. margin.

        '''# GET best width size #'''
        master_folder_size_array = ns_def.get_folder_width_size(master_folder_tuple, master_style_shape_tuple, master_shape_tuple, min_tag_inches)
        #print('--- master_folder_size_array   slide_max_width_inches,master_width_size_y_grid,master_folder_size,slide_max_hight_inches,master_hight_size_y_grid---')
        #print(master_folder_size_array)

        ### get root folder tuple
        master_root_folder_tuple = ns_def.get_root_folder_tuple(self, master_folder_size_array, tmp_folder_name)

        #### clear <<ROOT_FOLDER>>
        clear_section_tuple = master_root_folder_tuple
        clear_section_tuple[1, 1] = '<<ROOT_FOLDER>>'
        if action_type == 'CREATE_L2_AREA':
            ns_def.clear_section_sheet(tmp_ws_name, ppt_meta_file, clear_section_tuple) # uncommented at 2.0

        ### write root folder
        write_to_section = '<<ROOT_FOLDER>>'
        if action_type == 'CREATE_L2_AREA':
            ns_def.write_excel_meta(master_root_folder_tuple, excel_file_path, worksheet_name, write_to_section, offset_row, offset_column) # uncommented at 2.0

        ''' input master_width_size_y_grid in master_folder_size_array to excel  for l2 device'''
        update_current_y_grid_array = current_y_grid_array

        #modify x grid
        tmp_sheet_position_folder_array = ns_def.convert_master_to_array(tmp_ws_name, ppt_meta_file, '<<POSITION_FOLDER>>')
        tmp_sheet_position_folder_tuple = ns_def.convert_array_to_tuple(tmp_sheet_position_folder_array)
        for tmp_tmp_sheet_position_folder_tuple in tmp_sheet_position_folder_tuple:
            for tmp_master_folder_size_array_2 in master_folder_size_array[2]:
                if tmp_sheet_position_folder_tuple[tmp_tmp_sheet_position_folder_tuple[0],tmp_tmp_sheet_position_folder_tuple[1]] == tmp_master_folder_size_array_2[1][0][0] and tmp_master_folder_size_array_2[1][0][0] != 10 and \
                    isinstance(tmp_sheet_position_folder_tuple[tmp_tmp_sheet_position_folder_tuple[0],tmp_tmp_sheet_position_folder_tuple[1]], str) == True and isinstance(tmp_master_folder_size_array_2[1][0][0], str) == True: # add at ver 2.3.4
                        tmp_num = 0
                        for tmp_current_y_grid_array in current_y_grid_array:
                            if tmp_current_y_grid_array[0] == tmp_tmp_sheet_position_folder_tuple[0] - 1 and tmp_master_folder_size_array_2[1][0][1] != 0:
                                #print(update_current_y_grid_array[tmp_num][1][tmp_tmp_sheet_position_folder_tuple[1] - 1], tmp_master_folder_size_array_2[1][0][1])
                                update_current_y_grid_array[tmp_num][1][tmp_tmp_sheet_position_folder_tuple[1] - 1] = tmp_master_folder_size_array_2[1][0][1]
                                break
                            tmp_num += 1

        #modify y grid
        tmp_num = 0
        for tmp_current_y_grid_array in current_y_grid_array:
            if isinstance(tmp_current_y_grid_array[1][0], int) == True or  isinstance(tmp_current_y_grid_array[1][0], float) == True:
                max_y_grid_current = tmp_current_y_grid_array[1][0]
                for tmp_tmp_current_y_grid_array in tmp_current_y_grid_array[1:]:
                    for tmp_master_folder_size_array_2 in master_folder_size_array[2]:
                        if tmp_master_folder_size_array_2[1][0][0] == tmp_tmp_current_y_grid_array[1]:
                            if max_y_grid_current < tmp_master_folder_size_array_2[1][0][2]:
                                max_y_grid_current = tmp_master_folder_size_array_2[1][0][2]
                                update_current_y_grid_array[tmp_num][1][0] = max_y_grid_current
                                break
            tmp_num += 1

        #print('--- update_current_y_grid_array ---',update_current_y_grid_array)
        update_current_y_grid_tuple = ns_def.convert_array_to_tuple(update_current_y_grid_array)

        ### clear <<POSITION_FOLDER>>
        clear_sheet_position_folder_array = ns_def.convert_master_to_array(tmp_ws_name, ppt_meta_file, '<<POSITION_FOLDER>>')
        clear_sheet_position_folder_tuple = ns_def.convert_array_to_tuple(clear_sheet_position_folder_array)
        ns_def.clear_section_sheet(tmp_ws_name, ppt_meta_file, clear_sheet_position_folder_tuple)

        # prepare all_silde max size
        if self.all_slide_max_width < master_root_folder_tuple[2, 7]:
            self.all_slide_max_width = master_root_folder_tuple[2, 7]
        if self.all_slide_max_hight < master_root_folder_tuple[2, 8]:
            self.all_slide_max_hight = master_root_folder_tuple[2, 8]

        ### write <<POSITION_FOLDER>>
        write_to_section = '<<POSITION_FOLDER>>'
        if action_type == 'CREATE_L2_AREA':
            ns_def.write_excel_meta(update_current_y_grid_tuple, excel_file_path, worksheet_name, write_to_section, offset_row, offset_column)  # uncommented at 2.0

            # adjust for per slide size
            self.root_width = self.all_slide_max_width
            self.root_hight = self.all_slide_max_hight

            #print('---- self.all_slide_max_width,self.all_slide_max_hight ---- ', self.all_slide_max_width, self.all_slide_max_hight)

        ### Create ppt
        self.output_diagram_path = self.outFileTxt_2_1.get()
        self.excel_file_path = ppt_meta_file
        self.worksheet_name = tmp_ws_name

        '''run ns_ddx_figure'''
        if action_type == 'CREATE_L2_AREA':
            self.l2_folder_name = tmp_folder_name
            self.all_tag_size_array = []
            ns_ddx_figure.ns_ddx_figure_run.__init__(self)  # exit()  # uncommented at 2.0


        # Remove _tmp_ sheet from excel master
        if action_type == 'CREATE_L2_AREA':
            ns_def.remove_excel_sheet(ppt_meta_file, tmp_ws_name) # uncommented at 2.0



        return ([action_type,[self.all_slide_max_width,self.all_slide_max_hight]])
