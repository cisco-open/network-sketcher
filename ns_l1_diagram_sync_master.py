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
import tkinter as tk ,tkinter.ttk as ttk,tkinter.filedialog, tkinter.messagebox
import ns_def ,ns_ddx_figure
import shutil

class  ns_l1_diagram_sync_master():
    def __init__(self):
        '''
        update Excel master data file from updated master data excel file and updated name array
        '''
        # parameter
        source_master_file = self.inFileTxt_92_2.get()
        target_master_file = self.excel_file_path
        source_excel_master_ws_name = 'Master_Data'
        target_excel_master_ws_name = 'Master_Data_tmp_'

        # self.updated_name_array made from ns_l1_create
        #print('self.updated_name_array   ' + str(self.updated_name_array))

        # convert from master to array and convert to tuple
        self.source_position_line_array = ns_def.convert_master_to_array(source_excel_master_ws_name, source_master_file, '<<POSITION_LINE>>')
        self.source_position_line_tuple = ns_def.convert_array_to_tuple(self.source_position_line_array)
        self.target_position_line_array = ns_def.convert_master_to_array(target_excel_master_ws_name, target_master_file, '<<POSITION_LINE>>')
        self.target_position_line_tuple = ns_def.convert_array_to_tuple(self.target_position_line_array)

        self.target_position_line_array_fix = []
        self.target_position_line_array_fix.append([2, ['From_Name', 'To_Name', 'From_Tag_Name', 'To_Tag_Name', 'From_Side(RIGHT/LEFT)', 'To_Side(RIGHT/LEFT)', 'Offset From_X (inches)', 'Offset From_Y (inches)', 'Offset To_X (inches)', 'Offset To_Y (inches)', 'Channel(inches)', 'Color(No) only)' \
            , 'From_Port_Name', 'From_Speed', 'From_Duplex', 'From_Port_Type', 'To__Port_Name', 'To_Speed', 'To_From_Duplex', 'To_From_Port_Type']])
        self.target_position_line_tuple_fix = ns_def.convert_array_to_tuple(self.target_position_line_array_fix)

        #print('---- self.source_position_line_tuple  ----')
        #print(self.source_position_line_tuple )

        #print('---- self.target_position_line_tuple ----')
        #print(self.target_position_line_tuple)

        #replace source position line tuple to updated name
        for tmp_source_position_line_tuple in self.source_position_line_tuple:
            if tmp_source_position_line_tuple[0] != 2 and tmp_source_position_line_tuple[0] != 1 and \
                    (tmp_source_position_line_tuple[1] == 1 or tmp_source_position_line_tuple[1] == 2):
                    for tmp_updated_name_array in self.updated_name_array:
                        if str(self.source_position_line_tuple[tmp_source_position_line_tuple]) == str(tmp_updated_name_array[0]):
                            #print('[Replaced device name] ' ,str(self.source_position_line_tuple[tmp_source_position_line_tuple]) + ' --> ' + str(tmp_updated_name_array[1]))
                            self.source_position_line_tuple[tmp_source_position_line_tuple] = str(tmp_updated_name_array[1])

        #print('---- replace device name , self.source_position_line_tuple  ----')
        #print(self.source_position_line_tuple )
        #print(self.target_position_line_tuple)

        used_source_array = []
        for tmp_target_position_line_tuple in self.target_position_line_tuple:
            flag_new_device = True

            if tmp_target_position_line_tuple[0] != 2 and tmp_target_position_line_tuple[0] != 1 and tmp_target_position_line_tuple[1] == 1:
                for tmp_source_position_line_tuple in self.source_position_line_tuple:
                    if tmp_source_position_line_tuple[0] != 2 and tmp_source_position_line_tuple[0] != 1 and tmp_source_position_line_tuple[1] == 1 and \
                                tmp_source_position_line_tuple[0] not in used_source_array:
                        ### normal line case ###
                        if str(self.target_position_line_tuple[tmp_target_position_line_tuple[0],tmp_target_position_line_tuple[1]]) == \
                                str(self.source_position_line_tuple[tmp_source_position_line_tuple[0],tmp_source_position_line_tuple[1]]) and \
                                str(self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1]+1]) == \
                                str(self.source_position_line_tuple[tmp_source_position_line_tuple[0], tmp_source_position_line_tuple[1]+1]):

                            #print('[match tuple]  ' + str(tmp_source_position_line_tuple) + str(self.source_position_line_tuple[tmp_source_position_line_tuple[0],tmp_source_position_line_tuple[1]])+ '  ' + str(self.source_position_line_tuple[tmp_source_position_line_tuple[0], tmp_source_position_line_tuple[1]+1])  + '  ' + str(tmp_source_position_line_tuple[1]))

                            # make fix tuple for target
                            for num in range(20):
                                if tmp_source_position_line_tuple[1] + num in [5,6,7,8,9,10]: # Fixed a bug. at ve 2.5.1e
                                    self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num] = self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num]
                                else:
                                    self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0],tmp_target_position_line_tuple[1] + num] = self.source_position_line_tuple[tmp_source_position_line_tuple[0],tmp_source_position_line_tuple[1] + num]

                            #mark used line
                            used_source_array.append(tmp_source_position_line_tuple[0])
                            flag_new_device = False
                            break

                        ### oppsite line case ###
                        if str(self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1]]) == \
                                str(self.source_position_line_tuple[tmp_source_position_line_tuple[0], tmp_source_position_line_tuple[1]+1]) and \
                                str(self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1]+1]) == \
                                str(self.source_position_line_tuple[tmp_source_position_line_tuple[0], tmp_source_position_line_tuple[1]]):
                            #print('[match oppsite]  ' + str(tmp_source_position_line_tuple)  +  str(self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1]]) + '  '+str(self.target_position_line_tuple[tmp_target_position_line_tuple[0]+1, tmp_target_position_line_tuple[1]]) + '  ' + str(tmp_source_position_line_tuple[1]))

                            # make fix tuple for target
                            for num in range(20):
                                if tmp_source_position_line_tuple[1] + num in [3]:
                                    tmp_i = 1
                                    self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num + tmp_i] = self.source_position_line_tuple[tmp_source_position_line_tuple[0], tmp_source_position_line_tuple[1] + num]
                                if tmp_source_position_line_tuple[1] + num in [7,8]:
                                    tmp_i = 2
                                    self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num + tmp_i] = self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num]
                                if tmp_source_position_line_tuple[1] + num in [13,14,15,16]:
                                    tmp_i = 4
                                    self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num + tmp_i] = self.source_position_line_tuple[tmp_source_position_line_tuple[0], tmp_source_position_line_tuple[1] + num]
                                if tmp_source_position_line_tuple[1] + num in [4]:
                                    tmp_i = -1
                                    self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num + tmp_i] = self.source_position_line_tuple[tmp_source_position_line_tuple[0], tmp_source_position_line_tuple[1] + num]
                                if tmp_source_position_line_tuple[1] + num in [9, 10]:
                                    tmp_i = -2
                                    self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num + tmp_i] = self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num]
                                if tmp_source_position_line_tuple[1] + num in [17, 18, 19, 20]:
                                    tmp_i = -4
                                    self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num + tmp_i] = self.source_position_line_tuple[tmp_source_position_line_tuple[0], tmp_source_position_line_tuple[1] + num]
                                # bug fix. at ver 2.5.1e
                                if tmp_source_position_line_tuple[1] + num in [1, 2, 11, 12]:
                                    tmp_i = 0
                                    self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num + tmp_i] = self.source_position_line_tuple[tmp_source_position_line_tuple[0], tmp_source_position_line_tuple[1] + num]
                                if tmp_source_position_line_tuple[1] + num in [5, 6]:
                                    tmp_i = 0
                                    self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num + tmp_i] = self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num]

                            # mark used line
                            used_source_array.append(tmp_source_position_line_tuple[0])
                            flag_new_device = False
                            break

                # if new device or line add to tuple
                if flag_new_device == True:
                    for num in range(20):
                        if str(self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num]) != '':
                            if tmp_source_position_line_tuple[1] + num in [7, 8 ,9 ,10]:
                                self.target_position_line_tuple_fix[tmp_source_position_line_tuple[0], tmp_source_position_line_tuple[1] + num] = self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num]
                            else:
                                self.target_position_line_tuple_fix[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num] = self.target_position_line_tuple[tmp_target_position_line_tuple[0], tmp_target_position_line_tuple[1] + num]

        #overwrite positon line in Master data
        #print('--- self.target_position_line_tuple_fix ---')
        #print(self.target_position_line_tuple_fix)

        offset_row = 0
        offset_column = 0
        write_to_section = '<<POSITION_LINE>>'
        ns_def.overwrite_excel_meta(self.target_position_line_tuple_fix, self.excel_file_path, target_excel_master_ws_name, write_to_section, offset_row, offset_column)

        #file rename
        if os.path.isfile(self.inFileTxt_92_2_2.get()) == True:
            os.remove(self.inFileTxt_92_2_2.get())
        #os.rename(self.inFileTxt_92_2.get(), self.inFileTxt_92_2_2.get())
        #os.rename(self.excel_file_path, self.inFileTxt_92_2.get())


