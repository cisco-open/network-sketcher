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

class  ns_l2_table_sync_master():
    def __init__(self):
        '''
        update Excel master data file from l2 table excel file
        '''
        # parameter
        l2_table_ws_name = 'L2 Table'
        l2_table_file = self.inFileTxt_L2_2_1.get() 
        excel_master_ws_name = 'Master_Data'
        excel_master_ws_name_l2 = 'Master_Data_L2'
        excel_maseter_file = self.inFileTxt_L2_2_2.get()
        excel_maseter_file_backup = self.inFileTxt_L2_2_2_2.get()

        #check L2 Table sheet in Excel file
        input_excel_l2_table = openpyxl.load_workbook(l2_table_file)
        ws_list = input_excel_l2_table.get_sheet_names()
        if l2_table_ws_name not in ws_list:
            tkinter.messagebox.showwarning(title="L2 Table not in [L2_TABLE] file", message="Please check L2 Table sheet in [L2_TABLE] file.")
            return

        #backup L2 Table Excel file
        '''excel_maseter_file_backup = self.inFileTxt_L2_2_2_backup
        if os.path.isfile(excel_maseter_file_backup) == True:
            os.remove(excel_maseter_file_backup)
            shutil.copy(excel_maseter_file, excel_maseter_file_backup)
        else:
            shutil.copy(excel_maseter_file, excel_maseter_file_backup)'''

        #get L2 Table Excel file
        l2_table_array = []
        l2_table_array = ns_def.convert_excel_to_array(l2_table_ws_name, l2_table_file, 3)
        print('--- l2_table_array ---')
        #print(l2_table_array)

        tmp_current_value1 = '_dummy_'
        tmp_current_value2 = '_dummy_'
        new_l2_table_array = []
        for tmp_l2_table_array in l2_table_array:
            tmp_l2_table_array[1][2] = ''   #delete excel function
            tmp_l2_table_array[1][4] = ''

            if tmp_l2_table_array[1][0] == '':
                tmp_l2_table_array[1][0] = tmp_current_value1
            else:
                if tmp_l2_table_array[1][0] != tmp_current_value1:
                    tmp_current_value1 = tmp_l2_table_array[1][0]

            if tmp_l2_table_array[1][1] == '':
                tmp_l2_table_array[1][1] = tmp_current_value2
            else:
                if tmp_l2_table_array[1][1] != tmp_current_value2:
                    tmp_current_value2 = tmp_l2_table_array[1][1]

            new_l2_table_array.append(tmp_l2_table_array)

        sorted_new_l2_table_array = []

        for tmp_new_l2_table_array in new_l2_table_array:
            tmp_new_l2_table_array[1].extend(['','','','','','','',''])
            del tmp_new_l2_table_array[1][8:]

            # add physical port id for sort
            if_value = ns_def.get_if_value(tmp_new_l2_table_array[1][3])
            tmp_new_l2_table_array[1].extend([if_value])
            tmp_new_l2_table_array[1].extend([str(ns_def.split_portname(tmp_new_l2_table_array[1][3])[0])])

            # add logical port id for sort
            if tmp_new_l2_table_array[1][5] != '':
                if_value = ns_def.get_if_value(tmp_new_l2_table_array[1][5])
                tmp_new_l2_table_array[1].extend([if_value])
                tmp_new_l2_table_array[1].extend([str(ns_def.split_portname(tmp_new_l2_table_array[1][5])[0])])

            #replace ' ' to '' in l2segment description
            tmp_new_l2_table_array[1][6] = tmp_new_l2_table_array[1][6].replace(' ','')

            #if '\n' in vport name, change to changed name.
            if '\n' in tmp_new_l2_table_array[1][5]:
                tmp_split_vport_name = tmp_new_l2_table_array[1][5].split('\n')
                tmp_new_l2_table_array[1][5] = tmp_split_vport_name[1]

            #sort l2segments
            tmp_array = tmp_new_l2_table_array[1][6].split(',')
            tmp_last_array = []
            for tmp_tmp_array in tmp_array:
                if tmp_tmp_array != '':
                    tmp_last_array.append(tmp_tmp_array)
            tmp_array = tmp_last_array
            tmp_array.sort()
            tmp_str = str(tmp_array).replace('\'','').replace('[','').replace(']','').replace(' ','')
            tmp_new_l2_table_array[1][6] = tmp_str

            #replace ' ' to '' in Directory l2segment description
            tmp_new_l2_table_array[1][7] = tmp_new_l2_table_array[1][7].replace(' ','')

            #sort Direcotry l2segments
            tmp_array = tmp_new_l2_table_array[1][7].split(',')
            tmp_last_array = []
            for tmp_tmp_array in tmp_array:
                if tmp_tmp_array != '':
                    tmp_last_array.append(tmp_tmp_array)
            tmp_array = tmp_last_array
            tmp_array.sort()
            tmp_str = str(tmp_array).replace('\'','').replace('[','').replace(']','')
            tmp_new_l2_table_array[1][7] = tmp_str

            sorted_new_l2_table_array.append(tmp_new_l2_table_array[1])

        #sort sorted_new_l2_table_array
        sorted_new_l2_table_array = sorted(sorted_new_l2_table_array , reverse=False, key=lambda x: (x[0],x[1],x[9],x[8],x[5],x[7],x[6]))  # sort l2 table

        tmp_row_num = 3
        last_l2_table_array = []
        last_l2_table_array.append([1, ['<<L2_TABLE>>']])
        last_l2_table_array.append([2, ['Area', 'Device Name', 'Port Mode', 'Port Name', 'Virtual Port Mode', 'Virtual Port Name', 'Connected L2 Segment Name(Comma Separated)',  'L2 Name directly received by L3 Virtual Port (Comma Separated)']])
        for tmp_sorted_new_l2_table_array in sorted_new_l2_table_array:
            del tmp_sorted_new_l2_table_array[8:]
            last_l2_table_array.append([tmp_row_num,tmp_sorted_new_l2_table_array])
            tmp_row_num += 1

        print('--- last_l2_table_array ---')
        #print(last_l2_table_array)

        last_l2_table_tuple = {}
        last_l2_table_tuple = ns_def.convert_array_to_tuple(last_l2_table_array)

        #delete L2 Table sheet
        ns_def.remove_excel_sheet(excel_maseter_file, excel_master_ws_name_l2)
        #create L2 Table sheet
        ns_def.create_excel_sheet(excel_maseter_file, excel_master_ws_name_l2)
        #write tuple to excel master data
        ns_def.write_excel_meta(last_l2_table_tuple, excel_maseter_file, excel_master_ws_name_l2, '_template_', 0, 0)

        input_excel_l2_table.close()

        return (True)
