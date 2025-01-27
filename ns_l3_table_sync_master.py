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

class  ns_l3_table_sync_master():
    def __init__(self):
        '''
        update Excel master data file from l3 table excel file
        '''
        # parameter
        l3_table_ws_name = 'L3 Table'
        l3_table_file = self.inFileTxt_L3_2_1.get() 
        excel_master_ws_name = 'Master_Data'
        excel_master_ws_name_l3 = 'Master_Data_L3'
        excel_maseter_file = self.inFileTxt_L3_2_2.get()
        excel_maseter_file_backup = self.inFileTxt_L3_2_2_2.get()

        #check L3 Table sheet in Excel file
        input_excel_l3_table = openpyxl.load_workbook(l3_table_file)
        ws_list = input_excel_l3_table.sheetnames
        if l3_table_ws_name not in ws_list:
            tkinter.messagebox.showwarning(title="L3 Table not in [L3_TABLE] file", message="Please check L3 Table sheet in [L3_TABLE] file.")
            return

        #get L3 Table Excel file
        l3_table_array = []
        l3_table_array = ns_def.convert_excel_to_array(l3_table_ws_name, l3_table_file, 3)
        #print('--- l3_table_array ---')
        #print(l3_table_array)

        tmp_current_value1 = '_dummy_'
        tmp_current_value2 = '_dummy_'
        new_l3_table_array = []
        for tmp_l3_table_array in l3_table_array:
            if tmp_l3_table_array[1][0] == '':
                tmp_l3_table_array[1][0] = tmp_current_value1
            else:
                if tmp_l3_table_array[1][0] != tmp_current_value1:
                    tmp_current_value1 = tmp_l3_table_array[1][0]

            if tmp_l3_table_array[1][1] == '':
                tmp_l3_table_array[1][1] = tmp_current_value2
            else:
                if tmp_l3_table_array[1][1] != tmp_current_value2:
                    tmp_current_value2 = tmp_l3_table_array[1][1]

            new_l3_table_array.append(tmp_l3_table_array)
        sorted_new_l3_table_array = []

        for tmp_new_l3_table_array in new_l3_table_array:
            tmp_new_l3_table_array[1].extend(['','','','','','','',''])
            del tmp_new_l3_table_array[1][7:]
            # add if id for sort
            if_value = ns_def.get_if_value(tmp_new_l3_table_array[1][2])
            tmp_new_l3_table_array[1].extend([if_value])
            tmp_new_l3_table_array[1].extend([str(ns_def.split_portname(tmp_new_l3_table_array[1][2])[0])])

            #replace ' ' to '' in L3 Instance Name , IP Address
            tmp_new_l3_table_array[1][3] = tmp_new_l3_table_array[1][3].replace(' ','')
            tmp_new_l3_table_array[1][4] = tmp_new_l3_table_array[1][4].replace(' ','')

            tmp_ip_address_array = []
            work_new_l3_table_array = tmp_new_l3_table_array[1][4].split(',')
            for tmp_tmp_new_l3_table_array in work_new_l3_table_array:
                #check IP Addresses
                if ns_def.check_ip_format(tmp_tmp_new_l3_table_array) == 'IPv4':
                    tmp_ip_address_array.append([tmp_tmp_new_l3_table_array,ns_def.get_ipv4_value(tmp_tmp_new_l3_table_array)])
                    tmp_ip_address_array = sorted(tmp_ip_address_array , reverse=False, key=lambda x: (x[1]))  # sort ipv4 value

            #sort IP Addresses array
            #print('--- tmp_ip_address_array ---')
            #print(tmp_ip_address_array)

            change_tmp_ip_address_array = []
            for tmp_tmp_ip_address_array in tmp_ip_address_array:
                change_tmp_ip_address_array.append(tmp_tmp_ip_address_array[0])

            tmp_new_l3_table_array[1][4] = str(change_tmp_ip_address_array).replace('[','').replace(']','').replace('\'','').replace(' ','')
            sorted_new_l3_table_array.append(tmp_new_l3_table_array[1])

        #sort sorted_new_l3_table_array
        #print(sorted_new_l3_table_array)
        sorted_new_l3_table_array = sorted(sorted_new_l3_table_array , reverse=False, key=lambda x: (x[0], x[1], x[3], x[8], x[7], x[4]))  # sort l3 table

        tmp_row_num = 3
        last_l3_table_array = []
        last_l3_table_array.append([1, ['<<L3_TABLE>>']])
        last_l3_table_array.append([2, ['Area', 'Device Name', 'L3 IF Name','L3 Instance Name', 'IP Address / Subnet mask (Comma Separated)','[VPN] Target Device Name (Comma Separated)','[VPN] Target L3 Port Name (Comma Separated)']])
        for tmp_sorted_new_l3_table_array in sorted_new_l3_table_array:
            del tmp_sorted_new_l3_table_array[7:]
            last_l3_table_array.append([tmp_row_num,tmp_sorted_new_l3_table_array])
            tmp_row_num += 1

        #print('--- last_l3_table_array ---')
        #print(last_l3_table_array)

        last_l3_table_tuple = {}
        last_l3_table_tuple = ns_def.convert_array_to_tuple(last_l3_table_array)

        #delete L3 Table sheet
        ns_def.remove_excel_sheet(excel_maseter_file, excel_master_ws_name_l3)
        #create L3 Table sheet
        ns_def.create_excel_sheet(excel_maseter_file, excel_master_ws_name_l3)
        #write tuple to excel master data
        ns_def.write_excel_meta(last_l3_table_tuple, excel_maseter_file, excel_master_ws_name_l3, '_template_', 0, 0)

        input_excel_l3_table.close()

        return (True)
