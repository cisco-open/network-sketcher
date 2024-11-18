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
import sys, os, subprocess ,webbrowser , shutil
from pptx import Presentation
import ns_def,ns_ddx_figure, ns_egt_maker

class ns_modify_master_l3vpn():
    def __init__(self):
        print('--- ns_modify_master_l3vpn ---')

        # GET backup master file parameter
        # parameter
        ws_name = 'Master_Data'
        ws_l2_name = 'Master_Data_L2'
        ws_l3_name = 'Master_Data_L3'
        tmp_ws_name = '_tmp_'
        ppt_meta_file = str(self.excel_maseter_file_backup)
        print(self.excel_maseter_file_backup)

        # convert from master to array and convert to tuple
        #self.position_folder_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_FOLDER>>')
        #self.position_shape_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_SHAPE>>')
        #self.root_folder_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<ROOT_FOLDER>>')
        self.position_line_array = ns_def.convert_master_to_array(ws_name, ppt_meta_file, '<<POSITION_LINE>>')
        #self.position_folder_tuple = ns_def.convert_array_to_tuple(self.position_folder_array)
        #self.position_shape_tuple = ns_def.convert_array_to_tuple(self.position_shape_array)
        #self.root_folder_tuple = ns_def.convert_array_to_tuple(self.root_folder_array)
        #self.root_folder_tuple = ns_def.convert_array_to_tuple(self.root_folder_array)
        self.position_line_tuple = ns_def.convert_array_to_tuple(self.position_line_array)

        self.l2_table_array = ns_def.convert_master_to_array(ws_l2_name, ppt_meta_file, '<<L2_TABLE>>')
        self.l2_table_array_tuple = ns_def.convert_array_to_tuple(self.l2_table_array)
        self.l3_table_array = ns_def.convert_master_to_array(ws_l3_name, ppt_meta_file, '<<L3_TABLE>>')

        # print('---- self.position_folder_tuple ----')
        # print(self.position_folder_tuple)
        # print('---- self.position_folder_array ----')
        # print(self.position_folder_array)
        # print('---- self.position_shape_tuple ----')
        # print(self.position_shape_tuple)
        # print('---- self.position_shape_array ----')
        # print(self.position_shape_array)

        #print(self.l2_table_array)
        #print(self.l3_table_array)

        ### GET candidate VPN IF in Maseter_data_L2
        l2candidate_list = []
        l2candidate_array = []
        for row in self.l2_table_array:
            while len(row[1]) < 9:
                row[1].append('')

            values = row[1]
            if len(values) > 7 and values[5] != '' and values[2] == '' and values[3] == '' and values[6] == '' and values[7] == '':
                l2candidate_list.append([values[1],values[5]])
                l2candidate_array.append(row)

        #print(l2candidate_list)
        #print(l2candidate_array)

        ### GET candidate VPN IF in Maseter_data_L3
        l3candidate_array = []
        update_original_list = []
        for row in self.l3_table_array:
            while len(row[1]) < 9:
                row[1].append('')

            values = row[1]
            #print([values[1],values[2],values[5],values[6]])

            original_list = [values[1],values[2],values[5],values[6]]
            # Split the third and fourth elements by commas
            devices = original_list[2].split(',')
            vpns = original_list[3].split(',')

            paired_lists = []
            for i in range(len(devices)):
                # Create a new list for each pair and append it to paired_lists
                paired_lists.append([values[1],values[2],devices[i], vpns[i]])

                if len(values) > 7 and [values[1],values[2]] in l2candidate_list and  [devices[i],vpns[i]] in l2candidate_list:
                    l3candidate_array.append([values[1],values[2],devices[i],vpns[i]])

        #print(l3candidate_array)

        ### make meta data for L1 Master_Data sheet
        #print(self.position_line_array )
        max_value = max(item[0] for item in self.position_line_array)
        add_value = max_value

        add_vpn_position_line_array = self.position_line_array

        for tmp_l3candidate_array in l3candidate_array:
            add_value += 1
            ifname1 = ns_def.adjust_portname(tmp_l3candidate_array[1])
            ifname2 = ns_def.adjust_portname(tmp_l3candidate_array[3])
            add_vpn_position_line_array.append([add_value, [tmp_l3candidate_array[0], tmp_l3candidate_array[2], str(ifname1[0] + ' ' + ifname1[2]), str(ifname2[0] + ' ' + ifname2[2]), '', '', '', '', '', '', '', '', ifname1[1], 'Auto', 'Auto', '1000BASE-T', ifname2[1], 'Auto', 'Auto', '1000BASE-T']])

        #print(add_vpn_position_line_array)
        self.add_vpn_position_line_tuple = ns_def.convert_array_to_tuple(add_vpn_position_line_array)

        ### write L1 Master data
        ns_def.clear_section_sheet(ws_name, ppt_meta_file, self.position_line_tuple)
        ns_def.write_excel_meta(self.add_vpn_position_line_tuple, ppt_meta_file, ws_name, '<<POSITION_LINE>>',0, 0)

        ### make meta data for L2 Master_Data sheet
        update_l2_table_array = []

        l2candidate_array = []
        for item in l3candidate_array:
            l2candidate_array.append([item[0], item[1]])
            l2candidate_array.append([item[2], item[3]])

        #print(l2candidate_array)
        self.vpn_hostname_if_list = l2candidate_array

        for tmp_l2_table_array in self.l2_table_array:
            if len(tmp_l2_table_array[1]) >= 6:
                if [tmp_l2_table_array[1][1],tmp_l2_table_array[1][5]] in l2candidate_array:
                    #print([tmp_l2_table_array[1][1],tmp_l2_table_array[1][5]])
                    update_l2_table_array.append([tmp_l2_table_array[0], [tmp_l2_table_array[1][0], tmp_l2_table_array[1][1], '', tmp_l2_table_array[1][5], '', '', '', '', '']])

                else:
                    update_l2_table_array.append(tmp_l2_table_array)

        #print(update_l2_table_array)
        self.update_l2_table_array_tuple = ns_def.convert_array_to_tuple(update_l2_table_array)

        ### write L2 Master data
        ns_def.clear_section_sheet(ws_l2_name , ppt_meta_file, self.l2_table_array_tuple)
        ns_def.write_excel_meta(self.update_l2_table_array_tuple, ppt_meta_file, ws_l2_name , '<<L2_TABLE>>',0, 0)


class  ns_write_vpns_on_l1():
    def __init__(self):
        #parameter
        ws_l3_name = 'Master_Data_L3'
        excel_maseter_file = self.inFileTxt_L3_3_1.get()

        print('--- ns_vpns_on_l1_create ---')
        #### GET self.l3_table_array ####
        self.l3_table_array = ns_def.convert_master_to_array(ws_l3_name, excel_maseter_file, '<<L3_TABLE>>')
        #print('--- self.l3_table_array  ---')
        #print(self.l3_table_array )

        ### update self.l3_table_array for Comma Separated
        self.update_l3_table_array = []

        for tmp_l3_table_array in self.l3_table_array:
            if tmp_l3_table_array[0] != 1 and tmp_l3_table_array[0] != 2:
                if len(tmp_l3_table_array[1]) == 7 and tmp_l3_table_array[1][5] != '' and tmp_l3_table_array[1][6] != '':
                    dummy_5 = str(tmp_l3_table_array[1][5]).split(',')
                    dummy_6 = str(tmp_l3_table_array[1][6]).split(',')
                    if len(dummy_5) >= 2 and len(dummy_5) == len(dummy_6):
                        #print(dummy_5, dummy_6)
                        for i,name in enumerate(dummy_5):
                            self.update_l3_table_array.append([tmp_l3_table_array[0], [tmp_l3_table_array[1][0], tmp_l3_table_array[1][1], tmp_l3_table_array[1][2], tmp_l3_table_array[1][3], tmp_l3_table_array[1][4], str(dummy_5[i]).strip(' '), str(dummy_6[i]).strip(' ')]])
                    else:
                        self.update_l3_table_array.append([tmp_l3_table_array[0], [tmp_l3_table_array[1][0], tmp_l3_table_array[1][1], tmp_l3_table_array[1][2], tmp_l3_table_array[1][3], tmp_l3_table_array[1][4], str(tmp_l3_table_array[1][5]).strip(' '), str(tmp_l3_table_array[1][6]).strip(' ')]])
                else:
                    self.update_l3_table_array.append(tmp_l3_table_array)
            else:
                self.update_l3_table_array.append(tmp_l3_table_array)

        print('--- self.update_l3_table_array ---')
        #print(self.update_l3_table_array)

        ### make a vpn table ###
        self.l3_vpn_table_array = []
        for tmp_l3_table_array  in self.update_l3_table_array:
            if tmp_l3_table_array[0] != 1 and tmp_l3_table_array[0] != 2:
                tmp_tmp_l3_vpn_table_array = self.update_l3_table_array
                if len(tmp_l3_table_array[1]) == 7 and tmp_l3_table_array[1][5] != '' and tmp_l3_table_array[1][6] != '':
                    ### check target vpn exists ###
                    for tmp_tmp_tmp_l3_vpn_table_array in tmp_tmp_l3_vpn_table_array:
                        if tmp_tmp_tmp_l3_vpn_table_array[0] != 1 and tmp_tmp_tmp_l3_vpn_table_array[0] != 2:
                            if tmp_tmp_tmp_l3_vpn_table_array[1][1] == tmp_l3_table_array[1][5] and tmp_tmp_tmp_l3_vpn_table_array[1][2] == tmp_l3_table_array[1][6]:
                                self.l3_vpn_table_array.append([tmp_l3_table_array[1][1],tmp_l3_table_array[1][2],tmp_l3_table_array[1][5],tmp_l3_table_array[1][6]])

        #print(self.l3_vpn_table_array)
        if len(self.l3_vpn_table_array) == 0:
            print('vpn count ---> 0')
            return

        self.shape = Presentation(self.output_ppt_file)

        ### get GREEN color's shape list from self.position_style_shape_array
        green_color_shape_list = []
        for tmp_position_style_shape_array in self.position_style_shape_array:
            if tmp_position_style_shape_array[0] != 1 and tmp_position_style_shape_array[1][4] == 'GREEN':
                green_color_shape_list.append(tmp_position_style_shape_array[1][0])
        print('--- green_color_shape_list ---')
        #print(green_color_shape_list)

        ### get vpn info ###
        for tmp_l3_vpn_table_array in self.l3_vpn_table_array:
            for from_coord_list in self.coord_list:
                if from_coord_list[0] == tmp_l3_vpn_table_array[0]:
                    for to_coord_list in self.coord_list:
                        if to_coord_list[0] == tmp_l3_vpn_table_array[2]:
                            #print('---from_coord_list---')
                            #print(from_coord_list,to_coord_list)

                            #### defalut values ####
                            line_type = 'VPN'
                            inche_from_connect_x = from_coord_list[2]
                            inche_from_connect_y = from_coord_list[5]
                            inche_to_connect_x = to_coord_list[2]
                            inche_to_connect_y = to_coord_list[5]
                            flag_vertical_overlap = True
                            flag_horizontal_overlap = True
                            flag_line_curve = False
                            buffer_updown_overlap = 0.5 # inches
                            hight_vpn_curve = 0.4  # inches

                            ### set shape's up/down position ###
                            if from_coord_list[6] + buffer_updown_overlap < to_coord_list[4]:
                                inche_from_connect_y = from_coord_list[6]
                                inche_to_connect_y = to_coord_list[4]
                                flag_vertical_overlap = False
                            elif from_coord_list[4] - buffer_updown_overlap > to_coord_list[6]:
                                inche_from_connect_y = from_coord_list[4]
                                inche_to_connect_y = to_coord_list[6]
                                flag_vertical_overlap = False

                            ### set shape's right/left position ###
                            if flag_vertical_overlap == False:
                                if from_coord_list[3] < to_coord_list[1]:
                                    inche_from_connect_x = from_coord_list[3] - ((from_coord_list[3] - from_coord_list[2]) / 2)
                                    inche_to_connect_x = to_coord_list[1] + ((to_coord_list[3] - to_coord_list[2]) / 2)
                                    flag_horizontal_overlap = False
                                elif from_coord_list[1] > to_coord_list[3]:
                                    inche_from_connect_x = from_coord_list[1] + ((from_coord_list[3] - from_coord_list[2]) / 2)
                                    inche_to_connect_x = to_coord_list[3] - ((to_coord_list[3] - to_coord_list[2]) / 2)
                                    flag_horizontal_overlap = False
                            elif flag_vertical_overlap == True:
                                ### check to whatever line should be curve ###
                                for sandwich_coord_list in self.coord_list:
                                    if (sandwich_coord_list[5] >  from_coord_list[4] - (buffer_updown_overlap / 2 ) and sandwich_coord_list[5] < to_coord_list[6] + (buffer_updown_overlap / 2 ) and sandwich_coord_list[2] >  from_coord_list[2] and sandwich_coord_list[2] < to_coord_list[2]) or \
                                        (sandwich_coord_list[5] >  to_coord_list[4] - (buffer_updown_overlap / 2 ) and sandwich_coord_list[5] < from_coord_list[6] + (buffer_updown_overlap / 2 ) and sandwich_coord_list[2] >  to_coord_list[2] and sandwich_coord_list[2] < from_coord_list[2]):
                                        if sandwich_coord_list[0] in green_color_shape_list:
                                            #print('--- sandwich_coord_list ---',sandwich_coord_list)
                                            #print(from_coord_list, to_coord_list)
                                            flag_line_curve = True

                                if from_coord_list[3] < to_coord_list[1]:
                                    inche_from_connect_x = from_coord_list[3]
                                    inche_to_connect_x = to_coord_list[1]
                                    flag_horizontal_overlap = False
                                elif from_coord_list[1] > to_coord_list[3]:
                                    inche_from_connect_x = from_coord_list[1]
                                    inche_to_connect_x = to_coord_list[3]
                                    flag_horizontal_overlap = False

                            ### write vpns ###
                            if flag_line_curve == False:
                                ns_ddx_figure.extended.add_line(self,line_type,inche_from_connect_x,inche_from_connect_y,inche_to_connect_x,inche_to_connect_y)
                            elif flag_line_curve == True:
                                line_type = 'VPN_curve'
                                ns_ddx_figure.extended.add_line(self, line_type, inche_from_connect_x, inche_from_connect_y, abs((inche_to_connect_x - inche_from_connect_x) / 2) + inche_from_connect_x, inche_to_connect_y - hight_vpn_curve)
                                ns_ddx_figure.extended.add_line(self, line_type, inche_to_connect_x,inche_to_connect_y, abs((inche_to_connect_x - inche_from_connect_x) / 2) + inche_from_connect_x, inche_to_connect_y - hight_vpn_curve)


        self.active_ppt.save(self.output_ppt_file)