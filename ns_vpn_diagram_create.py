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
from pptx import Presentation
import ns_def,ns_ddx_figure

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