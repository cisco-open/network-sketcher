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

        print('---ns_vpns_on_l1_create---')
        #### GET self.l3_table_array ####
        self.l3_table_array = ns_def.convert_master_to_array(ws_l3_name, excel_maseter_file, '<<L3_TABLE>>')
        #print('--- self.l3_table_array  ---')
        #print(self.l3_table_array )

        ### make a vpn table ###
        self.l3_vpn_table_array = []
        for tmp_l3_table_array  in self.l3_table_array:
            if tmp_l3_table_array[0] != 1 and tmp_l3_table_array[0] != 2:
                tmp_tmp_l3_vpn_table_array = self.l3_table_array
                if len(tmp_l3_table_array[1]) == 7 and tmp_l3_table_array[1][5] != '' and tmp_l3_table_array[1][6] != '':
                    ### check target vpn exists ###
                    for tmp_tmp_tmp_l3_vpn_table_array in tmp_tmp_l3_vpn_table_array:
                        if tmp_tmp_tmp_l3_vpn_table_array[0] != 1 and tmp_tmp_tmp_l3_vpn_table_array[0] != 2:
                            if tmp_tmp_tmp_l3_vpn_table_array[1][1] == tmp_l3_table_array[1][5] and tmp_tmp_tmp_l3_vpn_table_array[1][2] == tmp_l3_table_array[1][6]:
                                self.l3_vpn_table_array.append([tmp_l3_table_array[1][1],tmp_l3_table_array[1][2],tmp_l3_table_array[1][5],tmp_l3_table_array[1][6]])

        print(self.l3_vpn_table_array)
        if len(self.l3_vpn_table_array) == 0:
            print('vpn count ---> 0')
            return

        self.shape = Presentation(self.output_ppt_file)

        ### get vpn info ###
        for tmp_l3_vpn_table_array in self.l3_vpn_table_array:
            for tmp_coord_list in self.coord_list:
                if tmp_coord_list[0] == tmp_l3_vpn_table_array[0]:
                    for tmp_tmp_coord_list in self.coord_list:
                        if tmp_tmp_coord_list[0] == tmp_l3_vpn_table_array[2]:
                            print('---tmp_coord_list---')
                            print(tmp_coord_list,tmp_tmp_coord_list)

                            ### write vpns ###
                            line_type = 'VPN'
                            inche_from_connect_x = tmp_coord_list[2]
                            inche_from_connect_y = tmp_coord_list[5]
                            inche_to_connect_x = tmp_tmp_coord_list[2]
                            inche_to_connect_y = tmp_tmp_coord_list[5]
                            ns_ddx_figure.extended.add_line(self,line_type,inche_from_connect_x,inche_from_connect_y,inche_to_connect_x,inche_to_connect_y)

        self.active_ppt.save(self.output_ppt_file)