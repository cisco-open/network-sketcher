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
from collections import Counter
import tkinter as tk ,tkinter.ttk
import ipaddress

class  auto_ip_addressing():
    def get_folder_list(self):
        print('--- get_folder_list ---')
        #parameter
        ws_name = 'Master_Data'
        excel_maseter_file = self.inFileTxt_L2_3_1.get()

        # GET Folder and wp name List
        self.folder_wp_name_array = ns_def.get_folder_wp_array_from_master(ws_name, excel_maseter_file)
        #print('---- folder_wp_name_array ----')
        #print(self.folder_wp_name_array)

        return_array = self.folder_wp_name_array[0]
        if len(self.folder_wp_name_array[1]) >= 1:
            return_array.append("_WAN(Way_Point)_")
        return return_array

    def get_auto_ip_param(self,target_area_name):
        print('--- get_auto_ip_param ---')
        #print(target_area_name)

        if target_area_name == "_WAN(Way_Point)_":
            target_area_name = 'N/A'

        '''get values of Master Data'''
        #parameter
        ws_name = 'Master_Data'
        ws_l2_name = 'Master_Data_L2'
        ws_l3_name = 'Master_Data_L3'
        excel_maseter_file = self.inFileTxt_L3_3_1.get()

        self.result_get_l2_broadcast_domains =  ns_def.get_l2_broadcast_domains.run(self,excel_maseter_file)  ## 'self.update_l2_table_array, device_l2_boradcast_domain_array, device_l2_directly_l3vport_array, device_l2_other_array, marged_l2_broadcast_group_array'

        #print('--- self.update_l2_table_array ---')
        #print(self.result_get_l2_broadcast_domains[0])
        #print('--- self.device_l2_boradcast_domain_array ---')
        #print(self.result_get_l2_broadcast_domains[1])
        #self.device_l2_boradcast_domain_array = self.result_get_l2_broadcast_domains[1]
        #print('--- device_l2_directly_l3vport_array ---')
        #print(self.result_get_l2_broadcast_domains[2])
        #self.device_l2_directly_l3vport_array = self.result_get_l2_broadcast_domains[2]
        #print('--- device_l2_other_array ---')
        #print(self.result_get_l2_broadcast_domains[3])
        self.device_l2_other_array = self.result_get_l2_broadcast_domains[3]
        #print('--- marged_l2_broadcast_group_array ---')
        #print(self.result_get_l2_broadcast_domains[4])
        self.marged_l2_broadcast_group_array = self.result_get_l2_broadcast_domains[4]
        #print('--- self.target_l2_broadcast_group_array ---')
        #print(self.target_l2_broadcast_group_array)

        self.l3_table_array = ns_def.convert_master_to_array(ws_l3_name, excel_maseter_file, '<<L3_TABLE>>')
        #print('--- self.l3_table_array  ---')
        #print(self.l3_table_array )

        # check ip address exists in target area
        flag_no_ipaddress = True
        add_l3_table_array = []
        for index, tmp_l3_table_array  in enumerate(self.l3_table_array):
            str(tmp_l3_table_array).replace(' ', '')
            if index >= 2:
                if tmp_l3_table_array[1][0] == target_area_name:
                    #print(tmp_l3_table_array[1])
                    if len(tmp_l3_table_array[1]) == 5:
                        flag_no_ipaddress = False

                if len(tmp_l3_table_array[1]) == 5:
                    if ',' in str(tmp_l3_table_array[1][4]):
                        #print('--- tmp_l3_table_array ', str(tmp_l3_table_array))
                        tmp_tmp_l3_table_array= str(tmp_l3_table_array[1][4]).split(',')
                        for tmp_add_array in tmp_tmp_l3_table_array:
                            tmp_tmp_tmp_l3_table_array = tmp_l3_table_array
                            tmp_tmp_tmp_l3_table_array[1][4] = tmp_add_array
                            #print('--- tmp_tmp_tmp_l3_table_array ', tmp_tmp_tmp_l3_table_array)
                            self.l3_table_array.append([tmp_tmp_tmp_l3_table_array[0],[tmp_tmp_tmp_l3_table_array[1][0],tmp_tmp_tmp_l3_table_array[1][1],tmp_tmp_tmp_l3_table_array[1][2],tmp_tmp_tmp_l3_table_array[1][3],tmp_tmp_tmp_l3_table_array[1][4]]])

        #print ('flag_no_ipaddress', str(flag_no_ipaddress))

        outside_ip_address_list = []
        inside_ip_address_list = []
        full_ip_address_list = []
        for index, tmp_l3_table_array in enumerate(self.l3_table_array):
            if index >= 2:
                if len(tmp_l3_table_array[1]) == 5 and ',' not in str(tmp_l3_table_array[1][4]):
                    full_ip_address_list.append(str(tmp_l3_table_array[1][4]).replace(' ', ''))

                if tmp_l3_table_array[1][0] != target_area_name:
                    if len(tmp_l3_table_array[1]) == 5 and ',' not in str(tmp_l3_table_array[1][4]):
                        # print(tmp_l3_table_array[1][4])
                        first_octet = int(str(tmp_l3_table_array[1][4]).split('.')[0])
                        second_octet = int(str(tmp_l3_table_array[1][4]).split('.')[1])
                        third_octet = int(str(tmp_l3_table_array[1][4]).split('.')[2])

                        outside_ip_address_list.append(str(first_octet) + '.' + str(second_octet)+ '.' + str(third_octet) + '.0')
                else:
                    if len(tmp_l3_table_array[1]) == 5 and ',' not in str(tmp_l3_table_array[1][4]):
                        # print(tmp_l3_table_array[1][4])
                        first_octet = int(str(tmp_l3_table_array[1][4]).split('.')[0])
                        second_octet = int(str(tmp_l3_table_array[1][4]).split('.')[1])
                        third_octet = int(str(tmp_l3_table_array[1][4]).split('.')[2])

                        inside_ip_address_list.append(str(first_octet) + '.' + str(second_octet) + '.' + str(third_octet) + '.0')

        if flag_no_ipaddress == True:
            current_ip_address_list = outside_ip_address_list
        else:
            current_ip_address_list = inside_ip_address_list

        #print(current_ip_address_list )
        if current_ip_address_list != []:
            word_counts = Counter(current_ip_address_list)
            most_common_word, most_common_count = word_counts.most_common(1)[0]
            print(f"--- most_common_word: {most_common_word} (most_common_count: {most_common_count})")
        else:
            most_common_word = '10.0.0.0'

        use_network = ''
        '''get starting ip address'''
        # set starting ip address
        start_ip = ipaddress.IPv4Address(most_common_word)

        # count for 192 , 172 , 10
        if most_common_word.startswith('192.168.'):
            increase_count = 256
        elif most_common_word.startswith('172.'):
            for i in range(16, 32):
                if most_common_word.startswith('172.' + str(i) + '.'):
                    increase_count = 256 * 16
        elif most_common_word.startswith('10.'):
            increase_count = 256 * 256
        else:
            start_ip = ipaddress.IPv4Address('10.0.0.0')
            increase_count = 256 * 256

        # output start network address(CIDR)
        flag_1st_third_octet = True

        for _ in range(increase_count):
            # Convert IP address to byte array
            ip_bytes = start_ip.packed
            # Get the third octet and increase by 1
            third_octet = ip_bytes[2] + 1
            # If the third octet exceeds 255, the second octet is also increased
            if third_octet > 255:
                second_octet = ip_bytes[1] + 1
                third_octet = 0  # Omitted if the second octet exceeds 255
            else:
                second_octet = ip_bytes[1]

            if flag_1st_third_octet == True:
                third_octet -= 1
                flag_1st_third_octet = False

            # Build a new IP address
            new_ip_bytes = bytearray(ip_bytes)
            new_ip_bytes[1] = second_octet
            new_ip_bytes[2] = third_octet
            start_ip = ipaddress.IPv4Address(bytes(new_ip_bytes))
            #print(start_ip)

            if len(full_ip_address_list) == 0:
                use_network = '10.0.0.0/24'
                break

            flag_found_use_network = True
            for tmp_full_ip_address_list in full_ip_address_list:

                # Define a network to determine
                network1 = ipaddress.ip_network(str(start_ip) +'/24')
                network2 = ipaddress.ip_network(str(tmp_full_ip_address_list), strict=False)

                # Determine if network1 and network2 overlap
                if network1.overlaps(network2):
                    #print(f"{network1} overlaps with {network2}")
                    flag_found_use_network = False
                    break
                else:
                    #print(f"{network1} does not overlap with {network2}")
                    use_network = network1

            if flag_found_use_network == True:
                break

        if flag_no_ipaddress == False:
            use_network = str(most_common_word) + str('/24')
        #print(use_network)

        # set the value to GUI entry
        self.sub3_4_3_entry_1.delete(0, tkinter.END)
        self.sub3_4_3_entry_1.insert(0, use_network)

    def run_auto_ip(self,target_area_name):
        print('--- run_auto_ip ---')
        l3_segment_group_array = ns_def.get_l3_segments(self)
        #print(l3_segment_group_array)

        '''Create existing IP address list'''
        exist_ip_list = []
        for tmp_l3_segment_group_array in l3_segment_group_array:
            for tmp_tmp_l3_segment_group_array in tmp_l3_segment_group_array:
                if tmp_tmp_l3_segment_group_array[4] != '':
                    exist_ip_list.append(tmp_tmp_l3_segment_group_array[4])
        print('--- exist_ip_list ---')
        #print(exist_ip_list)

        # Set to store unique networks without duplicates
        unique_networks = set()

        # Extract networks from each CIDR notation and add them to the set
        for cidr in exist_ip_list:
            # Remove any whitespace (if present in the input)
            cidr = cidr.strip()
            # Create an IPv4Network object
            network = ipaddress.IPv4Network(cidr, strict=False)
            # Add the network to the set
            unique_networks.add(network)

        # Print the list of unique network addresses in CIDR notation
        print("--- Unique network addresses in CIDR notation: ---")
        #for network in sorted(unique_networks):
        #    print(network.with_prefixlen)

        '''calc ip address'''
        ip_assigned_l3_segment_group_array = []
        ip_address_exists_array = []
        ip_address_exists_array_sub = []

        for tmp_l3_segment_group_array in l3_segment_group_array:
            flag_way_point_exists = False
            for tmp_tmp_l3_segment_group_array in tmp_l3_segment_group_array:
                if tmp_tmp_l3_segment_group_array[0] == 'N/A':
                    flag_way_point_exists = True

            if (tmp_l3_segment_group_array[0][0] == target_area_name and flag_way_point_exists == False) or (target_area_name == 'N/A' and flag_way_point_exists == True):
                #check to ip address exists
                flag_ip_address_exists = False

                for tmp_tmp_l3_segment_group_array in tmp_l3_segment_group_array:
                    if tmp_tmp_l3_segment_group_array[4] != '':
                        ip_address_exists = tmp_tmp_l3_segment_group_array[4]
                        ip_address_exists_array.append(tmp_tmp_l3_segment_group_array)
                        ip_address_exists_array_sub.append(tmp_tmp_l3_segment_group_array[4])
                        ip_address_exists_subnet = ip_address_exists.split('/')
                        flag_ip_address_exists = True

                #calc ip address subnet
                required_ip_num = len(tmp_l3_segment_group_array) + int(self.sub3_4_2_entry_1.get())
                # Add 2 for network address and broadcast address
                required_ip_num += 2
                # Find the number of bits in the subnet mask
                subnet_mask_bits = 32
                while True:
                    if 2 ** (32 - subnet_mask_bits) >= required_ip_num:
                        break
                    subnet_mask_bits -= 1

                # Display subnet mask in CIDR format
                cidr_notation = f"/{subnet_mask_bits}"
                #print(f"Required subnet mask in CIDR notation: {cidr_notation}")
                start_ip = str(self.sub3_4_3_entry_1.get()).split('/')[0]

                if flag_ip_address_exists == True:
                    network_exists = ipaddress.ip_network(ip_address_exists, strict=False)
                    start_ip = network_exists.network_address
                    cidr_notation = str('/') + str(ip_address_exists_subnet[1])

                '''check overlap'''
                # Check if a specific IP range overlaps with any of the unique networks
                #print(str(start_ip) + str(cidr_notation))
                specific_range = ipaddress.IPv4Network(str(start_ip) + str(cidr_notation))

                while True:
                    # Check for overlaps
                    overlap = any(specific_range.overlaps(network) for network in unique_networks)
                    if not overlap or flag_ip_address_exists == True:
                        break

                    # Calculate the broadcast address of the initial subnet
                    broadcast_address = specific_range.broadcast_address
                    # Calculate the first IP address of the next subnet by adding 1 to the broadcast address
                    next_subnet_network_address = broadcast_address + 1
                    # Create the next subnet based on the calculated network address
                    next_subnet = ipaddress.IPv4Network(f'{next_subnet_network_address}{str(cidr_notation)}')
                    specific_range = next_subnet
                #print(f'specific_range: {specific_range}')
                # Get all available hosts in the subnet
                subnet = ipaddress.IPv4Network(specific_range)
                available_hosts = list(subnet.hosts())

                ''' Descending order '''
                if str(self.combo3_4_4_1.get()) == "Descending order":
                    available_hosts = sorted(available_hosts, reverse=True)

                ''' Assign each host IP address from the subnet'''
                #for index, host in enumerate(available_hosts, start=1):
                #    print(f"Host {index}: {host}")

                ip_assign_num = 0
                pre_ip_assigned_l3_segment_group_array = []
                #print(ip_address_exists_array)


                for tmp_tmp_l3_segment_group_array in tmp_l3_segment_group_array:
                    if str(self.combo3_4_6_1.get()) == "Reassign within the same subnet":
                        #print("Reassign within the same subnet")
                        pre_ip_assigned_l3_segment_group_array.append([tmp_tmp_l3_segment_group_array[0],tmp_tmp_l3_segment_group_array[1],tmp_tmp_l3_segment_group_array[2],tmp_tmp_l3_segment_group_array[3],str(available_hosts[ip_assign_num]) + str(cidr_notation)])
                        ip_assign_num += 1
                    elif str(self.combo3_4_6_1.get()) == "Keep existing IP address":
                        #print("Keep existing IP address")
                        while True:
                            if tmp_tmp_l3_segment_group_array in ip_address_exists_array:
                                pre_ip_assigned_l3_segment_group_array.append(tmp_tmp_l3_segment_group_array)
                                ip_assign_num += 1
                                break
                            elif str(available_hosts[ip_assign_num]) + str(cidr_notation) in ip_address_exists_array_sub:
                                ip_assign_num += 1
                            else:
                                pre_ip_assigned_l3_segment_group_array.append([tmp_tmp_l3_segment_group_array[0], tmp_tmp_l3_segment_group_array[1], tmp_tmp_l3_segment_group_array[2], tmp_tmp_l3_segment_group_array[3], str(available_hosts[ip_assign_num]) + str(cidr_notation)])
                                ip_assign_num += 1
                                break

                ip_assigned_l3_segment_group_array.append(pre_ip_assigned_l3_segment_group_array)
                #print(f"pre_ip_assigned_l3_segment_group_array: {pre_ip_assigned_l3_segment_group_array}")

                unique_networks.add(subnet)

        print('--- ip_assigned_l3_segment_group_array ---')
        #print(ip_assigned_l3_segment_group_array)

        '''
        Update to the Master file
        '''
        ws_l3_name = 'Master_Data_L3'
        excel_maseter_file = self.inFileTxt_L3_3_1.get()
        self.l3_table_array = ns_def.convert_master_to_array(ws_l3_name, excel_maseter_file, '<<L3_TABLE>>')

        updated_l3_table_array = []
        for tmp_l3_table_array in self.l3_table_array:
            #print(tmp_l3_table_array)
            if tmp_l3_table_array[1][0] != target_area_name and (tmp_l3_table_array[0] == 1 or tmp_l3_table_array[0] == 2):
                updated_l3_table_array.append(tmp_l3_table_array)
            else:
                flag_l3_if_match = False
                for tmp_ip_assigned_l3_segment_group_array in ip_assigned_l3_segment_group_array:
                    for tmp_tmp_tmp_ip_assigned_l3_segment_group_array in tmp_ip_assigned_l3_segment_group_array:

                        if tmp_tmp_tmp_ip_assigned_l3_segment_group_array[1] == tmp_l3_table_array[1][1]  and tmp_tmp_tmp_ip_assigned_l3_segment_group_array[2] == tmp_l3_table_array[1][2]:
                            flag_l3_if_match = True
                            if len(tmp_l3_table_array) != 5:
                                updated_l3_table_array.append([tmp_l3_table_array[0],tmp_tmp_tmp_ip_assigned_l3_segment_group_array])
                            else:
                                updated_l3_table_array.append(tmp_l3_table_array)
                if flag_l3_if_match == False:
                    updated_l3_table_array.append(tmp_l3_table_array)

        print('--- updated_l3_table_array ---')
        #print(updated_l3_table_array)

        '''
        write the Master file
        '''
        last_l3_table_tuple = {}
        last_l3_table_tuple = ns_def.convert_array_to_tuple(updated_l3_table_array)

        excel_master_ws_name_l3 = 'Master_Data_L3'
        # delete L3 Table sheet
        ns_def.remove_excel_sheet(excel_maseter_file, excel_master_ws_name_l3)
        # create L3 Table sheet
        ns_def.create_excel_sheet(excel_maseter_file, excel_master_ws_name_l3)
        # write tuple to excel master data
        ns_def.write_excel_meta(last_l3_table_tuple, excel_maseter_file, excel_master_ws_name_l3, '_template_', 0, 0)








