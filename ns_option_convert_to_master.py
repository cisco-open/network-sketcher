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
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.util import Inches,Cm,Pt
import sys, os, re, io
import numpy as np
import math
import ns_def
import openpyxl
import yaml
from ciscoconfparse import CiscoConfParse

class  ns_option_convert_to_master_svg():
    def __init__(self):
        from svg.path import parse_path
        from svg.path.path import Line
        from xml.dom import minidom

        #parameter
        tmp_ppt_width = 30  # inches
        tmp_ppt_hight = 15  # inches

        path_array_inches = []
        device_array = []

        ### read the SVG file
        doc = minidom.parse(str(self.inFileTxt_1a_1.get()))
        path_strings = [path.getAttribute('d') for path
                        in doc.getElementsByTagName('path')]
        svg_width = [svg.getAttribute('width') for svg
                        in doc.getElementsByTagName('svg')]
        svg_hight = [svg.getAttribute('height') for svg
                        in doc.getElementsByTagName('svg')]

        ### create device_array
        for g in doc.getElementsByTagName('g'):
            if g.getAttribute('transform') != '':
                if len(g.getElementsByTagName('text')) == 2:
                    device_array.append([str(g.getAttribute('transform')), str(g.getElementsByTagName('text')[0].firstChild.nodeValue) + str(g.getElementsByTagName('text')[1].firstChild.nodeValue)])
                elif len(g.getElementsByTagName('text')) == 1:
                    device_array.append( [str(g.getAttribute('transform')), str(g.getElementsByTagName('text')[0].firstChild.nodeValue)] )

        #print(device_array)

        doc.unlink()

        '''create ppt'''
        ppt = Presentation()
        ppt.slide_width = Inches(tmp_ppt_width)
        ppt.slide_height = Inches(tmp_ppt_hight)

        slide_layout_5 = ppt.slide_layouts[5]
        slide = ppt.slides.add_slide(slide_layout_5)

        shapes = slide.shapes

        ### adjust ratio of ppt size
        ppt_ratio_x = float(tmp_ppt_width) / float(svg_width[0])
        ppt_ratio_y = float(tmp_ppt_hight)/ float(svg_hight[0])

        #print(svg_width[0], svg_hight[0], ppt_ratio_x , ppt_ratio_y)

        '''create the patch array'''
        for path_string in path_strings:
            if path_string != '':   # update for python 3.10
                path = parse_path(path_string)
                start_x = 0
                start_y = 0
                end_x = 0
                end_y = 0
                start_bit = True
                if 'Path()' not in str(path):
                    for e in path:
                        if isinstance(e, Line):
                            if start_bit == True:
                                start_x = e.start.real
                                start_y = e.start.imag
                                start_bit = False
                            end_x = e.end.real
                            end_y = e.end.imag

                    path_array_inches.append([start_x * ppt_ratio_x, start_y * ppt_ratio_y, end_x * ppt_ratio_x, end_y * ppt_ratio_y])
            #print(path_array_inches)

        '''add connectors'''
        for tmp_path_array_inches in path_array_inches:
            shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(tmp_path_array_inches[0]), Inches(tmp_path_array_inches[1]), \
                             Inches(tmp_path_array_inches[2]), Inches(tmp_path_array_inches[3]))

        '''add shapes'''
        for tmp_device_array in device_array:
            tmp_array = str(tmp_device_array[0]).split()
            print(tmp_array)
            tmp_1st_array = []
            tmp_3rd_array = []
            tmp_1st_array = eval(str((re.findall(r'\((.+)\)', tmp_array[0]))).replace('\'', ''))
            tmp_3rd_array = eval(str((re.findall(r'\((.+)\)', tmp_array[2]))).replace('\'', ''))

            shape_left = tmp_1st_array[0] * ppt_ratio_x
            shape_top = tmp_1st_array[1] * ppt_ratio_y
            shape_width = abs(tmp_3rd_array[0] * ppt_ratio_x)
            shape_hight = abs(tmp_3rd_array[1] * ppt_ratio_y)
            shapes = slide.shapes
            shapes = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(shape_left), Inches(shape_top), Inches(shape_width), Inches(shape_hight))
            shapes.adjustments[0] = 0.0
            shapes.text = tmp_device_array[1]

        '''add root folder'''
        shapes = slide.shapes
        shapes = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0), Inches(0), Inches(tmp_ppt_width), Inches(tmp_ppt_hight))
        shapes.adjustments[0] = 0.0
        shapes.text = '_tmp_'

        ppt.save("./_tmp_tmp_tmp_.pptx")


class  ns_option_convert_to_master_yaml():
    def __init__(self):

        #parameter
        tmp_ppt_width = 30  # inches
        tmp_ppt_hight = 15  # inches

        path_array_inches = []
        device_array = []
        path_strings = []

        ### read the yaml file
        with open(str(self.full_filepath), 'r') as yml:
            config = yaml.safe_load(yml)
            #print(config)

        tmp_min_x = 999999
        tmp_min_y = 999999
        tmp_max_x = 0
        tmp_max_y = 0
        for tmp_array in config['nodes']:
            #print(tmp_array['label'],tmp_array['x'],tmp_array['y'],tmp_array['id'])
            if tmp_array['x'] < tmp_min_x:
                tmp_min_x = tmp_array['x']
            if tmp_array['y'] < tmp_min_y:
                tmp_min_y = tmp_array['y']

        for tmp_array in config['nodes']:
            #print(tmp_array['label'],tmp_array['x'],tmp_array['y'])
            device_array.append([tmp_array['label'], int (tmp_array['x'] - tmp_min_x ), int(tmp_array['y'] - tmp_min_y),tmp_array['id']])
            if int (tmp_array['x'] - tmp_min_x ) > tmp_max_x:
                tmp_max_x = int (tmp_array['x'] - tmp_min_x )
            if int(tmp_array['y'] - tmp_min_y) > tmp_max_y:
                tmp_max_y = int(tmp_array['y'] - tmp_min_y)

        #print(device_array)

        for tmp_array in config['links']:
            for i in device_array:
                if i[3] == tmp_array['n1']:
                    tmp_tmp_label_start = i[0]
                if i[3] == tmp_array['n2']:
                    tmp_tmp_label_end = i[0]

            path_strings.append([tmp_array['id'],tmp_array['i1'],tmp_array['i2'],tmp_tmp_label_start,tmp_tmp_label_end])
        #print(path_strings)

        '''create ppt'''
        ppt = Presentation()
        ppt.slide_width = Inches(tmp_ppt_width)
        ppt.slide_height = Inches(tmp_ppt_hight)

        slide_layout_5 = ppt.slide_layouts[5]
        slide = ppt.slides.add_slide(slide_layout_5)

        shapes = slide.shapes

        ### adjust ratio of ppt size
        ppt_ratio_x = float(tmp_ppt_width) / float(tmp_max_x)
        ppt_ratio_y = float(tmp_ppt_hight)/ float(tmp_max_y)

        '''create the path array'''
        for path_string in path_strings:
            start_x = 0
            start_y = 0
            end_x = 0
            end_y = 0

            for d in device_array:
                if path_string[3] == d[0]:
                    start_x = d[1]
                    start_y = d[2]

                if path_string[4] == d[0]:
                    end_x = d[1]
                    end_y = d[2]


            path_array_inches.append([start_x * ppt_ratio_x  * 0.5 + 0.5, start_y * ppt_ratio_y  * 0.5 + 0.5, end_x * ppt_ratio_x  * 0.5 + 0.5, end_y * ppt_ratio_y  * 0.5 + 0.5])
        #print(path_array_inches)

        '''add connectors'''
        for tmp_path_array_inches in path_array_inches:
            shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(tmp_path_array_inches[0]), Inches(tmp_path_array_inches[1]), \
                             Inches(tmp_path_array_inches[2]), Inches(tmp_path_array_inches[3]))

        '''add shapes'''
        for tmp_device_array in device_array:
            tmp_1st_array = []
            tmp_3rd_array = []
            tmp_1st_array.append([tmp_device_array[1],tmp_device_array[2]])
            tmp_3rd_array.append([5,30])
            #print(tmp_1st_array[0])
            shape_left = float(tmp_1st_array[0][0]) * ppt_ratio_x * 0.5 + 0.5
            shape_top = float(tmp_1st_array[0][1]) * ppt_ratio_y * 0.5 + 0.5
            #print(tmp_3rd_array[0] , ppt_ratio_x)
            shape_width = abs(tmp_3rd_array[0][0] * ppt_ratio_x)
            shape_hight = abs(tmp_3rd_array[0][1] * ppt_ratio_y)


            shapes = slide.shapes
            shapes = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(shape_left), Inches(shape_top), Inches(shape_width), Inches(shape_hight))
            shapes.adjustments[0] = 0.0
            shapes.text = tmp_device_array[0]

        '''add root folder'''
        shapes = slide.shapes
        shapes = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0), Inches(0), Inches(tmp_ppt_width), Inches(tmp_ppt_hight))
        shapes.adjustments[0] = 0.0
        shapes.text = config['lab']['title']

        ppt.save("./_tmp_tmp_tmp_.pptx")


class  ns_overwrite_line_to_master_yaml():
    def __init__(self):
        import ns_def
        path_strings = []
        device_array = []

        ### read the yaml file
        with open(str(self.full_filepath), 'r') as yml:
            config = yaml.safe_load(yml)
            #print(config)

        #get node data
        for tmp_array in config['nodes']:
            #print(tmp_array)
            device_array.append([tmp_array['label'],tmp_array['id'],tmp_array['interfaces']])
        #print(device_array)

        #get line data and map to path_strings array
        for tmp_array in config['links']:
            for i in device_array:
                if i[1] == tmp_array['n1']:
                    tmp_tmp_label_start = i[0]

                    #get interface name
                    for b in i[2]:
                        if b['id'] == tmp_array['i1']:
                            tmp_tmp_int_start = b['label']

                if i[1] == tmp_array['n2']:
                    tmp_tmp_label_end = i[0]

                    #get interface name
                    for b in i[2]:
                        if b['id'] == tmp_array['i2']:
                            tmp_tmp_int_end = b['label']

            path_strings.append([tmp_array['id'],tmp_tmp_int_start,tmp_tmp_int_end,tmp_tmp_label_start,tmp_tmp_label_end])
        #print(path_strings)

        #adjust CML's interface format to NS
        new_line_array = []
        tmp_i = 1
        for tmp_path in path_strings:
            #get string part and number part from path_strings
            new_line_array.append([ns_def.adjust_portname(tmp_path[1]),ns_def.adjust_portname(tmp_path[2]),tmp_path[3],tmp_path[4],tmp_i ])
            tmp_i+=1

        #print(new_line_array)

        '''overwrite Master Data File'''
        new_master_line_array = []
        converted_tuple = {}
        tmp_used_array = []
        ppt_meta_file = str(self.excel_file_path)
        master_line_array = ns_def.convert_master_to_array('Master_Data', ppt_meta_file, '<<POSITION_LINE>>')

        for tmp_master_line_array in master_line_array:
            if tmp_master_line_array[0] >= 3:
                #print (tmp_master_line_array)
                ### Change values at each line ###
                tmp_master_line_array[1][13] = 'Unknown'
                tmp_master_line_array[1][14] = 'Unknown'
                tmp_master_line_array[1][15] = 'Unknown'
                tmp_master_line_array[1][17] = 'Unknown'
                tmp_master_line_array[1][18] = 'Unknown'
                tmp_master_line_array[1][19] = 'Unknown'
                #print(tmp_master_line_array)
                for tmp_new_line_array in new_line_array:
                    if tmp_master_line_array[1][0] == tmp_new_line_array[2] and tmp_master_line_array[1][1] == tmp_new_line_array[3] and tmp_new_line_array[4] not in tmp_used_array:
                        tmp_master_line_array[1][2] = str(tmp_new_line_array[0][0] + ' ' + str(tmp_new_line_array[0][2]))
                        tmp_master_line_array[1][3] = str(tmp_new_line_array[1][0] + ' ' + str(tmp_new_line_array[1][2]))
                        tmp_master_line_array[1][12] = str(tmp_new_line_array[0][1])
                        tmp_master_line_array[1][16] = str(tmp_new_line_array[1][1])
                        tmp_used_array.append(tmp_new_line_array[4])
                        #print(tmp_used_array)
                        break
                #print(tmp_master_line_array)
                new_master_line_array.append(tmp_master_line_array)


        converted_tuple = ns_def.convert_array_to_tuple(new_master_line_array)
        ns_def.overwrite_excel_meta(converted_tuple, self.excel_file_path, 'Master_Data', '<<POSITION_LINE>>', 0,0)


class  ns_l3_config_to_master_yaml():
    def __init__(self):
        # parameter
        l3_table_ws_name = 'Master_Data_L3'
        l3_table_file = self.full_filepath
        target_node_definition_ios = ['iosv','csr1000v','iosvl2','cat8000v']
        target_node_definition_asa = ['asav']
        target_node_definition_iosxr = ['iosxrv9000']

        #get L3 Table Excel file
        l3_table_array = []
        l3_table_array = ns_def.convert_excel_to_array(l3_table_ws_name, l3_table_file, 3)
        print('--- l3_table_array ---')
        #print(l3_table_array)

        '''get L3 ipaddress from yaml'''
        ### read the yaml file
        with open(str(self.yaml_full_filepath), 'r') as yml:
            config = yaml.safe_load(yml)
        # print(config)

        config_array = []
        last_int_array = []
        overwrite_l3_table_array = []

        print('--- [label, node_definition, id, configuration] ---')
        for tmp_config in config['nodes']:
            config_array.append([tmp_config['label'], tmp_config['node_definition'], tmp_config['id'], tmp_config['configuration']])
            # print(tmp_config['configuration'])
            if tmp_config['node_definition'] in target_node_definition_ios or tmp_config['node_definition'] in target_node_definition_asa or tmp_config['node_definition'] in target_node_definition_iosxr:
                '''
                CiscoConfParse
                '''
                CONFIG = tmp_config['configuration']

                if tmp_config['node_definition'] in target_node_definition_ios:
                    parse = CiscoConfParse(CONFIG.splitlines(), syntax='ios', factory=True)
                elif tmp_config['node_definition'] in target_node_definition_asa:
                    parse = CiscoConfParse(CONFIG.splitlines(), syntax='asa', factory=True)
                elif tmp_config['node_definition'] in target_node_definition_iosxr:
                    parse = CiscoConfParse(CONFIG.splitlines(), syntax='iosxr', factory=True)


                int_array = [[tmp_config['label'], tmp_config['node_definition'], tmp_config['id']]]
                dummy_array = []

                for tmp_parse in parse.find_objects('^interface\s'):
                    int_char = list(str(tmp_parse.interface_object))
                    int_char_2 = str(tmp_parse.interface_object)

                    for i, tmp_char in enumerate(str(tmp_parse.interface_object)):
                        if re.fullmatch('[0-9]+', tmp_char):
                            # print(str(i),tmp_char)
                            int_char.insert(i, ' ')
                            int_char_2 = str("".join(int_char))
                            break

                    if str(tmp_parse.ipv4_addr) != '':
                        dummy_array.append([int_char_2, tmp_parse.ipv4_addr + '/' + str(tmp_parse.ipv4_masklength)])

                int_array.append(dummy_array)
                #print(int_array)
                last_int_array.append(int_array)

        for tmp_l3_table_array in l3_table_array:
            #print(tmp_l3_table_array[1])
            for tmp_last_int_array in last_int_array:
                if tmp_l3_table_array[1][1] == tmp_last_int_array[0][0]:
                    #print(tmp_last_int_array[0][0] , tmp_l3_table_array[1][1])
                    if tmp_last_int_array[0][0] == tmp_l3_table_array[1][1]:
                        flag_l3_exist = False
                        for tmp_tmp_last_int_array in tmp_last_int_array[1]:
                            if tmp_tmp_last_int_array[0] == tmp_l3_table_array[1][2]:
                                #print('--- L3 address match ---   ' + str(tmp_last_int_array) , str(tmp_l3_table_array))
                                if len(tmp_l3_table_array[1]) == 3:
                                    tmp_l3_table_array[1].append('')
                                tmp_l3_table_array[1].append(tmp_tmp_last_int_array[1])
                                overwrite_l3_table_array.append(tmp_l3_table_array)
                                flag_l3_exist = True
                        if flag_l3_exist == False:
                            overwrite_l3_table_array.append(tmp_l3_table_array)

        print('--- overwrite_l3_table_array ---')
        #print(overwrite_l3_table_array)


        # write to master file
        last_overwrite_l3_table_tuple = {}
        last_overwrite_l3_table_tuple = ns_def.convert_array_to_tuple(overwrite_l3_table_array)
        print('--- last_overwrite_l3_table_tuple ---')
        #print(last_overwrite_l3_table_tuple)

        master_excel_meta = last_overwrite_l3_table_tuple
        excel_file_path = self.full_filepath
        worksheet_name = l3_table_ws_name
        section_write_to = '<<L3_TABLE>>'
        offset_row = 2
        offset_column = 0
        ns_def.overwrite_excel_meta(master_excel_meta, excel_file_path, worksheet_name, section_write_to, offset_row, offset_column)




