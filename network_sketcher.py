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
from tkinterdnd2 import *
import sys, os, subprocess ,webbrowser ,openpyxl
import ns_def,network_sketcher_dev,ns_sync_between_layers,ns_attribute_table_sync_master
import ns_extensions
import ns_vpn_diagram_create


class ns_front_run():
    '''
    Main Panel
    '''
    def __init__(self):
        #add cli flow at ver 2.3.1
        if len(sys.argv) > 1:
            import network_sketcher_cli
            self.argv_array = sys.argv[1:]
            network_sketcher_cli.ns_cli_run.__init__(self.argv_array)
            exit()

        self.click_value = ''
        self.click_value_2nd = ''
        self.click_value_3rd = ''
        self.click_value_VPN = ''
        self.root = TkinterDnD.Tk()
        self.root.title("Network Sketcher  ver 2.4.0")
        self.root.geometry("490x200+100+100")
        
        # Notebook
        nb = ttk.Notebook()

        # create Tabs
        tab_x1 = tk.Frame(nb)
        tab_x2 = tk.Frame(nb)

        nb.add(tab_x1, text='  Main Panel  ', padding=5)
        nb.add(tab_x2, text='  Help  ', padding=5)
        nb.pack(expand=1, fill='both')

        '''
        tab_x1  <<Main Panel>>
        '''
        self.main1_1 = tk.LabelFrame(tab_x1, text="    Please input a file for Network Sketcher     ", font=("", 14), height=1, background="#F2FDE3")
        self.main1_1.grid(row=0, column=1, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.main1_1_label_1 = tk.Label(self.main1_1, text="", background="#F2FDE3")
        self.main1_1_label_1 .grid(row=1, column=0, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.text = tk.StringVar()
        self.text.set("              drag and drop here (*.pptx;*.xlsx;*.yaml)")
        self.main1_1_label_4 = tk.Label(self.main1_1, textvariable=self.text, font=("", 10), background="#F2FDE3")
        self.main1_1_label_4.grid(row=4, column=1, columnspan=3, sticky='W', padx=5, pady=2)

        self.main1_1_label_6 = tk.Label(self.main1_1, text="", background="#F2FDE3")
        self.main1_1_label_6.grid(row=5, column=1, sticky='W', padx=5, pady=2)

        self.main1_1_entry_1 = tk.Entry(self.main1_1)
        self.main1_1_entry_1.grid(row=7, column=1, sticky="WE", pady=3, ipadx=90)
        self.main1_1_button_1 = tk.Button(self.main1_1, text="Browse ...", command=lambda: self.click_action_main1_1('self.main1_1_button_1'))
        self.main1_1_button_1.grid(row=7, column=2, sticky='W', padx=5, pady=2)
        self.main1_1_button_1 = tk.Button(self.main1_1, text="Submit", command=lambda: self.click_action_main1_1('self.main1_1_button_2'))
        self.main1_1_button_1.grid(row=7, column=3, sticky='W', padx=0, pady=2)

        #drag and drop
        self.entry_name_main1_1 = 'self.main1_1_entry_1'
        self.main1_1.drop_target_register(DND_FILES)
        self.main1_1.dnd_bind("<<Drop>>", self.drop_main1_1 ,self.entry_name_main1_1)

        ### Help
        #Help_1_label_1 = tk.Label(tab_x2, text="Version 2.1.0", background="#FFFFFF")
        #Help_1_label_1.grid(row=0, column=0, sticky='W', padx=5, pady=2)

        Help_1 = tk.LabelFrame(tab_x2, text="    Online User Guide     ", font=("", 14), height=1, background="#FFFFFF")
        Help_1.grid(row=1, column=0, columnspan=3, sticky='W', padx=5, pady=5, ipadx=10, ipady=5)

        Help_1_button_2 = tk.Button(Help_1, text="English", font=("", 14), command=lambda: self.click_action_main1_1('self.help_1_button_2'))
        Help_1_button_2.grid(row=1, column=1, sticky='W', padx=20, pady=2 , ipadx=15,ipady=0)

        Help_1_button_1 = tk.Button(Help_1, text="Japanese", font=("", 14), command=lambda: self.click_action_main1_1('self.help_1_button_1'))
        Help_1_button_1.grid(row=1, column=2, sticky='W', padx=20, pady=2 , ipadx=5 ,ipady=0)

        Help_1_1 = tk.Label(tab_x2, font=("", 10), text="Copyright 2023 Cisco Systems, Inc. and its affiliates  \n  SPDX-License-Identifier: Apache-2.0", background='#FFFFFF')
        Help_1_1.grid(column=0, row=3)

        # main loop
        self.root.mainloop()


    def drop_main1_1(self, event):
        if event:
            event.data = event.data.replace('{', '').replace('}', '')
            if event.data.endswith('.pptx') or event.data.endswith('.xlsx') or event.data.endswith('.yaml'):
                exec(self.entry_name_main1_1 + '.delete(0, tkinter.END)')
                exec(self.entry_name_main1_1 + '.insert(tk.END, event.data)')
                self.filename = os.path.basename(event.data)
                self.full_filepath = event.data
                self.text.set(self.filename)
                self.click_action_main1_1('self.main1_1_button_2')
            else:
                self.text.set('[ERROR] ' + 'Please input a file compatible with NS')
                self.main1_1_label_4 = tk.Label(self.main1_1, textvariable=self.text, font=("", 10), background="#FBE5D6")
                self.main1_1_label_4.grid(row=4, column=1, columnspan=7, sticky='W', padx=5, pady=2)


    def click_action_main1_1(self,click_value):
        if click_value == 'self.main1_1_button_1': # select browse
            fTyp = [("", "*.pptx;*.xlsx;*.yaml")]
            iDir = os.path.abspath(os.path.dirname(sys.argv[0]))
            self.full_filepath = tk.filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
            self.filename = os.path.basename(self.full_filepath)
            exec(self.entry_name_main1_1 + '.delete(0, tkinter.END)')
            exec(self.entry_name_main1_1 + '.insert(tk.END, self.full_filepath)')
            self.text.set(self.filename)
            self.click_action_main1_1('self.main1_1_button_2')

        if click_value == 'self.main1_1_button_2': # run submit on Main Panel
            file_type_array = ns_def.check_file_type(self.main1_1_entry_1.get())

            if file_type_array[0] == 'ERROR':
                self.text.set('[ERROR] ' + file_type_array[1])
                self.main1_1_label_4 = tk.Label(self.main1_1, textvariable=self.text, font=("", 10), background="#FBE5D6")
                self.main1_1_label_4.grid(row=4, column=1, columnspan=7, sticky='W', padx=5, pady=2)

            elif file_type_array[0] == 'PPT_SKECH':
                self.main1_1_label_4 = tk.Label(self.main1_1, textvariable=self.text, font=("", 10), background="#F2FDE3")
                self.main1_1_label_4.grid(row=4, column=1, columnspan=7, sticky='W', padx=5, pady=2)
                ns_front_run.sub_ppt_sketch_1(self,file_type_array)

            elif file_type_array[0] == 'EXCEL_MASTER':
                #print(file_type_array)
                self.main1_1_label_4 = tk.Label(self.main1_1, textvariable=self.text, font=("", 10), background="#F2FDE3")
                self.main1_1_label_4.grid(row=4, column=1, columnspan=7, sticky='W', padx=5, pady=2)
                ns_front_run.sub_excel_master_1(self, file_type_array)

            elif file_type_array[0] == 'EXCEL_DEVICE':
                #check attribute Table sheet in Excel file at ver 2.4.0
                input_device_table = openpyxl.load_workbook(str(self.main1_1_entry_1.get()))
                ws_list = input_device_table.sheetnames
                if 'Attribute' not in ws_list:
                    tkinter.messagebox.showinfo('info', 'The \'Attribute\' sheet, which was added in Ver. 2.4, is missing from the Dervice file. Please export the device file again from the master file.')
                    return

                #print(file_type_array)
                self.main1_1_label_4 = tk.Label(self.main1_1, textvariable=self.text, font=("", 10), background="#F2FDE3")
                self.main1_1_label_4.grid(row=4, column=1, columnspan=7, sticky='W', padx=5, pady=2)
                ns_front_run.sub_excel_device_1(self, file_type_array)

            elif file_type_array[0] == 'YAML_CML':
                #print(file_type_array)
                self.main1_1_label_4 = tk.Label(self.main1_1, textvariable=self.text, font=("", 10), background="#F2FDE3")
                self.main1_1_label_4.grid(row=4, column=1, columnspan=7, sticky='W', padx=5, pady=2)

                network_sketcher_dev.ns_front_run.click_action(self,'1-4b')

            else:
                self.text.set('[ERROR] Please enter a file compatible with NS')
                self.main1_1_label_4 = tk.Label(self.main1_1, textvariable=self.text, font=("", 10), background="#FBE5D6")
                self.main1_1_label_4.grid(row=4, column=1, columnspan=7, sticky='W', padx=5, pady=2)

        if click_value == 'self.help_1_button_1':
            webbrowser.open('https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/Japanese/User_Guide%5BJP%5D.md')

        if click_value == 'self.help_1_button_2':
            webbrowser.open('https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/User_Guide%5BEN%5D.md')

    '''
    Sketch Panel
    '''
    def sub_ppt_sketch_1(self,file_type_array):
        local_filename = self.filename
        local_fullpath = self.full_filepath
        push_array = [self.filename,self.full_filepath]

        self.sub1_1 = tk.Toplevel()
        self.sub1_1.title('Sketch Panel')
        self.root.update_idletasks()
        #print(self.root.winfo_width(),self.root.winfo_height(),self.root.winfo_x(),self.root.winfo_y() )  # width, height , x , y
        geo =  str(self.root.winfo_width()) + 'x' + str(self.root.winfo_height()) + '+' + str(self.root.winfo_x()) + '+' + str(self.root.winfo_y() + self.root.winfo_height() + 30)
        self.sub1_1.geometry(geo)

        self.sub1_0_label_1 = tk.Label(self.sub1_1, text=local_filename, font=("", 12), background="#FFFFFF")
        self.sub1_0_label_1 .grid(row=0, column=0, columnspan=7, sticky='W', padx=5, pady=5, ipadx=30, ipady=5)

        self.sub1_2 = tk.LabelFrame(self.sub1_1, text='Create Starter set', font=("", 14), height=1, background="#FFF9E7")
        self.sub1_2.grid(row=1, column=0, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.sub1_2_label_1 = tk.Button(self.sub1_2, text=" Master file\n Device file", font=("", 12), command=lambda: self.click_action_sub1_1('self.sub1_1_button_1',push_array))
        self.sub1_2_label_1.grid(row=3, column=1, sticky='W', padx=20, pady=20, ipadx=25, ipady=0)


        self.sub1_3 = tk.LabelFrame(self.sub1_1, text='Update to the Master file', font=("", 14), height=1, background="#FFF9E7")
        self.sub1_3.grid(row=1, column=1, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.text_sub1_3 = tk.StringVar()
        self.text_sub1_3.set(" drag and drop here ([MASTER]*.xlsx)")
        self.sub1_3_label_4 = tk.Label(self.sub1_3, textvariable=self.text_sub1_3, font=("", 10), background="#FFF9E7")
        self.sub1_3_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)

        self.sub1_3_entry_1 = tk.Entry(self.sub1_3)
        self.sub1_3_entry_1.grid(row=4, column=1, sticky="WE", padx=5, pady=3)
        self.sub1_3_button_1 = tk.Button(self.sub1_3, text="Browse ...", command=lambda: self.click_action_sub1_1('self.sub1_1_button_2',push_array))
        self.sub1_3_button_1.grid(row=4, column=2, sticky='W', padx=5, pady=2)
        self.sub1_3_button_1 = tk.Button(self.sub1_3, text="Submit", command=lambda: self.click_action_sub1_1('self.sub1_1_button_3',push_array))
        self.sub1_3_button_1.grid(row=4, column=3, sticky='W', padx=0, pady=2)

        #drag and drop
        self.entry_name_sub1_3 = 'self.sub1_3_entry_1'
        self.sub1_3.drop_target_register(DND_FILES)
        self.sub1_3.dnd_bind("<<Drop>>", self.drop_sub1_3 ,self.entry_name_sub1_3)

    def drop_sub1_3(self, event):
        if event:
            event.data = event.data.replace('{', '').replace('}', '')
            if event.data.endswith('.xlsx'):
                exec(self.entry_name_sub1_3 + '.delete(0, tkinter.END)')
                exec(self.entry_name_sub1_3 + '.insert(tk.END, event.data)')
                self.filename = os.path.basename(event.data)
                self.full_filepath = event.data
                self.text_sub1_3.set(self.filename)
                push_array = [self.filename, self.full_filepath]
                self.click_action_sub1_1('self.sub1_1_button_3',push_array)
            else:
                self.text_sub1_3.set('[ERROR] ' + 'Please input a file corresponding to NS')
                self.sub1_3_label_4 = tk.Label(self.sub1_3, textvariable=self.text_sub1_3, font=("", 10), background="#FBE5D6")
                self.sub1_3_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)


    def click_action_sub1_1(self, click_value,push_array):
        if click_value == 'self.sub1_1_button_1':  # Create the Master file
            full_filepath = self.main1_1_entry_1.get()
            iDir = os.path.abspath(os.path.dirname(full_filepath))

            #pre-defined for dev parameter
            self.outFileTxt_1_1 = tk.Entry(self.sub1_1)
            self.outFileTxt_1_2 = tk.Entry(self.sub1_1)
            self.inFileTxt_1_1 = tk.Entry(self.sub1_1)
            self.inFileTxt_2_1 = tk.Entry(self.sub1_1)
            self.outFileTxt_2_1 = tk.Entry(self.sub1_1)
            self.outFileTxt_2_2 = tk.Entry(self.sub1_1)
            self.outFileTxt_2_3 = tk.Entry(self.sub1_1)
            self.outFileTxt_2_4 = tk.Entry(self.sub1_1)
            self.inFileTxt_11_1 = tk.Entry(self.sub1_1)
            self.outFileTxt_11_2 = tk.Entry(self.sub1_1)
            self.inFileTxt_L2_1_1 = tk.Entry(self.sub1_1)
            self.inFileTxt_L3_1_1 = tk.Entry(self.sub1_1)

            #input for dev parameter
            basename_without_ext = os.path.splitext(os.path.basename(full_filepath))[0]
            self.outFileTxt_1_1.delete(0, tkinter.END)
            self.outFileTxt_1_1.insert(tk.END, iDir + ns_def.return_os_slash() + '[L1_DIAGRAM]AllAreasTag_' + basename_without_ext + '.pptx')
            self.outFileTxt_1_2.delete(0, tkinter.END)
            self.outFileTxt_1_2.insert(tk.END, iDir + ns_def.return_os_slash() + '[MASTER]' + basename_without_ext + '.xlsx')


            self.inFileTxt_2_1.delete(0, tkinter.END)
            self.inFileTxt_2_1.insert(tk.END, self.outFileTxt_1_2.get())
            self.inFileTxt_1_1.delete(0, tkinter.END)
            self.inFileTxt_1_1.insert(tk.END,self.main1_1_entry_1.get())
            self.inFileTxt_L2_1_1.delete(0, tkinter.END)
            self.inFileTxt_L2_1_1.insert(tk.END, iDir + ns_def.return_os_slash() + '[MASTER]' + basename_without_ext + '.xlsx')
            self.inFileTxt_L3_1_1.delete(0, tkinter.END)
            self.inFileTxt_L3_1_1.insert(tk.END, iDir + ns_def.return_os_slash() + '[MASTER]' + basename_without_ext + '.xlsx')


            self.inFileTxt_11_1.delete(0, tkinter.END)
            self.inFileTxt_11_1.insert(tk.END, self.full_filepath)
            self.outFileTxt_11_2.delete(0, tkinter.END)
            self.outFileTxt_11_2.insert(tk.END, iDir + ns_def.return_os_slash() + '[DEVICE]' + basename_without_ext + '.xlsx')

            ### run 1-4 in network_sketcher_dev ,  create l1 master file and sheet
            if self.click_value_2nd != 'self.sub1_1_button_3':
                self.click_value = '1-4'
                network_sketcher_dev.ns_front_run.click_action(self,'1-4')
            else:
                self.inFileTxt_L2_1_1.delete(0, tkinter.END)
                self.inFileTxt_L2_1_1.insert(tk.END, self.full_filepath)
                self.inFileTxt_L3_1_1.delete(0, tkinter.END)
                self.inFileTxt_L3_1_1.insert(tk.END, self.full_filepath)

            ### run L2-1-2 in network_sketcher_dev ,  add l2 master sheet
            self.click_value = 'L2-1-2'
            network_sketcher_dev.ns_front_run.click_action(self,'L2-1-2')

            # remove exist L2/ file
            #if os.path.isfile(self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L2_TABLE]')) == True:  # fixed ns-005 at 2.2.1(b)
            #    os.remove(self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L2_TABLE]'))  # fixed ns-005 at 2.2.1(b)

            ### run L3-1-2 in network_sketcher_dev ,  add l3 master sheet
            self.click_value = 'L3-1-2'
            network_sketcher_dev.ns_front_run.click_action(self,'L3-1-2')

            # remove exist L3/ file
            #if os.path.isfile(self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L3_TABLE]')) == True:  # fixed ns-005 at 2.2.1(b)
            #    os.remove(self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L3_TABLE]')) # fixed ns-005 at 2.2.1(b)

            ###Create the device file
            self.click_value_2nd = 'self.sub1_1_button_1'
            self.click_action_sub('self.self.sub2_5_button_3', push_array)
            self.click_value_2nd = ''

            ### open master panel
            if self.click_value_3rd != 'self.sub1_1_button_3':
                file_type_array = ['EXCEL_MASTER','EXCEL_MASTER']
                self.full_filepath = self.outFileTxt_1_2.get()
                self.filename = os.path.basename(self.full_filepath)
                ns_front_run.sub_excel_master_1(self, file_type_array)

        if click_value == 'self.sub1_1_button_2':  # select browse
            fTyp = [("","*.xlsx")]
            iDir = os.path.abspath(os.path.dirname(sys.argv[0]))
            self.full_filepath = tk.filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
            self.filename = os.path.basename(self.full_filepath)
            exec(self.entry_name_sub1_3 + '.delete(0, tkinter.END)')
            exec(self.entry_name_sub1_3 + '.insert(tk.END, self.full_filepath)')
            self.text_sub1_3.set(self.filename)
            self.click_action_sub1_1('self.sub1_1_button_3',push_array)

        if click_value == 'self.sub1_1_button_3':  # run submit on Sketch Panel
            file_type_array = ns_def.check_file_type(self.sub1_3_entry_1.get())

            if file_type_array[0] == 'ERROR':
                self.text_sub1_3.set('[ERROR] ' + file_type_array[1])
                self.sub1_3_label_4 = tk.Label(self.sub1_3, textvariable=self.text_sub1_3, font=("", 10), background="#FBE5D6")
                self.sub1_3_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)

            elif file_type_array[0] == 'EXCEL_MASTER':
                self.click_value_2nd = 'self.sub1_1_button_3'
                self.click_value_3rd = 'self.sub1_1_button_3'
                self.sub1_3_label_4 = tk.Label(self.sub1_3, textvariable=self.text_sub1_3, font=("", 10), background="#FFF9E7")
                self.sub1_3_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)
                #ns_front_run.sub_excel_master_1(self, file_type_array)
                #print('--- Update to the Master file ---')

                ### pre-defined for dev parameter
                self.inFileTxt_92_1 = tk.Entry(self.sub1_3)
                self.inFileTxt_92_2 = tk.Entry(self.sub1_3)
                self.inFileTxt_92_2_2 = tk.Entry(self.sub1_3)
                self.inFileTxt_2_1 = tk.Entry(self.sub1_3)
                self.outFileTxt_2_1 = tk.Entry(self.sub1_3)
                self.outFileTxt_2_2 = tk.Entry(self.sub1_3)
                self.outFileTxt_2_3 = tk.Entry(self.sub1_3)
                self.outFileTxt_2_4 = tk.Entry(self.sub1_3)

                ### input for dev parameter
                full_filepath_master = self.sub1_3_entry_1.get()
                full_filepath_sketch = self.main1_1_entry_1.get()
                iDir = os.path.abspath(os.path.dirname(full_filepath_master))
                basename_without_ext = os.path.splitext(os.path.basename(full_filepath_master))[0]

                self.inFileTxt_92_1.delete(0, tkinter.END)
                self.inFileTxt_92_1.insert(tk.END, full_filepath_sketch)
                self.inFileTxt_92_2.delete(0, tkinter.END)
                self.inFileTxt_92_2.insert(tk.END, full_filepath_master)
                self.inFileTxt_92_2_2.delete(0, tkinter.END)
                self.inFileTxt_92_2_2.insert(tk.END, iDir + ns_def.return_os_slash() + basename_without_ext + '_backup' + '.xlsx')

                ###check Master file open
                ns_def.check_file_open(full_filepath_master)

                ###create backup master file
                ns_def.get_backup_filename(full_filepath_master)

                '''Sketch file Sync to Master'''
                ### run 92-3 for dev , l1_sketch sync with L1_master file
                self.click_value = '92-3'
                network_sketcher_dev.ns_front_run.click_action(self, '92-3')

                ### device name that is updated in l1_sketch sync with master file
                ns_sync_between_layers.l1_sketch_device_name_sync_with_l2l3_master(self)

                ### L1 Master data update to L2 Master data
                ns_sync_between_layers.l1_master_device_and_line_sync_with_l2l3_master(self)
                self.click_value_2nd = ''
                self.click_value_3rd = ''

                # remove exist L3/ file
                if os.path.isfile(self.outFileTxt_11_2.get().replace('[MASTER]', '')) == True:
                    os.remove(self.outFileTxt_11_2.get().replace('[MASTER]', ''))

                ### open master panel
                file_type_array = ['EXCEL_MASTER', 'EXCEL_MASTER']
                self.full_filepath = full_filepath_master
                self.filename = os.path.basename(self.full_filepath)
                ns_front_run.sub_excel_master_1(self, file_type_array)

            else:
                self.text_sub1_3.set('[ERROR] Please input the Master file')
                self.sub1_3_label_4 = tk.Label(self.sub1_3, textvariable=self.text_sub1_3, font=("", 10), background="#FBE5D6")
                self.sub1_3_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)



    '''
    Master Panel
    '''
    def sub_excel_master_1(self,file_type_array):
        local_fullpath = self.full_filepath
        local_filename = self.filename
        push_array = [self.filename,self.full_filepath]

        if self.click_value_2nd == 'self.sub1_1_button_3':
            local_fullpath = self.sub1_3_entry_1.get()
            local_filename = os.path.basename(local_fullpath)
            push_array = [local_filename, local_fullpath]

        self.sub2_1 = tk.Toplevel()
        self.sub2_1.title('Master Panel')
        self.root.update_idletasks()
        #print(self.root.winfo_width(),self.root.winfo_height(),self.root.winfo_x(),self.root.winfo_y() )  # width, height , x , y
        geo =  str(self.root.winfo_width() + 180) + 'x' + str(self.root.winfo_height() + 120) + '+' + str(self.root.winfo_x() + self.root.winfo_width()) + '+' + str(self.root.winfo_y())
        self.sub2_1.geometry(geo)

        self.sub2_0_label_1 = tk.Label(self.sub2_1, text=local_filename, font=("", 12), background="#FFFFFF")
        self.sub2_0_label_1 .grid(row=0, column=0, columnspan=7, sticky='W', padx=5, pady=5, ipadx=30, ipady=5)

        ### pre-defined for dev parameter
        self.outFileTxt_1_2 = tk.Entry(self.sub2_1)
        self.outFileTxt_2_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_2_2 = tk.Entry(self.sub2_1)
        self.outFileTxt_2_3 = tk.Entry(self.sub2_1)
        self.outFileTxt_2_4 = tk.Entry(self.sub2_1)
        self.inFileTxt_2_1 = tk.Entry(self.sub2_1)
        self.click_value_dummy = 'dummy'
        self.inFileTxt_L2_3_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L2_3_4_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L2_3_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L2_3_2 = tk.Entry(self.sub2_1)
        self.inFileTxt_L3_3_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L3_3_4_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L3_3_5_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L3_4_2_2 = tk.Entry(self.sub2_1)
        self.inFileTxt_11_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_11_2 = tk.Entry(self.sub2_1)
        self.outFileTxt_11_3 = tk.Entry(self.sub2_1) # for a bug fix at 2.2.1(c)
        self.inFileTxt_L2_1_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L2_1_4_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L2_1_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L2_1_2 = tk.Entry(self.sub2_1)
        self.inFileTxt_L3_1_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L3_1_4_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L3_1_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_L3_1_2 = tk.Entry(self.sub2_1)

        ### input for dev parameter
        basename_without_ext = os.path.splitext(os.path.basename(local_fullpath))[0]
        iDir = os.path.abspath(os.path.dirname(local_fullpath))

        self.outFileTxt_2_1.delete(0, tkinter.END)
        self.outFileTxt_2_1.insert(tk.END, iDir + ns_def.return_os_slash() + '[L1_DIAGRAM]PerArea_' + basename_without_ext.replace('[MASTER]', '') + '.pptx')
        self.outFileTxt_2_2.delete(0, tkinter.END)
        self.outFileTxt_2_2.insert(tk.END, iDir + ns_def.return_os_slash() + '[L1_DIAGRAM]PerAreaTag_' + basename_without_ext.replace('[MASTER]', '') + '.pptx')
        self.outFileTxt_2_3.delete(0, tkinter.END)
        self.outFileTxt_2_3.insert(tk.END, iDir + ns_def.return_os_slash() + '[L1_DIAGRAM]AllAreas_' + basename_without_ext.replace('[MASTER]', '') + '.pptx')
        self.outFileTxt_2_4.delete(0, tkinter.END)
        self.outFileTxt_2_4.insert(tk.END, iDir + ns_def.return_os_slash() + '[L1_DIAGRAM]AllAreasTag_' + basename_without_ext.replace('[MASTER]', '') + '.pptx')
        self.inFileTxt_2_1.delete(0, tkinter.END)
        self.inFileTxt_2_1.insert(tk.END, local_fullpath)
        self.inFileTxt_L2_3_1.delete(0, tkinter.END)
        self.inFileTxt_L2_3_1.insert(tk.END, local_fullpath)
        self.inFileTxt_L3_3_1.delete(0, tkinter.END)
        self.inFileTxt_L3_3_1.insert(tk.END, local_fullpath)
        self.inFileTxt_11_1.delete(0, tkinter.END)
        self.inFileTxt_11_1.insert(tk.END, local_fullpath)
        self.outFileTxt_11_2.delete(0, tkinter.END)
        self.outFileTxt_11_2.insert(tk.END, iDir + ns_def.return_os_slash() + '[DEVICE]' + basename_without_ext + '.xlsx')
        self.inFileTxt_L2_1_1.delete(0, tkinter.END)
        self.inFileTxt_L2_1_1.insert(tk.END, local_fullpath)
        self.outFileTxt_L2_1_4_1.delete(0, tkinter.END)
        self.outFileTxt_L2_1_4_1.insert(tk.END, local_fullpath)
        self.outFileTxt_L2_1_1.delete(0, tkinter.END)
        self.outFileTxt_L2_1_1.insert(tk.END, local_fullpath)
        self.outFileTxt_L2_1_2.delete(0, tkinter.END)
        self.outFileTxt_L2_1_2.insert(tk.END, local_fullpath)
        self.inFileTxt_L3_1_1.delete(0, tkinter.END)
        self.inFileTxt_L3_1_1.insert(tk.END, local_fullpath)
        self.outFileTxt_L3_1_4_1.delete(0, tkinter.END)
        self.outFileTxt_L3_1_4_1.insert(tk.END, local_fullpath)
        self.outFileTxt_L3_1_1.delete(0, tkinter.END)
        self.outFileTxt_L3_1_1.insert(tk.END, local_fullpath)
        self.outFileTxt_L3_1_2.delete(0, tkinter.END)
        self.outFileTxt_L3_1_2.insert(tk.END, local_fullpath)

        ### run 2-4-x for dev , Create L1 diagram
        self.sub2_2x = tk.LabelFrame(self.sub2_1, text='Create Diagram files', font=("", 14), height=1, background="#FBE5D6")
        self.sub2_2x.grid(row=1, column=0, columnspan=7, sticky='W', padx=5, pady=0, ipadx=2, ipady=0)

        self.sub2_2 = tk.LabelFrame(self.sub2_2x, text='Layer1 Diagram', font=("", 14), height=1, background="#FEF6F0")
        self.sub2_2.grid(row=1, column=0, columnspan=7, sticky='W', padx=2, pady=2, ipadx=5, ipady=2)

        self.sub2_2_button_3 = tk.Button(self.sub2_2, text="All Areas", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'2-4-3'))
        self.sub2_2_button_3.grid(row=2, column=1, sticky='WE', padx=5, pady=2, ipadx=15)
        self.sub2_2_button_4 = tk.Button(self.sub2_2, text="All Areas with IF Tag", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'2-4-4'))
        self.sub2_2_button_4.grid(row=2, column=2, sticky='WE', padx=5, pady=2)
        self.sub2_2_button_1 = tk.Button(self.sub2_2, text="Per Area", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'2-4-1'))
        self.sub2_2_button_1.grid(row=2, column=3, sticky='WE', padx=5, pady=2, ipadx=15)
        self.sub2_2_button_2 = tk.Button(self.sub2_2, text="Per Area with IF Tag", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'2-4-2'))
        self.sub2_2_button_2.grid(row=2, column=4, sticky='WE', padx=5, pady=2)

        ### run L2-3-x for dev , Create L2 diagram
        self.sub2_3 = tk.LabelFrame(self.sub2_2x, text='Layer2 Diagram', font=("", 14), height=1, background="#FEF6F0")
        self.sub2_3.grid(row=4, column=0, sticky='W', padx=1, pady=0, ipadx=1, ipady=2)

        ## Add at ve 2.3.0(b)
        optionL2_3_6 = ns_extensions.ip_report.get_folder_list(self)
        global variableL2_3_6
        variableL2_3_6 = tk.StringVar()
        self.comboL2_3_6 = ttk.Combobox(self.sub2_3 , values=optionL2_3_6, textvariable=variableL2_3_6, font=("", 12), state='readonly' , width=20)
        self.comboL2_3_6.set(str(optionL2_3_6[0]))
        self.comboL2_3_6.option_add("*TCombobox*Listbox.Font", 12)
        self.comboL2_3_6.grid(row=0, column=0, sticky='WE', padx=1, pady=5, ipady=0, ipadx=8, columnspan=3)

        self.sub2_3_button_1 = tk.Button(self.sub2_3, text="Per Area", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'L2-3-2'))
        self.sub2_3_button_1.grid(row=6, column=1, sticky='WE', padx=0, pady=2, ipadx=0)

        ### run L3-3-x for dev , Create L3 diagram
        self.sub2_4 = tk.LabelFrame(self.sub2_2x, text='Layer3 Diagram', font=("", 14), height=1, background="#FEF6F0")
        self.sub2_4.grid(row=4, column=2, sticky='W', padx=1, pady=0, ipadx=1, ipady=2)

        self.sub2_4_empty1 = tk.LabelFrame(self.sub2_4, text='', font=("", 14), width=10)
        self.sub2_4_empty1 .grid(row=1, column=1, sticky='WE', padx=0, pady=0, ipadx=1)

        self.sub2_4_button_1 = tk.Button(self.sub2_4, text="All Areas", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'L3-4-1')) # add button at ver 2.3.0
        self.sub2_4_button_1.grid(row=1, column=2, sticky='WE', padx=1, pady=2, ipadx=20)

        self.sub2_4_button_1 = tk.Button(self.sub2_4, text="Per Area", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'L3-3-2'))
        self.sub2_4_button_1.grid(row=2, column=2, sticky='WE', padx=1, pady=2, ipadx=20)

        ### run xx-xx for dev , Create VPN diagram
        self.sub2_6 = tk.LabelFrame(self.sub2_2x, text='VPN Diagram', font=("", 14), height=1, background="#FEF6F0")
        self.sub2_6.grid(row=4, column=4, sticky='W', padx=1, pady=0, ipadx=5, ipady=2)

        self.sub2_6_empty1 = tk.LabelFrame(self.sub2_6, text='', font=("", 14), width=15)
        self.sub2_6_empty1 .grid(row=10, column=0, sticky='WE', padx=0, pady=0, ipadx=1)

        self.sub2_6_button_3 = tk.Button(self.sub2_6, text="VPNs on L1", font=("", 12), command=lambda: self.click_action_sub('self.self.sub2_6_button_1', push_array))
        self.sub2_6_button_3.grid(row=10, column=1, sticky='WE', padx=1, pady=2, ipadx=5)

        self.sub2_6_button_4 = tk.Button(self.sub2_6, text="VPNs on L3", font=("", 12), command=lambda: self.click_action_sub('self.self.sub2_6_button_2', push_array))
        self.sub2_6_button_4.grid(row=11, column=1, sticky='WE', padx=1, pady=2, ipadx=5)

        ### Attribute combobox at ver 2.4.0
        self.ATTR_1_2 = tk.LabelFrame(self.sub2_2x, text='Attribute ', font=("", 12), height=1, background="#FEF6F0")
        self.ATTR_1_2.grid(row=4, column=5, sticky='W', padx=3, pady=3, ipadx=0, ipady=0)

        optionATTR_1_1 = ns_def.get_attribute_title_list(self, self.inFileTxt_L2_3_1.get())
        global variableATTR_1_1
        variableATTR_1_1 = tk.StringVar()
        self.comboATTR_1_1 = ttk.Combobox(self.ATTR_1_2, values=optionATTR_1_1, textvariable=variableATTR_1_1, font=("", 12), state='readonly',width=10)
        self.comboATTR_1_1.set(str(optionATTR_1_1[0]))
        self.comboATTR_1_1.option_add("*TCombobox*Listbox.Font", 12)
        self.comboATTR_1_1.grid(row=0, column=0, sticky='N', padx=1, pady=1, ipady=0, ipadx=3)
        self.comboATTR_1_1.bind("<<ComboboxSelected>>", self.on_combobox_select)
        #print(self.comboATTR_1_1.get())

        self.attribute_tuple1_1 = ns_def.get_global_attribute_tuple(self.inFileTxt_L2_3_1.get(), self.comboATTR_1_1.get())

        ### run 11-4 for dev , Export to the Device file
        self.sub2_0_label_2 = tk.Label(self.sub2_1, text='', font=("", 1))
        self.sub2_0_label_2 .grid(row=7, column=0, columnspan=7, sticky='W', padx=0, pady=0, ipadx=0, ipady=0)

        self.sub2_5 = tk.LabelFrame(self.sub2_1, text='Export to the Device file', font=("", 14), height=1, background="#DFC9EF")
        self.sub2_5.grid(row=8, column=0, sticky='W', padx=5, pady=0, ipadx=5, ipady=2)

        push_array = []
        self.sub2_5_button_3 = tk.Button(self.sub2_5, text="Device file", font=("", 12), command=lambda: self.click_action_sub('self.self.sub2_5_button_3', push_array))
        self.sub2_5_button_3.grid(row=10, column=1, sticky='WE', padx=50, pady=2, ipadx=15)

        '''
        Extensions
        '''
        self.sub3_3 = tk.LabelFrame(self.sub2_1, text='Extensions', font=("", 14), height=1, background="#C2E2EC")
        self.sub3_3.grid(row=8, column=1, sticky='W', padx=1, pady=0, ipadx=5, ipady=5)

        self.sub3_3_button_1 = tk.Button(self.sub3_3, text="Auto IP Addressing", font=("", 12), command=lambda: ns_front_run.sub_master_extention_1(self))
        self.sub3_3_button_1.grid(row=0, column=0, sticky='WE', padx=15, pady=2, ipadx=5)

        self.sub3_3_button_2 = tk.Button(self.sub3_3, text="Report", font=("", 12), command=lambda: ns_front_run.sub_master_extention_2(self))
        self.sub3_3_button_2.grid(row=0, column=1, sticky='WE', padx=5, pady=2, ipadx=15)

    def on_combobox_select(self, event):
        self.attribute_tuple1_1 = ns_def.get_global_attribute_tuple(self.inFileTxt_L2_3_1.get(), self.comboATTR_1_1.get())
    def sub_master_extention_1(self):
        local_filename = self.filename
        local_fullpath = self.full_filepath
        push_array = [self.filename,self.full_filepath]

        self.sub3_4 = tk.Toplevel()
        self.sub3_4.title('Auto IP Addressing')
        self.root.update_idletasks()
        #print(self.root.winfo_width(),self.root.winfo_height(),self.root.winfo_x(),self.root.winfo_y() )  # width, height , x , y
        geo =  str(self.root.winfo_width() + 95) + 'x' + str(self.root.winfo_height() + 250) + '+' + str(self.root.winfo_x() + self.root.winfo_width() - 250) + '+' + str(self.root.winfo_y() )
        self.sub3_4.geometry(geo)

        self.sub3_4_0 = tk.Label(self.sub3_4, text=local_filename, font=("", 12), background="#FFFFFF")
        self.sub3_4_0 .grid(row=0, column=0, columnspan=7, sticky='W', padx=5, pady=5, ipadx=30, ipady=5)

        # 1.Select Area
        self.sub3_4_1 = tk.LabelFrame(self.sub3_4, text='Auto IP Addressing', font=("", 16), height=1, background="#C2E2EC")
        self.sub3_4_1.grid(row=1, column=0, columnspan=5, sticky='W', padx=5, pady=0, ipadx=3, ipady=0)

        self.sub3_4_1_1 = tk.Label(self.sub3_4_1, text='- Select Area (Required)', font=("", 16), background="#E8F4F8")
        self.sub3_4_1_1 .grid(row=0, column=0, sticky='W', padx=5, pady=0, ipadx=5, ipady=0)

        option3_4_1_1 = ns_extensions.auto_ip_addressing.get_folder_list(self)
        global variable3_4_1_1
        variable3_4_1_1 = tk.StringVar()
        self.combo3_4_1_1 = ttk.Combobox(self.sub3_4_1 , values=option3_4_1_1, textvariable=variable3_4_1_1, font=("", 12), state='readonly')
        self.combo3_4_1_1.set("<Select Area>")
        self.combo3_4_1_1.option_add("*TCombobox*Listbox.Font", 12)
        self.combo3_4_1_1.grid(row=0, column=1, sticky='WE', padx=5, pady=15, ipady=2, ipadx=15)
        self.combo3_4_1_1.bind("<<ComboboxSelected>>", lambda event: ns_extensions.auto_ip_addressing.get_auto_ip_param(self,self.combo3_4_1_1.get()))

        # IP Address Range Settings
        self.sub3_4_x = tk.LabelFrame(self.sub3_4_1, text='Range Settings(Option)', font=("", 14), height=1, background="#E8F4F8")
        self.sub3_4_x.grid(row=2, column=0, columnspan=5, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.sub3_4_3 = tk.Label(self.sub3_4_x, text='- Starting point of IP address network (CIDR):', font=("", 12), background="#E8F4F8")
        self.sub3_4_3 .grid(row=0, column=0, columnspan=3, sticky='W', padx=5, pady=0, ipadx=5, ipady=0)
        self.sub3_4_3_entry_1 = tk.Entry(self.sub3_4_x, font=("",12))
        self.sub3_4_3_entry_1.grid(row=0, column=8, sticky="WE", padx=5, pady=0, ipadx=5)

        self.sub3_4_3_entry_1.insert(0, '')
        self.sub3_4_3_entry_1['justify'] = tkinter.CENTER

        self.sub3_4_2 = tk.Label(self.sub3_4_x, text='- Number of free IP addresses in each segment:', font=("", 12), background="#E8F4F8")
        self.sub3_4_2 .grid(row=1, column=0, columnspan=12, sticky='W', padx=5, pady=0, ipadx=5, ipady=0)
        self.sub3_4_2_entry_1 = tk.Entry(self.sub3_4_x, font=("",12))
        self.sub3_4_2_entry_1.grid(row=1, column=8, sticky="WE", padx=5, pady=0, ipadx=5)
        self.sub3_4_2_entry_1.insert(0,'2')
        self.sub3_4_2_entry_1['justify'] = tkinter.CENTER

        # IP address numbering rules
        self.sub3_4_4 = tk.LabelFrame(self.sub3_4_1, text='Numbering rules(Option)', font=("", 14), height=1, background="#E8F4F8")
        self.sub3_4_4.grid(row=3, column=0, columnspan=5, sticky='W', padx=5, pady=0, ipadx=5, ipady=0)

        self.sub3_4_4_1 = tk.Label(self.sub3_4_4, text='- Ascending or descending order:', font=("", 12), background="#E8F4F8")
        self.sub3_4_4_1 .grid(row=0, column=0, columnspan=5, sticky='W', padx=5, pady=0, ipadx=5, ipady=0)

        option3_4_4_1 = ["Ascending order", "Descending order"]
        global variable3_4_4_1
        variable3_4_4_1 = tk.StringVar()
        self.combo3_4_4_1 = ttk.Combobox(self.sub3_4_4 , values=option3_4_4_1, textvariable=variable3_4_4_1, font=("", 12), state='readonly')
        self.combo3_4_4_1.current(0)
        self.combo3_4_4_1.option_add("*TCombobox*Listbox.Font", 12)
        self.combo3_4_4_1.grid(row=0, column=5, sticky='WE', padx=5, pady=2, ipadx=15)
        self.combo3_4_4_1.bind("<<ComboboxSelected>>")

        self.sub3_4_4_2 = tk.Label(self.sub3_4_4, text='   * [e.g.] Ascending : 1 -> 2 -> 3 ... , Descending : 254 -> 253 -> 251 ... ', font=("", 10), background="#E8F4F8")
        self.sub3_4_4_2 .grid(row=1, column=0, columnspan=8, sticky='W', padx=5, pady=0, ipadx=5, ipady=0)


        # Support functions
        self.sub3_4_6 = tk.LabelFrame(self.sub3_4_1, text='Completion of missing IP addresses(Option)', font=("", 14), height=1, background="#E8F4F8")
        self.sub3_4_6.grid(row=4, column=0, columnspan=5, sticky='W', padx=5, pady=5, ipadx=5, ipady=0)

        self.sub3_4_6_1 = tk.Label(self.sub3_4_6, text='- Within the same layer 3 segment:', font=("", 12), background="#E8F4F8")
        self.sub3_4_6_1 .grid(row=0, column=0, columnspan=5, sticky='W', padx=5, pady=0, ipadx=5, ipady=0)

        option3_4_6_1 = [ "Keep existing IP address","Reassign within the same subnet"]
        global variable3_4_6_1
        variable3_4_6_1 = tk.StringVar()
        self.combo3_4_6_1 = ttk.Combobox(self.sub3_4_6 , values=option3_4_6_1, textvariable=variable3_4_6_1, font=("", 12), state='readonly')
        self.combo3_4_6_1.current(0)
        self.combo3_4_6_1.option_add("*TCombobox*Listbox.Font", 12)
        self.combo3_4_6_1.grid(row=0, column=5, sticky='WE', padx=5, pady=2, ipadx=25)
        self.combo3_4_6_1.bind("<<ComboboxSelected>>")

        # Run
        self.sub3_4_button_1 = tk.Button(self.sub3_4_1, text=" Run IP Addressing ", font=("", 14), command=lambda: self.click_action_sub('self.sub3_4_button_1',self.combo3_4_1_1.get()))
        self.sub3_4_button_1.grid(row=6, column=0, sticky='W', padx=30, pady=10)

    def sub_master_extention_2(self):
        local_filename = self.filename
        local_fullpath = self.full_filepath
        push_array = [self.filename,self.full_filepath]

        self.sub3_5 = tk.Toplevel()
        self.sub3_5.title('Report')
        self.root.update_idletasks()
        #print(self.root.winfo_width(),self.root.winfo_height(),self.root.winfo_x(),self.root.winfo_y() )  # width, height , x , y
        geo =  str(self.root.winfo_width() - 230) + 'x' + str(self.root.winfo_height() - 60) + '+' + str(self.root.winfo_x() + self.root.winfo_width() +150) + '+' + str(self.root.winfo_y() + 50 )
        self.sub3_5.geometry(geo)

        self.sub3_5_0 = tk.Label(self.sub3_5, text=local_filename, font=("", 12), background="#FFFFFF")
        self.sub3_5_0 .grid(row=0, column=0, columnspan=7, sticky='W', padx=5, pady=5, ipadx=30, ipady=15)

        # IP Address report
        self.sub3_5_1 = tk.LabelFrame(self.sub3_5, text='Export', font=("", 14), height=1, background="#C2E2EC")
        self.sub3_5_1.grid(row=1, column=0, columnspan=5, sticky='W', padx=5, pady=0, ipadx=3, ipady=0)

        # Export to the IP Address table
        self.sub3_5_button_1 = tk.Button(self.sub3_5_1, text=" IP Address table ", font=("", 12), command=lambda: self.click_action_sub('self.sub3_5_button_1','dummy'))
        self.sub3_5_button_1.grid(row=6, column=0, sticky='W', padx=20, pady=5)

    def click_action_sub(self, click_value, target_area_name):
        if click_value == 'self.sub3_5_button_1':  # select IP address table
            ###export_ip_report
            ns_extensions.ip_report.export_ip_report(self, target_area_name)
            ns_def.messagebox_file_open(str(self.outFileTxt_11_3.get()))

        if click_value == 'self.sub3_4_button_1':  # select Run
            #change target area name to N/A
            if target_area_name == '_WAN(Way_Point)_':
                target_area_name = 'N/A'

            ###check Master file open
            ns_def.check_file_open(self.inFileTxt_L3_3_1.get())

            ###create backup master file
            ns_def.get_backup_filename(self.inFileTxt_L3_3_1.get())

            ###run_auto_ip
            ns_extensions.auto_ip_addressing.run_auto_ip(self,target_area_name)

            ### messagebox
            tkinter.messagebox.showinfo(title='Complete', message='[MASTER] file has been updated.')

        if click_value == 'self.self.sub2_5_button_3':  # Create Device file
            ### check file open
            if ns_def.check_file_open(str(self.outFileTxt_11_2.get()).replace('[MASTER]','')) == True:
                return ()

            ### create device file and L1 Table
            self.click_value = '11-4'
            network_sketcher_dev.ns_front_run.click_action(self, '11-4')
            # run x-x for dev , Create L2 Table
            self.click_value = 'L2-1-2'
            network_sketcher_dev.ns_front_run.click_action(self, 'L2-1-2')
            # run x-x for dev , Create L3 Table
            self.click_value = 'L3-1-2'
            network_sketcher_dev.ns_front_run.click_action(self, 'L3-1-2')

            # run x-x for dev , Create Attribute Table add to ver 2.4.0
            self.click_value = 'ATTR-1-1'
            network_sketcher_dev.ns_front_run.click_action(self, 'ATTR-1-1')

            if self.click_value_2nd != 'self.sub1_1_button_1' and self.click_value_2nd != 'self.sub3_1_button_3':
                ns_def.messagebox_file_open(str(self.outFileTxt_11_2.get()).replace('[MASTER]',''))

        if click_value == 'self.self.sub2_6_button_1':  # Click "VPNs on L1"
            #print('--- Click "VPNs on L1" ---')
            ### create L1 Table with [VPNs_on_L1]]
            self.click_value = 'VPN-1-1'
            network_sketcher_dev.ns_front_run.click_action(self, '2-4-3')

            ### Write VPNs on L1 ###
            ns_vpn_diagram_create.ns_write_vpns_on_l1.__init__(self)

            ns_def.messagebox_file_open(self.output_ppt_file) #Add at Ver 2.3.1(a)

        if click_value == 'self.self.sub2_6_button_2':  # Click "VPNs on L3"
            #print('--- Click VPNs on L3 ---')
            self.click_value = 'L3-4-1'
            self.click_value_VPN = 'VPN-1-3'

            ### Modify Master file for L3 vpn ###
            #ns_vpn_diagram_create.ns_modify_master_l3vpn.__init__(self)

            ### Create L3 All Areas with l3vpn master file ###
            network_sketcher_dev.ns_front_run.click_action(self, 'L3-4-1')

            ### reset initial value
            self.click_value_VPN = ''


    '''
    Device Panel
    '''
    def sub_excel_device_1(self,file_type_array):
        local_filename = self.filename
        local_fullpath = self.full_filepath
        push_array = [self.filename,self.full_filepath]

        self.sub3_1 = tk.Toplevel()
        self.sub3_1.title('Device Panel')
        self.root.update_idletasks()
        #print(self.root.winfo_width(),self.root.winfo_height(),self.root.winfo_x(),self.root.winfo_y() )  # width, height , x , y
        geo =  str(self.root.winfo_width() - 180) + 'x' + str(self.root.winfo_height() - 20) + '+' + str(self.root.winfo_x() + self.root.winfo_width()) + '+' + str(self.root.winfo_y() + self.root.winfo_height() + 30)
        self.sub3_1.geometry(geo)

        self.sub3_0_label_1 = tk.Label(self.sub3_1, text=local_filename, font=("", 12), background="#FFFFFF")
        self.sub3_0_label_1 .grid(row=0, column=0, columnspan=7, sticky='W', padx=5, pady=5, ipadx=30, ipady=5)

        self.sub3_2 = tk.LabelFrame(self.sub3_1, text='Update to the Master file', font=("", 14), height=1, background="#E5F4F7")
        self.sub3_2.grid(row=1, column=0, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.text_sub3_1 = tk.StringVar()
        self.text_sub3_1.set("      drag and drop here ([MASTER]*.xlsx)")
        self.sub3_1_label_4 = tk.Label(self.sub3_2, textvariable=self.text_sub3_1, font=("", 10), background="#E5F4F7")
        self.sub3_1_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)

        self.sub3_1_entry_1 = tk.Entry(self.sub3_2)
        self.sub3_1_entry_1.grid(row=4, column=1, sticky="WE", padx=5, pady=3, ipadx=15)
        self.sub3_1_button_1 = tk.Button(self.sub3_2, text="Browse ...", command=lambda: self.click_action_sub3_1('self.sub3_1_button_2',push_array))
        self.sub3_1_button_1.grid(row=4, column=2, sticky='W', padx=5, pady=2)
        self.sub3_1_button_1 = tk.Button(self.sub3_2, text="Submit", command=lambda: self.click_action_sub3_1('self.sub3_1_button_3',push_array))
        self.sub3_1_button_1.grid(row=4, column=3, sticky='W', padx=0, pady=2)

        #drag and drop
        self.entry_name_sub3_1 = 'self.sub3_1_entry_1'
        self.sub3_1.drop_target_register(DND_FILES)
        self.sub3_1.dnd_bind("<<Drop>>", self.drop_sub3_1 ,self.entry_name_sub3_1)

    def drop_sub3_1(self, event):
        if event:
            event.data = event.data.replace('{', '').replace('}', '')
            if event.data.endswith('.xlsx'):
                exec(self.entry_name_sub3_1 + '.delete(0, tkinter.END)')
                exec(self.entry_name_sub3_1 + '.insert(tk.END, event.data)')
                self.filename = os.path.basename(event.data)
                self.full_filepath = event.data
                self.text_sub3_1.set(self.filename)
                push_array = [self.filename, self.full_filepath]
                self.click_action_sub3_1('self.sub3_1_button_3',push_array)
            else:
                self.text_sub3_1.set('[ERROR] ' + 'Please input a file corresponding to NS')
                self.sub3_1_label_4 = tk.Label(self.sub3_1, textvariable=self.text_sub3_1, font=("", 10), background="#FBE5D6")
                self.sub3_1_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)

    def click_action_sub3_1(self, click_value,push_array):
        if click_value == 'self.sub3_1_button_2':  # select browse
            fTyp = [("","*.xlsx")]
            iDir = os.path.abspath(os.path.dirname(sys.argv[0]))
            self.full_filepath = tk.filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
            self.filename = os.path.basename(self.full_filepath)
            exec(self.entry_name_sub3_1 + '.delete(0, tkinter.END)')
            exec(self.entry_name_sub3_1 + '.insert(tk.END, self.full_filepath)')
            self.text_sub3_1.set(self.filename)
            self.click_action_sub1_1('self.sub3_1_button_3',push_array)

        if click_value == 'self.sub3_1_button_3':  # run submit on Main Panel
            file_type_array = ns_def.check_file_type(self.sub3_1_entry_1.get())

            if file_type_array[0] == 'ERROR':
                self.text_sub3_1.set('[ERROR] ' + file_type_array[1])
                self.sub3_1_label_4 = tk.Label(self.sub3_1, textvariable=self.text_sub3_1, font=("", 10), background="#FBE5D6")
                self.sub3_1_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)

            elif file_type_array[0] == 'EXCEL_MASTER':
                self.sub3_1_label_4 = tk.Label(self.sub3_1, textvariable=self.text_sub3_1, font=("", 10), background="#E5F4F7")
                self.sub3_1_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)
                #ns_front_run.sub_excel_master_1(self, file_type_array)
                #print('--- Update to the Master file ---')

                ### pre-defined for dev parameter
                self.inFileTxt_11_1 = tk.Entry(self.sub3_1)
                self.outFileTxt_11_2 = tk.Entry(self.sub3_1)
                self.inFileTxt_12_1 = tk.Entry(self.sub3_1)
                self.inFileTxt_12_1_2 = tk.Entry(self.sub3_1)
                self.outFileTxt_12_3_1 = tk.Entry(self.sub3_1)
                self.inFileTxt_12_2 = tk.Entry(self.sub3_1)
                self.inFileTxt_12_2_2 = tk.Entry(self.sub3_1)
                self.inFileTxt_12_2_3 = tk.Entry(self.sub3_1)
                self.inFileTxt_L2_2_1 = tk.Entry(self.sub3_1)
                self.inFileTxt_L2_2_2 = tk.Entry(self.sub3_1)
                self.inFileTxt_L2_2_2_2 = tk.Entry(self.sub3_1)
                self.inFileTxt_L3_2_1 = tk.Entry(self.sub3_1)
                self.inFileTxt_L3_2_2 = tk.Entry(self.sub3_1)
                self.inFileTxt_L3_2_2_2 = tk.Entry(self.sub3_1)
                self.inFileTxt_L3_1_1 = tk.Entry(self.sub3_1)
                self.inFileTxt_L2_1_1 = tk.Entry(self.sub3_1)

                ### input for dev parameter
                full_filepath_master = self.sub3_1_entry_1.get()
                full_filepath_device = self.main1_1_entry_1.get()
                iDir = os.path.abspath(os.path.dirname(full_filepath_master))
                basename_without_ext = os.path.splitext(os.path.basename(full_filepath_master))[0]
                basename_without_ext_device = os.path.splitext(os.path.basename(full_filepath_device))[0]

                self.inFileTxt_12_1.delete(0, tkinter.END)
                self.inFileTxt_12_1.insert(tk.END, full_filepath_device)

                self.inFileTxt_12_2.delete(0, tkinter.END)
                self.inFileTxt_12_2.insert(tk.END, full_filepath_master)

                self.inFileTxt_12_2_2.delete(0, tkinter.END)
                self.inFileTxt_12_2_2.insert(tk.END, iDir + ns_def.return_os_slash() + basename_without_ext + '_backup' + '.xlsx')

                # make ppt diagram backup path
                self.inFileTxt_12_2_3.delete(0, tkinter.END)
                self.inFileTxt_12_2_3.insert(tk.END, iDir + ns_def.return_os_slash() + '[L1_DIAGRAM]AllAreasTag_' + str(basename_without_ext).replace('[MASTER]', '') + '_backup' + '.pptx')

                # SET ppt diagram file path
                self.outFileTxt_12_3_1.delete(0, tkinter.END)
                self.outFileTxt_12_3_1.insert(tk.END, iDir + ns_def.return_os_slash() + '[L1_DIAGRAM]AllAreasTag_' + str(basename_without_ext).replace('[MASTER]', '') + '.pptx')

                # SET Device file path
                self.outFileTxt_11_2.delete(0, tkinter.END)
                self.outFileTxt_11_2.insert(tk.END, iDir + ns_def.return_os_slash() + '[DEVICE]' + basename_without_ext_device + '.xlsx')

                # SET L2 path
                self.inFileTxt_L2_2_1.delete(0, tkinter.END)
                self.inFileTxt_L2_2_1.insert(tk.END, full_filepath_device)
                self.inFileTxt_L2_2_2.delete(0, tkinter.END)
                self.inFileTxt_L2_2_2.insert(tk.END, full_filepath_master)
                self.inFileTxt_L2_2_2_backup = iDir + ns_def.return_os_slash() + os.path.splitext(os.path.basename(self.inFileTxt_L2_2_2.get()))[0] + '_backup' + '.xlsx'
                self.inFileTxt_L2_2_2_2.delete(0, tkinter.END)
                self.inFileTxt_L2_2_2_2.insert(tk.END, iDir + ns_def.return_os_slash() + os.path.splitext(os.path.basename(self.inFileTxt_L2_2_2.get()))[0] + '_backup' + '.xlsx')

                # SET L3 path
                self.inFileTxt_L3_2_1.delete(0, tkinter.END)
                self.inFileTxt_L3_2_1.insert(tk.END, full_filepath_device)
                self.inFileTxt_L3_2_2.delete(0, tkinter.END)
                self.inFileTxt_L3_2_2.insert(tk.END, full_filepath_master)
                self.inFileTxt_L3_2_2_backup = iDir + ns_def.return_os_slash() + os.path.splitext(os.path.basename(self.inFileTxt_L3_2_2.get()))[0] + '_backup' + '.xlsx'
                self.inFileTxt_L3_2_2_2.delete(0, tkinter.END)
                self.inFileTxt_L3_2_2_2.insert(tk.END, iDir + ns_def.return_os_slash() + os.path.splitext(os.path.basename(self.inFileTxt_L3_2_2.get()))[0] + '_backup' + '.xlsx')
                self.inFileTxt_L3_1_1.delete(0, tkinter.END)
                self.inFileTxt_L3_1_1.insert(tk.END, full_filepath_master)
                self.inFileTxt_L2_1_1.delete(0, tkinter.END)
                self.inFileTxt_L2_1_1.insert(tk.END, full_filepath_master)

                ### check file open
                ns_def.check_file_open(full_filepath_master)

                ###create backup master file
                ns_def.get_backup_filename(full_filepath_master)

                ### l1_device_port_name_sync_with_l1_master
                print('--- Layer1 sync ---')
                self.click_value = '12-3'
                network_sketcher_dev.ns_front_run.click_action(self, '12-3')

                ### l2_device_table_sync_with_l2_master
                print('--- Layer2 sync ---')
                self.click_value = 'L2-2-3'
                network_sketcher_dev.ns_front_run.click_action(self, 'L2-2-3')

                ### l3_device_table_sync_with_l3_master
                print('--- Layer3 sync ---')
                self.click_value = 'L3-2-3'
                network_sketcher_dev.ns_front_run.click_action(self, 'L3-2-3')

                ### l1_device_port_name_sync_with_l2l3_master
                print('--- Port sync(1/3) ---')
                ns_sync_between_layers.l1_device_port_name_sync_with_l2l3_master(self)

                ### l2_device_port_name_sync_with_l3_master
                print('--- Port sync(2/3) ---')
                ns_sync_between_layers.l2_device_port_name_sync_with_l3_master(self)

                ### l2_master_sync_with_l3_master
                print('--- Port sync(3/3) ---')
                ns_sync_between_layers.l2_device_table_sync_with_l3_master(self)

                # attribute table sync to master at ver 2.4.0
                print('--- Attribute sync ---')
                ns_attribute_table_sync_master.ns_attribute_table_sync_master.__init__(self)

                filename = os.path.basename(full_filepath_device)
                ret = tkinter.messagebox.askyesno('Complete', 'Would you like to re-export the Device file?\n\n' + filename)
                if ret == True:
                    ### check file open
                    if ns_def.check_file_open(full_filepath_device) == False:
                        self.click_value_2nd = 'self.sub3_1_button_3'
                        self.click_action_sub('self.self.sub2_5_button_3', push_array)
                        self.click_value_2nd = ''
                        if ns_def.return_os_slash() == '\\\\':  # add ver 2.1.1 for bug fix on Mac OS
                            #print(' # add ver 2.1.1 for bug fix on Mac OS', ns_def.return_os_slash())
                            subprocess.Popen(full_filepath_device, shell=True)

                ### open master panel
                file_type_array = ['EXCEL_MASTER', 'EXCEL_MASTER']
                self.full_filepath = full_filepath_master
                self.filename = os.path.basename(self.full_filepath)
                ns_front_run.sub_excel_master_1(self, file_type_array)

            else:
                self.text_sub3_1.set('[ERROR] Please input the Master file')
                self.sub3_1_label_4 = tk.Label(self.sub3_1, textvariable=self.text_sub3_1, font=("", 10), background="#FBE5D6")
                self.sub3_1_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)

if __name__ == '__main__':
    ns_front_run()
