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
import sys, os, subprocess ,webbrowser
import ns_def,network_sketcher_dev,ns_sync_between_layers
import ns_vpn_diagram_create


class ns_front_run():
    '''
    Main Panel
    '''
    def __init__(self):
        self.click_value = ''
        self.click_value_2nd = ''
        self.click_value_3rd = ''
        self.root = TkinterDnD.Tk()
        self.root.title("Network Sketcher  ver 2.1.0")
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
            webbrowser.open('https://github.com/yuhsukeogawa/network-sketcher/blob/main/User_Guide/Japanese/User_Guide%5BJP%5D.md')

        if click_value == 'self.help_1_button_2':
            webbrowser.open('https://github.com/yuhsukeogawa/network-sketcher/blob/main/User_Guide/English/User_Guide%5BEN%5D.md')

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
            if os.path.isfile(self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L2_TABLE]')) == True:
                os.remove(self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L2_TABLE]'))

            ### run L3-1-2 in network_sketcher_dev ,  add l3 master sheet
            self.click_value = 'L3-1-2'
            network_sketcher_dev.ns_front_run.click_action(self,'L3-1-2')

            # remove exist L3/ file
            if os.path.isfile(self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L3_TABLE]')) == True:
                os.remove(self.inFileTxt_L2_1_1.get().replace('[MASTER]', '[L3_TABLE]'))

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
                print('Update to the Master file')

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
        geo =  str(self.root.winfo_width() + 100) + 'x' + str(self.root.winfo_height() + 140) + '+' + str(self.root.winfo_x() + self.root.winfo_width()) + '+' + str(self.root.winfo_y())
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
        self.inFileTxt_11_1 = tk.Entry(self.sub2_1)
        self.outFileTxt_11_2 = tk.Entry(self.sub2_1)
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
        self.sub2_2 = tk.LabelFrame(self.sub2_1, text='Create the L1 Diagram file', font=("", 14), height=1, background="#FBE5D6")
        self.sub2_2.grid(row=1, column=0, columnspan=7, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.sub2_2_button_3 = tk.Button(self.sub2_2, text="All Areas", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'2-4-3'))
        self.sub2_2_button_3.grid(row=2, column=1, sticky='WE', padx=5, pady=2, ipadx=15)
        self.sub2_2_button_4 = tk.Button(self.sub2_2, text="All Areas with IF Tag", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'2-4-4'))
        self.sub2_2_button_4.grid(row=2, column=2, sticky='WE', padx=5, pady=2)
        self.sub2_2_button_1 = tk.Button(self.sub2_2, text="Per Area", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'2-4-1'))
        self.sub2_2_button_1.grid(row=2, column=3, sticky='WE', padx=5, pady=2, ipadx=15)
        self.sub2_2_button_2 = tk.Button(self.sub2_2, text="Per Area with IF Tag", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'2-4-2'))
        self.sub2_2_button_2.grid(row=2, column=4, sticky='WE', padx=5, pady=2)


        ### run L2-3-x for dev , Create L2 diagram
        self.sub2_3 = tk.LabelFrame(self.sub2_1, text='Create the L2 Diagram file', font=("", 14), height=1, background="#FBE5D6")
        self.sub2_3.grid(row=4, column=0, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.sub2_3_button_1 = tk.Button(self.sub2_3, text="Per Area", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'L2-3-2'))
        self.sub2_3_button_1.grid(row=6, column=1, sticky='WE', padx=5, pady=2, ipadx=15)
        self.sub2_3_button_2 = tk.Button(self.sub2_3, text="Per Device", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'L2-3-3'))
        self.sub2_3_button_2.grid(row=6, column=2, sticky='WE', padx=5, pady=2)


        ### run L3-3-x for dev , Create L3 diagram
        self.sub2_4 = tk.LabelFrame(self.sub2_1, text='Create the L3 Diagram file', font=("", 14), height=1, background="#FBE5D6")
        self.sub2_4.grid(row=4, column=3, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.sub2_4_button_1 = tk.Button(self.sub2_4, text="Per Area", font=("", 12), command=lambda: network_sketcher_dev.ns_front_run.click_action(self,'L3-3-2'))
        self.sub2_4_button_1.grid(row=6, column=3, sticky='WE', padx=50, pady=2, ipadx=15)


        ### run 11-4 for dev , Create Device file
        self.sub2_0_label_2 = tk.Label(self.sub2_1, text='', font=("", 6))
        self.sub2_0_label_2 .grid(row=7, column=0, columnspan=7, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.sub2_5 = tk.LabelFrame(self.sub2_1, text='Export to the Device file', font=("", 14), height=1, background="#DFC9EF")
        self.sub2_5.grid(row=8, column=0, columnspan=7, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        push_array = []
        self.sub2_5_button_3 = tk.Button(self.sub2_5, text="Device file", font=("", 12), command=lambda: self.click_action_sub('self.self.sub2_5_button_3', push_array))

        self.sub2_5_button_3.grid(row=10, column=1, sticky='WE', padx=50, pady=2, ipadx=15)

        ### run xx-xx for dev , Create VPN diagram
        self.sub2_6 = tk.LabelFrame(self.sub2_1, text='Create the VPN diagram file', font=("", 14), height=1, background="#FFF2CC")
        self.sub2_6.grid(row=8, column=1, columnspan=7, sticky='W', padx=5, pady=2, ipadx=5, ipady=2)

        self.sub2_6_button_3 = tk.Button(self.sub2_6, text="VPNs on L1", font=("", 12), command=lambda: self.click_action_sub('self.self.sub2_6_button_1', push_array))
        self.sub2_6_button_3.grid(row=10, column=1, sticky='WE', padx=50, pady=2, ipadx=15)

        #self.sub2_6_button_4 = tk.Button(self.sub2_6, text="VPN only", font=("", 12), command=lambda: self.click_action_sub('self.self.sub2_6_button_2', push_array))
        #self.sub2_6_button_4.grid(row=10, column=2, sticky='WE', padx=5, pady=2, ipadx=15)



    def click_action_sub(self, click_value, push_array):
        if click_value == 'self.self.sub2_5_button_3':  # select browse
            ### check file open
            if ns_def.check_file_open(str(self.outFileTxt_11_2.get()).replace('[MASTER]','')) == True:
                return ()

            ### create device file
            self.click_value = '11-4'
            network_sketcher_dev.ns_front_run.click_action(self, '11-4')
            # run x-x for dev , Create L2 Table file
            self.click_value = 'L2-1-2'
            network_sketcher_dev.ns_front_run.click_action(self, 'L2-1-2')
            # run x-x for dev , Create L3 Table file
            self.click_value = 'L3-1-2'
            network_sketcher_dev.ns_front_run.click_action(self, 'L3-1-2')

            if self.click_value_2nd != 'self.sub1_1_button_1' and self.click_value_2nd != 'self.sub3_1_button_3':
                ns_def.messagebox_file_open(str(self.outFileTxt_11_2.get()).replace('[MASTER]',''))

        if click_value == 'self.self.sub2_6_button_1':  # Click "VPNs on L1"
            print('--- Click "VPNs on L1" ---')
            ### create L1 Table with [VPNs_on_L1]]
            self.click_value = 'VPN-1-1'
            network_sketcher_dev.ns_front_run.click_action(self, '2-4-3')

            ### Write VPNs on L1 ###
            ns_vpn_diagram_create.ns_write_vpns_on_l1.__init__(self)

        if click_value == 'self.self.sub2_6_button_2':  # Click "VPN only"
            print('--- Click "VPN only ---')

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
        geo =  str(self.root.winfo_width() - 180) + 'x' + str(self.root.winfo_height()) + '+' + str(self.root.winfo_x() + self.root.winfo_width()) + '+' + str(self.root.winfo_y() + self.root.winfo_height() + 30)
        self.sub3_1.geometry(geo)

        self.sub3_0_label_1 = tk.Label(self.sub3_1, text=local_filename, font=("", 12), background="#FFFFFF")
        self.sub3_0_label_1 .grid(row=0, column=0, columnspan=7, sticky='W', padx=5, pady=5, ipadx=30, ipady=5)

        self.sub3_1 = tk.LabelFrame(self.sub3_1, text='Update to the Master file', font=("", 14), height=1, background="#E5F4F7")
        self.sub3_1.grid(row=1, column=1, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        self.text_sub3_1 = tk.StringVar()
        self.text_sub3_1.set("      drag and drop here ([MASTER]*.xlsx)")
        self.sub3_1_label_4 = tk.Label(self.sub3_1, textvariable=self.text_sub3_1, font=("", 10), background="#E5F4F7")
        self.sub3_1_label_4.grid(row=2, column=1, columnspan=3, sticky='W', padx=5, pady=20)

        self.sub3_1_entry_1 = tk.Entry(self.sub3_1)
        self.sub3_1_entry_1.grid(row=4, column=1, sticky="WE", padx=5, pady=3, ipadx=15)
        self.sub3_1_button_1 = tk.Button(self.sub3_1, text="Browse ...", command=lambda: self.click_action_sub3_1('self.sub3_1_button_2',push_array))
        self.sub3_1_button_1.grid(row=4, column=2, sticky='W', padx=5, pady=2)
        self.sub3_1_button_1 = tk.Button(self.sub3_1, text="Submit", command=lambda: self.click_action_sub3_1('self.sub3_1_button_3',push_array))
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
                print('Update to the Master file')

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
                if ns_def.check_file_open(full_filepath_master) == True:
                    return ()

                ###create backup master file
                ns_def.get_backup_filename(full_filepath_master)

                ### l1_device_port_name_sync_with_l1_master
                self.click_value = '12-3'
                network_sketcher_dev.ns_front_run.click_action(self, '12-3')

                ### l2_device_table_sync_with_l2_master
                self.click_value = 'L2-2-3'
                network_sketcher_dev.ns_front_run.click_action(self, 'L2-2-3')

                ### l3_device_table_sync_with_l3_master
                self.click_value = 'L3-2-3'
                network_sketcher_dev.ns_front_run.click_action(self, 'L3-2-3')

                ### l1_device_port_name_sync_with_l2l3_master
                ns_sync_between_layers.l1_device_port_name_sync_with_l2l3_master(self)

                ### l2_device_port_name_sync_with_l3_master
                ns_sync_between_layers.l2_device_port_name_sync_with_l3_master(self)

                ### l2_master_sync_with_l3_master
                ns_sync_between_layers.l2_device_table_sync_with_l3_master(self)

                filename = os.path.basename(full_filepath_device)
                ret = tkinter.messagebox.askyesno('Complete', 'Would you like to re-export and open the Device file?\n\n' + filename)
                if ret == True:
                    ### check file open
                    if ns_def.check_file_open(full_filepath_device) == False:
                        self.click_value_2nd = 'self.sub3_1_button_3'
                        self.click_action_sub('self.self.sub2_5_button_3', push_array)
                        self.click_value_2nd = ''
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
