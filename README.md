
https://github.com/user-attachments/assets/9ff207f8-c6b3-4584-b166-98ae4e4c8297

NoLang (no-lang.com)
Otologic (https://otologic.jp) CC BY 4.0

<p align="center">
  <img src="https://github.com/user-attachments/assets/cc82082d-c4a5-4f13-90f5-adaf162202b2" alt="image" />
</p>

# Network Sketcher
**Network Sketcher generates network configuration diagrams in PowerPoint and manages configuration information in Excel. Additionally, exporting a AI ​​context can be used to generate config files using LLM.**
* Automatic generation of each configuration document by metadatization of network configuration information
* Automated synchronization between documents
* Minimize maintenance and training load by automatic generation of common formats
* Facilitate automatic analysis, AI utilization, and inter-system collaboration by metadatization of configuration information.
* Template support for equipment configuration
![image](https://github.com/user-attachments/assets/9f497061-08ee-4c78-9040-d5b37d2f3e69)

![image](https://github.com/cisco-open/network-sketcher/assets/13013736/240ddee0-823d-472f-87d4-8ae7eb1fff7d)

# DEMO
https://github.com/cisco-open/network-sketcher/assets/13013736/b76ec8fa-44ad-4d02-a7c2-579f67ad24a9

# New Features
- Ver 2.6.1<br>
[Network Sketcher Ver 2.6.1 supported the creation of a network configuration with LLM from scratch](https://github.com/cisco-open/network-sketcher/releases/tag/Ver2.6.1)

<img width="1423" height="806" alt="image" src="https://github.com/user-attachments/assets/761072de-d64b-4772-bdc7-6224f53fddd8" />


- Ver 2.6.0<br>

[Network Sketcher Ver 2.6.0 now supports master file conversion from Visio, draw.io, NetBox, and CML data to Network Sketcher.](https://github.com/cisco-open/network-sketcher/wiki/1%E2%80%904.-Convert-data-from-other-systems-into-master-files-(Visio,-draw.io,-NetBox-and-CML)_en)

<img alt="image" src="https://github.com/user-attachments/assets/436a1462-bdf7-49cf-bc4f-235be6cb7d42" />
Although Network Sketcher now supports multiple formats, it is not intended to replace the main drawing tool, but rather aims for mutually beneficial development.

    
- Ver 2.5.0
  - [Communication flow management functionality has been added.](https://github.com/cisco-open/network-sketcher/wiki/9%E2%80%901.Exporting-Flow-files)
![image](https://github.com/user-attachments/assets/8683c172-505e-4af8-a87a-dc1a1a86a121)

# Limitations
- IPv4 only. IPv6 is not supported.
- A DEVICE file contains multiple sheets, but only one sheet should be updated at a time. Simultaneous synchronization of multiple sheet updates is not supported.
- Do not use Network Skecher on master files in your One Drive folder.
- Deleting Layer 1 links using the GUI cannot identify individual interfaces and will delete more Layer 2 data than intended. Use the CLI command (delete l1_link) to delete Layer 1 links.
 
# Requirement
- __Network Sketcher supports cross-platform. Works with Windows, Mac OS, and Linux.__
  - MAC OS may not display well in Dark mode.
- __Python ver 3.x__
- __Software that can edit .pptx and .xlsx files__
  - Microsoft Powerpoint and Excel are the best
  - Google Slides and Spreadsheets import/export functionality is available. Excel functions display will show an error, but it works fine.
  - Libre Office and Softmaker office cannot be used.

# Installation
```bash
git clone https://github.com/cisco-open/network-sketcher/
cd network-sketcher
python3 -m pip install -r requirements.txt
python3 network_sketcher.py
```
or
```bash
#Download via browser
https://github.com/cisco-open/network-sketcher/archive/refs/heads/main.zip

#Unzip the ZIP file and execute the following in the prompt of the folder
python3 -m pip install -r requirements.txt
python3 network_sketcher.py
```

# Installation Supplement
 * Alternative to “python -m pip install -r requirements.txt”
```bash
python3 -m pip install tkinterdnd2
python3 -m pip install "openpyxl>=3.1.3,<=3.1.5"
python3 -m pip install python-pptx
python3 -m pip install ipaddress
python3 -m pip install numpy
python3 -m pip install pyyaml
python3 -m pip install ciscoconfparse
python3 -m pip install networkx
python3 -m pip install svg.path
```

* Mac OS requires the following additional installation.
```bash
brew install tcl-tk
brew install tkdnd
```
* Ubuntu requires the following additional installation.<br>
  GUI drag and drop doesn't work on Ubuntu, you need to compile tkdnd from source or use "Browse" and "Submit".
```bash
sudo apt-get install python3-tk
```
   

# User Guide
| Language  | Link |
| ------------- | ------------- |
| English  | [Link](https://github.com/cisco-open/network-sketcher/wiki/User_Guide%5BEN%5D) |
| Japanese  | [Link](https://github.com/cisco-open/network-sketcher/wiki/User_Guide%5BJP%5D) |
<br>
 
# How to create the exe file for Windows using pyinstaller
 ```bash
pyinstaller.exe [file path]/network_sketcher.py --onefile --collect-data tkinterdnd2 --additional-hooks-dir  [file path] --clean --add-data "./ns_extensions_cmd_list.txt;." --add-data "./ns_logo.png;."
 ```
<br>

# Performance Measurement Summary

| Item                                               | 16 Devices (4x4) | 64 Devices (8x8) | 256 Devices (16x16) | 1024 Devices (32x32) |
|----------------------------------------------------|-----------:|-----------:|------------:|-------------:|
| Master file creation (Reflect only L1 information in the no_data master file. Connect adjacent devices. Measure command execution time.)                               | 1m55s      | 7m36s      | 50m31s      | TBD          |
| Layer 1 diagram generation (All Areas with tags)   | 2s         | 5s         | 27s         | TBD          |
| Layer 2 diagram generation                         | 3s         | 14s        | 9m22s       | TBD          |
| Layer 3 diagram generation (All Areas)             | 13s        | 29s        | 2m42s       | TBD          |
| Device file export                                 | 5s         | 14s        | 51s         | TBD          |

---
Status: Python code optimization, multi-processor support, and GPU support are not yet implemented. <br>
Network Skecher Ver: 2.6.1a <br>
Test environment: Intel Core Ultra 7 (1.70 GHz), 32.0 GB RAM, Windows 11 Enterprise <br>

# GUI vs. CLI Feature Support Matrix

| Feature Item | GUI | CLI (AI Context) |
| --- | --- | --- |
| Create master file from PowerPoint rough sketch | ✅ | ❌ |
| Convert master files from Visio, Draw.io, NetBox, CML | ✅ | ❌ |
| Area placement | ✅ (automatic) | ✅ (user-specified) |
| Create / delete / modify areas | ✅ | ✅ |
| Place / create / delete / modify devices | ✅ | ✅ |
| Place / create / delete / modify waypoints | ✅ | ✅ |
| Add Layer 1 connections | ✅ | ✅ |
| Delete Layer 1 connections | ⚠️ (port cannot be specified) | ✅ |
| Change Layer 1 port names | ✅ | ✅ |
| Change Layer 1 connection details (e.g., duplex) | ✅ | ❌ |
| Change Layer 2 segments (VLAN) | ✅ | ✅ |
| Add / delete virtual ports (SVI, loopback, port-channel) | ✅ | ✅ |
| Change IP addresses / Layer 3 instances (VRF) | ✅ | ✅ |
| Change attributes | ✅ | ❌ |
| Add / delete VPNs | ✅ | ❌ |
| Flow management | ✅ | ❌ |
| Export various reports | ✅ | ❌ |
| Export empty master files (no data) | ❌ | ✅ |
| Export AI context files | ✅ | ✅ |
| Export device files | ✅ | ❌ |
| Generate L1/L2/L3 topology diagrams | ✅ | ❌ |

# SAMPLE
## Input ppt file (rough sketch)
![image](https://github.com/user-attachments/assets/35e13f4b-d81e-42df-a036-b018b47a199a)

[Sample-rough5.pptx](https://github.com/user-attachments/files/18668273/Sample-rough5.pptx)

## Output
### L1 figure(PPT)
![image](https://github.com/user-attachments/assets/e28aef48-411c-48fe-8700-336b298a658f)

[[L1_DIAGRAM]AllAreasTag_Sample.figure5.pptx](https://github.com/user-attachments/files/18611145/L1_DIAGRAM.AllAreasTag_Sample.figure5.pptx)

### L2 figure(PPT)
![image](https://github.com/user-attachments/assets/8a62d5ed-f244-4e87-85a0-89925aaa339f)

[[L2_DIAGRAM]DC-TOP1_Sample.figure5.pptx](https://github.com/user-attachments/files/18611147/L2_DIAGRAM.DC-TOP1_Sample.figure5.pptx)

### L3 figure(PPT)
![image](https://github.com/user-attachments/assets/0e0b6e8c-628b-4af5-a20a-a940eab4877a)

[[L3_DIAGRAM]AllAreas_old_Sample.figure5.pptx](https://github.com/user-attachments/files/18611149/L3_DIAGRAM.AllAreas_old_Sample.figure5.pptx)

### Device table(Excel)
![image](https://github.com/user-attachments/assets/33f95b5c-03d3-419e-bbee-5786efe9deb7)

[[DEVICE]Sample.figure5.xlsx](https://github.com/user-attachments/files/18611140/DEVICE.Sample.figure5.xlsx)

### Master file(Excel)
[[MASTER]Sample.figure5.xlsx](https://github.com/user-attachments/files/18918397/MASTER_Sample.figure5.zip)

# Author
 
* Yusuke Ogawa - Security Architect, Cisco | CCIE#17583
 
# License
SPDX-License-Identifier: Apache-2.0

Copyright 2023  Cisco Systems, Inc. and its affiliates

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.





























