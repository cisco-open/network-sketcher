


# Network Sketcher
**Network Sketcher makes network design and configuration management faster, more accurate, and easier. Simply create your first rough sketch and it will automatically generate L1L2L3 network diagrams and management tables. Network components are integrated into one master file, so when you update the management table (device file), all related network diagrams and management tables are automatically updated as well.**

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
- Ver 2.3.4
  - [Summary Diagram is automatically added in L1 All Areas](https://github.com/cisco-open/network-sketcher/releases/tag/Ver2.3.4)
- Ver 2.3.3
  - [Two patterns are created: area-oriented and connection-oriented](https://github.com/cisco-open/network-sketcher/releases/tag/Ver2.3.3)
- Ver 2.3.2
  - ["VPNs on L3” has been added to the VPN configuration diagram to reflect VPN segments in the Layer 3 logical configuration diagram.](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/6-2.%20VPN%20setting.md)
- Ver 2.3.1
  - [Added the ability to output information in the master file via CLI](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/8-1.%20show%20commands.md)
- Ver 2.3.0
  - [Added the ability to create an overall L3 diagram (All Areas)](https://github.com/cisco-open/network-sketcher/releases/tag/Ver2.3.0)
- Ver 2.2.2
  - [Automated device placement adjustment in L1 diagram](https://github.com/user-attachments/assets/8014132b-4e38-422a-9ab0-3ee397e27b1e)
- Ver 2.2.1
  - [IP address table](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/7-3.%20Export%20IP%20address%20table.md) :[(Sample output file) ](https://github.com/user-attachments/files/15784654/IP_TABLE.Sample.figure5.xlsx)
- Ver 2.2.0
  - [Automatic IP address assignment](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/7-2.%20Automatic%20IP%20address%20assignment.md)
- Ver 2.1.0
  - [VPN Diagram](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/6-1.%20Generation%20of%20VPN%20Diagram%20.md) 
  - [VPN Configuration](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/6-2.%20VPN%20setting.md) 
  - [Cross-platform support (Windows, Linux, Mac OS)](https://github.com/user-attachments/assets/04ab332d-c876-44ff-83ef-b98d64d24b1f)
  - [Drawing beyond maximum PowerPoint size](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/A-1.%20Procedure%20for%20pasting%20PPT%20figures%20that%20exceed%20the%20maximum%20paper%20size%20into%20Excel.md)
  - [Import of yaml file from CML(Cisco Modeling Labs) diagrams](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/7-1.%20Convert%20CML%20configuration%20file%20(yaml)%20to%20Network%20Sketcher%20master%20file.md) 

# Limitations
- IPv4 only. IPv6 is not supported.
- Importing CML yaml files is only supported in L1 and L3(IP addresses of physical interfaces of ios, asa, iosxr, and csr1k).
 
# Requirement
- __Network Sketcher supports cross-platform. Works with Windows, Mac OS, and Linux.__
- __Python ver 3.x__
- __Software that can edit .pptx and .xlsx files__
  - Microsoft Powerpoint and Excel are best
  - Google Slides and Spreadsheets import/export functionality is available. Excel functions display will show an error, but it works fine.
  - Libre Office and Softmaker office cannot be used.

# Installation
```bash
git clone https://github.com/cisco-open/network-sketcher/
cd network-sketcher
python -m pip install -r requirements.txt
python network_sketcher.py
```
or
```bash
#Download via browser
https://github.com/cisco-open/network-sketcher/archive/refs/heads/main.zip

#Unzip the ZIP file and execute the following in the prompt of the folder
python -m pip install -r requirements.txt
python network_sketcher.py
```

# Installation Supplement
 * Alternative to “python -m pip install -r requirements.txt”
```bash
python -m pip install tkinterdnd2
python -m pip install tkinterdnd2-universal
python -m pip install openpyxl
python -m pip install python-pptx
python -m pip install ipaddress
python -m pip install numpy
python -m pip install pyyaml
python -m pip install ciscoconfparse
```

* Mac OS requires the following additional installation.
```bash
brew install tcl-tk
```

* Linux requires the following additional installation.
```bash
sudo apt-get install python3-tk
```

# User Guide
| Lang  | Link |
| ------------- | ------------- |
| English  | [Link](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/User_Guide%5BEN%5D.md) |
| 日本語  | [Link](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/Japanese/User_Guide%5BJP%5D.md) |
 
# How to create the exe file for Windows using pyinstaller
 ```bash
pyinstaller.exe [file path]/network_sketcher.py --onefile --collect-data tkinterdnd2 --noconsole --additional-hooks-dir  [file path] --clean
 ```

# SAMPLE
## Input ppt file (rough sketch)
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/87e792ec-f0d6-47f9-96ad-1dfda4de5228)

[Sample-rough-sketch.pptx](https://github.com/cisco-open/network-sketcher/files/12298813/Sample-rough-sketch.pptx)

## Output
### Device table(Excel)
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/f806bbe0-9a53-4030-aca1-f4df7a616a34)

[DEVICE.Sample.figure5.xlsx](https://github.com/cisco-open/network-sketcher/files/12298814/DEVICE.Sample.figure5.xlsx)

### L1 figure(PPT)
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/4bc62cdd-ca32-4bc0-a70a-ceaa24faf9e8)

[L1_DIAGRAM.AllAreasTag_Sample.figure5.pptx](https://github.com/cisco-open/network-sketcher/files/12298815/L1_DIAGRAM.AllAreasTag_Sample.figure5.pptx)

### L2 figure(PPT)
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/0fd450b2-02df-42cc-9750-b9e43cf35525)

[L2_DIAGRAM.PerArea_Sample.figure5.pptx](https://github.com/cisco-open/network-sketcher/files/12298817/L2_DIAGRAM.PerArea_Sample.figure5.pptx)

### L3 figure(PPT)
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/78519067-f42c-4f3f-ac1d-602f769adddd)

[L3_DIAGRAM.PerArea_Sample.figure5.pptx](https://github.com/cisco-open/network-sketcher/files/12298818/L3_DIAGRAM.PerArea_Sample.figure5.pptx)

### Master file(Excel)
[MASTER.Sample.figure5.xlsx](https://github.com/cisco-open/network-sketcher/files/12298821/MASTER.Sample.figure5.xlsx)


# Author
 
* Yusuke Ogawa
* CCIE# 17583
* Security Architect @ Cisco
 
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
