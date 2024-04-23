


# Network Sketcher
**Network Skecher makes network design and configuration management faster, more accurate, and easier. Simply create your first rough sketch and it will automatically generate L1L2L3 network diagrams and management tables. Network components are integrated into one master file, so when you update the management table (device file), all related network diagrams and management tables are automatically updated as well.**
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/240ddee0-823d-472f-87d4-8ae7eb1fff7d)

# DEMO
https://github.com/cisco-open/network-sketcher/assets/13013736/b76ec8fa-44ad-4d02-a7c2-579f67ad24a9

# New Features
- Ver 2.2.0
  - [Automatic IP address assignment](https://github.com/cisco-open/network-sketcher/blob/ver_2.2.0/User_Guide/English/7-2.%20Automatic%20IP%20address%20assignment.md)
- Ver 2.1.0
  - [VPN Diagram](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/6-1.%20Generation%20of%20VPN%20Diagram%20.md) 
  - [VPN Configuration](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/6-2.%20VPN%20setting.md) 
  - Cross-platform support (Windows, Linux, Mac OS)
  - [Drawing beyond maximum PowerPoint size](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/A-1.%20Procedure%20for%20pasting%20PPT%20figures%20that%20exceed%20the%20maximum%20paper%20size%20into%20Excel.md)
  - [Import of yaml file from CML(Cisco Modeling Labs) diagrams](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/7-1.%20Convert%20CML%20configuration%20file%20(yaml)%20to%20Network%20Sketcher%20master%20file.md) 

# Limitations
- IPv4 only. IPv6 is not supported.
- Importing CML yaml files is only supported in L1 and L3(IP addresses of physical interfaces of ios, asa, iosxr, and csr1k).
- VPN diagrams are only supported for Layer 1 and not reflected in Layer 2 and Layer 3 diagrams.
 
# Requirement
- __Python 3.x__
- __Network Sketcher supports cross-platform. Works with Windows, Mac OS, and Linux.__

# Installation
 * Required
```bash
pip install tkinterdnd2
pip install tkinterdnd2-universal
pip install openpyxl
pip install python-pptx
pip install ipaddress
pip install numpy
pip install pyyaml
pip install ciscoconfparse
```
or
```bash
pip install -r requirements.txt
```

* Mac OS requires the following additional installation.
```bash
brew install tcl-tk
```

* Linux requires the following additional installation.
```bash
sudo apt-get install python3-tk
```
 
# Usage
 
```bash
git clone https://github.com/cisco-open/network-sketcher/
cd network-sketcher
python network_sketcher.py
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


# User Guide
| Lang  | Link |
| ------------- | ------------- |
| English  | [Link](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/English/User_Guide%5BEN%5D.md) |
| 日本語  | [Link](https://github.com/cisco-open/network-sketcher/blob/main/User_Guide/Japanese/User_Guide%5BJP%5D.md) |
 
# Note
* How to make the exe file for Windows using pyinstaller
 ```bash
pyinstaller.exe [file path]/network_sketcher.py --onefile --collect-data tkinterdnd2 --noconsole --additional-hooks-dir  [file path] --clean
 ```

# Author
 
* Yusuke Ogawa
* CCIE# 17583
* Security Architect @ Cisco
* yuogawa@cisco.com
 
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
