


# Network Sketcher　
Network Sketcher that helps make network design and configuration management faster, more accurate, and easier.
Simply create a rough sketch first, and the system will automatically generate L1L2L3 network diagrams and management tables.
Since the network components are consolidated in one master file, updating the management table (device file) automatically updates all related network diagrams and management tables as well.
![image](https://github.com/yuhsukeogawa/network-sketcher/assets/13013736/3e106a72-ad5d-4e58-b7a6-e783a6dcafc5)


# DEMO
https://github.com/yuhsukeogawa/network-sketcher/assets/13013736/3d2a4c8b-a783-468a-ab56-d528b0af1ac3

# Requirement
 
__Currently, the OS that this tool runs on is Windows only. Linux and MAC OS cannot be used because file paths are not compatible now.__
 
* tkinterdnd2
* openpyxl
* python-pptx
* ipaddress
* numpy
 
# Installation
 
```bash
pip install tkinterdnd2
pip install openpyxl
pip install python-pptx
pip install ipaddress
pip install numpy
```
or
```bash
pip install -r requirements.txt
```
 
# Usage
 
```bash
git clone https://github.com/yuhsukeogawa/network-sketcher/
cd network-sketcher
python network_sketcher.py
```

# SAMPLE
## Input ppt file (rough sketch)
![image](https://github.com/yuhsukeogawa/network-sketcher/assets/13013736/43f4af0c-8414-4aea-921b-f06e88e9bbdd)
[Sample-rough-sketch.pptx](https://github.com/yuhsukeogawa/network-sketcher/files/12158449/Sample-rough-sketch.pptx)

## Output
### Device table(Excel)
![image](https://github.com/yuhsukeogawa/network-sketcher/assets/13013736/72614b8f-0519-4e2f-942f-bed2c10c1cde)
[[DEVICE]Sample figure5.xlsx](https://github.com/yuhsukeogawa/network-sketcher/files/12158439/DEVICE.Sample.figure5.xlsx)

### L1 figure(PPT)
![image](https://github.com/yuhsukeogawa/network-sketcher/assets/13013736/7908f403-0347-4a08-adeb-1157ed2e5b1a)
[[L1_DIAGRAM]AllAreasTag_Sample figure5.pptx](https://github.com/yuhsukeogawa/network-sketcher/files/12158436/L1_DIAGRAM.AllAreasTag_Sample.figure5.pptx)

### L2 figure(PPT)
![image](https://github.com/yuhsukeogawa/network-sketcher/assets/13013736/dc67d7a5-dc55-4544-a0ac-7baf7b4ce04d)
[[L2_DIAGRAM]PerArea_Sample figure5.pptx](https://github.com/yuhsukeogawa/network-sketcher/files/12158437/L2_DIAGRAM.PerArea_Sample.figure5.pptx)

### L3 figure(PPT)
![image](https://github.com/yuhsukeogawa/network-sketcher/assets/13013736/ead89f2b-879b-4c24-a850-84fded20dbb9)
[[L3_DIAGRAM]PerArea_Sample figure5.pptx](https://github.com/yuhsukeogawa/network-sketcher/files/12158438/L3_DIAGRAM.PerArea_Sample.figure5.pptx)

### Master file(Excel)
[[MASTER]Sample figure5.xlsx](https://github.com/yuhsukeogawa/network-sketcher/files/12158535/MASTER.Sample.figure5.xlsx)


# User Guide
| Lang  | Link |
| ------------- | ------------- |
| English  | [Link](https://github.com/yuhsukeogawa/network-sketcher/blob/main/User_Guide/English/User_Guide%5BEN%5D.md) |
| 日本語  | [Link](https://github.com/yuhsukeogawa/network-sketcher/blob/main/User_Guide/Japanese/User_Guide%5BJP%5D.md) |
 
# Note
* How to make the exe file for Windows using pyinstaller
 ```bash
pyinstaller.exe [file path]/network_sketcher.py --onefile --collect-data tkinterdnd2 --noconsole --additional-hooks-dir  [file path] --clean
 ```

# Author
 
* Yusuke Ogawa
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
