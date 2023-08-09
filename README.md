


# Network Sketcher　
Network Sketcher that helps make network design and configuration management faster, more accurate, and easier.
Simply create a rough sketch first, and the system will automatically generate L1L2L3 network diagrams and management tables.
Since the network components are consolidated in one master file, updating the management table (device file) automatically updates all related network diagrams and management tables as well.
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/240ddee0-823d-472f-87d4-8ae7eb1fff7d)



# DEMO
https://github.com/cisco-open/network-sketcher/assets/13013736/56015621-93ef-4df5-8816-e79488df7f53



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
git clone https://github.com/cisco-open/network-sketcher/
cd network-sketcher
python network_sketcher.py
```

# SAMPLE
## Input ppt file (rough sketch)
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/0ef66092-0995-47fa-8849-e1f640a35ab4)
[Sample-rough-sketch.pptx](https://github.com/cisco-open/network-sketcher/files/12298813/Sample-rough-sketch.pptx)

## Output
### Device table(Excel)
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/2c657c04-329a-4049-8181-399051aba91c)
[DEVICE.Sample.figure5.xlsx](https://github.com/cisco-open/network-sketcher/files/12298814/DEVICE.Sample.figure5.xlsx)

### L1 figure(PPT)
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/69fdbac5-8482-4486-ab47-c1bd01cc11e2)
[L1_DIAGRAM.AllAreasTag_Sample.figure5.pptx](https://github.com/cisco-open/network-sketcher/files/12298815/L1_DIAGRAM.AllAreasTag_Sample.figure5.pptx)

### L2 figure(PPT)
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/fd98050b-30f4-4544-9cac-62f848b82833)
[L2_DIAGRAM.PerArea_Sample.figure5.pptx](https://github.com/cisco-open/network-sketcher/files/12298817/L2_DIAGRAM.PerArea_Sample.figure5.pptx)

### L3 figure(PPT)
![image](https://github.com/cisco-open/network-sketcher/assets/13013736/bc63ec4e-fb46-4f05-b28c-8395adf6aaaa)
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
