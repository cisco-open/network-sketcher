#!/bin/bash

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

if ! command -v docker &> /dev/null
then
    echo "Docker is not installed. Please install Docker."
    exit 1
fi

docker run --rm -v "$(pwd):/src/" ghcr.io/batonogov/pyinstaller-windows -c \
  "pip install -r requirements.txt && \
  pyinstaller network_sketcher.py --onefile --collect-data tkinterdnd2 --noconsole --additional-hooks-dir . --clean && \
  ls -la /src/dist/&& \
  mv /src/dist/network_sketcher.exe network_sketcher.exe && \
  rm -rf __pycache__/ build/ dist/ network_sketcher.spec"

