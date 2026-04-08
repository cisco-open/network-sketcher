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

"""
L3 SVG style definitions and rendering helpers.

Provides shape/line style dictionaries and SVG primitive renderers
used by nsm_l3_svg_create to convert captured coordinate data to SVG.

Style values are derived from nsm_ddx_figure.extended.add_shape / add_line
to ensure pixel-accurate correspondence with the PowerPoint output.

Shape adjustments[0] in PPTX -> rx ratio in SVG:
  0.0      -> 0.0 (sharp corners)
  0.0001   -> 0.0001 (near-sharp)
  0.0008   -> 0.0008
  0.015    -> 0.015
  0.2002   -> 0.2002 (moderately rounded)
  0.30045  -> 0.30 (well-rounded)
  0.50445  -> 0.50 (circular ends)
"""

DPI = 96
PT_TO_PX = DPI / 72.0
FONT_FAMILY = 'Calibri, Segoe UI, Arial, sans-serif'


# ---- L3 shape styles ----
# Keys match shape_type values passed to extended.add_shape.
# Values sourced from nsm_ddx_figure.py extended.add_shape (lines 2724-2979).
#
# fill:   (R,G,B) or None (background/transparent)
# stroke: (R,G,B) or None (no border)
# sw:     stroke-width in points
# rx:     corner-radius ratio (adjustments[0] value from PPTX)
# fs:     font-size in points
# tc:     text color (R,G,B)
# align:  'left' or None (center)

L3_STYLES = {
    # DEVICE_NORMAL: fill=(235,241,222) default, overridden by attribute color at runtime
    'DEVICE_NORMAL':      {'fill': (235, 241, 222), 'stroke': (0, 0, 0),       'sw': 1.0,  'rx': 0.0001,  'tc': (0, 0, 0), 'fs': 8},
    # DEVICE_L3_INSTANCE: fill=(250,251,247) default, overridden by lightened attribute color
    'DEVICE_L3_INSTANCE': {'fill': (250, 251, 247), 'stroke': (0, 0, 0),       'sw': 0.5,  'rx': 0.0008,  'tc': (0, 0, 0), 'fs': 6, 'align': 'left'},
    # WAY_POINT_NORMAL: fill=(220,230,242) default, overridden by attribute color
    'WAY_POINT_NORMAL':   {'fill': (220, 230, 242), 'stroke': (0, 0, 0),       'sw': 1.0,  'rx': 0.2002,  'tc': (0, 0, 0), 'fs': 8},
    # L3_INSTANCE: rounded purple-tinted box
    'L3_INSTANCE':        {'fill': (230, 224, 236), 'stroke': (0, 0, 0),       'sw': 1.0,  'rx': 0.2007,  'tc': (0, 0, 0), 'fs': 6},
    # TAG_NORMAL: white pill with black border (adjustments=0.50445)
    'TAG_NORMAL':         {'fill': (255, 255, 255), 'stroke': (0, 0, 0),       'sw': 0.5,  'rx': 0.50,    'tc': (0, 0, 0), 'fs': 4},
    # IP_ADDRESS_TAG: no fill, no visible border
    'IP_ADDRESS_TAG':     {'fill': None,            'stroke': None,            'sw': 0.75, 'rx': 0.0,     'tc': (0, 0, 0), 'fs': 6, 'align': 'left'},
    # L3_SEGMENT_GRAY: light gray rounded segment bar
    'L3_SEGMENT_GRAY':    {'fill': (249, 249, 249), 'stroke': (0, 0, 0),       'sw': 0.75, 'rx': 0.30,    'tc': (0, 0, 0), 'fs': 4},
    # L3_SEGMENT_VPN: purple-tinted segment bar
    'L3_SEGMENT_VPN':     {'fill': (248, 243, 251), 'stroke': (112, 48, 160),  'sw': 0.75, 'rx': 0.30,    'tc': (0, 0, 0), 'fs': 4},
    # FOLDER_NORMAL: transparent fill, gray dashed border
    'FOLDER_NORMAL':      {'fill': None,            'stroke': (205, 205, 205), 'sw': 1.0,  'rx': 0.015,   'tc': (0, 0, 0), 'fs': 10},
    # OUTLINE_NORMAL: white fill, black border (outermost box)
    'OUTLINE_NORMAL':     {'fill': (255, 255, 255), 'stroke': (0, 0, 0),       'sw': 1.0,  'rx': 0.0,     'tc': (0, 0, 0), 'fs': 6},
}


# ---- L3 line styles ----
# Keys match line_type values passed to extended.add_line.
# Values sourced from nsm_ddx_figure.py extended.add_line (lines 2981-3050).

L3_LINE_STYLES = {
    # L3_SEGMENT: thick black line with diamond markers at both ends
    'L3_SEGMENT':        {'stroke': (0, 0, 0),       'sw': 2.5,  'marker_start': 'diamond', 'marker_end': 'diamond'},
    # L3_SEGMENT-L3IF: thin black line from tag to segment, diamond at end
    'L3_SEGMENT-L3IF':   {'stroke': (0, 0, 0),       'sw': 0.7,  'marker_end': 'diamond'},
    # L3_SEGMENT-VPN: purple line from VPN tag to segment, diamond at end
    'L3_SEGMENT-VPN':    {'stroke': (112, 48, 160),   'sw': 0.7,  'marker_end': 'diamond'},
    # L3_INSTANCE: purple connecting line between tag and instance box
    'L3_INSTANCE':       {'stroke': (96, 74, 123),    'sw': 0.7},
}
