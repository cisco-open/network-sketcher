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
Subprocess worker for parallel SVG device rendering.

This module has a proper ``if __name__ == '__main__':`` guard,
allowing ProcessPoolExecutor to spawn child processes safely
on Windows even when the parent is a Flask web server.

Protocol (stdin/stdout, JSON):
    Input:  {"devices": [...], "style_index": {...}, "attribute_colors": {...}, "num_workers": N}
    Output: {"svg": "<rendered svg fragment>"}
"""

import json
import sys
import os
from concurrent.futures import ProcessPoolExecutor


def _in(inches):
    return round(inches * 96, 2)


def _rgb(r, g, b):
    return f'rgb({int(r)},{int(g)},{int(b)})'


def _escape(text):
    if text is None:
        return ''
    s = str(text)
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


COLOR_MAP = {
    'ORANGE': [253, 234, 218],
    'BLUE': [220, 230, 242],
    'GREEN': [235, 241, 222],
    'GRAY': [242, 242, 242],
}


def _resolve_color(name, style_index, attribute_colors):
    if '_AIR_' in str(name):
        return [255, 255, 255]
    tag = name
    if '<' in str(name):
        tag = '<' + str(name).split('<')[-1].split('>')[0] + '>'
    color = None
    entry = style_index.get(tag) or style_index.get(name)
    if entry and len(entry) > 3 and entry[3]:
        color = COLOR_MAP.get(entry[3])
    if attribute_colors and tag in attribute_colors:
        rgb = attribute_colors[tag]
        if isinstance(rgb, (list, tuple)) and len(rgb) == 3:
            color = [int(rgb[0]), int(rgb[1]), int(rgb[2])]
    return color


def _render_chunk(args):
    """Render a chunk of devices to SVG string. Runs in worker process."""
    devices, style_index, attribute_colors = args
    pt = 96 / 72.0
    parts = []
    for dev in devices:
        name = dev['name']
        dn = dev['display_name']
        px, py, pw, ph = _in(dev['x']), _in(dev['y']), _in(dev['w']), _in(dev['h'])
        rn = dev.get('roundness', 0.0)
        rx = _in(round(rn * min(dev['w'], dev['h']) * 0.5, 2))
        is_air = '_AIR_' in str(name)
        if is_air:
            parts.append(
                f'<rect x="{px}" y="{py}" width="{pw}" height="{ph}" '
                f'rx="{rx}" ry="{rx}" class="device-air"/>\n')
        else:
            fc = _resolve_color(name, style_index, attribute_colors)
            fill = _rgb(*fc) if fc else 'none'
            parts.append(
                f'<rect x="{px}" y="{py}" width="{pw}" height="{ph}" '
                f'rx="{rx}" ry="{rx}" fill="{fill}" class="device"/>\n')
            tx = round(px + pw / 2, 2)
            ty = round(py + ph / 2, 2)
            parts.append(
                f'<text x="{tx}" y="{ty}" class="device-text">{_escape(dn)}</text>\n')
    return ''.join(parts)


def main():
    raw = sys.stdin.buffer.read()
    data = json.loads(raw)

    devices = data['devices']
    style_index = data['style_index']
    attribute_colors = data['attribute_colors']
    num_workers = data.get('num_workers', max(1, (os.cpu_count() or 1) - 1))

    if len(devices) < 200 or num_workers <= 1:
        result = _render_chunk((devices, style_index, attribute_colors))
    else:
        chunk_size = max(1, len(devices) // num_workers)
        chunks = [devices[i:i + chunk_size] for i in range(0, len(devices), chunk_size)]
        chunk_args = [(c, style_index, attribute_colors) for c in chunks]
        with ProcessPoolExecutor(max_workers=num_workers) as pool:
            results = list(pool.map(_render_chunk, chunk_args))
        result = ''.join(results)

    out = json.dumps({'svg': result})
    sys.stdout.buffer.write(out.encode('utf-8'))
    sys.stdout.buffer.flush()


if __name__ == '__main__':
    main()
