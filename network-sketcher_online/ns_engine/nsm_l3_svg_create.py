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
L3 SVG diagram creation module — python-pptx free rendering.

Uses mock PPTX objects so the existing L3 layout engine
(nsm_l3_diagram_create) runs its full coordinate computation
without creating any real PPTX shapes or files.

Key technique: monkey-patch ``Presentation`` in nsm_l3_diagram_create's
module namespace with ``_MockPresentation``, and replace
``extended.add_shape`` / ``add_line`` with capture-only stubs.

The captured coordinates are rendered to SVG with z-order sorting that
mirrors the PowerPoint layer management.
"""

import os
import shutil
import tempfile

import nsm_ddx_figure


# ===================================================================
# Mock PPTX objects — zero python-pptx shape creation
# ===================================================================

class _MockElement:
    tag = 'mock'
    def getparent(self):
        return None


class _MockSpTree(list):
    """Silently accepts remove / insert of mock elements."""
    def remove(self, elem):
        try:
            super().remove(elem)
        except ValueError:
            pass

    def insert(self, idx, elem):
        try:
            super().insert(idx, elem)
        except Exception:
            pass


class _MockAdjustments:
    def __setitem__(self, k, v): pass
    def __getitem__(self, k): return 0.0
    def __bool__(self): return False


class _MockColorFormat:
    rgb = None
    brightness = 0.0


class _MockFillFormat:
    def __init__(self):
        self.fore_color = _MockColorFormat()
    def solid(self): pass
    def background(self): pass


class _MockLineFormat:
    def __init__(self):
        self.color = _MockColorFormat()
        self.width = 0
        self.fill = _MockFillFormat()


class _MockFont:
    def __init__(self):
        self.color = _MockColorFormat()
        self.name = 'Calibri'
        self.size = 0


class _MockParagraph:
    def __init__(self):
        self.font = _MockFont()
        self.alignment = None


class _MockTextFrame:
    def __init__(self):
        self.paragraphs = [_MockParagraph()]
        self.margin_top = 0
        self.margin_bottom = 0
        self.margin_left = 0
        self.margin_right = 0
        self.vertical_anchor = None


class _MockShadow:
    inherit = False


class _MockShape:
    def __init__(self, text=''):
        self._element = _MockElement()
        self.adjustments = _MockAdjustments()
        self.fill = _MockFillFormat()
        self.line = _MockLineFormat()
        self.text_frame = _MockTextFrame()
        self.shadow = _MockShadow()
        self.text = str(text)
        self.left = 0
        self.top = 0
        self.width = 0
        self.height = 0

    @property
    def has_text_frame(self):
        return True


class _MockShapeCollection:
    def __init__(self):
        self._spTree = _MockSpTree()
        self.title = _MockShape()

    def add_shape(self, *a, **kw):
        return _MockShape()

    def add_connector(self, *a, **kw):
        return _MockShape()


class _MockSlide:
    def __init__(self):
        self.shapes = _MockShapeCollection()


class _MockSlideCollection:
    def __init__(self):
        self._slides = []

    def add_slide(self, layout=None):
        s = _MockSlide()
        self._slides.append(s)
        return s

    def __iter__(self):
        return iter(self._slides)

    def __len__(self):
        return len(self._slides)


class _MockPresentation:
    """Drop-in replacement for python-pptx Presentation.
    Accepts all attribute assignments; slides return mocks."""

    def __init__(self, path=None):
        self.slide_width = 0
        self.slide_height = 0
        self.slide_layouts = [None] * 10
        self._slide_col = _MockSlideCollection()

    @property
    def slides(self):
        return self._slide_col

    def save(self, path):
        pass


# ===================================================================
# Capture-only wrappers (no original call, no PPTX shapes)
# ===================================================================

def _make_svg_add_shape(capture_list, enabled_flag):
    def wrapper(self, shape_type, shape_left, shape_top,
                shape_width, shape_hight, shape_text):
        if enabled_flag[0]:
            capture_list.append(
                ('shape', shape_type, shape_left, shape_top,
                 shape_width, shape_hight, str(shape_text)))

        if (getattr(self, 'click_value_l3', '') == 'L3-4-1'
                and getattr(self, 'flag_re_create', False)
                and not getattr(self, 'flag_second_page', False)):
            self.add_shape_write_array.append(
                [shape_type, shape_left, shape_top,
                 shape_width, shape_hight, shape_text])

        mock = _MockShape(shape_text)
        self.shape = mock
        sp = getattr(getattr(self, 'slide', None), 'shapes', None)
        if sp is not None:
            sp._spTree.append(mock._element)
    return wrapper


def _make_svg_add_line(capture_list, enabled_flag):
    def wrapper(self, line_type, x1, y1, x2, y2):
        if enabled_flag[0]:
            capture_list.append(('line', line_type, x1, y1, x2, y2))
        self.shape = _MockShape()
    return wrapper


# ===================================================================
# Z-order (mirrors _spTree.insert positions in PPTX engine)
# ===================================================================

_Z_ORDER = {
    'OUTLINE_NORMAL':     0,
    'FOLDER_NORMAL':      1,
    'DEVICE_NORMAL':      2,
    'DEVICE_L3_INSTANCE': 2,
    'WAY_POINT_NORMAL':   2,
    'L3_SEGMENT_GRAY':    3,
    'L3_SEGMENT_VPN':     3,
    'L3_INSTANCE':        4,
    'TAG_NORMAL':         5,
    'IP_ADDRESS_TAG':     6,
}


# ===================================================================
# Main class
# ===================================================================

class nsm_l3_svg_create:
    def __init__(self):
        ppt_meta_file = str(self.inFileTxt_L3_3_1.get())
        self.full_filepath = ppt_meta_file

        capture_list = []
        capture_enabled = [False]

        orig_add_shape = nsm_ddx_figure.extended.add_shape
        orig_add_line = nsm_ddx_figure.extended.add_line

        nsm_ddx_figure.extended.add_shape = _make_svg_add_shape(
            capture_list, capture_enabled)
        nsm_ddx_figure.extended.add_line = _make_svg_add_line(
            capture_list, capture_enabled)

        tmp_master_path = None
        l3_type = getattr(self, '_l3_svg_type', 'all_areas')

        try:
            import nsm_l3_diagram_create

            _orig_Presentation = nsm_l3_diagram_create.Presentation
            nsm_l3_diagram_create.Presentation = _MockPresentation

            try:
                if l3_type == 'per_area':
                    _run_per_area(self, capture_list, capture_enabled,
                                  nsm_l3_diagram_create)
                else:
                    tmp_master_path = _run_all_areas(
                        self, ppt_meta_file, capture_list, capture_enabled,
                        nsm_l3_diagram_create)
            finally:
                nsm_l3_diagram_create.Presentation = _orig_Presentation

        except Exception as e:
            import traceback
            traceback.print_exc()
        finally:
            nsm_ddx_figure.extended.add_shape = orig_add_shape
            nsm_ddx_figure.extended.add_line = orig_add_line
            if tmp_master_path:
                try:
                    if os.path.exists(tmp_master_path):
                        os.remove(tmp_master_path)
                except OSError:
                    pass

        # --- Build SVG from captured data ---
        attr_colors = getattr(self, 'attribute_tuple1_1', {})
        out = getattr(self, 'output_svg_file', None)

        if l3_type == 'per_area':
            # Generate one SVG per area (split by OUTLINE_NORMAL boundaries)
            area_groups = _split_capture_by_area(capture_list)
            saved_paths = []
            base_no_ext = os.path.splitext(out)[0] if out else None

            for i, area_items in enumerate(area_groups):
                shapes = [it[1:] for it in area_items if it[0] == 'shape']
                lines_data = [it[1:] for it in area_items if it[0] == 'line']
                if not shapes and not lines_data:
                    continue

                area_name = _area_name_from_capture(area_items)
                # Skip spurious groups that have no FOLDER_NORMAL (no area name)
                if not area_name:
                    continue
                safe_name = _safe_area_name_l3(area_name)
                title_text = f'[L3] {area_name}'

                slide_w, slide_h = _compute_content_extent(shapes, lines_data)
                svg = _shapes_to_svg(shapes, lines_data, slide_w, slide_h,
                                     title_text, attr_colors)

                if svg and base_no_ext:
                    area_path = base_no_ext + '_' + safe_name + '.svg'
                    with open(area_path, 'w', encoding='utf-8') as f:
                        f.write(svg)
                    saved_paths.append(area_path)

            self._per_area_svg_files = saved_paths
        else:
            shapes = [it[1:] for it in capture_list if it[0] == 'shape']
            lines_data = [it[1:] for it in capture_list if it[0] == 'line']

            # Keep standard slide minimum for all-areas output
            slide_w, slide_h = _compute_content_extent(
                shapes, lines_data, min_w=13.4, min_h=7.5)

            title_text = ''
            try:
                title_text = self.slide.shapes.title.text
            except Exception:
                pass
            if not title_text:
                title_text = '[L3] All Areas'

            svg = _shapes_to_svg(shapes, lines_data, slide_w, slide_h,
                                  title_text, attr_colors)

            if out and svg:
                with open(out, 'w', encoding='utf-8') as f:
                    f.write(svg)
                try:
                    print(f'L3 SVG saved: {out}')
                except UnicodeEncodeError:
                    pass

def _safe_area_name_l3(name):
    """Sanitize an area name for use in a filename."""
    import re as _re
    safe = _re.sub(r'[\\/*?:"<>|]', '-', str(name))
    safe = safe.strip('. ')
    return safe or 'Area'


def _split_capture_by_area(capture_list):
    """Split capture list into per-area groups.

    In l3_area_create, the CREATE pass draws shapes in this order:
      DEVICE_NORMAL/TAG_NORMAL/WAY_POINT_NORMAL → FOLDER_NORMAL →
      L3_SEGMENT_GRAY/connector LINES → OUTLINE_NORMAL → IP_ADDRESS_TAG → L3_INSTANCE lines

    Because OUTLINE_NORMAL is drawn AFTER the device shapes, a naive split at
    OUTLINE_NORMAL would place each area's OUTLINE and IP labels into the NEXT
    area's group (off-by-one).

    This corrected version splits at OUTLINE_NORMAL boundaries but then moves
    the OUTLINE_NORMAL + its trailing IP labels / L3-instance lines back to the
    group they logically belong to (the one that contains the corresponding
    FOLDER_NORMAL with the same area name).
    """
    # Step 1: raw split at each OUTLINE_NORMAL
    # raw_groups[0] = shapes before first OUTLINE  (= area-0 devices, no OUTLINE)
    # raw_groups[k] = [OUTLINE_{k-1}, IP_{k-1}, L3inst_{k-1}, area-k devices ...]
    raw_groups = []
    current = []
    for item in capture_list:
        if item[0] == 'shape' and item[1] == 'OUTLINE_NORMAL':
            raw_groups.append(current)
            current = [item]
        else:
            current.append(item)
    if current:
        raw_groups.append(current)

    if not raw_groups:
        return []

    def _split_tail(items):
        """
        Given a raw group that starts with OUTLINE_NORMAL, separate it into:
          tail  – OUTLINE_NORMAL + immediately following IP_ADDRESS_TAG shapes
                  and L3_INSTANCE lines  (all belong to the *previous* area)
          head  – everything else (start of the *next* area's device shapes)
        """
        if not items or not (items[0][0] == 'shape' and
                              items[0][1] == 'OUTLINE_NORMAL'):
            return [], items
        i = 1
        while i < len(items):
            it = items[i]
            if it[0] == 'shape' and it[1] == 'IP_ADDRESS_TAG':
                i += 1
            elif it[0] == 'line':
                i += 1
            else:
                break
        return items[:i], items[i:]

    # Step 2: rebuild correct per-area groups
    areas = []
    for i, raw in enumerate(raw_groups):
        if i == 0:
            # First raw group has no OUTLINE – it is purely area-0 device shapes
            areas.append(list(raw))
        else:
            tail, head = _split_tail(raw)
            # Attach tail (OUTLINE + IPs + L3-instance lines) to the *previous* area
            if areas:
                areas[-1].extend(tail)
            # head is the start of the current area
            areas.append(head)

    return [a for a in areas if a]


def _area_name_from_capture(area_items):
    """Extract the area name from FOLDER_NORMAL shape text in an area's capture."""
    for item in area_items:
        if item[0] == 'shape' and item[1] == 'FOLDER_NORMAL':
            txt = item[6] if len(item) > 6 else ''
            if txt:
                return str(txt)
    return ''


def _run_per_area(ctx, capture_list, cap_en, mod):
    ctx.output_ppt_file = ''
    ctx.click_value = 'L3-3-2'
    ctx.click_value_l3 = ''
    ctx.flag_re_create = False
    ctx.flag_second_page = False
    ctx.add_shape_array = []
    ctx.add_shape_write_array = []
    ctx.vpn_hostname_if_list = []
    ctx.per_index2_before_array = [0.0]
    ctx.per_index2_after_array = [0.0]
    ctx.y_grid_segment_array = []
    ctx.update_start_area_array = []
    if not hasattr(ctx, 'click_value_VPN'):
        ctx.click_value_VPN = ''

    cap_en[0] = True
    try:
        mod.nsm_l3_diagram_create.__init__(ctx)
    except UnicodeEncodeError:
        pass


def _compute_content_extent(shapes, lines, margin=1.0, min_w=0.0, min_h=0.0):
    """Compute SVG dimensions from actual content bounding box (inches).

    Unlike PowerPoint (capped at 56x56 inches), SVG has no size limit.
    Returns (width, height) in inches with margin.
    min_w/min_h can be set to 13.4/7.5 for full-slide output; per-area
    output passes 0.0 so the canvas tightly wraps the actual content.
    """
    max_x, max_y = min_w, min_h

    for sh in shapes:
        _, x, y, sw, sh_ = sh[0], sh[1], sh[2], sh[3], sh[4]
        right = x + sw
        bottom = y + sh_
        if right > max_x:
            max_x = right
        if bottom > max_y:
            max_y = bottom

    for ln in lines:
        _, x1, y1, x2, y2 = ln
        for x in (x1, x2):
            if x > max_x:
                max_x = x
        for y in (y1, y2):
            if y > max_y:
                max_y = y

    return max_x + margin, max_y + margin


def _run_all_areas(ctx, ppt_meta_file, capture_list, cap_en, mod):
    iDir = os.path.dirname(ppt_meta_file) or os.getcwd()
    base = os.path.splitext(os.path.basename(ppt_meta_file))[0]
    ext = os.path.splitext(ppt_meta_file)[1] or '.xlsx'
    tmp_master = os.path.join(
        iDir, base.replace('[MASTER]', '__TMP_SVG__[MASTER]') + ext)

    ctx.output_ppt_file = ''
    ctx.excel_maseter_file_backup = tmp_master
    ctx.add_shape_array = []
    ctx.add_shape_write_array = []
    if not hasattr(ctx, 'click_value_VPN'):
        ctx.click_value_VPN = ''

    import tkinter as tk
    ctx.outFileTxt_L3_3_4_1.delete(0, tk.END)
    ctx.outFileTxt_L3_3_4_1.insert(tk.END, '')
    ctx.outFileTxt_L3_3_5_1.delete(0, tk.END)
    ctx.outFileTxt_L3_3_5_1.insert(tk.END, tmp_master)

    if os.path.isfile(tmp_master):
        os.remove(tmp_master)

    # For NSM files, read the area count directly from the master (Parquet only,
    # no Excel I/O) so we can skip create_master_file_one_area when there is only
    # one area.  For xlsx (or other formats), always call create_master_file_one_area
    # to preserve the original behaviour without adding any Excel operations.
    if str(ppt_meta_file).lower().endswith('.nsm'):
        import nsm_def as _nsm_def
        _fw = _nsm_def.get_folder_wp_array_from_master('Master_Data', ppt_meta_file)
        folder_count = len(_fw[0]) if _fw else 0
        if folder_count <= 1:
            shutil.copy2(ppt_meta_file, tmp_master)
        else:
            mod.create_master_file_one_area.__init__(ctx)
    else:
        mod.create_master_file_one_area.__init__(ctx)

    ctx.vpn_hostname_if_list = []
    ctx.click_value = 'L3-3-2'
    ctx.click_value_l3 = 'L3-4-1'
    ctx.global_wp_array = []
    ctx.update_start_area_array = []
    ctx.y_grid_segment_array = []
    ctx.flag_second_page = False
    ctx.flag_re_create = False
    ctx.per_index2_before_array = [0.0]
    ctx.per_index2_after_array = [0.0]

    cap_en[0] = False
    try:
        mod.nsm_l3_diagram_create.__init__(ctx)
    except UnicodeEncodeError:
        pass

    ctx._l3_data_cache = {
        'result': ctx.result_get_l2_broadcast_domains,
        'update_l2': ctx.update_l2_table_array,
        'target_groups': ctx.target_l2_broadcast_group_array,
        'pos_folder_arr': ctx.position_folder_array,
        'pos_shape_arr': ctx.position_shape_array,
        'pos_line_arr': ctx.position_line_array,
        'pos_style_arr': ctx.position_style_shape_array,
        'pos_tag_arr': ctx.position_tag_array,
        'root_folder_arr': ctx.root_folder_array,
        'pos_folder_t': ctx.position_folder_tuple,
        'pos_shape_t': ctx.position_shape_tuple,
        'pos_line_t': ctx.position_line_tuple,
        'pos_style_t': ctx.position_style_shape_tuple,
        'pos_tag_t': ctx.position_tag_tuple,
        'root_folder_t': ctx.root_folder_tuple,
        'folder_wp': ctx.folder_wp_name_array,
        'l2_table_arr': getattr(ctx, 'l2_table_array', []),
        'l3_table_arr': getattr(ctx, 'l3_table_array', []),
        'device_list': getattr(ctx, 'device_list_array', []),
        'wp_list': getattr(ctx, 'wp_list_array', []),
        'all_shape_list': getattr(ctx, 'all_shape_list_array', []),
        'update_l3': ctx.update_l3_table_array,
        'l3_by_dev': ctx.l3_rows_by_device,
        'l3_by_dev_if': ctx.l3_rows_by_device_if,
        'l2_by_area': ctx.l2_rows_by_area,
        'grp_by_member': ctx.groups_by_member,
        'members_by_grp': ctx.members_by_group,
    }

    capture_list.clear()
    cap_en[0] = True
    ctx.flag_re_create = True

    try:
        mod.nsm_l3_diagram_create.__init__(ctx)
    except UnicodeEncodeError:
        pass

    ctx._l3_data_cache = None

    return tmp_master


# ===================================================================
# SVG renderer
# ===================================================================

def _shapes_to_svg(shapes, lines, slide_w, slide_h, title, attr_colors):
    from nsm_ddx_svg_l3 import L3_STYLES, L3_LINE_STYLES

    shapes = [s for _, s in sorted(
        enumerate(shapes),
        key=lambda p: (_Z_ORDER.get(p[1][0], 5), p[0]))]

    DPI = 96

    def _in(v):
        return round(v * DPI, 2)

    def _pt(v):
        return round(v * DPI / 72.0, 2)

    def _rgb(r, g, b):
        return f'rgb({int(r)},{int(g)},{int(b)})'

    def _esc(t):
        if t is None:
            return ''
        return str(t).replace('&', '&amp;').replace('<', '&lt;').replace(
            '>', '&gt;').replace('"', '&quot;')

    FF = 'Calibri, Segoe UI, Arial, sans-serif'
    w, h = _in(slide_w), _in(slide_h)

    p = [
        f'<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{w}" height="{h}" viewBox="0 0 {w} {h}">\n'
        f'<defs>\n'
        f'<marker id="diamond" viewBox="0 0 10 10" refX="5" refY="5" '
        f'markerWidth="8" markerHeight="8" orient="auto-start-reverse">'
        f'<path d="M5 0 L10 5 L5 10 L0 5 Z" fill="black"/></marker>\n'
        f'<marker id="diamond-purple" viewBox="0 0 10 10" refX="5" refY="5" '
        f'markerWidth="8" markerHeight="8" orient="auto-start-reverse">'
        f'<path d="M5 0 L10 5 L5 10 L0 5 Z" fill="rgb(112,48,160)"/>'
        f'</marker>\n'
        f'<style>text{{font-family:{FF};}}</style>\n'
        f'</defs>\n'
        f'<rect width="{w}" height="{h}" fill="white"/>\n'
    ]

    if title:
        p.append(f'<text x="{_in(0.28)}" y="{_in(0.5)}" '
                 f'font-size="{_pt(18)}px" fill="black">{_esc(title)}</text>\n')

    for sh in shapes:
        st, x, y, sw, sh_, txt = sh
        sty = L3_STYLES.get(st) or L3_STYLES.get('DEVICE_NORMAL')
        if sty is None:
            continue

        fill = sty.get('fill')
        stroke = sty.get('stroke')
        sw_ = sty.get('sw', 1.0)
        rxr = sty.get('rx', 0.0)
        fs = sty.get('fs', 6)
        tc = sty.get('tc', (0, 0, 0))
        align = sty.get('align')

        if st == 'DEVICE_NORMAL' and txt in attr_colors:
            fill = tuple(attr_colors[txt])
        elif st == 'DEVICE_L3_INSTANCE' and txt in attr_colors:
            b = attr_colors[txt]
            fill = (min(255, b[0]+15), min(255, b[1]+10), min(255, b[2]+25))
        elif st == 'WAY_POINT_NORMAL' and txt in attr_colors:
            fill = tuple(attr_colors[txt])

        fs_ = _rgb(*fill) if fill else 'none'
        ss_ = _rgb(*stroke) if stroke else 'none'
        rx = _in(rxr * min(sw, sh_) * 0.5) if rxr else 0

        if fill is not None or stroke is not None:
            p.append(
                f'<rect x="{_in(x)}" y="{_in(y)}" '
                f'width="{_in(sw)}" height="{_in(sh_)}" '
                f'rx="{rx}" ry="{rx}" '
                f'fill="{fs_}" stroke="{ss_}" '
                f'stroke-width="{_pt(sw_)}"/>\n')

        if txt and str(txt).strip():
            tcs = _rgb(*tc) if tc else 'black'
            fsp = _pt(fs)
            anc, bl = 'middle', 'central'
            tx = _in(x + sw / 2)
            ty = _in(y + sh_ / 2)
            if align == 'left':
                anc = 'start'
                tx = _in(x + 0.05)
            if st == 'FOLDER_NORMAL':
                ty = _in(y + 0.15)
                bl = 'hanging'
            p.append(
                f'<text x="{tx}" y="{ty}" font-size="{fsp}px" '
                f'fill="{tcs}" text-anchor="{anc}" '
                f'dominant-baseline="{bl}">{_esc(txt)}</text>\n')

    for ln in lines:
        lt, x1, y1, x2, y2 = ln
        ls = L3_LINE_STYLES.get(lt)
        if ls is None:
            continue
        sc = _rgb(*ls['stroke'])
        lw = _pt(ls['sw'])
        mk = ''
        vpn = 'VPN' in lt
        if ls.get('marker_start'):
            mk += f' marker-start="url(#{"diamond-purple" if vpn else "diamond"})"'
        if ls.get('marker_end'):
            mk += f' marker-end="url(#{"diamond-purple" if vpn else "diamond"})"'
        p.append(
            f'<line x1="{_in(x1)}" y1="{_in(y1)}" '
            f'x2="{_in(x2)}" y2="{_in(y2)}" '
            f'stroke="{sc}" stroke-width="{lw}"{mk}/>\n')

    p.append('</svg>\n')
    return ''.join(p)


if __name__ == '__main__':
    nsm_l3_svg_create()
