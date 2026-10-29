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

"""Helpers for the '_UNTAG_' L2 segment suffix.

'B_UNTAG_' means VLAN B is carried untagged (native) on that port. On one
port, 'B' and 'B_UNTAG_' are the same VLAN and the '_UNTAG_' form wins; a port
carries at most one '_UNTAG_' segment. On one device they are the same VLAN.
Across an L1 link, untagged segments connect to untagged segments (any name)
and tagged segments connect to the same tagged name.

Only the standard library is used so that any module can import this.
"""

UNTAG_SUFFIX = '_UNTAG_'


def is_untag(name):
    return str(name).endswith(UNTAG_SUFFIX)


def base_name(name):
    name = str(name)
    return name[:-len(UNTAG_SUFFIX)] if name.endswith(UNTAG_SUFFIX) else name


def is_valid_name(name):
    """A segment name must not be empty, and '_UNTAG_' needs a base name."""
    name = str(name)
    return name != '' and base_name(name) != ''


def split_names(cell):
    """Split a comma separated cell into names (spaces removed, empties dropped)."""
    if cell is None:
        return []
    text = str(cell).replace(' ', '').replace('\u3000', '')
    return [part for part in text.split(',') if part != '']


def normalize_port_segments(names):
    """Apply the same-port rule to names already stored on one port.

    Returns (names, error). Names with the same base collapse into one entry,
    preferring the '_UNTAG_' form at the position of the first occurrence.
    error is set when two '_UNTAG_' names with different bases are present.
    """
    result = []
    index_by_base = {}
    for name in names:
        base = base_name(name)
        if base in index_by_base:
            pos = index_by_base[base]
            if is_untag(name) and not is_untag(result[pos]):
                result[pos] = name
            continue
        index_by_base[base] = len(result)
        result.append(name)
    untags = [name for name in result if is_untag(name)]
    error = None
    if len(untags) > 1:
        error = 'Only one _UNTAG_ segment is allowed per port: ' + ', '.join(untags)
    return result, error


def normalize_input_segments(names):
    """normalize_port_segments for user input: also rejects a bare '_UNTAG_'."""
    for name in names:
        if not is_valid_name(name):
            return list(names), f'Invalid L2 segment name: {name}'
    return normalize_port_segments(names)


def merge_port_segments(existing, new_names):
    """Add new_names to the segments of one port.

    Returns (result, added, unchanged, error). result keeps the existing order
    (a tagged name replaced by its '_UNTAG_' form stays in place) and appends
    new names. When error is set, result equals the normalized existing list
    and nothing was added.
    """
    result, error = normalize_port_segments(list(existing))
    if error:
        return result, [], [], error
    original = list(result)
    added = []
    unchanged = []
    for name in new_names:
        if not is_valid_name(name):
            return original, [], [], f'Invalid L2 segment name: {name}'
        base = base_name(name)
        match = next((i for i, cur in enumerate(result) if base_name(cur) == base), None)
        other_untag = next((cur for cur in result if is_untag(cur) and base_name(cur) != base), None)
        if match is None:
            if is_untag(name) and other_untag is not None:
                return original, [], [], (f'Only one _UNTAG_ segment is allowed per port: '
                                          f'{other_untag} is already set, cannot add {name}')
            result.append(name)
            added.append(name)
        elif result[match] == name or (not is_untag(name) and is_untag(result[match])):
            unchanged.append(name)
        else:
            if other_untag is not None:
                return original, [], [], (f'Only one _UNTAG_ segment is allowed per port: '
                                          f'{other_untag} is already set, cannot add {name}')
            result[match] = name
            added.append(name)
    return result, added, unchanged, None


def add_l2_segment_to_cell(current_value, name):
    """Add one name to a column 6 / 7 cell for the single add commands.

    Returns (new_value, status). status is 'added', 'exists', or an error
    message. Without '_UNTAG_' on either side the result is exactly what the
    commands produced before (append, then plain sort).
    """
    current_value = '' if current_value is None else str(current_value)
    segs = [seg.strip() for seg in current_value.split(',')] if current_value else []
    if is_untag(name) and base_name(name) == '':
        return current_value, f'Invalid L2 segment name: {name}'
    if not is_untag(name) and not any(is_untag(seg) for seg in segs):
        if name in segs:
            return current_value, 'exists'
        joined = name if current_value.strip() == '' else current_value + ',' + name
        return ','.join(sorted(seg.strip() for seg in joined.split(','))), 'added'
    result, added, _unchanged, error = merge_port_segments([seg for seg in segs if seg], [name])
    if error:
        return current_value, error
    if not added:
        return current_value, 'exists'
    return ','.join(sorted(result)), 'added'


def find_segment_to_delete(names, name):
    """The stored name a delete command removes: exact match first, then same VLAN."""
    if name in names:
        return name
    return find_by_base(names, name)


def find_by_base(names, name):
    """Return the stored name on a port that is the same VLAN as name, or None."""
    base = base_name(name)
    return next((cur for cur in names if base_name(cur) == base), None)


def device_canonical_map(segment_lists):
    """{base name: box name} for one device from its column 6 name lists.

    The box name is 'B_UNTAG_' when any port of the device carries B untagged,
    otherwise the name itself. Only bases that need rewriting are included.
    """
    untagged_bases = set()
    seen = set()
    for names in segment_lists:
        for name in names:
            seen.add(name)
            if is_untag(name):
                untagged_bases.add(base_name(name))
    mapping = {}
    for base in untagged_bases:
        if base in seen:
            mapping[base] = base + UNTAG_SUFFIX
    return mapping


def canonical_names(names, mapping):
    """Rewrite names with mapping and drop duplicates, keeping the first position."""
    result = []
    for name in names:
        name = mapping.get(name, name)
        if name not in result:
            result.append(name)
    return result


def canonicalize_l2_cells(rows, device_index=1, segment_index=6):
    """Rewrite column 6 of L2 table rows in place to the per-device box name.

    rows are lists (the row values). Only cells whose value changes are
    rewritten, so masters without '_UNTAG_' stay byte-identical. Returns the
    number of rewritten cells.
    """
    lists_by_device = {}
    for row in rows:
        if len(row) > segment_index and row[segment_index] not in ('', None):
            lists_by_device.setdefault(row[device_index], []).append(split_names(row[segment_index]))
    maps = {}
    for device, lists in lists_by_device.items():
        if not any(is_untag(name) for names in lists for name in names):
            continue  # devices without '_UNTAG_' are left exactly as stored
        port_lists = [normalize_port_segments(names)[0] for names in lists]
        maps[device] = device_canonical_map(port_lists)
    changed = 0
    for row in rows:
        if len(row) <= segment_index or row[device_index] not in maps:
            continue
        cell = row[segment_index]
        if cell in ('', None):
            continue
        names = normalize_port_segments(split_names(cell))[0]
        names = canonical_names(names, maps[row[device_index]])
        new_cell = ','.join(names)
        if new_cell != cell and names != split_names(cell):
            row[segment_index] = new_cell
            changed += 1
    return changed


def segments_connect(local_name, opp_name):
    """Two segments across an L1 link: untagged with untagged, or same tagged name."""
    if is_untag(local_name) and is_untag(opp_name):
        return True
    if not is_untag(local_name) and not is_untag(opp_name):
        return local_name == opp_name
    return False
