"""Compact rendering of `show` output for the AI Context file.

The AI Context bundles the output of every `show` command, and on large masters
that dump dominates the file (461 KB of the 612 KB produced for a 302-device
master). The redundancy is structural rather than accidental:

  * `show attribute` repeats the same handful of RGB triples in thousands of
    cells,
  * `show l1_interface` repeats one (speed, duplex, media) combination on almost
    every row and carries an abbreviation that the engine derives from the
    interface name,
  * `show device_interface` and `show device` restate information that is fully
    contained in `show l1_interface` / `show area_device`,
  * several sections carry columns that are empty on every row.

Every transformation here is reversible, and :func:`expand_section` implements
the inverse so the reversibility can be tested rather than asserted. The `show`
commands themselves are untouched -- this module only reformats their output
while the AI Context is being assembled.
"""

import ast

# Sections dropped entirely because they can be rebuilt from another section.
# Maps command -> the note that replaces it.
DERIVABLE_SECTIONS = {
    'show device_interface':
        "[omitted] The physical ports of each device are the rows of "
        "show_l1_interface grouped by device name; waypoint devices are listed "
        "separately in show_waypoint_interface. Only the per-device port "
        "ordering of this section is not reproduced, and it carries no meaning.",
    'show device':
        "[omitted] The full device list is the union of the device lists in "
        "show_area_device.",
}

_ABBREV_BY_NAME = {
    'GigabitEthernet': 'GE',
    'Ethernet': 'E',
    'FastEthernet': 'FE',
    'TenGigabitEthernet': 'TE',
}


def derive_interface_abbrev(full_name):
    """Reproduce the abbreviation the engine stores for an interface name.

    Mirrors ``nsm_def.adjust_portname``: the name is split at the first digit,
    the alphabetic part is mapped through a small table (falling back to its
    first two characters) and the numeric part is appended after a space.
    Returns None for names the rule does not cover.
    """
    if not any(ch.isdigit() for ch in full_name):
        return None
    collapsed = full_name.replace(' ', '')
    head, tail = '', ''
    seen_digit = False
    for ch in collapsed:
        if ch.isdigit():
            seen_digit = True
        if seen_digit:
            tail += ch
        else:
            head += ch
    if len(head) <= 1:
        abbrev = head
    else:
        abbrev = _ABBREV_BY_NAME.get(head, head[:2])
    return (abbrev + ' ' + tail).strip()


def _parse(text):
    """Return the parsed list for a show output, or None if it is not one."""
    text = (text or '').strip()
    if not text.startswith('['):
        return None
    try:
        value = ast.literal_eval(text)
    except (ValueError, SyntaxError):
        return None
    return value if isinstance(value, list) else None


def _is_row_table(value):
    """True for a flat list of equal-length lists of scalars (a table)."""
    if not value or not all(isinstance(r, list) for r in value):
        return False
    width = len(value[0])
    if width < 2 or any(len(r) != width for r in value):
        return False
    return all(not isinstance(c, (list, tuple)) for r in value for c in r)


def _most_common(values):
    counts = {}
    for v in values:
        counts[v] = counts.get(v, 0) + 1
    return max(counts.items(), key=lambda kv: kv[1])[0]


# --- show attribute ---------------------------------------------------------

def _compact_attribute(rows):
    """Replace every attribute cell whose color is its column default with the
    bare value, and declare the defaults once in a legend."""
    if len(rows) < 2:
        return None
    header, body = rows[0], rows[1:]
    width = len(header)

    parsed = []           # per row: list of (value, rgb) or None for non-cells
    colors_by_col = {}
    for row in body:
        cells = []
        for idx, cell in enumerate(row):
            pair = None
            if idx > 0 and isinstance(cell, str) and cell.startswith('['):
                try:
                    value, rgb = ast.literal_eval(cell)
                except (ValueError, SyntaxError, TypeError):
                    pair = None
                else:
                    if isinstance(rgb, list):
                        pair = (value, tuple(rgb))
                        colors_by_col.setdefault(idx, []).append(pair[1])
            cells.append(pair)
        parsed.append(cells)

    if not colors_by_col:
        return None

    defaults = {idx: _most_common(seen) for idx, seen in colors_by_col.items()}

    out_rows = [header]
    for row, cells in zip(body, parsed):
        new_row = []
        for idx, cell in enumerate(row):
            pair = cells[idx]
            if pair is not None and pair[1] == defaults.get(idx):
                new_row.append(pair[0])
            else:
                new_row.append(cell)
        out_rows.append(new_row)

    legend = ', '.join(
        "%d %s=%s" % (idx + 1, _name(header, idx, width), list(rgb))
        for idx, rgb in sorted(defaults.items()))
    notes = [
        "[compact] A bare 'VALUE' cell uses its column's default color; cells "
        "that differ keep the full \"['VALUE', [R, G, B]]\" form.",
        "[compact] Default color per 1-based column: " + legend + ".",
    ]
    return notes, out_rows


def _name(header, idx, width):
    if idx < width and isinstance(header[idx], str) and header[idx]:
        return repr(header[idx])
    return '(unnamed)'


def _expand_attribute(notes, rows):
    defaults = _parse_color_legend(notes)
    header, body = rows[0], rows[1:]
    out = [header]
    for row in body:
        new_row = []
        for idx, cell in enumerate(row):
            if idx > 0 and isinstance(cell, str) and not cell.startswith('['):
                rgb = defaults.get(idx)
                new_row.append("['%s', %s]" % (cell, list(rgb)))
            else:
                new_row.append(cell)
        out.append(new_row)
    return out


def _parse_color_legend(notes):
    for note in notes:
        marker = 'Default color per 1-based column: '
        if marker in note:
            body = note.split(marker, 1)[1].rstrip('.')
            result = {}
            for entry in body.split('], '):
                key, _, rgb = entry.partition('=')
                if not rgb.endswith(']'):
                    rgb += ']'
                result[int(key.split()[0]) - 1] = tuple(ast.literal_eval(rgb))
            return result
    return {}


# --- show l1_interface ------------------------------------------------------

_L1_COLUMNS = 6  # device, abbrev, full name, speed, duplex, media


def _compact_l1_interface(rows):
    if not rows or any(len(r) != _L1_COLUMNS for r in rows):
        return None

    drop_abbrev = all(derive_interface_abbrev(r[2]) == r[1] for r in rows)
    default = _most_common([tuple(r[3:6]) for r in rows])

    out_rows = []
    for row in rows:
        head = [row[0], row[2]] if drop_abbrev else [row[0], row[1], row[2]]
        if tuple(row[3:6]) == default:
            out_rows.append(head)
        else:
            out_rows.append(head + list(row[3:6]))

    notes = [
        "[compact] A short row omits (Speed, Duplex, Port_Type), which is "
        "%s for that interface; a long row spells all three out."
        % (list(default),),
    ]
    if drop_abbrev:
        notes.append(
            "[compact] The abbreviated interface column is omitted. It is "
            "derived from the interface name: split at the first digit, map "
            "GigabitEthernet/Ethernet/FastEthernet/TenGigabitEthernet to "
            "GE/E/FE/TE (any other name keeps its first two characters), then "
            "append the digits after a space -- 'GigabitEthernet 1/0/11' -> "
            "'GE 1/0/11'.")
    return notes, out_rows


def _expand_l1_interface(notes, rows):
    default = _parse_default_tuple(notes)
    drop_abbrev = any('abbreviated interface column is omitted' in n for n in notes)
    out = []
    for row in rows:
        head_len = 2 if drop_abbrev else 3
        head, tail = list(row[:head_len]), list(row[head_len:])
        device = head[0]
        full = head[-1]
        abbrev = derive_interface_abbrev(full) if drop_abbrev else head[1]
        out.append([device, abbrev, full] + (tail if tail else list(default)))
    return out


def _parse_default_tuple(notes):
    for note in notes:
        marker = 'which is '
        if marker in note and 'Speed' in note:
            body = note.split(marker, 1)[1]
            body = body.split(' for that interface')[0]
            return ast.literal_eval(body)
    return []


# --- generic: drop columns that are empty on every row ----------------------

def _compact_empty_columns(rows, header=None):
    width = len(rows[0])
    empty = [idx for idx in range(width)
             if all(r[idx] == '' for r in rows)]
    if not empty:
        return None
    keep = [idx for idx in range(width) if idx not in empty]
    out_rows = [[r[idx] for idx in keep] for r in rows]
    notes = [
        "[compact] Column(s) %s were empty on every row and are omitted; each "
        "row now has %d fields." % (
            ', '.join(str(i + 1) for i in empty), len(keep)),
    ]
    return notes, out_rows


def _expand_empty_columns(notes, rows):
    empty = _parse_dropped_columns(notes)
    if not empty:
        return rows
    out = []
    for row in rows:
        values = list(row)
        new_row = []
        for idx in range(len(values) + len(empty)):
            if idx in empty:
                new_row.append('')
            else:
                new_row.append(values.pop(0))
        out.append(new_row)
    return out


def _parse_dropped_columns(notes):
    for note in notes:
        marker = 'Column(s) '
        if marker in note:
            body = note.split(marker, 1)[1].split(' were empty')[0]
            return {int(tok) - 1 for tok in body.split(',')}
    return set()


# --- public API -------------------------------------------------------------

_COMPACTORS = {
    'show attribute': _compact_attribute,
    'show l1_interface': _compact_l1_interface,
}

_EXPANDERS = {
    'show attribute': _expand_attribute,
    'show l1_interface': _expand_l1_interface,
}


def compact_section(command, text):
    """Return the compact rendering of one show output.

    Falls back to the original text whenever the output does not have the
    expected shape, so an engine change can never corrupt the AI Context.
    """
    if command in DERIVABLE_SECTIONS:
        return DERIVABLE_SECTIONS[command]

    rows = _parse(text)
    if rows is None or not rows or not all(isinstance(r, list) for r in rows):
        return text

    compactor = _COMPACTORS.get(command)
    result = compactor(rows) if compactor else None
    if result is None and _is_row_table(rows):
        result = _compact_empty_columns(rows)
    if result is None:
        return text

    notes, out_rows = result
    compact_text = '\n'.join(notes) + '\n' + repr(out_rows)
    return compact_text if len(compact_text) < len(text) else text


def expand_section(command, text):
    """Inverse of :func:`compact_section`, for verifying reversibility."""
    lines = (text or '').splitlines()
    notes = [ln for ln in lines if ln.startswith('[compact]')]
    if not notes:
        return text
    payload = '\n'.join(ln for ln in lines if not ln.startswith('[compact]'))
    rows = _parse(payload)
    if rows is None:
        return text
    expander = _EXPANDERS.get(command, _expand_empty_columns)
    return repr(expander(notes, rows))


def compact_show_outputs(results_by_command):
    """Compact a whole {command: output} mapping."""
    return {cmd: compact_section(cmd, text)
            for cmd, text in results_by_command.items()}


def read_reference(path):
    """Return the static command reference document verbatim."""
    try:
        with open(str(path), 'r', encoding='utf-8') as f:
            return f.read()
    except OSError:
        return ''
