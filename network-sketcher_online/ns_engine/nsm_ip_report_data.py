"""IP report tabs data builder (shared by Online edition and Local MCP).

This module computes the two extra Device Table tabs introduced in v3.1.2a:

  - **IP Address_Summary** -- per-area CIDR aggregation via
    ``ipaddress.collapse_addresses``.
  - **IP Address_List** -- one row per distinct (IP, Mask, Network, Device,
    L3 IF, L3 Instance, Area) tuple, sorted by IP value.

It mirrors the data extraction logic of the Offline / Online engines'
``nsm_extensions.ip_report.export_ip_report`` (which writes an
``[IP_REPORT]*.xlsx`` file) but is a *pure function* operating on
already-loaded ``.nsm``-derived structures. Notably, this module:

  - never reads the filesystem;
  - never imports ``openpyxl`` or any Excel-related code;
  - takes no ``master_path`` argument;

This makes it structurally impossible for the Device Table HTML path to
fall back to an Excel reader.
"""

from __future__ import annotations

import ipaddress
import logging
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# L3_TABLE normalisation
# ---------------------------------------------------------------------------

# L3_TABLE columns (row > 2):
#   [0] Area
#   [1] Device Name
#   [2] L3 Port Name (a.k.a. L3 IF Name)
#   [3] L3 Instance Name
#   [4] IP Address / Subnet mask (Comma Separated for multiple IPs)
#   [5] VPN Target Device Name (unused here)
#   [6] VPN Target L3 Port Name (unused here)
#
# As written by nsm_l3_table_from_master, repeated values in cols [0]/[1]/[2]
# are blanked out within a contiguous run (visual grouping), so we forward-fill
# those columns when extracting.


L3Record = Tuple[str, str, str, str, str]  # (area, device, l3_if, l3_instance, ip_with_subnet)


def _safe_cell(value) -> str:
    """Return a stripped string for a cell value (None / '' both -> '')."""
    if value is None:
        return ''
    return str(value).strip()


def _extract_l3_rows(l3_raw: Sequence) -> List[L3Record]:
    """Normalise the raw L3_TABLE section into a flat list of records.

    Performs three operations:
      1. Forward-fills the Area / Device / L3 IF columns within contiguous
         visual groups (the L3_TABLE writer blanks repeated values).
      2. Splits comma-separated IP values into one record per IP.
      3. Drops rows whose IP column is empty (no IP -> not interesting for
         either tab).

    Args:
        l3_raw: result of ``convert_master_to_array('Master_Data_L3',
            master_path, '<<L3_TABLE>>')`` -- a list of ``[row_index, row]``.

    Returns:
        list of ``(area, device, l3_if, l3_instance, ip_with_subnet)``.
    """
    records: List[L3Record] = []
    current_area = ''
    current_device = ''
    current_l3if = ''

    for entry in l3_raw:
        if not isinstance(entry, list) or len(entry) < 2:
            continue
        row_index = entry[0]
        row = entry[1]
        # Skip <RANGE> (row 1) and <HEADER> (row 2); only process data rows.
        if not isinstance(row_index, int) or row_index <= 2:
            continue
        if not isinstance(row, list) or not row:
            continue
        # Skip the terminating <END> marker row.
        if row[0] == '<END>':
            continue

        area = _safe_cell(row[0]) if len(row) > 0 else ''
        device = _safe_cell(row[1]) if len(row) > 1 else ''
        l3if = _safe_cell(row[2]) if len(row) > 2 else ''
        l3inst = _safe_cell(row[3]) if len(row) > 3 else ''
        ip_with_subnet = _safe_cell(row[4]) if len(row) > 4 else ''

        # Forward-fill the visually-blanked columns.
        if area:
            current_area = area
        else:
            area = current_area
        if device:
            current_device = device
        else:
            device = current_device
        if l3if:
            current_l3if = l3if
        else:
            l3if = current_l3if

        # No IP -> not relevant for IP_REPORT tabs.
        if not ip_with_subnet:
            continue

        # Split comma-separated IPs (e.g. "10.0.0.1/24, 192.168.1.1/24").
        for raw_ip in ip_with_subnet.split(','):
            ip_norm = raw_ip.strip()
            if not ip_norm:
                continue
            records.append((area, device, l3if, l3inst, ip_norm))

    return records


# ---------------------------------------------------------------------------
# IP Address_List tab
# ---------------------------------------------------------------------------

IP_LIST_HEADERS = [
    'IP Address', 'Mask', 'Network Address',
    'Device Name', 'L3 IF Name', 'L3 Instance', 'Area',
]


def _decompose_ip(ip_with_subnet: str) -> Tuple[str, str, str, str]:
    """Decompose an IP/subnet string into (ip, mask, network, sort_key).

    Matches the behaviour of nsm_extensions.ip_report.export_ip_report
    (L740-757): malformed values (no '/' or ipaddress raises) get the
    placeholder '[None]' and sort to the very end via the sentinel key
    '255255255255'.
    """
    ip = '[None]'
    mask = '[None]'
    network = '[None]'
    sort_key = '255255255255'

    if '/' not in ip_with_subnet:
        return ip, mask, network, sort_key
    try:
        net = ipaddress.ip_network(ip_with_subnet, strict=False)
        iface = ipaddress.ip_interface(ip_with_subnet)
        ip = str(iface.ip)
        mask = str(iface.netmask)
        _ip_part, prefix = ip_with_subnet.split('/', 1)
        network = f'{net.network_address}/{prefix.strip()}'
        # IPv4-only: zero-pad each octet to 3 digits for lexical ordering
        # that matches numeric ordering.
        try:
            sort_key = ''.join(f'{int(o):03}' for o in ip.split('.'))
        except ValueError:
            # IPv6 -- ipaddress will have raised before we reach here, but be
            # defensive: fall back to lexical IP comparison.
            sort_key = ip
    except (ValueError, ipaddress.AddressValueError,
            ipaddress.NetmaskValueError) as exc:
        logger.debug('IP decompose failed for %r: %s', ip_with_subnet, exc)
        # Keep the '[None]' placeholders.

    return ip, mask, network, sort_key


def _build_ip_address_list_rows(l3_records: Iterable[L3Record]
                                ) -> List[List[str]]:
    """Build the IP Address_List tab rows.

    Each record is decomposed to (IP, Mask, Network, Device, L3 IF,
    L3 Instance, Area). Duplicates (same 7-tuple) collapse to one row.
    The Offline edition stores L3 Instance as a single space ' ' when the
    source value is empty; we replicate that so xlsx-vs-tab comparisons stay
    byte-aligned.
    """
    seen: set = set()
    unsorted: List[Tuple[str, List[str]]] = []

    for area, device, l3if, l3inst, ip_with_subnet in l3_records:
        ip, mask, network, sort_key = _decompose_ip(ip_with_subnet)
        instance = l3inst if l3inst else ' '
        row = [ip, mask, network, device, l3if, instance, area]
        key = tuple(row)
        if key in seen:
            continue
        seen.add(key)
        unsorted.append((sort_key, row))

    # Sort by numeric IP order ascending (sentinel keeps invalid IPs at end).
    unsorted.sort(key=lambda x: x[0])
    return [row for _, row in unsorted]


# ---------------------------------------------------------------------------
# IP Address_Summary tab
# ---------------------------------------------------------------------------

IP_SUMMARY_HEADERS = ['Area', 'Summary(CIDR)']


def _collapse_for_area(cidrs: Iterable[str]) -> List[str]:
    """Run ipaddress.collapse_addresses on a list of CIDR strings, skipping
    any value that cannot be parsed as a network. Returns the collapsed
    CIDR list sorted by network address."""
    networks = []
    for cidr in cidrs:
        if '/' not in cidr:
            continue
        try:
            networks.append(ipaddress.ip_network(cidr, strict=False))
        except (ValueError, ipaddress.AddressValueError,
                ipaddress.NetmaskValueError) as exc:
            logger.debug('skipping unparseable CIDR %r: %s', cidr, exc)
            continue
    if not networks:
        return []
    try:
        collapsed = ipaddress.collapse_addresses(networks)
    except TypeError:
        # Mixing IPv4 and IPv6 -- bucket and collapse separately to mirror
        # the user's expectation of seeing both summarised.
        v4 = [n for n in networks if n.version == 4]
        v6 = [n for n in networks if n.version == 6]
        collapsed_list: List = []
        if v4:
            collapsed_list.extend(ipaddress.collapse_addresses(v4))
        if v6:
            collapsed_list.extend(ipaddress.collapse_addresses(v6))
        collapsed = collapsed_list
    return [str(n) for n in collapsed]


def _build_summary_rows(l3_records: Iterable[L3Record],
                        area_list: Sequence[str]) -> List[List[str]]:
    """Build the IP Address_Summary tab rows.

    Per-area CIDR aggregation. Per user spec (chosen May 2026), the Area
    column is populated on every row (sort/filter/CSV friendly), in contrast
    to the Offline xlsx which leaves it blank on the 2nd+ row within the
    same area.
    """
    records_list = list(l3_records)
    rows: List[List[str]] = []
    for area in area_list:
        area_cidrs = [r[4] for r in records_list if r[0] == area]
        if not area_cidrs:
            continue
        for cidr in _collapse_for_area(area_cidrs):
            rows.append([area, cidr])
    return rows


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_ip_report_tabs(device_area_map: Dict[str, str],
                         area_list: Sequence[str],
                         l3_raw: Sequence) -> List[dict]:
    """Compute the two IP report tabs for the Device Table HTML.

    Args:
        device_area_map: ``{device_name: area}`` already produced from the
            ``Master_Data`` / ``<<POSITION_SHAPE>>`` section by the caller.
            Currently unused (the L3 records carry their own Area column),
            but kept in the signature so future extensions (e.g. joining
            with non-L3 devices) do not need a public API change.
        area_list: ordered list of real area names (POSITION_SHAPE-derived,
            with waypoint pseudo-areas like ``_N/A_`` already excluded by
            the caller).
        l3_raw: result of ``convert_master_to_array('Master_Data_L3',
            master_path, '<<L3_TABLE>>')``.

    Returns:
        ``[{id: 'ips', label: 'IP Address_Summary', headers, rows},
           {id: 'ipl', label: 'IP Address_List',    headers, rows}]``
    """
    # device_area_map is accepted to keep the signature stable for future use;
    # mark it as intentionally unused without spamming the logs.
    _ = device_area_map

    l3_records = _extract_l3_rows(l3_raw)
    summary_rows = _build_summary_rows(l3_records, area_list)
    list_rows = _build_ip_address_list_rows(l3_records)

    return [
        {
            'id': 'ips',
            'label': 'IP Address_Summary',
            'headers': IP_SUMMARY_HEADERS,
            'rows': summary_rows,
        },
        {
            'id': 'ipl',
            'label': 'IP Address_List',
            'headers': IP_LIST_HEADERS,
            'rows': list_rows,
        },
    ]
