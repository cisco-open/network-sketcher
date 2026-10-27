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

"""IPv4 / IPv6 helpers for the L3 table "IP Address / Subnet mask" column.

Only the standard library is imported here, so pure modules such as
nsm_ip_report_data can use these helpers without loading tkinter or openpyxl.
"""

import ipaddress

IPV4 = 'IPv4'
IPV6 = 'IPv6'
IPV6_EXCLUDED = 'IPv6_EXCLUDED'
NOT_MATCH = 'NOT_MATCH'


def _ipv4_octets(text):
    """Return the four octets of 'a.b.c.d/prefix' as ints, or None."""
    if text.count('.') != 3 or text.count('/') != 1:
        return None
    address, prefix = text.split('/')
    if not prefix.isdecimal() or not 1 <= int(prefix) <= 32:
        return None
    octets = address.split('.')
    if not all(octet.isdecimal() and 0 <= int(octet) <= 255 for octet in octets):
        return None
    return [int(octet) for octet in octets]


def _ipv6_interface(text):
    """Return an IPv6Interface for 'addr/prefix', or None.

    A zone index ('%eth0') is rejected: it only names a local link.
    """
    if '%' in text or '/' not in text or ':' not in text:
        return None
    prefix = text.split('/', 1)[1]
    if not prefix.isdecimal() or not 1 <= int(prefix) <= 128:
        return None
    try:
        return ipaddress.IPv6Interface(text)
    except ValueError:
        return None


def ipv6_exclusion_reason(text):
    """Return why an IPv6 value is not drawn, or '' when it is a target.

    Values with a zone index or in a range that never identifies a configured
    interface address in a network diagram are excluded.
    """
    text = str(text).strip()
    if '%' in text:
        return 'zone index (%)'
    iface = _ipv6_interface(text)
    if iface is None:
        return ''
    ip = iface.ip
    if ip.ipv4_mapped is not None:
        return 'IPv4-mapped (::ffff:0:0/96)'
    if ip.is_unspecified:
        return 'unspecified (::/128)'
    if ip.is_loopback:
        return 'loopback (::1/128)'
    if ip.is_link_local:
        return 'link-local (fe80::/10)'
    if ip.is_site_local:
        return 'site-local (fec0::/10)'
    if ip.is_multicast:
        return 'multicast (ff00::/8)'
    return ''


def check_ip_format(char_ip_address):
    """Classify one 'address/prefix' value.

    Returns 'IPv4', 'IPv6', 'IPv6_EXCLUDED' or 'NOT_MATCH'.
    """
    text = str(char_ip_address).strip()
    if _ipv4_octets(text) is not None:
        return IPV4
    if '%' in text and ':' in text:
        return IPV6_EXCLUDED
    if _ipv6_interface(text) is None:
        return NOT_MATCH
    if ipv6_exclusion_reason(text):
        return IPV6_EXCLUDED
    return IPV6


def normalize_ip(char_ip_address):
    """Return the stored form of a valid value.

    IPv4 keeps the text as written (spaces removed). IPv6 uses the lower-case
    compressed form so that the same address always compares equal.
    """
    text = str(char_ip_address).replace(' ', '').replace('\u3000', '').strip()
    iface = _ipv6_interface(text)
    if iface is not None:
        return str(iface)
    return text


def get_ip_sort_key(char_ip_address):
    """Numeric order with IPv4 first, then IPv6, then anything unparseable.

    For IPv4 this matches the zero-padded octet order of get_ipv4_value.
    """
    text = str(char_ip_address).strip()
    octets = _ipv4_octets(text)
    if octets is not None:
        value = (octets[0] << 24) | (octets[1] << 16) | (octets[2] << 8) | octets[3]
        return (4, value, '')
    iface = _ipv6_interface(text)
    if iface is not None:
        return (6, int(iface.ip), '')
    return (99, 0, text)


def get_ip_version_key(char_ip_address):
    """Key for a stable sort that only moves IPv6 values behind the others."""
    text = str(char_ip_address).strip()
    if _ipv4_octets(text) is None and _ipv6_interface(text) is not None:
        return 1
    return 0


def ipv6_host_label(iface):
    """Host part shown on an L3 interface tag for an IPv6Interface.

    Mirrors the IPv4 rule, which drops the octets covered by the mask
    (/24 -> '.1'): whole 16-bit hextets covered by the prefix are dropped,
    leading zero hextets of the rest are omitted, and '::' marks the cut.
    2001:db8:1000:10::1/64 -> '::1'.
    """
    prefixlen = iface.network.prefixlen
    if prefixlen < 16:
        return str(iface.ip)
    hextets = [int(part, 16) for part in iface.ip.exploded.split(':')]
    tail = hextets[min(prefixlen, 112) // 16:]
    while len(tail) > 1 and tail[0] == 0:
        tail.pop(0)
    return '::' + ':'.join(f'{hextet:x}' for hextet in tail)
