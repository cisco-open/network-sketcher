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
Network Sketcher Online Service

Provides a web interface for generating network diagrams (L1/L2/L3)
and device files from Network Sketcher master files.

Usage:
    python ns_web_start.py

Open the URL shown at startup (configured in ns_web_config.json) in your browser.
"""

import os
import sys
import ssl
import socket
import uuid
import subprocess
import zipfile
import re
import shutil
import shlex
import logging
import logging.handlers
import datetime
import time
import json
import threading
import concurrent.futures
from pathlib import Path
from urllib.parse import quote as url_quote

from flask import (
    Flask, request, jsonify, send_file,
    make_response, abort, Response, stream_with_context
)

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
NS_DIR = PROJECT_DIR / 'network-sketcher_offline'

os.environ['NS_WEB_SERVER'] = '1'

from ns_engine.nsm_adapter import bootstrap as _engine_bootstrap
_engine_bootstrap()
from ns_engine.nsm_adapter import run_cli as _engine_run_cli
from ns_engine.nsm_adapter import run_cli_isolated as _engine_run_cli_isolated
UPLOAD_DIR = BASE_DIR / 'uploads'
STATIC_DIR = BASE_DIR / 'static'

CONFIG_PATH = BASE_DIR / 'ns_web_config.json'
HELP_MSG_PATH = BASE_DIR / 'help_msg.json'

def _load_config():
    """Load configuration from ns_web_config.json."""
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        raw = json.load(f)
    cfg = {}
    for key, val in raw.items():
        if key.startswith('_'):
            continue
        cfg[key] = val['value'] if isinstance(val, dict) and 'value' in val else val
    return cfg

def _load_help_messages():
    """Load help messages from help_msg.json."""
    try:
        with open(HELP_MSG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

# Command reference source (single source of truth shared with AI Context and
# the Local MCP nsm://commands resource).
CMD_REF_PATH = BASE_DIR / 'ns_engine' / 'nsm_extensions_cmd_list.txt'
_cmd_ref_cache = None

# Verbs the Update Master "Run" actually executes. Any other verb (currently
# only `export`) is skipped at run time ("(skipped) ..."), so such commands are
# also excluded from the Command Reference viewer. Single source of truth:
# /run_commands references this same set.
RUNNABLE_VERBS = {'add', 'delete', 'rename', 'show'}


def _normalize_quotes(text):
    """Normalize smart/curly quotes to straight quotes.

    Some examples in nsm_extensions_cmd_list.txt use typographic quotes
    (U+2018/U+2019 single, U+201C/U+201D double). The Run validator expects
    straight quotes, so the insertable syntax must be normalized.
    """
    if not text:
        return text
    return (text
            .replace('\u2018', "'").replace('\u2019', "'")
            .replace('\u201c', '"').replace('\u201d', '"'))


def _load_command_reference():
    """Parse nsm_extensions_cmd_list.txt into a structured command reference.

    Returns a list of categories:
        [{'category': str, 'items': [{'id', 'title', 'body', 'syntax'}]}]

    Only the command portion is included; everything from the first
    ``## RULE`` heading onward (architecture rules) is excluded. The result
    is cached in module global ``_cmd_ref_cache`` (parsed once).
    """
    global _cmd_ref_cache
    if _cmd_ref_cache is not None:
        return _cmd_ref_cache

    try:
        with open(CMD_REF_PATH, 'r', encoding='utf-8') as f:
            raw = f.read()
    except OSError:
        _cmd_ref_cache = []
        return _cmd_ref_cache

    lines = raw.splitlines()

    # Cut off at the first architecture-rule heading (## RULE ...).
    cut = len(lines)
    for i, ln in enumerate(lines):
        if ln.startswith('## RULE'):
            cut = i
            break
    lines = lines[:cut]

    import re as _re

    _cat_re = _re.compile(r'^##\s+(Show|Add|Delete|Rename|Export)\s+Commands\s+reference\s*$', _re.I)
    _quoting_re = _re.compile(r'^##\s+IMPORTANT:\s*Quoting', _re.I)

    # Split into ## blocks (heading + body lines until next ## heading).
    blocks = []  # (heading_text, [body_lines])
    cur_head = None
    cur_body = []
    for ln in lines:
        if ln.startswith('## '):
            if cur_head is not None:
                blocks.append((cur_head, cur_body))
            cur_head = ln[3:].strip()
            cur_body = []
        elif ln.startswith('# '):
            # Top-level title / TOC: ignore as a standalone block.
            if cur_head is not None:
                blocks.append((cur_head, cur_body))
                cur_head = None
                cur_body = []
        else:
            if cur_head is not None:
                cur_body.append(ln)
    if cur_head is not None:
        blocks.append((cur_head, cur_body))

    def _extract_syntax(body_lines):
        """Return the first ```bash fenced block's content (normalized)."""
        in_fence = False
        collected = []
        for ln in body_lines:
            stripped = ln.strip()
            if not in_fence:
                if stripped.startswith('```'):
                    in_fence = True
                continue
            if stripped.startswith('```'):
                break
            collected.append(ln)
        syntax = '\n'.join(collected).strip()
        return _normalize_quotes(syntax)

    def _slug(text):
        return _re.sub(r'[^a-z0-9]+', '-', text.lower()).strip('-')

    categories = []
    current = None
    for head, body in blocks:
        if _quoting_re.match('## ' + head):
            # Quoting rules: informational section (no command/syntax).
            body_text = '\n'.join(body).strip()
            categories.append({
                'category': 'Quoting Rules',
                'items': [{
                    'id': 'quoting-rules',
                    'title': 'Quoting Rules for All Commands',
                    'body': body_text,
                    'syntax': '',
                }],
            })
            current = None
            continue

        cat_m = _cat_re.match('## ' + head)
        if cat_m:
            current = {'category': cat_m.group(1).capitalize(), 'items': []}
            categories.append(current)
            continue

        # Individual command block (only counted when under a category).
        if current is not None:
            # Exclude commands the Run action would skip (verb not runnable,
            # currently only `export`). The verb is the first token of the
            # heading, e.g. "export l1_diagram" -> "export".
            verb = head.split()[0].lower() if head.split() else ''
            if verb not in RUNNABLE_VERBS:
                continue
            body_text = '\n'.join(body).strip()
            current['items'].append({
                'id': _slug(head),
                'title': head,
                'body': body_text,
                'syntax': _extract_syntax(body),
            })

    # Drop empty categories (e.g. Export once all its items are excluded).
    categories = [c for c in categories if c['items']]
    _cmd_ref_cache = categories
    return _cmd_ref_cache

_cfg = _load_config()
_help_msgs = _load_help_messages()

def _resolve_path(relative_or_absolute):
    """Resolve a path that may be relative to BASE_DIR or absolute."""
    p = Path(relative_or_absolute)
    if p.is_absolute():
        return p
    return BASE_DIR / p

if UPLOAD_DIR.is_dir():
    for _d in UPLOAD_DIR.iterdir():
        if _d.is_dir():
            shutil.rmtree(str(_d), ignore_errors=True)
        elif _d.is_file() and _d.suffix == '.zip':
            _d.unlink(missing_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)

MAX_FILE_SIZE = _cfg['max_file_size_mb'] * 1024 * 1024
_raw_timeout = _cfg.get('subprocess_timeout', 600)
SUBPROCESS_TIMEOUT = int(_raw_timeout) if _raw_timeout else None  # None = no limit

logging.basicConfig(
    level=getattr(logging, _cfg.get('log_level', 'INFO').upper(), logging.INFO),
    format='%(asctime)s [%(levelname)s] %(message)s'
)
logger = logging.getLogger(__name__)

# --- Daily-rotating access log ---
LOGS_DIR = BASE_DIR / 'logs'
LOGS_DIR.mkdir(exist_ok=True)
ACCESS_LOG_RETENTION_DAYS = _cfg.get('access_log_retention_days', 90)


def _cleanup_old_access_logs():
    """Delete access log files older than the configured retention period."""
    cutoff = datetime.date.today() - datetime.timedelta(days=ACCESS_LOG_RETENTION_DAYS)
    for f in LOGS_DIR.glob('access_*.log'):
        try:
            date_str = f.stem.replace('access_', '')
            file_date = datetime.datetime.strptime(date_str, '%Y%m%d').date()
            if file_date < cutoff:
                f.unlink()
                logger.info('Deleted old access log: %s', f.name)
        except (ValueError, OSError):
            pass


class _DailyFileHandler(logging.Handler):
    """File handler that switches to a new file each day (access_YYYYMMDD.log)."""

    def __init__(self, log_dir):
        super().__init__()
        self._log_dir = Path(log_dir)
        self._current_date = None
        self._stream = None
        self._lock_file = threading.Lock()

    def _open_for_date(self, d):
        if self._stream:
            self._stream.close()
        fname = f'access_{d:%Y%m%d}.log'
        self._stream = open(self._log_dir / fname, 'a', encoding='utf-8')
        self._current_date = d

    def emit(self, record):
        with self._lock_file:
            today = datetime.date.today()
            if self._current_date != today:
                self._open_for_date(today)
                _cleanup_old_access_logs()
            try:
                self._stream.write(self.format(record) + '\n')
                self._stream.flush()
            except Exception:
                self.handleError(record)

    def close(self):
        with self._lock_file:
            if self._stream:
                self._stream.close()
                self._stream = None
        super().close()


access_logger = logging.getLogger('access')
access_logger.setLevel(logging.INFO)
access_logger.propagate = False
_daily_handler = _DailyFileHandler(LOGS_DIR)
_daily_handler.setFormatter(logging.Formatter('%(message)s'))
access_logger.addHandler(_daily_handler)

_cleanup_old_access_logs()

app = Flask(__name__, static_folder=str(STATIC_DIR))
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

HEARTBEAT_TIMEOUT = _cfg['heartbeat_timeout']
CLEANUP_INTERVAL = _cfg['cleanup_interval']


@app.after_request
def _log_access(response):
    path = request.path
    if path.startswith('/heartbeat/'):
        return response

    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    client_ip = request.remote_addr or '-'
    method = request.method
    status = response.status_code

    extra = ''
    upload_file = getattr(request, '_ns_upload_filename', None)
    download_file = getattr(request, '_ns_download_filename', None)
    if upload_file:
        extra += f' upload="{upload_file}"'
    if download_file:
        extra += f' download="{download_file}"'

    access_logger.info(
        '%s %s %s %s %s%s', now, client_ip, method, path, status, extra
    )
    return response


@app.after_request
def _set_security_headers(response):
    response.headers.setdefault('X-Content-Type-Options', 'nosniff')
    # device_preview is embedded as a thumbnail iframe from the same origin.
    # Allow same-origin framing only for that route; all others remain fully denied.
    if request.path.startswith('/device_preview/'):
        response.headers.setdefault('X-Frame-Options', 'SAMEORIGIN')
        response.headers.setdefault(
            'Content-Security-Policy',
            "default-src 'self'; script-src 'self' 'unsafe-inline'; "
            "style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; "
            "font-src 'self'; connect-src 'self'; frame-ancestors 'self'"
        )
    else:
        response.headers.setdefault('X-Frame-Options', 'DENY')
        response.headers.setdefault(
            'Content-Security-Policy',
            "default-src 'self'; script-src 'self' 'unsafe-inline'; "
            "style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; "
            "font-src 'self'; connect-src 'self'; frame-ancestors 'none'"
        )
    response.headers.setdefault('Referrer-Policy', 'strict-origin-when-cross-origin')
    response.headers.setdefault(
        'Strict-Transport-Security', 'max-age=31536000; includeSubDomains'
    )
    return response


_heartbeats = {}
_heartbeats_lock = threading.Lock()

# In-memory SVG content cache: {job_id: {filename: bytes}}
# Populated by run_ns_command_subprocess immediately after _wait_readable confirms
# the file is accessible.  The /svg_raw/ route serves from this cache so that
# subsequent browser requests are never affected by Windows AV file locks.
_svg_mem_cache: dict = {}
_svg_mem_cache_lock = threading.Lock()

# Tracks the last attribute value used when SVG thumbnails were generated for
# a given job + scope. Used by svg_grid_stream to skip CLI re-execution when
# attribute is unchanged and the cache is warm.
#   {job_id: {scope: attr_string}}
# where scope is either '__initial__' (the all-areas / Device Tables initial
# generation) or '<area_name>' (a single per-area generation triggered by the
# dropdown). attr_string may be '' if no attribute was selected.
_svg_grid_last_attr: dict = {}
_svg_grid_last_attr_lock = threading.Lock()

# Sentinel scope key for the initial (Device Tables + All Areas) generation.
_SVG_GRID_INITIAL_SCOPE = '__initial__'


# When an SVG file's root <svg> element declares a width/height larger than
# what Chromium's image rasterizer can handle (~16384 px on either axis,
# https://crbug.com/431005 and related), the browser silently renders an
# <img src="...svg"> as a blank box. The largest L2/L3 All Areas diagrams
# for ~1000-device networks easily exceed this limit (e.g. height=29922
# for L3 with 1024 devices). To keep thumbnails visible in every browser,
# /svg_raw/ accepts a `?thumb=1` query parameter that serves a normalized
# copy of the SVG with width/height scaled down to fit inside the safe
# canvas limit. The viewBox attribute is kept untouched so the SVG still
# scales to whatever CSS size the <img> element was given.
_SVG_THUMB_MAX_DIM = 2048
_SVG_ROOT_TAG_RE = re.compile(rb'<svg\b([^>]*)>', re.IGNORECASE)
_SVG_WIDTH_RE = re.compile(r'\bwidth\s*=\s*"([0-9.]+)(?:px)?"', re.IGNORECASE)
_SVG_HEIGHT_RE = re.compile(r'\bheight\s*=\s*"([0-9.]+)(?:px)?"', re.IGNORECASE)


def _normalize_svg_for_thumb(svg_bytes):
    """Return SVG bytes whose root <svg> width/height fit inside ``_SVG_THUMB_MAX_DIM``.

    Only the first occurrence of the root <svg> tag is rewritten. If the file
    has no width/height attributes (viewBox-only) or already fits, the input
    is returned unchanged. The function is safe on malformed input — any
    parsing failure short-circuits back to the original bytes.
    """
    try:
        m = _SVG_ROOT_TAG_RE.search(svg_bytes, 0, 4096)
        if not m:
            return svg_bytes
        attrs_text = m.group(1).decode('utf-8', errors='replace')
        wm = _SVG_WIDTH_RE.search(attrs_text)
        hm = _SVG_HEIGHT_RE.search(attrs_text)
        if not wm or not hm:
            return svg_bytes
        w = float(wm.group(1))
        h = float(hm.group(1))
        if w <= 0 or h <= 0:
            return svg_bytes
        longest = max(w, h)
        if longest <= _SVG_THUMB_MAX_DIM:
            return svg_bytes
        scale = _SVG_THUMB_MAX_DIM / longest
        new_w = round(w * scale, 2)
        new_h = round(h * scale, 2)
        new_attrs = _SVG_WIDTH_RE.sub(f'width="{new_w}"', attrs_text, count=1)
        new_attrs = _SVG_HEIGHT_RE.sub(f'height="{new_h}"', new_attrs, count=1)
        new_tag = ('<svg' + new_attrs + '>').encode('utf-8')
        return svg_bytes[:m.start()] + new_tag + svg_bytes[m.end():]
    except Exception:
        return svg_bytes


def touch_heartbeat(job_id):
    with _heartbeats_lock:
        _heartbeats[job_id] = time.time()


def remove_heartbeat(job_id):
    with _heartbeats_lock:
        _heartbeats.pop(job_id, None)


def _cleanup_loop():
    while True:
        time.sleep(CLEANUP_INTERVAL)
        try:
            now = time.time()
            expired = []
            with _heartbeats_lock:
                for jid, ts in list(_heartbeats.items()):
                    if now - ts > HEARTBEAT_TIMEOUT:
                        expired.append(jid)
                for jid in expired:
                    del _heartbeats[jid]
            for jid in expired:
                job_dir = UPLOAD_DIR / jid
                if job_dir.is_dir():
                    shutil.rmtree(str(job_dir), ignore_errors=True)
                    logger.info('Cleaned up expired job: %s', jid)
                with _svg_mem_cache_lock:
                    _svg_mem_cache.pop(jid, None)
                with _svg_grid_last_attr_lock:
                    _svg_grid_last_attr.pop(jid, None)
        except Exception as e:
            logger.warning('Cleanup error: %s', e)


_cleanup_thread = threading.Thread(target=_cleanup_loop, daemon=True)
_cleanup_thread.start()


def validate_master_filename(filename):
    if not filename:
        return False, 'Filename is empty'
    filename = os.path.basename(filename)
    if not filename:
        return False, 'Filename is empty after sanitization'
    if not filename.endswith('.xlsx') and not filename.endswith('.nsm'):
        return False, 'Only .xlsx and .nsm files are supported'
    if not filename.startswith('[MASTER]'):
        return False, 'Filename must start with [MASTER]'
    return True, ''


def sanitize_job_id(job_id):
    if not re.match(r'^[a-f0-9\-]{36}$', job_id):
        return None
    return job_id


def find_master_file(work_dir):
    for f in os.listdir(work_dir):
        if f.startswith('[MASTER]') and (f.endswith('.xlsx') or f.endswith('.nsm')):
            return f
    return None


def get_active_master(work_dir):
    active_file = os.path.join(work_dir, '.active_master')
    if os.path.exists(active_file):
        with open(active_file, 'r', encoding='utf-8') as f:
            name = f.read().strip()
        if name and os.path.exists(os.path.join(work_dir, name)):
            return name
    return find_master_file(work_dir)


def set_active_master(work_dir, filename):
    active_file = os.path.join(work_dir, '.active_master')
    with open(active_file, 'w', encoding='utf-8') as f:
        f.write(filename)


def run_ns_command(args):
    engine_dir = str(BASE_DIR / 'ns_engine')
    logger.debug('Running (in-process): %s', ' '.join(args))

    try:
        result = _engine_run_cli(args, cwd=engine_dir)
        logger.debug('Exit code: %d', result.returncode)
        if result.stdout:
            logger.debug('stdout: %s', result.stdout[:500])
        if result.stderr:
            logger.warning('stderr: %s', result.stderr[:500])
        return result
    except Exception as e:
        logger.error('Command failed: %s', e)
        return None


def format_size(size_bytes):
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f'{size_bytes:.1f} {unit}'
        size_bytes /= 1024
    return f'{size_bytes:.1f} TB'


def collect_generated_files(work_dir, master_filename):
    files = []
    for f in sorted(os.listdir(work_dir)):
        if f == master_filename:
            continue
        if '__TMP__' in f:
            continue
        if f.startswith('.'):
            continue
        if f.startswith('[MASTER]') and f.endswith('.xlsx'):
            continue
        filepath = os.path.join(work_dir, f)
        if not os.path.isfile(filepath):
            continue
        size = os.path.getsize(filepath)
        file_type = 'unknown'
        if f.startswith('[DEVICE]'):
            file_type = 'device'
        elif f.startswith('[L1_DIAGRAM]'):
            file_type = 'l1'
        elif f.startswith('[L2_DIAGRAM]'):
            file_type = 'l2'
        elif f.startswith('[L3_DIAGRAM]'):
            file_type = 'l3'
        files.append({
            'name': f,
            'size': size,
            'size_human': format_size(size),
            'type': file_type,
        })
    return files


# ---------- Routes ----------

@app.route('/')
def index():
    help_data = {k: v.get('text', '') for k, v in _help_msgs.items()}
    help_json = json.dumps(help_data, ensure_ascii=True)
    html = HTML_TEMPLATE.replace(
        '{{PARALLEL_LIMIT}}', str(_cfg['parallel_limit'])
    ).replace(
        '{{AI_CONTEXT_BTN_LABEL}}', _cfg['ai_context_button_label']
    ).replace(
        '{{AI_CONTEXT_BTN_URL}}', _cfg['ai_context_button_url']
    ).replace(
        '{{HELP_DATA_JSON}}', help_json
    )
    resp = make_response(html)
    resp.headers['Content-Type'] = 'text/html; charset=utf-8'
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    return resp


@app.route('/static/logo.png')
def logo():
    logo_path = STATIC_DIR / 'logo.png'
    if logo_path.is_file():
        return send_file(str(logo_path), mimetype='image/png')
    abort(404)


@app.route('/command_reference')
def command_reference():
    """Return the parsed Network Sketcher CLI command reference as JSON.

    Lazily fetched by the in-app Command Reference drawer on first open, so
    the (sizeable) reference does not weigh down the initial page load.
    """
    data = _load_command_reference()
    resp = jsonify(data)
    # Reference content is static per server build; allow client/proxy cache.
    resp.headers['Cache-Control'] = 'public, max-age=3600'
    return resp


@app.route('/create_empty_master', methods=['POST'])
def create_empty_master():
    import tempfile
    tmp_dir = tempfile.mkdtemp()
    dummy_path = os.path.join(tmp_dir, '[MASTER]dummy.xlsx')
    with open(dummy_path, 'wb') as f:
        f.write(b'')
    result = run_ns_command([
        'export', 'master_file_nodata',
        '--master', dummy_path,
    ])
    output_path = os.path.join(tmp_dir, '[MASTER]no_data.xlsx')
    try:
        os.remove(dummy_path)
    except OSError:
        pass
    if os.path.isfile(output_path):
        return send_file(
            output_path,
            as_attachment=True,
            download_name='[MASTER]no_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    msg = ''
    if result:
        msg = (result.stdout or '') + (result.stderr or '')
    return jsonify({'error': msg or 'Failed to create empty master file'}), 500


@app.route('/heartbeat/<job_id>', methods=['POST'])
def heartbeat(job_id):
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)
    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        return jsonify({'alive': False}), 404
    touch_heartbeat(job_id)
    return jsonify({'alive': True})


@app.route('/restore/<job_id>')
def restore_session(job_id):
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)
    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        return jsonify({'valid': False}), 404

    touch_heartbeat(job_id)

    master_filename = get_active_master(str(work_dir))
    if not master_filename:
        return jsonify({'valid': False, 'reason': 'No master file'}), 404

    basename = master_filename.replace('[MASTER]', '').replace('.xlsx', '').replace('.nsm', '')

    masters = sorted([
        f for f in os.listdir(str(work_dir))
        if f.startswith('[MASTER]') and (f.endswith('.xlsx') or f.endswith('.nsm')) and f != master_filename
    ])

    generated = collect_generated_files(str(work_dir), master_filename)

    meta = {}
    meta_path = work_dir / '.meta.json'
    if meta_path.is_file():
        try:
            with open(str(meta_path), 'r', encoding='utf-8') as mf:
                meta = json.load(mf)
        except Exception:
            pass

    return jsonify({
        'valid': True,
        'master_filename': master_filename,
        'basename': basename,
        'updated_masters': masters,
        'generated_files': generated,
        'areas': meta.get('areas', []),
        'device_count': meta.get('device_count', 0),
        'link_count': meta.get('link_count', 0),
        'attribute_titles': meta.get('attribute_titles', []),
    })


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file selected'}), 400

    file = request.files['file']
    filename = os.path.basename(file.filename or '')
    valid, msg = validate_master_filename(filename)
    if not valid:
        return jsonify({'error': msg}), 400

    job_id = str(uuid.uuid4())
    work_dir = UPLOAD_DIR / job_id
    work_dir.mkdir(parents=True, exist_ok=True)

    master_path = work_dir / filename
    file.save(str(master_path))

    if filename.endswith('.xlsx'):
        from ns_engine.nsm_io import xlsx_to_nsm
        nsm_filename = filename.rsplit('.', 1)[0] + '.nsm'
        nsm_path = work_dir / nsm_filename
        try:
            xlsx_to_nsm(str(master_path), str(nsm_path))
            filename = nsm_filename
            logger.info('Converted .xlsx to .nsm: %s', nsm_path)
        except Exception as e:
            logger.error('Failed to convert to .nsm: %s', e)
            return jsonify({'error': f'Failed to convert to .nsm: {e}'}), 500

    set_active_master(str(work_dir), filename)

    touch_heartbeat(job_id)
    request._ns_upload_filename = filename
    logger.info('Uploaded: %s -> %s (job: %s)',
                filename, work_dir / filename, job_id)
    return jsonify({
        'job_id': job_id,
        'filename': filename,
    })


def run_ns_command_isolated(args, work_dir, master_filename):
    """Run NS CLI with an isolated copy of the master file.

    Each parallel task gets its own subdirectory with a private master
    copy so that concurrent processes never contend on the same file.
    For .nsm masters, the embedded xlsx is extracted and used directly
    to avoid openpyxl compatibility issues during export.
    """
    task_dir = work_dir / f'_task_{uuid.uuid4().hex[:8]}'
    task_dir.mkdir(exist_ok=True)

    try:
        task_master = task_dir / master_filename
        shutil.copy2(str(work_dir / master_filename), str(task_master))

        isolated_args = []
        replace_next = False
        for a in args:
            if replace_next:
                isolated_args.append(str(task_master))
                replace_next = False
            elif a == '--master':
                isolated_args.append(a)
                replace_next = True
            else:
                isolated_args.append(a)

        result = run_ns_command(isolated_args)

        for f in task_dir.iterdir():
            if f.name == master_filename:
                continue
            if f.name.startswith('__TMP__') or f.name.startswith('_tmp_'):
                continue
            dest = work_dir / f.name
            try:
                shutil.move(str(f), str(dest))
            except Exception:
                logger.warning('Could not move %s to %s', f, dest)

        return result
    except Exception as e:
        import traceback
        logger.error('run_ns_command_isolated error: %s\n%s', e, traceback.format_exc())
        return None
    finally:
        shutil.rmtree(str(task_dir), ignore_errors=True)


def _wait_readable(path, timeout=10.0, interval=0.15):
    """Wait until a file can be opened for reading.

    Handles transient Windows file locks caused by AV scanners (e.g. Defender)
    that scan newly-moved files before releasing them for normal access.
    Returns True if readable within timeout, False otherwise.
    """
    import time as _t
    deadline = _t.time() + timeout
    while _t.time() < deadline:
        try:
            with open(str(path), 'rb') as f:
                f.read(1)
            return True
        except (PermissionError, OSError):
            _t.sleep(interval)
    return False


def run_ns_command_subprocess(args, work_dir, master_filename):
    """Run NS CLI as a separate subprocess for true parallelism.

    Unlike run_ns_command_isolated (which uses run_cli with the global
    _cli_lock), each call here launches a fresh Python interpreter.
    This bypasses the lock entirely, enabling genuine concurrent execution
    across multiple CPU cores.
    """
    task_dir = work_dir / f'_task_{uuid.uuid4().hex[:8]}'
    task_dir.mkdir(exist_ok=True)
    try:
        task_master = task_dir / master_filename
        shutil.copy2(str(work_dir / master_filename), str(task_master))

        isolated_args = []
        replace_next = False
        for a in args:
            if replace_next:
                isolated_args.append(str(task_master))
                replace_next = False
            elif a == '--master':
                isolated_args.append(a)
                replace_next = True
            else:
                isolated_args.append(a)

        worker = BASE_DIR / 'ns_engine' / '_ns_worker.py'
        # On Windows, suppress the console window that subprocess.run creates
        # by default for Python scripts.
        extra_kwargs = {}
        if sys.platform == 'win32':
            extra_kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
        proc = subprocess.run(
            [sys.executable, str(worker)] + isolated_args,
            cwd=str(task_dir),
            capture_output=True,
            text=True,
            timeout=SUBPROCESS_TIMEOUT,
            **extra_kwargs,
        )
        if proc.returncode != 0:
            logger.warning(
                'subprocess worker FAILED rc=%d\nstderr=%s\nstdout=%s',
                proc.returncode,
                (proc.stderr or '')[:1000],
                (proc.stdout or '')[:500],
            )
        else:
            logger.debug('subprocess worker OK rc=%d', proc.returncode)

        moved_file_names = []  # track successfully-moved file names for race-free SVG lookup
        for f in task_dir.iterdir():
            if f.name == master_filename:
                continue
            if f.name.startswith('__TMP__') or f.name.startswith('_tmp_'):
                continue
            dest = work_dir / f.name
            try:
                shutil.move(str(f), str(dest))
                if _wait_readable(dest):
                    moved_file_names.append(f.name)
                    # Cache SVG content in memory so /svg_raw/ can serve it
                    # without touching the file again (avoids Windows AV re-lock).
                    if f.name.lower().endswith('.svg'):
                        try:
                            with open(str(dest), 'rb') as _cf:
                                svg_bytes = _cf.read()
                            job_id = work_dir.name
                            with _svg_mem_cache_lock:
                                if job_id not in _svg_mem_cache:
                                    _svg_mem_cache[job_id] = {}
                                _svg_mem_cache[job_id][f.name] = svg_bytes
                        except Exception:
                            pass
                else:
                    logger.warning('File not readable after 10s: %s', dest)
            except Exception:
                logger.warning('Could not move %s to %s', f, dest)
        # Attach moved_file_names to proc so _run_task can use them for race-free SVG lookup
        proc.moved_file_names = moved_file_names
        return proc
    except subprocess.TimeoutExpired:
        logger.error(
            'run_ns_command_subprocess: timed out after 600s. args=%s', args)
        # Return a dummy result object that callers can handle like a failed proc
        class _TimeoutResult:
            returncode = -1
            stdout = ('[ERROR] 処理がタイムアウトしました (600秒超過)。\n'
                      'データセットが大きすぎる可能性があります。\n'
                      'L3 All Areas の場合はエリア数またはデバイス数を減らすか、'
                      'Per Area でご確認ください。')
            stderr = 'TimeoutExpired'
            moved_file_names = []
        return _TimeoutResult()
    except Exception as exc:
        import traceback
        logger.error('run_ns_command_subprocess error: %s\n%s', exc, traceback.format_exc())
        return None
    finally:
        shutil.rmtree(str(task_dir), ignore_errors=True)


def _run_show_command(show_cmd, master_path):
    """Run a single show command and return its stdout as a single array string."""
    args = show_cmd.strip().split()
    result = run_ns_command(args + ['--master', master_path, '--one_msg'])
    if result and result.stdout:
        return result.stdout.strip()
    return ''


def generate_ai_context_parallel(work_dir, master_filename):
    """Generate AI Context file by running show commands in parallel."""
    master_path = str(work_dir / master_filename)
    show_commands = _cfg.get('ai_context_show_commands', [])
    limit = _cfg.get('parallel_limit', 5)

    task_dirs = []
    futures_map = {}

    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=limit) as executor:
            for idx, cmd in enumerate(show_commands):
                task_dir = work_dir / f'_aictx_{uuid.uuid4().hex[:8]}'
                task_dir.mkdir(exist_ok=True)
                task_master = task_dir / master_filename
                shutil.copy2(master_path, str(task_master))
                task_dirs.append(task_dir)

                future = executor.submit(_run_show_command, cmd, str(task_master))
                futures_map[future] = (idx, cmd)

            results_by_idx = {}
            for future in concurrent.futures.as_completed(futures_map):
                idx, cmd = futures_map[future]
                try:
                    results_by_idx[idx] = future.result()
                except Exception as e:
                    logger.warning('AI context show command failed: %s - %s', cmd, e)
                    results_by_idx[idx] = f'[ERROR] {cmd}: {e}'
    finally:
        for td in task_dirs:
            shutil.rmtree(str(td), ignore_errors=True)

    content = "'''\nBasic response policy\n'''\n"
    content += '* You are a network specialist and technical consultant at Cisco.\n'
    content += '* You provide specific, logical answers to broad and technical questions or consultations, including your reasoning, and you possess a high level of analytical ability.\n'
    content += '* A customer has provided output from the OSS tool "Network Sketcher" using a show command.\n\n'
    content += "'''\nAll data in the master file\n'''\n"

    for idx in range(len(show_commands)):
        label = show_commands[idx].replace(' ', '_')
        content += f'** {label}\n{results_by_idx.get(idx, "")}\n'
    content += '\n'

    cmd_list_path = BASE_DIR / 'ns_engine' / 'nsm_extensions_cmd_list.txt'
    if cmd_list_path.is_file():
        with open(str(cmd_list_path), 'r', encoding='utf-8') as f:
            content += f.read()

    basename = master_filename.replace('[MASTER]', '').replace('.xlsx', '').replace('.nsm', '')
    out_path = work_dir / f'[AI_Context]{basename}.txt'
    with open(str(out_path), 'w', encoding='utf-8') as f:
        f.write(content + '\n')

    return out_path.is_file()


@app.route('/generate_step/<job_id>/<step>', methods=['POST'])
def generate_step(job_id, step):
    job_id = sanitize_job_id(job_id)
    if not job_id:
        return jsonify({'error': 'Invalid job ID'}), 400

    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        return jsonify({'error': 'Job not found'}), 404

    master_filename = get_active_master(str(work_dir))
    if not master_filename:
        return jsonify({'error': 'Master file not found'}), 404

    master_path = str(work_dir / master_filename)

    if step == 'init':
        init_commands = {
            'area':      ['show', 'area', '--master', master_path],
            'device':    ['show', 'device', '--master', master_path],
            'l1_link':   ['show', 'l1_link', '--master', master_path],
            'attribute': ['show', 'attribute', '--master', master_path],
        }

        init_results = {}
        for key, cmd in init_commands.items():
            init_results[key] = run_ns_command(cmd)

        areas = []
        r = init_results.get('area')
        if r and r.returncode == 0:
            for line in r.stdout.strip().split('\n'):
                line = line.strip()
                if line and not line.startswith('[') and not line.endswith('_wp_'):
                    areas.append(line)

        device_count = 0
        r = init_results.get('device')
        if r and r.returncode == 0:
            for line in r.stdout.strip().split('\n'):
                line = line.strip()
                if line and not line.startswith('['):
                    device_count += 1

        link_count = 0
        r = init_results.get('l1_link')
        if r and r.returncode == 0:
            for line in r.stdout.strip().split('\n'):
                line = line.strip()
                if line and line.startswith('[['):
                    link_count += 1

        attribute_titles = []
        r = init_results.get('attribute')
        if r and r.returncode == 0:
            lines = r.stdout.strip().split('\n')
            if lines:
                header = lines[0].strip()
                if header.startswith('[') and header.endswith(']'):
                    import ast as _ast
                    try:
                        cols = _ast.literal_eval(header)
                        attribute_titles = [c for c in cols
                                            if c and c != 'Device Name']
                    except Exception:
                        pass

        meta = {
            'areas': areas,
            'device_count': device_count,
            'link_count': link_count,
            'attribute_titles': attribute_titles,
        }
        try:
            with open(str(work_dir / '.meta.json'), 'w', encoding='utf-8') as mf:
                json.dump(meta, mf)
        except Exception:
            pass

        return jsonify(dict(success=True, **meta))

    elif step == 'device_file':
        result = run_ns_command_isolated(
            ['export', 'device_file', '--master', master_path],
            work_dir, master_filename,
        )
        if result and '[ERROR]' not in (result.stdout or ''):
            return jsonify({'success': True})
        msg = ''
        if result:
            msg = (result.stdout or '') + (result.stderr or '')
        return jsonify({'success': False, 'message': msg})

    elif step == 'l1_diagram':
        l1_type = request.args.get('type', 'all_areas_tag')
        if l1_type not in ('all_areas', 'all_areas_tag', 'per_area', 'per_area_tag'):
            return jsonify({'success': False, 'message': 'Invalid L1 type'}), 400
        cmd_args = [
            'export', 'l1_diagram',
            '--master', master_path,
            '--type', l1_type,
        ]
        attr = request.args.get('attribute', '')
        if attr:
            cmd_args += ['--attribute', attr]
        fmt = request.args.get('format', 'pptx')
        if fmt == 'svg':
            cmd_args += ['--format', 'svg']
        result = run_ns_command_isolated(cmd_args, work_dir, master_filename)
        if result and '[ERROR]' not in (result.stdout or ''):
            return jsonify({'success': True})
        msg = ''
        if result:
            msg = (result.stdout or '') + (result.stderr or '')
        return jsonify({'success': False, 'message': msg})

    elif step == 'l1_preview_svg':
        cmd_args = [
            'export', 'l1_diagram',
            '--master', master_path,
            '--type', 'all_areas_tag',
            '--format', 'svg',
        ]
        attr = request.args.get('attribute', '')
        if attr:
            cmd_args += ['--attribute', attr]
        try:
            result = run_ns_command_isolated(cmd_args, work_dir, master_filename)
            if result and '[ERROR]' not in (result.stdout or ''):
                return jsonify({'success': True})
            msg = ''
            if result:
                msg = (result.stdout or '') + (result.stderr or '')
            return jsonify({'success': False, 'message': msg})
        except Exception as e:
            import traceback
            logger.error('L1 SVG preview error: %s\n%s', e, traceback.format_exc())
            return jsonify({'success': False, 'message': str(e)})

    elif step == 'l3_preview_svg':
        cmd_args = [
            'export', 'l3_diagram',
            '--master', master_path,
            '--type', 'all_areas',
            '--format', 'svg',
        ]
        attr = request.args.get('attribute', '')
        if attr:
            cmd_args += ['--attribute', attr]
        try:
            result = run_ns_command_isolated(cmd_args, work_dir, master_filename)
            if result and '[ERROR]' not in (result.stdout or ''):
                return jsonify({'success': True})
            msg = ''
            if result:
                msg = (result.stdout or '') + (result.stderr or '')
            return jsonify({'success': False, 'message': msg})
        except Exception as e:
            import traceback
            logger.error('L3 SVG preview error: %s\n%s', e, traceback.format_exc())
            return jsonify({'success': False, 'message': str(e)})

    elif step == 'l3_per_area_preview_svg':
        cmd_args = [
            'export', 'l3_diagram',
            '--master', master_path,
            '--type', 'per_area',
            '--format', 'svg',
        ]
        attr = request.args.get('attribute', '')
        if attr:
            cmd_args += ['--attribute', attr]
        try:
            result = run_ns_command_isolated(cmd_args, work_dir, master_filename)
            if result and '[ERROR]' not in (result.stdout or ''):
                area_svgs = _find_per_area_svgs(work_dir, '[L3_DIAGRAM]PerArea_')
                if area_svgs:
                    return jsonify({'success': True,
                                    'first_svg': area_svgs[0],
                                    'filter_prefix': _common_prefix(area_svgs)})
                return jsonify({'success': True})
            msg = ''
            if result:
                msg = (result.stdout or '') + (result.stderr or '')
            return jsonify({'success': False, 'message': msg})
        except Exception as e:
            import traceback
            logger.error('L3 Per Area SVG preview error: %s\n%s', e, traceback.format_exc())
            return jsonify({'success': False, 'message': str(e)})

    elif step == 'l1_per_area_tag_preview_svg':
        cmd_args = [
            'export', 'l1_diagram',
            '--master', master_path,
            '--type', 'per_area_tag',
            '--format', 'svg',
        ]
        attr = request.args.get('attribute', '')
        if attr:
            cmd_args += ['--attribute', attr]
        try:
            result = run_ns_command_isolated(cmd_args, work_dir, master_filename)
            if result and '[ERROR]' not in (result.stdout or ''):
                area_svgs = _find_per_area_svgs(work_dir, '[L1_DIAGRAM]PerAreaTag_')
                if area_svgs:
                    return jsonify({'success': True,
                                    'first_svg': area_svgs[0],
                                    'filter_prefix': _common_prefix(area_svgs)})
                return jsonify({'success': True})
            msg = ''
            if result:
                msg = (result.stdout or '') + (result.stderr or '')
            return jsonify({'success': False, 'message': msg})
        except Exception as e:
            import traceback
            logger.error('L1 Per Area Tag SVG preview error: %s\n%s', e, traceback.format_exc())
            return jsonify({'success': False, 'message': str(e)})

    elif step == 'l2_preview_svg':
        area = request.args.get('area', '')
        if not area:
            return jsonify({'success': False, 'message': 'Area not specified'})
        cmd_args = [
            'export', 'l2_diagram',
            '--master', master_path,
            '--area', area,
            '--format', 'svg',
        ]
        attr = request.args.get('attribute', '')
        if attr:
            cmd_args += ['--attribute', attr]
        try:
            result = run_ns_command_isolated(cmd_args, work_dir, master_filename)
            if result and '[ERROR]' not in (result.stdout or ''):
                return jsonify({'success': True})
            msg = ''
            if result:
                msg = (result.stdout or '') + (result.stderr or '')
            return jsonify({'success': False, 'message': msg})
        except Exception as e:
            import traceback
            logger.error('L2 SVG preview error: %s\n%s', e, traceback.format_exc())
            return jsonify({'success': False, 'message': str(e)})

    elif step == 'l2_diagram':
        area = request.args.get('area', '')
        if not area:
            return jsonify({'success': False, 'message': 'Area not specified'})
        try:
            l2_cmd = [
                'export', 'l2_diagram',
                '--master', master_path,
                '--area', area,
            ]
            fmt = request.args.get('format', 'pptx')
            if fmt == 'svg':
                l2_cmd += ['--format', 'svg']
            result = run_ns_command_isolated(l2_cmd, work_dir, master_filename)
            if result and '[ERROR]' not in (result.stdout or ''):
                return jsonify({'success': True})
            msg = ''
            if result:
                msg = (result.stdout or '') + (result.stderr or '')
            return jsonify({'success': False, 'message': msg})
        except Exception as e:
            import traceback
            logger.error('L2 diagram error: %s\n%s', e, traceback.format_exc())
            return jsonify({'success': False, 'message': str(e)})

    elif step == 'l3_diagram':
        l3_type = request.args.get('type', 'all_areas')
        if l3_type not in ('all_areas', 'per_area'):
            return jsonify({'success': False, 'message': 'Invalid L3 type'}), 400
        cmd_args = [
            'export', 'l3_diagram',
            '--master', master_path,
            '--type', l3_type,
        ]
        attr = request.args.get('attribute', '')
        if attr:
            cmd_args += ['--attribute', attr]
        fmt = request.args.get('format', 'pptx')
        if fmt == 'svg':
            cmd_args += ['--format', 'svg']
        try:
            result = run_ns_command_isolated(cmd_args, work_dir, master_filename)
            if result and '[ERROR]' not in (result.stdout or ''):
                return jsonify({'success': True})
            msg = ''
            if result:
                msg = (result.stdout or '') + (result.stderr or '')
            return jsonify({'success': False, 'message': msg})
        except Exception as e:
            import traceback
            logger.error('L3 diagram error: %s\n%s', e, traceback.format_exc())
            return jsonify({'success': False, 'message': str(e)})

    elif step == 'ai_context_file':
        try:
            ok = generate_ai_context_parallel(work_dir, master_filename)
            if ok:
                return jsonify({'success': True})
            return jsonify({'success': False, 'message': 'Failed to create AI context file'})
        except Exception as e:
            logger.error('AI context generation error: %s', e)
            return jsonify({'success': False, 'message': str(e)})

    return jsonify({'error': 'Unknown step'}), 400


def _run_deferred_sync(master_path):
    """Run L2/L3 sync that was deferred during batch command execution.

    Safety net: should not normally be reached because the last syncable
    command in a batch always runs with skip_sync=False.
    """
    logger.info('Running deferred L2/L3 sync for %s', os.path.basename(master_path))
    try:
        from ns_engine import nsm_adapter
        nsm_adapter.run_cli([
            'export', 'master_file_backup', '--master', master_path,
        ], cwd=os.path.dirname(master_path))
    except Exception as e:
        logger.warning('Deferred sync fallback: %s', e)


# --- CLI command failure detection (for the Update Master "Run" results) ---
# The engine signals failure inconsistently: most paths print the bracketed
# marker, but in two casings ([ERROR] and [Error]); others print a bare
# "Error:" line or a specific rejection phrase with no marker at all; and
# uncaught exceptions surface only via a non-zero return code / stderr.
# A stdout-only "[ERROR]" check therefore misclassified many real failures as
# success (green). The matchers below are intentionally narrow (anchored /
# bracketed / fixed phrases) to avoid flagging legitimate success output.
_ERR_BRACKET_RE = re.compile(r'\[error\]', re.IGNORECASE)   # [ERROR] or [Error]
_ERR_LINE_RE = re.compile(r'(?im)^\s*Error:')               # bare "Error:" line start
_FAILURE_PHRASES = (
    'Input must start with',
    '[WARNING] No ',
    'No matching entry found',
    'is already used in an existing link',
    'IP Address format is invalid',
    'Invalid virtual port name',
    'Invalid portchannel name',
    'Validation error:',
    'Not found in the argument',
    'Not found in arguments',
    'ERROR and STOP',
    'cannot find section',
    # bulk partial-failure summaries (some entries failed; surface as failure)
    'with errors',
    'error(s)',
    'Not found:',
)


def _command_failed(result, stdout):
    """Return True if a CLI command result represents a failure.

    Catches: missing result, non-zero return code (uncaught exceptions /
    TypeError paths that print only to stderr), the bracketed error marker in
    either casing, a bare "Error:" line, and known rejection/partial-failure
    phrases. Deliberately ignores idempotent "No change made" no-ops.
    """
    if result is None:
        return True
    if getattr(result, 'returncode', 0) != 0:
        return True
    if _ERR_BRACKET_RE.search(stdout) or _ERR_LINE_RE.search(stdout):
        return True
    return any(p in stdout for p in _FAILURE_PHRASES)


@app.route('/run_commands/<job_id>', methods=['POST'])
def run_commands(job_id):
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        abort(404)

    master_filename = get_active_master(str(work_dir))
    if not master_filename:
        return jsonify({'success': False, 'message': 'Master file not found'})

    data = request.get_json(silent=True) or {}
    commands_text = data.get('commands', '').strip()
    if not commands_text:
        return jsonify({'success': False, 'message': 'No commands provided'})

    lines = [l.strip() for l in commands_text.splitlines() if l.strip() and not l.strip().startswith('#')]

    ALLOWED_VERBS = RUNNABLE_VERBS
    MUTATING_VERBS = {'add', 'delete', 'rename'}
    has_mutation = any(
        (shlex.split(l)[0] if l else '') in MUTATING_VERBS
        for l in lines
        if l and l.split()[0] in ALLOWED_VERBS
    )

    if has_mutation:
        version = data.get('version', 1)
        original_basename = find_master_file(str(work_dir)) or master_filename
        base_no_ext = os.path.splitext(original_basename)[0].replace('[MASTER]', '')
        base_no_ver = re.sub(r'_\d+$', '', base_no_ext)
        ext = os.path.splitext(master_filename)[1] or '.xlsx'
        new_master_name = f'[MASTER]{base_no_ver}_{version}{ext}'
        new_master_path = str(work_dir / new_master_name)
        try:
            shutil.copy2(str(work_dir / master_filename), new_master_path)
        except Exception as e:
            logger.warning('Failed to copy master for versioning: %s', e)
            return jsonify({'success': False, 'message': 'Failed to create versioned master'})
        target_master_path = new_master_path
    else:
        new_master_name = None
        target_master_path = str(work_dir / master_filename)

    SYNCABLE_SUBCMDS = {'l1_link_bulk', 'device_location'}

    parsed_lines = []
    for line in lines:
        try:
            parts = shlex.split(line)
        except ValueError:
            parts = line.split()
        parsed_lines.append((line, parts))

    def _subcmd(parts):
        if len(parts) >= 2:
            return parts[1]
        return ''

    def _needs_sync(parts):
        return len(parts) >= 2 and parts[0] in MUTATING_VERBS and _subcmd(parts) in SYNCABLE_SUBCMDS

    results = []
    errors = []
    pending_sync_master = None

    for idx, (line, parts) in enumerate(parsed_lines):
        if not parts:
            continue
        if parts[0] not in ALLOWED_VERBS:
            results.append({'command': line, 'success': True, 'skipped': True, 'output': f'Skipped (only add/delete/rename/show are supported)'})
            continue

        is_last = (idx == len(parsed_lines) - 1)
        next_parts = parsed_lines[idx + 1][1] if not is_last else []
        current_syncable = _needs_sync(parts)
        next_syncable = _needs_sync(next_parts) if next_parts else False

        skip_sync = current_syncable and next_syncable and not is_last

        cmd_args = parts + ['--master', target_master_path]
        if skip_sync:
            cmd_args.append('--skip_sync')
            pending_sync_master = target_master_path

        result = run_ns_command(cmd_args)
        stdout = (result.stdout or '') if result else ''
        stderr = (result.stderr or '') if result else ''
        ok = not _command_failed(result, stdout)
        results.append({'command': line, 'success': ok, 'output': stdout + stderr})
        if not ok:
            errors.append(line)

        if not skip_sync and pending_sync_master and current_syncable:
            pending_sync_master = None

    if pending_sync_master:
        _run_deferred_sync(pending_sync_master)

    if has_mutation and new_master_name:
        set_active_master(str(work_dir), new_master_name)
        # Invalidate SVG grid cache so next buildSvgGrid() regenerates from
        # updated master. Drop all scopes (initial + every per-area entry).
        with _svg_grid_last_attr_lock:
            _svg_grid_last_attr.pop(job_id, None)

    return jsonify({
        'success': len(errors) == 0,
        'results': results,
        'errors': errors,
        'updated_master': new_master_name,
    })


def _find_per_area_svgs(work_dir, prefix):
    """Return sorted list of per-area SVG filenames in work_dir matching prefix."""
    try:
        files = sorted([
            f for f in os.listdir(str(work_dir))
            if f.lower().endswith('.svg') and f.startswith(prefix)
        ])
        return files
    except Exception:
        return []


def _common_prefix(filenames):
    """Return the longest common prefix among filenames (for viewer filtering)."""
    if not filenames:
        return ''
    prefix = filenames[0]
    for name in filenames[1:]:
        while not name.startswith(prefix):
            prefix = prefix[:-1]
            if not prefix:
                return ''
    return prefix


def _safe_area_for_filename(name):
    """Convert area name to the safe form used in SVG filenames (matches L1/L3 svg create logic)."""
    safe = re.sub(r'[\\/*?:"<>|]', '-', str(name))
    safe = safe.strip('. ')
    return safe or 'Area'


def _find_svgs_for_cell(work_dir, cell_id, file_list=None):
    """Return sorted list of SVG filenames that belong to the given grid cell.

    If file_list is provided, match against that list instead of scanning work_dir.
    This avoids race conditions on Windows where os.listdir may miss recently-moved
    files when another thread is concurrently modifying the same directory.

    cell_id format:
      l1_all                  → [L1_DIAGRAM]AllAreasTag_ (single file)
      l1_per_area_<area>      → [L1_DIAGRAM]PerAreaTag_*_<safe_area>.svg (per-area file)
      l3_all                  → [L3_DIAGRAM]AllAreas_ (single file)
      l3_per_area_<area>      → [L3_DIAGRAM]PerArea_*_<safe_area>.svg (per-area file)
      l2_area_<area>          → [L2_DIAGRAM]<area>_ (single file)
      l2_all                  → [L2_DIAGRAM]AllAreas_ (single file)
    """
    if file_list is not None:
        files = [f for f in file_list if f.lower().endswith('.svg')]
    else:
        try:
            files = [f for f in os.listdir(str(work_dir)) if f.lower().endswith('.svg')]
        except Exception:
            return []

    if cell_id == 'l1_all':
        return sorted([f for f in files if f.startswith('[L1_DIAGRAM]AllAreasTag_')])
    if cell_id.startswith('l1_per_area_'):
        area = cell_id[len('l1_per_area_'):]
        safe = _safe_area_for_filename(area)
        return sorted([f for f in files
                        if f.startswith('[L1_DIAGRAM]PerAreaTag_') and
                        os.path.splitext(f)[0].endswith('_' + safe)])
    if cell_id == 'l3_all':
        return sorted([f for f in files if f.startswith('[L3_DIAGRAM]AllAreas_')])
    if cell_id.startswith('l3_per_area_'):
        area = cell_id[len('l3_per_area_'):]
        safe = _safe_area_for_filename(area)
        return sorted([f for f in files
                        if f.startswith('[L3_DIAGRAM]PerArea_') and
                        os.path.splitext(f)[0].endswith('_' + safe)])
    if cell_id == 'l2_all':
        return sorted([f for f in files if f.startswith('[L2_DIAGRAM]AllAreas_')])
    if cell_id.startswith('l2_area_'):
        area = cell_id[len('l2_area_'):]
        return sorted([f for f in files if f.startswith(f'[L2_DIAGRAM]{area}_')])
    return []


def _resolve_layer_files(work_dir, scope, file_list=None):
    """Map a viewer scope to the L1/L2/L3 SVG filenames for that scope.

    Used by ``/diagram_preview/<job_id>`` (live tabbed viewer) and the CLI
    bridge in ``/diagram_preview_html/<job_id>`` to translate a UI cell
    selection (e.g. clicking the L2 thumbnail of "Office_LAN") into a
    consistent triple of L1/L2/L3 filenames so the user can switch between
    layers within a single window.

    Args:
        work_dir: Path-like work directory for the job (used as fallback
                  when ``file_list`` is not supplied).
        scope: ``'all'`` for All Areas, or ``'area:<area_name>'`` for the
               per-area scope.
        file_list: Optional iterable of filenames to resolve against. When
                   omitted the directory is scanned. Mirrors the
                   ``_find_svgs_for_cell`` semantics.

    Returns:
        ``dict`` ``{'l1': filename or None, 'l2': ..., 'l3': ...}`` where
        ``None`` marks a layer whose SVG has not been generated yet (the
        live viewer falls back to a "Not generated yet" placeholder for
        that tab). Filenames are basenames within ``work_dir``.
    """
    scope = (scope or '').strip()
    if not scope:
        return {'l1': None, 'l2': None, 'l3': None}

    if scope == 'all':
        cell_ids = {'l1': 'l1_all', 'l2': 'l2_all', 'l3': 'l3_all'}
    elif scope.startswith('area:'):
        area = scope[len('area:'):]
        if not area:
            return {'l1': None, 'l2': None, 'l3': None}
        cell_ids = {
            'l1': 'l1_per_area_' + area,
            # NOTE: L2 per-area cell uses 'l2_area_<area>' (NOT 'l2_per_area_'),
            # matching _find_svgs_for_cell's existing contract. Renaming would
            # invalidate every cached client URL, so we keep this asymmetry.
            'l2': 'l2_area_' + area,
            'l3': 'l3_per_area_' + area,
        }
    else:
        return {'l1': None, 'l2': None, 'l3': None}

    resolved = {}
    for layer, cell_id in cell_ids.items():
        matches = _find_svgs_for_cell(work_dir, cell_id, file_list=file_list)
        # Multiple matches are theoretically possible (e.g. legacy + tag-less
        # variants) but in practice only one per cell is produced; we pick
        # the first deterministically so the live viewer is reproducible.
        resolved[layer] = matches[0] if matches else None
    return resolved


@app.route('/svg_grid_stream/<job_id>')
def svg_grid_stream(job_id):
    """SSE endpoint: generate SVG thumbnails and stream completion events.

    Two operating modes (selected via ``?mode=`` query):

      - ``mode=initial`` (default): produce the always-shown thumbnails
        (``l1_all``, ``l2_all`` and ``l3_all``) and trigger AI Context
        generation in parallel. No per-area diagrams are generated here --
        those wait for the user to pick an area in the dropdown.

      - ``mode=area``: requires an ``area`` query parameter. Generates the
        three per-area cells (``l1_per_area_<area>``, ``l3_per_area_<area>``,
        ``l2_area_<area>``) only. AI Context is not regenerated -- the
        initial-mode call has already produced it.

    Cache scoping: ``_svg_grid_last_attr`` is keyed by ``(job_id, scope)``
    where scope is ``__initial__`` or the area name. This lets attribute
    changes invalidate everything in lock-step while still allowing each
    area to be cached independently.
    """
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        abort(404)

    master_filename = get_active_master(str(work_dir))
    if not master_filename:
        abort(404)

    attr = request.args.get('attribute', '')
    mode = (request.args.get('mode') or 'initial').strip().lower()
    if mode not in ('initial', 'area'):
        abort(400)
    limit = _cfg.get('parallel_limit', 5)

    # Resolve the cache scope key + the list of tasks for this request.
    if mode == 'initial':
        scope = _SVG_GRID_INITIAL_SCOPE
        area_arg = None
    else:
        area_arg = (request.args.get('area') or '').strip()
        if not area_arg:
            abort(400)
        scope = area_arg

    # Build task list: (cell_id, cmd_args)
    def make_args(base, extra=None):
        a = list(base) + ['--master', str(work_dir / master_filename), '--format', 'svg']
        if extra:
            a += extra
        if attr and any(x in ('l1_diagram', 'l3_diagram') for x in base):
            a += ['--attribute', attr]
        return a

    # Build tasks: (result_cell_ids, cmd_args)
    # A single task may produce SVGs for multiple cells (e.g. per_area generates one SVG per area)
    def _task(cell_ids_or_single, cmd):
        return (cell_ids_or_single if isinstance(cell_ids_or_single, list) else [cell_ids_or_single], cmd)

    if mode == 'initial':
        tasks = [
            _task('l1_all', make_args(['export', 'l1_diagram', '--type', 'all_areas_tag'])),
            _task('l2_all', make_args(['export', 'l2_diagram', '--type', 'all_areas'])),
            _task('l3_all', make_args(['export', 'l3_diagram', '--type', 'all_areas'])),
        ]
    else:
        # Single-area on-demand path. CLI was extended in v3.1.2a to accept
        # --area for l1_diagram (per_area_tag) and l3_diagram (per_area).
        tasks = [
            _task(f'l1_per_area_{area_arg}',
                  make_args(['export', 'l1_diagram', '--type', 'per_area_tag',
                             '--area', area_arg])),
            _task(f'l3_per_area_{area_arg}',
                  make_args(['export', 'l3_diagram', '--type', 'per_area',
                             '--area', area_arg])),
            _task(f'l2_area_{area_arg}',
                  make_args(['export', 'l2_diagram', '--area', area_arg])),
        ]

    import base64 as _b64_svg

    def _serve_cell_from_cache(cell_ids):
        """Try to build per_cell/per_cell_b64 for cell_ids from _svg_mem_cache or disk.
        Returns (per_cell, per_cell_b64) if all cells are satisfied, else None."""
        with _svg_mem_cache_lock:
            job_cache = dict(_svg_mem_cache.get(job_id, {}))
        cached_names = list(job_cache.keys()) if job_cache else []

        per_cell = {}
        per_cell_b64 = {}
        for cid in cell_ids:
            # Try memory cache name list first (fast), then full disk scan as fallback
            found = _find_svgs_for_cell(work_dir, cid, file_list=cached_names) if cached_names else None
            if not found:
                found = _find_svgs_for_cell(work_dir, cid)
            if not found:
                return None  # This cell's SVG is not available yet
            per_cell[cid] = found
            fname = found[0]
            svg_bytes = job_cache.get(fname)
            if not svg_bytes:
                # Load from disk and warm the memory cache
                try:
                    with open(str(work_dir / fname), 'rb') as _df:
                        svg_bytes = _df.read()
                    with _svg_mem_cache_lock:
                        _svg_mem_cache.setdefault(job_id, {})[fname] = svg_bytes
                except Exception:
                    pass
            if svg_bytes and len(svg_bytes) <= 300 * 1024:
                # Normalize width/height before base64-inlining so the
                # decoded data: URL also stays under Chromium's 16384-px
                # canvas limit. The original bytes in _svg_mem_cache are
                # untouched and still served full-size by /svg_raw/.
                per_cell_b64[fname] = _b64_svg.b64encode(
                    _normalize_svg_for_thumb(svg_bytes)
                ).decode('ascii')
        return per_cell, per_cell_b64

    def _run_task(cell_ids, cmd_args):
        try:
            # Use subprocess execution for true parallelism (bypasses _cli_lock)
            result = run_ns_command_subprocess(cmd_args, work_dir, master_filename)
            success = result is not None and result.returncode == 0
            per_cell = {}
            per_cell_b64 = {}  # {filename: base64_str} for files <= 300 KB
            if success:
                # Use the list of files actually moved by the subprocess (race-free)
                # instead of scanning work_dir with os.listdir, which can miss entries
                # on Windows when another thread concurrently renames files there.
                moved_names = getattr(result, 'moved_file_names', None)
                for cid in cell_ids:
                    found = _find_svgs_for_cell(work_dir, cid, file_list=moved_names)
                    per_cell[cid] = found
                    if found:
                        svg_path = work_dir / found[0]
                        try:
                            if svg_path.stat().st_size <= 300 * 1024:
                                with open(str(svg_path), 'rb') as _f:
                                    # Normalize root <svg> width/height to keep
                                    # the inlined data: URL within Chromium's
                                    # 16384-px image rasterizer limit.
                                    per_cell_b64[found[0]] = _b64_svg.b64encode(
                                        _normalize_svg_for_thumb(_f.read())
                                    ).decode('ascii')
                        except Exception:
                            pass
            else:
                for cid in cell_ids:
                    per_cell[cid] = []
            return per_cell, success, per_cell_b64
        except Exception as exc:
            logger.error('svg_grid_stream task error: %s', exc)
            return {cid: [] for cid in cell_ids}, False, {}

    def _stream_events(per_cell, per_cell_b64, success=True):
        """Yield SSE data lines for a completed per_cell mapping."""
        for cell_id, svgs in per_cell.items():
            first_svg = svgs[0] if svgs else None
            event_dict = {
                'cell': cell_id,
                'files': svgs,
                'filter': '',
                'error': not success or len(svgs) == 0,
            }
            if first_svg and first_svg in per_cell_b64:
                event_dict['svg_b64'] = per_cell_b64[first_svg]
            yield f'data: {json.dumps(event_dict)}\n\n'

    def generate():
        try:
            # Flush HTTP response headers + keep the TCP/TLS connection alive
            # immediately so EventSource.onopen fires in the browser even when
            # the first real cell event is delayed (e.g. the L2 All Areas
            # subprocess can take 15-20s before yielding anything for large
            # masters). Without this, some Windows network stacks / proxies /
            # antivirus / TLS MITM layers treat the long silence as a dead
            # connection and drop it. SSE comment lines (starting with `:`)
            # are ignored by browsers but force Werkzeug to flush the
            # response so the client sees the connection as established.
            yield ': keepalive\n\n'

            basename = master_filename.replace('[MASTER]', '').replace('.xlsx', '').replace('.nsm', '')
            ai_filename_expected = f'[AI_Context]{basename}.txt'

            # --- Cache fast-path (per scope) ---
            # If the same attribute was used last time for *this scope* and all
            # the scope's SVGs still exist (memory or disk), serve everything
            # from cache without re-running the CLI.
            with _svg_grid_last_attr_lock:
                job_scopes = _svg_grid_last_attr.get(job_id) or {}
                last_attr = job_scopes.get(scope)
            if last_attr is not None and last_attr == attr:
                cache_results = []
                all_cached = True
                for cell_ids, _cmd in tasks:
                    cached = _serve_cell_from_cache(cell_ids)
                    if cached:
                        cache_results.append((cell_ids, cached))
                    else:
                        all_cached = False
                        break

                if all_cached:
                    for _cids, (per_cell, per_cell_b64) in cache_results:
                        yield from _stream_events(per_cell, per_cell_b64)
                    if mode == 'initial':
                        # AI context: reuse existing file if already generated
                        ai_filename = (ai_filename_expected
                                       if (work_dir / ai_filename_expected).exists()
                                       else None)
                        if not ai_filename:
                            try:
                                ai_ok = generate_ai_context_parallel(work_dir, master_filename)
                                ai_filename = ai_filename_expected if ai_ok else None
                            except Exception as exc:
                                logger.error('AI context error (cache path): %s', exc)
                        yield 'data: ' + json.dumps({'done': True, 'ai_filename': ai_filename}) + '\n\n'
                    else:
                        # Area mode: AI Context belongs to the initial scope,
                        # so just report which area finished.
                        yield 'data: ' + json.dumps({'done': True, 'area': area_arg}) + '\n\n'
                    return

            # --- Full generation path ---
            ai_done_event = threading.Event()
            ai_result = [None]

            if mode == 'initial':
                # Start AI context generation in a background thread so it runs
                # concurrently with SVG generation rather than sequentially.
                def _run_ai_context():
                    try:
                        ai_ok = generate_ai_context_parallel(work_dir, master_filename)
                        ai_result[0] = ai_filename_expected if ai_ok else None
                    except Exception as exc:
                        logger.error('AI context generation error in svg_grid_stream: %s', exc)
                    finally:
                        ai_done_event.set()

                threading.Thread(target=_run_ai_context, daemon=True).start()
            else:
                # Area mode never (re)generates AI Context.
                ai_done_event.set()

            with concurrent.futures.ThreadPoolExecutor(max_workers=limit) as executor:
                futures = {
                    executor.submit(_run_task, cell_ids, cmd_args): cell_ids
                    for cell_ids, cmd_args in tasks
                }
                # Replaced concurrent.futures.as_completed() with a polling
                # loop that emits an SSE heartbeat every few seconds while
                # tasks are still running. This keeps the connection alive
                # for slow tasks (e.g. the L2 All Areas subprocess for large
                # masters can complete after L1 but before L3) so the
                # browser's EventSource does not get torn down by idle-aware
                # network middleboxes (Windows TLS proxy, AV, corporate
                # firewall, etc.). Heartbeats are SSE comment lines (`:`)
                # and are silently ignored by the EventSource API.
                _SSE_HEARTBEAT_INTERVAL = 5.0
                pending = set(futures.keys())
                while pending:
                    done, pending = concurrent.futures.wait(
                        pending,
                        timeout=_SSE_HEARTBEAT_INTERVAL,
                        return_when=concurrent.futures.FIRST_COMPLETED,
                    )
                    if not done:
                        yield ': keepalive\n\n'
                        continue
                    for future in done:
                        per_cell, success, per_cell_b64 = future.result()
                        yield from _stream_events(per_cell, per_cell_b64, success)

            # Record the attribute for this scope so the next call can use cache fast-path
            with _svg_grid_last_attr_lock:
                _svg_grid_last_attr.setdefault(job_id, {})[scope] = attr

            if mode == 'initial':
                # Wait for AI context thread (usually already finished by now).
                # Send periodic heartbeats while waiting so the connection
                # stays alive even if AI Context outlasts the SVG cells.
                while not ai_done_event.wait(timeout=_SSE_HEARTBEAT_INTERVAL):
                    yield ': keepalive\n\n'
                yield 'data: ' + json.dumps({'done': True, 'ai_filename': ai_result[0]}) + '\n\n'
            else:
                yield 'data: ' + json.dumps({'done': True, 'area': area_arg}) + '\n\n'
        except GeneratorExit:
            pass
        except Exception as exc:
            logger.error('svg_grid_stream error: %s', exc)
            yield 'data: ' + json.dumps({'done': True, 'error': True}) + '\n\n'

    return Response(
        stream_with_context(generate()),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache',
            'X-Accel-Buffering': 'no',
            'Connection': 'keep-alive',
        }
    )


@app.route('/files/<job_id>')
def list_files(job_id):
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        abort(404)

    master_filename = get_active_master(str(work_dir)) or ''
    files = collect_generated_files(str(work_dir), master_filename)
    return jsonify({'files': files})


def _convert_svg_for_visio(svg_bytes: bytes) -> bytes:
    """CSS class rules をインライン SVG 属性に変換して Visio 互換 SVG を生成する。
    <style> ブロックを削除し、各要素の class 属性を対応する CSS プロパティとして
    インライン属性に展開することで Visio が正しく解釈できる SVG を生成する。
    元の SVG ファイルは変更しない。"""
    svg = svg_bytes.decode('utf-8', errors='replace')

    # Step1: <style> ブロックから CSS クラスルールを抽出
    css_class_rules = {}  # {class_name: {prop: val}}
    text_font_family = None
    style_m = re.search(r'<style[^>]*>(.*?)</style>', svg, re.DOTALL | re.IGNORECASE)
    if style_m:
        css_text = style_m.group(1)
        for m in re.finditer(r'\.([\w-]+)\s*\{([^}]*)\}', css_text):
            props = {}
            for decl in m.group(2).split(';'):
                if ':' in decl:
                    k, v = decl.split(':', 1)
                    props[k.strip()] = v.strip()
            if props:
                css_class_rules[m.group(1)] = props
        tm = re.search(r'(?:^|\s)text\s*\{([^}]*)\}', css_text, re.MULTILINE)
        if tm:
            for decl in tm.group(1).split(';'):
                if ':' in decl:
                    k, v = decl.split(':', 1)
                    if k.strip() == 'font-family':
                        text_font_family = v.strip()
        svg = svg[:style_m.start()] + svg[style_m.end():]  # <style> ブロック削除

    # Step2: class="X" を CSS プロパティに展開してインライン属性化
    def inline_class(m):
        tag = m.group(0)
        cls_name = m.group(1)
        props = css_class_rules.get(cls_name, {})
        tag = re.sub(r'\s+class="[^"]*"', '', tag)  # class 属性を削除
        new_attrs = [f'{p}="{v}"' for p, v in props.items() if f'{p}=' not in tag]
        if new_attrs:
            tag = re.sub(r'\s*(/?)>\s*$', ' ' + ' '.join(new_attrs) + r' \1>', tag.rstrip())
        return tag
    svg = re.sub(r'<\S[^<>]*\sclass="([\w-]+)"[^<>]*/?>', inline_class, svg, flags=re.DOTALL)

    # Step3: <text> 要素に font-family を付与 (未設定のもののみ)
    if text_font_family:
        ff = text_font_family
        svg = re.sub(
            r'(<text\b(?:(?!font-family)[^>])*?>)',
            lambda m: m.group(0) if 'font-family' in m.group(0)
                      else re.sub(r'(<text\b)', rf'\1 font-family="{ff}"', m.group(0)),
            svg, flags=re.DOTALL
        )

    # Step4: dominant-baseline を Visio 互換に変換 (Visio は SVG2 dominant-baseline 非対応)
    # "central": テキスト中心を y に合わせる → dy="0.35em" で下方シフトして近似
    # "hanging": テキスト上端を y に合わせる → dy="0.75em" で下方シフトして近似
    svg = re.sub(r'\bdominant-baseline="central"', 'dy="0.35em"', svg)
    svg = re.sub(r'\bdominant-baseline="hanging"', 'dy="0.75em"', svg)

    # Step5: orient="auto-start-reverse" を "auto" に変換 (SVG2 構文 → Visio は SVG1.1 のみ対応)
    svg = svg.replace('orient="auto-start-reverse"', 'orient="auto"')

    # Step6: marker-end / marker-start 属性を削除 (Visio でダイヤモンドマーカーがズレるため非表示)
    svg = re.sub(r'\s*marker-end="[^"]*"', '', svg)
    svg = re.sub(r'\s*marker-start="[^"]*"', '', svg)

    # Step7: 外枠 rect を削除 (白塗り+ストローク付きの大きな rect がVisioでズレるため)
    # 判定条件: stroke あり + 白塗り (white or rgb(255,255,255)) + SVG 寸法の 70% 以上
    _svg_dim = re.search(r'<svg\b[^>]+\bwidth="([0-9.]+)"[^>]+\bheight="([0-9.]+)"', svg)
    if _svg_dim:
        _svgw, _svgh = float(_svg_dim.group(1)), float(_svg_dim.group(2))
        def _drop_outer_rect(m):
            tag = m.group(0)
            if 'stroke=' not in tag:
                return tag
            if 'fill="rgb(255,255,255)"' not in tag and 'fill="white"' not in tag:
                return tag
            # 角丸なし (rx=0 または rx 属性なし) のみ対象
            rxm = re.search(r'\brx="([^"]*)"', tag)
            if rxm and float(rxm.group(1)) > 0:
                return tag
            wm = re.search(r'\bwidth="([0-9.]+)"', tag)
            hm = re.search(r'\bheight="([0-9.]+)"', tag)
            if wm and hm:
                if float(wm.group(1)) >= _svgw * 0.7 and float(hm.group(1)) >= _svgh * 0.7:
                    return ''
            return tag
        svg = re.sub(r'<rect\b[^>]*/>', _drop_outer_rect, svg)

    # Step8: <rect> の fill 欠落・fill="none" を fill="white" に変換
    # SVG のデフォルト fill は black のため、fill 属性がない rect も Visio では黒く表示される。
    # fill="none" (明示的な透明) と fill 属性なし (暗黙的な black) の両方を white に変換する。
    def _fix_none_fill_rect(m):
        tag = m.group(0)
        if 'fill="none"' in tag:
            return tag.replace('fill="none"', 'fill="white"')
        if 'fill=' not in tag:
            # fill 属性が存在しない → SVG デフォルト black → white を追加
            return tag.replace('<rect', '<rect fill="white"', 1)
        return tag
    svg = re.sub(r'<rect\b[^>]*/>', _fix_none_fill_rect, svg)

    return svg.encode('utf-8')


# ---------------------------------------------------------------------------
# Cisco-stencil substitution for the "Download for draw.io (stencil)" button.
# ---------------------------------------------------------------------------
# All shape classification and the device/WayPoint discrimination are driven
# entirely from master attribute columns and stencil_aliases.json. Color-based
# heuristics have been removed because users may freely edit the engine's
# fill-colour palette, which would otherwise silently break classification.
#
# Master attribute columns used:
#   - "Default"      : 'DEVICE' or 'WayPoint'. Decides whether a cell gets
#                      stencilized at all and short-circuits WayPoint -> cloud.
#   - "Stencil Type" : Optional explicit override. Either an mxgraph.* shape id
#                      (passthrough) or a friendly alias resolved via
#                      stencil_aliases.json 'exact' lookup.
#   - All others     : Free-text columns (Role/Model/OS/Type/...) scanned by
#                      stencil_aliases.json 'prefix' / 'contains' keyword
#                      lists when no explicit Stencil Type is provided.

# Master 'Default' column values that mark a stencilizable cell.
_STENCIL_DEFAULT_DEVICE = 'DEVICE'
_STENCIL_DEFAULT_WAYPOINT = 'WAYPOINT'
_STENCIL_DEFAULT_KINDS = (_STENCIL_DEFAULT_DEVICE, _STENCIL_DEFAULT_WAYPOINT)

# Master attribute columns excluded from the prefix/contains heuristic search
# because they are processed by dedicated steps in _resolve_stencil.
_STENCIL_RESERVED_COLS = {'Default'}

# Header names recognised as the master's "Stencil Type" column. Matched after
# normalisation (lowercase, no whitespace/hyphen/underscore/slash).
_STENCIL_HEADER_NAMES = {
    'stenciltype', 'stencil', 'drawiostencil', 'drawioshape',
    'shape', 'ステンシルタイプ', 'ステンシル',
}

# Cached alias index loaded from stencil_aliases.json (lazy + hot-reloaded
# based on file mtime). The cached 'index' is a dict with three sub-indices:
#   exact    : {normalised_alias: shape_id}            -- Stencil Type lookup
#   prefix   : {shape_id: [normalised_keyword, ...]}   -- name/attrs prefix
#   contains : {shape_id: [normalised_keyword, ...]}   -- name/attrs substring
# Access guarded by _stencil_alias_lock.
_stencil_alias_cache = {'mtime': 0.0, 'index': None, 'default': ''}
_stencil_alias_lock = threading.Lock()


def _normalise_alias(s):
    """Normalise a string for case/whitespace/hyphen/underscore-insensitive lookup."""
    if not s:
        return ''
    return re.sub(r'[\s\-_/]+', '', str(s)).lower()


def _empty_stencil_index():
    """Return a fresh empty structured stencil index."""
    return {'exact': {}, 'prefix': {}, 'contains': {}}


def _load_stencil_alias_index():
    """Load (and hot-reload) stencil_aliases.json.

    Returns:
        (index, default_shape) where index is a dict with three sub-indices:
            'exact'    -> {normalised_alias: shape_id}
            'prefix'   -> {shape_id: [normalised_keyword, ...]}
            'contains' -> {shape_id: [normalised_keyword, ...]}

    Supports both:
      - v2 schema (preferred): top-level 'shapes' object whose values are
        {'exact': [...], 'prefix': [...], 'contains': [...]}.
      - v1 schema (backward compatibility): top-level 'aliases' object whose
        values are flat keyword lists; each entry is treated as 'exact'-only.

    On any error returns (_empty_stencil_index(), '') so callers fall back
    to default_shape gracefully.
    """
    cfg_path = BASE_DIR / 'stencil_aliases.json'
    try:
        mtime = cfg_path.stat().st_mtime
    except OSError:
        return _empty_stencil_index(), ''
    with _stencil_alias_lock:
        cached = _stencil_alias_cache['index']
        if (mtime == _stencil_alias_cache['mtime']
                and isinstance(cached, dict)
                and cached.get('exact')):
            return cached, _stencil_alias_cache['default']
        try:
            with open(str(cfg_path), 'r', encoding='utf-8') as f:
                cfg = json.load(f)
        except (OSError, ValueError) as exc:
            logger.warning('Failed to load stencil_aliases.json: %s', exc)
            return _empty_stencil_index(), ''
        index = _empty_stencil_index()

        def _add_keywords(target_dict, shape_id, raw_list):
            if not raw_list:
                return
            bucket = target_dict.setdefault(shape_id, [])
            seen = set(bucket)
            for raw in raw_list:
                k = _normalise_alias(raw)
                if k and k not in seen:
                    bucket.append(k)
                    seen.add(k)

        shapes_def = cfg.get('shapes')
        if isinstance(shapes_def, dict):
            # v2 schema
            for shape_id, rules in shapes_def.items():
                if (not isinstance(shape_id, str)
                        or not shape_id.startswith('mxgraph.')
                        or not isinstance(rules, dict)):
                    continue
                for alias in rules.get('exact') or []:
                    key = _normalise_alias(alias)
                    if key and key not in index['exact']:
                        index['exact'][key] = shape_id
                _add_keywords(index['prefix'], shape_id, rules.get('prefix'))
                _add_keywords(index['contains'], shape_id, rules.get('contains'))
        else:
            # v1 backward compatibility: 'aliases' lists are exact-only.
            for shape_id, aliases in (cfg.get('aliases') or {}).items():
                if not isinstance(shape_id, str) or not shape_id.startswith('mxgraph.'):
                    continue
                for alias in aliases or []:
                    key = _normalise_alias(alias)
                    if key and key not in index['exact']:
                        index['exact'][key] = shape_id

        default_shape = cfg.get('default_shape') or 'mxgraph.cisco.servers.standard_host'
        _stencil_alias_cache['mtime'] = mtime
        _stencil_alias_cache['index'] = index
        _stencil_alias_cache['default'] = default_shape
        return index, default_shape


def _is_empty_attr(value):
    """True for empty/placeholder attribute values like '', '<EMPTY>', 'N/A'."""
    if value is None:
        return True
    s = str(value).strip()
    if not s:
        return True
    return s.lower() in {'<empty>', 'empty', 'n/a', 'na', 'none', '-'}


def _read_master_attributes(master_path):
    """Run 'show attribute' on the master and parse rows.

    Returns:
        (rows, headers) where rows = {device_name: {header: value}} and
        headers is the ordered list of column names (excluding 'Device Name').
        Falls back to ({}, []) if anything goes wrong; callers must tolerate.
    """
    if not master_path or not os.path.exists(master_path):
        return {}, []
    try:
        result = run_ns_command(['show', 'attribute', '--master', master_path])
    except Exception as exc:
        logger.warning('show attribute failed: %s', exc)
        return {}, []
    if not result or not getattr(result, 'stdout', None):
        return {}, []
    import ast as _ast
    headers = []
    rows = {}
    lines = [ln for ln in result.stdout.strip().split('\n') if ln.strip()]
    if not lines:
        return {}, []
    try:
        first = _ast.literal_eval(lines[0])
        if isinstance(first, list):
            headers = [str(c) for c in first]
    except (ValueError, SyntaxError):
        return {}, []
    if not headers or headers[0] != 'Device Name':
        return {}, []
    for ln in lines[1:]:
        try:
            parsed = _ast.literal_eval(ln)
        except (ValueError, SyntaxError):
            continue
        if not isinstance(parsed, list) or not parsed:
            continue
        device_name = str(parsed[0]).strip()
        if not device_name:
            continue
        attrs = {}
        for col_idx in range(1, min(len(parsed), len(headers))):
            cell = parsed[col_idx]
            # 'show attribute' emits each non-name cell as a Python-repr
            # STRING of the form "['<value>', [r, g, b]]", so we need a
            # second literal_eval to extract just the value. Tolerate the
            # rare case where the cell is already an unwrapped list.
            value = cell
            if isinstance(cell, str):
                stripped = cell.strip()
                if stripped.startswith('[') and stripped.endswith(']'):
                    try:
                        inner = _ast.literal_eval(stripped)
                        if isinstance(inner, list) and inner:
                            value = inner[0]
                        else:
                            value = inner
                    except (ValueError, SyntaxError):
                        value = cell
            elif isinstance(cell, list) and cell:
                value = cell[0]
            attrs[headers[col_idx]] = '' if value is None else str(value)
        rows[device_name] = attrs
    return rows, headers[1:]


def _resolve_stencil(name, attrs, alias_index, stencil_header):
    """Resolve final mxgraph shape id from master-driven sources only.

    Resolution priority (color-independent):

      Step 0 -- master 'Default' column == 'WayPoint' (case-insensitive)
                -> mxgraph.cisco.storage.cloud (replaces the legacy fill-colour
                heuristic; users can change palettes freely without affecting
                classification). Conversely, when Default == 'DEVICE' the
                cloud shape is removed from the candidate set in Step 3+4
                below, because a DEVICE row is by definition a real device
                (any cloud rendering must be requested explicitly via the
                Stencil Type column, which is processed in Step 1+2 first).

      Step 1 -- master 'Stencil Type' column begins with 'mxgraph.'
                -> passthrough verbatim (operator-supplied explicit shape id).

      Step 2 -- master 'Stencil Type' column matches an entry in the alias
                index 'exact' bucket (case-/whitespace-insensitive)
                -> mapped shape id.

      Step 3 -- longest 'prefix' match (any shape) against the normalised
                device name and the master's free-text attribute columns
                (Role/Model/OS/Type/...; Default and Stencil Type excluded).
                Used for short tokens such as 'fw', 'rt', 'ap', 'pc', 'sw'
                that should only match at the start of a name.

      Step 4 -- longest 'contains' match (any shape) against the same
                normalised sources. Used for full words such as 'firewall',
                'router', 'switch'.

      None   -- no signal found; caller applies default_shape.

    Tie-breaking: when two keywords have the same length, the one declared
    first in stencil_aliases.json wins (Python dict iteration order).
    """
    # Step 0a: WayPoint shortcut from the master Default column.
    default_val = ''
    if attrs:
        default_val = (attrs.get('Default') or '').strip().upper()
        if default_val == _STENCIL_DEFAULT_WAYPOINT:
            return 'mxgraph.cisco.storage.cloud'

    # Steps 1 + 2: explicit Stencil Type column override (always wins,
    # even for DEVICE rows that want a cloud rendering).
    if attrs and stencil_header:
        raw = (attrs.get(stencil_header) or '').strip()
        if raw and not _is_empty_attr(raw):
            if raw.startswith('mxgraph.'):
                return raw
            shape = alias_index['exact'].get(_normalise_alias(raw))
            if shape:
                return shape

    # Build the heuristic search corpus: device name + every free-text
    # attribute column except the reserved/Stencil-Type ones.
    sources = []
    if name:
        sources.append(name)
    if attrs:
        for col, val in attrs.items():
            if col in _STENCIL_RESERVED_COLS:
                continue
            if stencil_header and col == stencil_header:
                continue
            if val and not _is_empty_attr(val):
                sources.append(val)
    norm_sources = [s for s in (_normalise_alias(s) for s in sources) if s]
    if not norm_sources:
        return None

    # Step 0b: when the master explicitly marks this row as a DEVICE, drop
    # the cloud shape from the heuristic candidate set so that names or
    # role/model strings containing 'internet'/'sdwan'/'carrier'/'cloud'
    # cannot accidentally turn a real device into a cloud silhouette. If
    # the user really wants a cloud rendering for this row they can set
    # the Stencil Type column (handled by Step 1/2 above).
    cloud_shape = 'mxgraph.cisco.storage.cloud'
    exclude_cloud = (default_val == _STENCIL_DEFAULT_DEVICE)

    # Step 3: longest prefix match across all shapes.
    best = None
    for shape_id, prefixes in alias_index['prefix'].items():
        if exclude_cloud and shape_id == cloud_shape:
            continue
        for kw in prefixes:
            kw_len = len(kw)
            if best is not None and kw_len <= best[0]:
                continue
            for src in norm_sources:
                if src.startswith(kw):
                    best = (kw_len, shape_id)
                    break
    if best is not None:
        return best[1]

    # Step 4: longest substring (contains) match across all shapes.
    best = None
    for shape_id, contains in alias_index['contains'].items():
        if exclude_cloud and shape_id == cloud_shape:
            continue
        for kw in contains:
            kw_len = len(kw)
            if best is not None and kw_len <= best[0]:
                continue
            for src in norm_sources:
                if kw in src:
                    best = (kw_len, shape_id)
                    break
    if best is not None:
        return best[1]

    return None


def _build_cisco_style(shape, light=False):
    """Build an mxCell style string for a Cisco stencil that fits its bbox.

    Cloud-shaped stencils (typically used for WayPoint / WAN / Internet) are
    rendered with a white fill and a dark stroke so the cloud silhouette and
    the device label both stay legible against light page backgrounds.
    Other stencils use the Cisco-blue fill that matches the official palette.

    When light=True the whole stencil is rendered semi-transparent so that
    pre-existing diagram lines (e.g. dense L2 trunks/port-channels) remain
    visible behind the icon.
    """
    if shape == 'mxgraph.cisco.storage.cloud' or shape.endswith('.cloud'):
        fill_color = '#ffffff'
        stroke_color = '#036897'
    else:
        fill_color = '#036897'
        stroke_color = '#ffffff'
    style = (
        f'shape={shape};html=1;pointerEvents=1;dashed=0;'
        f'fillColor={fill_color};strokeColor={stroke_color};strokeWidth=2;'
        'verticalLabelPosition=bottom;verticalAlign=top;align=center;'
        'outlineConnect=0;'
    )
    if light:
        style += 'opacity=45;'
    return style


# Distinctive fillColors used by the Network Sketcher engine for L3 instance
# labels (e.g. "Default" / "VRF-1" / "MGMT") drawn inside device rectangles in
# L3 diagrams. These small rounded labels are the visible marker that a device
# hosts L3 instances; their presence is the signal we use to decide whether to
# render the Cisco stencil semi-transparent so the labels remain readable.
# Interface-name tags use #ffffff and are therefore not matched.
_L3_INSTANCE_FILLS = {'#e6e0ec'}


def _has_l3_instance_inside(cell, all_vertices):
    """True iff cell's bbox contains an L3 instance label cell.

    L3 instance labels are small rounded rectangles whose distinctive
    light-lavender fillColor (#e6e0ec) is reserved by the engine for this
    purpose. Containment alone is sufficient -- size/shape checks are
    intentionally omitted because the labels themselves are tiny (typically
    around 38x10 px) and would be excluded by any size threshold.

    Interface-name tags (white #ffffff fill) are filtered out by the
    fillColor requirement.
    """
    g = cell.find('mxGeometry')
    if g is None:
        return False
    try:
        x = float(g.get('x', 0)); y = float(g.get('y', 0))
        w = float(g.get('width', 0)); h = float(g.get('height', 0))
    except (ValueError, TypeError):
        return False
    if w <= 0 or h <= 0:
        return False
    tol = 0.5  # 1px tolerance for sub-pixel rounding

    for other in all_vertices:
        if other is cell:
            continue
        other_style = other.get('style') or ''
        m = re.search(r'fillColor=(#[0-9a-fA-F]{3,8})', other_style)
        if not m or m.group(1).lower() not in _L3_INSTANCE_FILLS:
            continue
        og = other.find('mxGeometry')
        if og is None:
            continue
        try:
            ox = float(og.get('x', 0)); oy = float(og.get('y', 0))
            ow = float(og.get('width', 0)); oh = float(og.get('height', 0))
        except (ValueError, TypeError):
            continue
        if ow <= 0 or oh <= 0:
            continue
        if (ox >= x - tol and oy >= y - tol
                and ox + ow <= x + w + tol
                and oy + oh <= y + h + tol):
            return True
    return False


def _apply_cisco_stencils(drawio_bytes, master_path, transparency='none'):
    """Replace device rectangles with Cisco stencil styles in-place.

    The original mxGeometry (x/y/width/height) is preserved so that connecting
    lines remain anchored to the device's previous bounding box.

    transparency:
      'none' -- all stencils opaque (default; used for L1 diagrams)
      'all'  -- all stencils 45% opaque (used for L2 dense diagrams so the
                underlying L2 segment / port-channel lines remain visible)
      'auto' -- per-device: light only when the device contains an L3 instance
                container (VRF / Default rectangle), so the inner labels remain
                visible behind the Cisco icon. Used for L3 diagrams.
    """
    import xml.etree.ElementTree as ET

    alias_index, default_shape = _load_stencil_alias_index()
    rows, headers = _read_master_attributes(master_path)
    stencil_header = next(
        (h for h in headers if _normalise_alias(h) in _STENCIL_HEADER_NAMES),
        None,
    )

    try:
        text = drawio_bytes.decode('utf-8')
    except UnicodeDecodeError:
        return drawio_bytes
    try:
        root = ET.fromstring(text)
    except ET.ParseError as exc:
        logger.warning('draw.io XML parse failed during stencilization: %s', exc)
        return drawio_bytes

    all_vertices = [c for c in root.iter('mxCell') if c.get('vertex') == '1']
    # Master-driven cell selection (color-independent). When the master
    # actually advertises a 'Default' column we use it as a strict gate so
    # that only rows explicitly marked DEVICE or WayPoint get stencilized.
    # If the column is absent (e.g. legacy masters), we fall back to "any
    # cell whose name matches a master row" so older files still work.
    has_default_col = 'Default' in headers

    for cell in all_vertices:
        name = (cell.get('value') or '').strip()
        if not name:
            continue
        attrs = rows.get(name)
        if not attrs:
            # Not a device in the master (page frame, area folder, annotation,
            # legend, etc.). Skip regardless of fill colour.
            continue
        if has_default_col:
            default_val = (attrs.get('Default') or '').strip().upper()
            if default_val not in _STENCIL_DEFAULT_KINDS:
                # 'Default' is set but to something other than DEVICE/WayPoint;
                # treat conservatively and skip rather than guess.
                continue
        shape = _resolve_stencil(
            name, attrs, alias_index, stencil_header
        ) or default_shape
        if transparency == 'all':
            is_light = True
        elif transparency == 'auto':
            is_light = _has_l3_instance_inside(cell, all_vertices)
        else:
            is_light = False
        cell.set('style', _build_cisco_style(shape, light=is_light))

    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            + ET.tostring(root, encoding='unicode')).encode('utf-8')


def _convert_svg_to_drawio(svg_bytes: bytes) -> bytes:
    """Convert a Network Sketcher SVG to draw.io (.drawio) XML format.

    Parses the SVG structure produced by Network Sketcher (both L3 inline-style
    and L1/L2 CSS-class-based SVGs) and maps each element to an mxCell in an
    mxGraphModel.  Text elements whose anchor falls within a rect are merged as
    the rect's value label.  Lines become floating edges (no source/target ids).

    Bug fixes vs. initial version:
    - CSS class rules from <style> block are now parsed and applied when inline
      attributes are absent (fixes missing stroke/fill on L1/L2 SVGs).
    - Edge mxGeometry now uses <mxPoint as="sourcePoint/targetPoint"> directly
      as mxGeometry children instead of the incorrect <Array> wrapper (fixes
      lines not rendering in draw.io).
    """
    import xml.etree.ElementTree as ET

    NS = 'http://www.w3.org/2000/svg'
    svg_text = svg_bytes.decode('utf-8', errors='replace')

    # --- helpers ---
    def _rgb_to_hex(rgb_str):
        """Convert 'rgb(r,g,b)' or named colours to #rrggbb. Returns None for 'none'."""
        if not rgb_str or rgb_str == 'none':
            return None
        rgb_str = rgb_str.strip()
        if rgb_str.startswith('#'):
            return rgb_str
        if rgb_str == 'white':
            return '#ffffff'
        if rgb_str == 'black':
            return '#000000'
        m = re.match(r'rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', rgb_str)
        if m:
            return '#{:02x}{:02x}{:02x}'.format(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        return '#000000'

    def _fv(el, attr, default=0.0):
        """Parse a float attribute from an element, return default if missing/invalid."""
        try:
            return float(el.get(attr, default))
        except (ValueError, TypeError):
            return float(default)

    def _parse_float_css(val_str, default=1.0):
        """Parse a CSS value like '2.0px' or '2.0' to float."""
        if not val_str:
            return default
        try:
            return float(str(val_str).replace('px', '').strip())
        except (ValueError, TypeError):
            return default

    def _text_in_rect(tx, ty, rx, ry, rw, rh):
        """Return True if text anchor (tx, ty) is inside the rect bounds."""
        return rx <= tx <= rx + rw and ry <= ty <= ry + rh

    # --- Step 1: extract CSS class rules from <style> block ---
    # This handles L1/L2 SVGs that use class="device", class="folder", etc.
    # Result: css_rules = {class_name: {'fill': ..., 'stroke': ..., 'stroke-width': ...}}
    css_rules = {}
    style_m = re.search(r'<style[^>]*>(.*?)</style>', svg_text, re.DOTALL | re.IGNORECASE)
    if style_m:
        css_text = style_m.group(1)
        for m in re.finditer(r'\.([\w-]+)\s*\{([^}]*)\}', css_text):
            props = {}
            for decl in m.group(2).split(';'):
                if ':' in decl:
                    k, v = decl.split(':', 1)
                    props[k.strip()] = v.strip()
            if props:
                css_rules[m.group(1)] = props

    def _css_prop(el, prop, fallback=None):
        """Get a CSS property for an element: inline attr first, then class rule."""
        # Inline attribute (L3 SVG style)
        inline = el.get(prop)
        if inline is not None:
            return inline
        # CSS class (L1/L2 SVG style)
        cls = el.get('class', '')
        for cls_name in cls.split():
            rule = css_rules.get(cls_name, {})
            if prop in rule:
                return rule[prop]
        return fallback

    def _parse_transform_rotation(el):
        """Extract rotation angle (degrees) from SVG transform='rotate(angle ...)'.
        Returns 0.0 if no rotation transform is present."""
        transform = el.get('transform', '')
        m = re.match(r'rotate\s*\(\s*([-\d.]+)', transform)
        if m:
            return float(m.group(1))
        return 0.0

    # --- Step 2: parse SVG ---
    svg_clean = re.sub(r'<\?xml[^>]*\?>', '', svg_text).strip()
    try:
        root = ET.fromstring(svg_clean)
    except ET.ParseError:
        return b'<mxfile><diagram><mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/></root></mxGraphModel></diagram></mxfile>'

    def _tag(el):
        t = el.tag
        if t.startswith('{'):
            t = t.split('}', 1)[1]
        return t

    rects = []
    texts = []
    lines = []
    for el in root.iter():
        tag = _tag(el)
        if tag == 'rect':
            # Skip the full-canvas background rect (class="bg") - draw.io has
            # its own white background; including it forces the bounding box to
            # be the full SVG canvas, preventing tight page-size calculation.
            cls = el.get('class', '')
            if 'bg' in cls.split():
                continue
            rects.append(el)
        elif tag == 'text':
            texts.append(el)
        elif tag == 'line':
            lines.append(el)

    # SVG canvas area used to detect decorative "page-frame" rects below.
    # Falls back to 0 (= disable frame detection) when the root <svg> lacks
    # explicit width/height (very rare for engine-emitted SVGs).
    _canvas_w = _fv(root, 'width', 0.0)
    _canvas_h = _fv(root, 'height', 0.0)
    _canvas_area = _canvas_w * _canvas_h
    # Threshold: any rect covering >= this fraction of the canvas is treated
    # as a decorative page frame (e.g. L3 inner page rect at ~77 % coverage).
    _FRAME_AREA_RATIO = 0.5
    _FRAME_FILL_TOKENS = {'white', 'rgb(255,255,255)', '#ffffff', 'none', ''}

    # --- Step 3: match text labels to rects ---
    texts_standalone = []   # texts not matched to any rect (e.g. diagram title)
    rect_data = []
    for el in rects:
        rx = _fv(el, 'x')
        ry = _fv(el, 'y')
        rw = _fv(el, 'width')
        rh = _fv(el, 'height')
        rotation = _parse_transform_rotation(el)
        # SVG class="tag" rects are interface-name labels; they must be rendered
        # AFTER lines so they appear in front (higher z-order in draw.io).
        # SVG class="folder"/"root" rects use dominant-baseline="hanging" text
        # (top-aligned), so their draw.io label must use verticalAlign=top.
        cls_parts = el.get('class', '').split()
        _fill_inline = (el.get('fill') or '').strip()

        # CSS-class-based (L1) tag detection
        is_tag = 'tag' in cls_parts
        # L2 SVGs use inline styles with no CSS classes.
        # Interface-name tags in L2: small height (≤20), rounded corners (rx>1),
        # white fill.  Purple-bg L2 segment rects have rgb(247,245,249) fill, so
        # they are excluded by the white-fill check.
        if not is_tag and not cls_parts and rh <= 20.0:
            _rx_val = _fv(el, 'rx', 0.0)
            if _rx_val > 1.0 and _fill_inline in ('white', 'rgb(255,255,255)'):
                is_tag = True

        # folder/root class → area box with top-centre label
        is_top_label = any(c in cls_parts for c in ('folder', 'root'))
        # top_label_align: 'center' for area boxes, 'left' for device containers
        top_label_align = 'center' if is_top_label else None

        # L2 device container: fill="rgb(250,251,247)" (beige), large rect
        # Device names use text-anchor="start" dominant-baseline="hanging" → top-left
        if not is_top_label and _fill_inline == 'rgb(250,251,247)':
            is_top_label = True
            top_label_align = 'left'

        # L3 FOLDER_NORMAL: no CSS class, transparent fill + light-gray stroke.
        # L3 SVGs do not emit class attributes; identify area boxes by their
        # specific fill/stroke combination (FOLDER_NORMAL style).
        if not is_top_label and not is_tag and _fill_inline == 'none':
            _stroke_inline = (el.get('stroke') or '').strip()
            if _stroke_inline == 'rgb(205,205,205)':
                is_top_label = True
                top_label_align = 'center'

        # Decorative "page-frame" detection. The L3 SVG includes an inner page
        # rectangle (e.g. 2467x1206 white-fill with thin black stroke) that is
        # large enough to contain ANY text point in the diagram. The original
        # "smallest containing rect" pairing would then incorrectly attach
        # free-floating IP labels (e.g. ".58") to this frame's value, badly
        # mispositioning them. Page-frame rects are still emitted as cells
        # (so the visible border survives) but are excluded from text matching
        # so labels fall through to the standalone-text path at their original
        # SVG coordinates.
        # Conservative criteria - all must hold:
        #   - not already classified as folder/area or tag
        #   - top-left is NOT (0, 0): preserves canvas-bg label behaviour for
        #     the diagram title (which is currently rendered as the bg cell's
        #     centre label and works as-is)
        #   - area >= 50 % of the SVG canvas area
        #   - fill is white or none (color-tinted device rects never qualify)
        is_frame = False
        if (not is_top_label and not is_tag
                and _canvas_area > 0
                and (rx != 0.0 or ry != 0.0)
                and rw * rh >= _FRAME_AREA_RATIO * _canvas_area
                and _fill_inline.lower() in _FRAME_FILL_TOKENS):
            is_frame = True

        rect_data.append({'x': rx, 'y': ry, 'w': rw, 'h': rh, 'el': el,
                          'label': '', 'rotation': rotation, 'font_size': None,
                          'font_color': None,
                          'is_tag': is_tag,
                          'is_top_label': is_top_label,
                          'top_label_align': top_label_align,
                          'is_frame': is_frame})

    for tel in texts:
        tx = _fv(tel, 'x')
        ty = _fv(tel, 'y')
        text_content = (tel.text or '').strip()
        if not text_content:
            continue
        best = None
        best_area = float('inf')
        for rd in rect_data:
            # Decorative page-frame rects are not eligible label hosts;
            # otherwise they would swallow free-floating IP labels that
            # don't fall inside any real device/area rect.
            if rd.get('is_frame'):
                continue
            if _text_in_rect(tx, ty, rd['x'], rd['y'], rd['w'], rd['h']):
                area = rd['w'] * rd['h']
                if area < best_area:
                    best_area = area
                    best = rd
        def _text_fs_and_color(tel):
            """Return (font_size, fill_color) for a text element.
            Checks inline font-size attribute first (L2 style), then CSS class."""
            fs = None
            cls = tel.get('class', '')
            for cls_name in cls.split():
                rule = css_rules.get(cls_name, {})
                if 'font-size' in rule:
                    fs = _parse_float_css(rule['font-size'], None)
                    break
            if fs is None:
                inline_fs = tel.get('font-size')
                if inline_fs:
                    fs = _parse_float_css(inline_fs, None)
            color = tel.get('fill', '')
            return fs, color

        if best is not None and not best['label']:
            best['label'] = text_content
            # Capture font-size and color from the matched text element.
            if best.get('font_size') is None or best.get('font_color') is None:
                fs, color = _text_fs_and_color(tel)
                if best.get('font_size') is None:
                    best['font_size'] = fs
                if best.get('font_color') is None:
                    best['font_color'] = color
        else:
            # Text not assigned to a rect (no containing rect, OR the containing
            # rect already has a label → e.g. L2 segment names like TESTVLAN-NAME
            # that sit just outside the interface tag but inside a larger device box).
            # Always preserve as standalone so no text is silently dropped.
            fs, color = _text_fs_and_color(tel)
            anchor = tel.get('text-anchor', 'middle')
            # SVG text y: adjust to the TOP of the text for draw.io coordinates.
            #   "hanging"  → y is already the top of the text
            #   "central"  → y is the vertical center; top = y - fs/2
            #   "auto"/default → y is the baseline; top ≈ y - fs
            # draw.io text cells add ~5 px internal top-padding even with
            # spacing=0, so subtract that to keep relative spacing identical to SVG.
            # 7.0 = zero gap; 5.0 = ~2 px visual gap (matches SVG appearance).
            _DRAWIO_TEXT_PAD = 5.0
            _dom_base = tel.get('dominant-baseline', 'auto')
            if _dom_base == 'hanging':
                sy = ty - _DRAWIO_TEXT_PAD
            elif _dom_base == 'central':
                sy = ty - (fs or 12) / 2.0 - _DRAWIO_TEXT_PAD
            else:
                sy = ty - (fs or 12) - _DRAWIO_TEXT_PAD
            texts_standalone.append({'x': tx, 'y': sy,
                                     'label': text_content, 'font_size': fs,
                                     'fill': color, 'anchor': anchor})

    # --- Step 3b: compute content bounding box and translation offset ---
    # Translate all coordinates so content starts at (MARGIN, MARGIN) and the
    # draw.io page is sized tightly around the content.  Without this, draw.io
    # opens with the full SVG canvas (e.g. 1992×1740 pt) at a small zoom level,
    # making the diagram appear tiny in the upper corner.
    MARGIN = 30  # padding (px) around content on all sides
    _bx = ([rd['x'] for rd in rect_data] + [rd['x'] + rd['w'] for rd in rect_data] +
           [_fv(el, 'x1') for el in lines] + [_fv(el, 'x2') for el in lines] +
           [st['x'] for st in texts_standalone])
    _by = ([rd['y'] for rd in rect_data] + [rd['y'] + rd['h'] for rd in rect_data] +
           [_fv(el, 'y1') for el in lines] + [_fv(el, 'y2') for el in lines] +
           [st['y'] for st in texts_standalone])
    if _bx and _by:
        _min_x, _min_y = min(_bx), min(_by)
        _max_x, _max_y = max(_bx), max(_by)
        ox = MARGIN - _min_x        # x translation offset
        oy = MARGIN - _min_y        # y translation offset
        page_w = int(_max_x - _min_x + 2 * MARGIN)
        page_h = int(_max_y - _min_y + 2 * MARGIN)
    else:
        svg_el0 = root if _tag(root) == 'svg' else root.find('.//{%s}svg' % NS) or root
        ox, oy = 0, 0
        page_w = int(_fv(svg_el0, 'width', 800))
        page_h = int(_fv(svg_el0, 'height', 600))

    # --- Step 4: build mxGraphModel XML ---
    mxfile = ET.Element('mxfile')
    diagram = ET.SubElement(mxfile, 'diagram')
    model = ET.SubElement(diagram, 'mxGraphModel')

    # Compute dx/dy to centre the page in a typical draw.io viewport.
    # draw.io rendering: screen_pos = (graph_pos + d) * scale
    # When draw.io opens a file it auto-fits the page to the available canvas.
    # The fit-scale = min(VP_W/page_w, VP_H/page_h).  Whichever dimension is
    # the "bottleneck" fills the viewport; the other dimension has slack that
    # can be centred by setting dx or dy appropriately.
    # Target viewport estimate: 800×550 px (1920×1080 screen minus sidebar/toolbar).
    _VP_W, _VP_H = 800, 550
    _aspect_w = page_w / _VP_W
    _aspect_h = page_h / _VP_H
    if _aspect_w >= _aspect_h:
        # Width is the bottleneck → horizontal fit; centre vertically with dy
        _fit_scale = _VP_W / page_w
        _dx = 0
        _dy = max(0, int((_VP_H - page_h * _fit_scale) / (2 * _fit_scale)))
    else:
        # Height is the bottleneck → vertical fit; centre horizontally with dx
        _fit_scale = _VP_H / page_h
        _dx = max(0, int((_VP_W - page_w * _fit_scale) / (2 * _fit_scale)))
        _dy = 0

    model.set('dx', str(_dx))
    model.set('dy', str(_dy))
    model.set('grid', '1')
    model.set('gridSize', '10')
    model.set('page', '1')
    model.set('pageScale', '1')
    model.set('pageWidth', str(page_w))
    model.set('pageHeight', str(page_h))

    root_el = ET.SubElement(model, 'root')
    ET.SubElement(root_el, 'mxCell', id='0')
    ET.SubElement(root_el, 'mxCell', id='1', parent='0')

    cell_id = 2

    def _add_rect_cell(rd):
        """Append an mxCell vertex for one rect entry and return the next cell id."""
        nonlocal cell_id
        el = rd['el']
        fill_str = _css_prop(el, 'fill', 'white')
        stroke_str = _css_prop(el, 'stroke', 'none')
        sw_raw = _css_prop(el, 'stroke-width', '1')
        rx_val = _fv(el, 'rx', 0.0)

        fill_hex = _rgb_to_hex(fill_str) or '#ffffff'
        stroke_hex = _rgb_to_hex(stroke_str)
        sw_float = _parse_float_css(sw_raw, 1.0)

        is_tag = rd.get('is_tag', False)
        is_top_label = rd.get('is_top_label', False)
        top_label_align = rd.get('top_label_align', 'center')  # 'center' or 'left'
        font_color = rd.get('font_color') or ''

        style_parts = [
            'shape=rectangle',
            f'rounded={"1" if rx_val > 0 else "0"}',
            f'fillColor={fill_hex}',
            f'strokeColor={stroke_hex}' if stroke_hex else 'strokeColor=none',
            f'strokeWidth={sw_float:.1f}',
            'html=1',
        ]
        if is_tag:
            # Interface-name tags: prevent text wrapping and allow the label to
            # overflow the narrow cell width so it always renders on one line.
            style_parts += ['whiteSpace=nowrap', 'overflow=visible', 'align=center']
        elif is_top_label:
            # Area/folder boxes: dominant-baseline="hanging" in SVG → top-aligned.
            # L2 device containers: text-anchor="start" → left-aligned.
            h_align = top_label_align or 'center'
            style_parts += ['whiteSpace=wrap', 'verticalAlign=top', f'align={h_align}']
        else:
            style_parts.append('whiteSpace=wrap')

        # Apply font color if captured from the SVG text element (e.g. red labels)
        if font_color:
            _fc_hex = _rgb_to_hex(font_color)
            if _fc_hex:
                style_parts.append(f'fontColor={_fc_hex}')

        # Apply SVG transform rotation to draw.io style (draw.io rotates around
        # cell centre, matching SVG rotate(angle cx cy) where cx/cy = rect centre).
        rotation = rd.get('rotation', 0.0)
        if rotation:
            style_parts.append(f'rotation={rotation:.2f}')
        # Apply font size from the matched text element's CSS class (tag-text: 8px)
        font_size = rd.get('font_size')
        if font_size is not None:
            style_parts.append(f'fontSize={int(round(font_size))}')
        style_str = ';'.join(style_parts) + ';'
        cell = ET.SubElement(root_el, 'mxCell',
                             id=str(cell_id),
                             value=rd['label'],
                             vertex='1',
                             parent='1',
                             style=style_str)
        geom = ET.SubElement(cell, 'mxGeometry',
                             x=str(round(rd['x'] + ox)),
                             y=str(round(rd['y'] + oy)),
                             width=str(round(rd['w'])),
                             height=str(round(rd['h'])))
        geom.set('as', 'geometry')
        cell_id += 1

    # --- rect → vertex mxCell ---
    # Render order: non-tag rects first (background areas, devices), then lines,
    # then tag rects last.  In draw.io, later mxCell elements are drawn on top,
    # so this ensures interface-name tags always appear in front of connection lines.
    for rd in rect_data:
        if not rd.get('is_tag'):
            _add_rect_cell(rd)

    # --- line → edge mxCell (floating, no source/target) ---
    # Fix: use <mxPoint as="sourcePoint/targetPoint"> directly as mxGeometry
    # children, NOT wrapped in <Array>.  The Array form is only for waypoints.
    for el in lines:
        x1 = _fv(el, 'x1')
        y1 = _fv(el, 'y1')
        x2 = _fv(el, 'x2')
        y2 = _fv(el, 'y2')
        stroke_str = _css_prop(el, 'stroke', 'black')
        sw_raw = _css_prop(el, 'stroke-width', '1')

        stroke_hex = _rgb_to_hex(stroke_str) or '#000000'
        sw_float = _parse_float_css(sw_raw, 1.0)

        style_str = (
            f'endArrow=none;startArrow=none;'
            f'strokeColor={stroke_hex};'
            f'strokeWidth={sw_float:.1f};'
            'html=1;'
        )
        cell = ET.SubElement(root_el, 'mxCell',
                             id=str(cell_id),
                             value='',
                             edge='1',
                             parent='1',
                             style=style_str)
        geom = ET.SubElement(cell, 'mxGeometry', relative='1')
        geom.set('as', 'geometry')

        sp = ET.SubElement(geom, 'mxPoint')
        sp.set('x', str(round(x1 + ox)))
        sp.set('y', str(round(y1 + oy)))
        sp.set('as', 'sourcePoint')

        tp = ET.SubElement(geom, 'mxPoint')
        tp.set('x', str(round(x2 + ox)))
        tp.set('y', str(round(y2 + oy)))
        tp.set('as', 'targetPoint')

        cell_id += 1

    # --- tag rects → vertex mxCell (rendered AFTER lines so they appear in front) ---
    for rd in rect_data:
        if rd.get('is_tag'):
            _add_rect_cell(rd)

    # --- standalone text → text mxCell ---
    # Includes diagram titles, L2 segment names (TESTVLAN-NAME, vlanXXX), and
    # any other text not matched to a rect (or whose rect already had a label).
    for st in texts_standalone:
        fs = st.get('font_size') or 12
        # Preserve SVG text colour (e.g. purple for L2 segment names)
        fill_color = st.get('fill', '')
        fc_hex = _rgb_to_hex(fill_color) if fill_color else None
        # Preserve text alignment (text-anchor: start → left, middle → center)
        anchor = st.get('anchor', 'middle')
        h_align = 'left' if anchor == 'start' else 'center'
        style_str = (
            f'text;html=1;align={h_align};verticalAlign=top;'
            'resizable=0;points=[];strokeColor=none;fillColor=none;'
            # Remove draw.io's default internal padding so the text appears
            # flush against the top of the cell – matching SVG positioning.
            'spacing=0;spacingTop=0;spacingBottom=0;'
            f'fontSize={int(round(fs))};'
        )
        if fc_hex:
            style_str += f'fontColor={fc_hex};'
        cell = ET.SubElement(root_el, 'mxCell',
                             id=str(cell_id),
                             value=st['label'],
                             vertex='1',
                             parent='1',
                             style=style_str)
        # Estimate cell width from text length and font size
        est_w = max(len(st['label']) * fs * 0.65, 80)
        geom = ET.SubElement(cell, 'mxGeometry',
                             x=str(round(st['x'] + ox)),
                             y=str(round(st['y'] + oy)),
                             width=str(round(est_w)),
                             height=str(round(fs * 1.8)))
        geom.set('as', 'geometry')
        cell_id += 1

    xml_str = ET.tostring(mxfile, encoding='unicode', xml_declaration=False)
    return ('<?xml version="1.0" encoding="UTF-8"?>\n' + xml_str).encode('utf-8')


@app.route('/download/<job_id>/<path:filename>')
def download_file(job_id, filename):
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    filepath = work_dir / filename

    try:
        filepath.resolve().relative_to(work_dir.resolve())
    except ValueError:
        abort(403)

    if filepath.is_file():
        request._ns_download_filename = filename
        return send_file(str(filepath), as_attachment=True)
    abort(404)


@app.route('/download_visio/<job_id>/<path:filename>')
def download_visio(job_id, filename):
    """Convert SVG to Visio-compatible format and serve as attachment."""
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    filepath = (work_dir / filename).resolve()
    try:
        filepath.relative_to(work_dir.resolve())
    except ValueError:
        abort(403)

    if not filename.lower().endswith('.svg'):
        abort(400)

    # メモリキャッシュ優先、なければディスク
    with _svg_mem_cache_lock:
        svg_bytes = _svg_mem_cache.get(job_id, {}).get(filename)
    if svg_bytes is None:
        if not filepath.is_file():
            abort(404)
        svg_bytes = filepath.read_bytes()

    visio_bytes = _convert_svg_for_visio(svg_bytes)
    # ファイル名の .svg 直前に _visio を付与
    visio_name = re.sub(r'\.svg$', '_visio.svg', filename, flags=re.IGNORECASE)
    if visio_name == filename:
        visio_name = filename + '_visio.svg'

    return Response(
        visio_bytes,
        mimetype='image/svg+xml',
        headers={'Content-Disposition': f'attachment; filename="{visio_name}"'}
    )


@app.route('/download_drawio/<job_id>/<path:filename>')
def download_drawio(job_id, filename):
    """Convert SVG to draw.io (.drawio) format and serve as attachment."""
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    filepath = (work_dir / filename).resolve()
    try:
        filepath.relative_to(work_dir.resolve())
    except ValueError:
        abort(403)

    if not filename.lower().endswith('.svg'):
        abort(400)

    with _svg_mem_cache_lock:
        svg_bytes = _svg_mem_cache.get(job_id, {}).get(filename)
    if svg_bytes is None:
        if not filepath.is_file():
            abort(404)
        svg_bytes = filepath.read_bytes()

    drawio_bytes = _convert_svg_to_drawio(svg_bytes)
    drawio_name = re.sub(r'\.svg$', '.drawio', filename, flags=re.IGNORECASE)
    if drawio_name == filename:
        drawio_name = filename + '.drawio'

    return Response(
        drawio_bytes,
        mimetype='application/xml',
        headers={'Content-Disposition': f'attachment; filename="{drawio_name}"'}
    )


@app.route('/download_drawio_stencil/<job_id>/<path:filename>')
def download_drawio_stencil(job_id, filename):
    """Convert SVG to draw.io and replace devices with Cisco stencils."""
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    filepath = (work_dir / filename).resolve()
    try:
        filepath.relative_to(work_dir.resolve())
    except ValueError:
        abort(403)

    if not filename.lower().endswith('.svg'):
        abort(400)

    master_filename = get_active_master(str(work_dir))
    if not master_filename:
        return Response(
            'Master file not found in this job. '
            'Stencil conversion requires the master file (.nsm or .xlsx).',
            status=400,
            mimetype='text/plain',
        )

    with _svg_mem_cache_lock:
        svg_bytes = _svg_mem_cache.get(job_id, {}).get(filename)
    if svg_bytes is None:
        if not filepath.is_file():
            abort(404)
        svg_bytes = filepath.read_bytes()

    drawio_bytes = _convert_svg_to_drawio(svg_bytes)
    if filename.startswith('[L2_DIAGRAM]'):
        transparency = 'all'
    elif filename.startswith('[L3_DIAGRAM]'):
        transparency = 'auto'
    else:
        transparency = 'none'
    try:
        stencil_bytes = _apply_cisco_stencils(
            drawio_bytes,
            str(work_dir / master_filename),
            transparency=transparency,
        )
    except Exception as exc:
        logger.error('Stencil conversion failed: %s', exc)
        return Response(
            f'Stencil conversion failed: {exc}',
            status=500,
            mimetype='text/plain',
        )

    out_name = re.sub(r'\.svg$', '_stencil.drawio', filename, flags=re.IGNORECASE)
    if out_name == filename:
        out_name = filename + '_stencil.drawio'

    return Response(
        stencil_bytes,
        mimetype='application/xml',
        headers={'Content-Disposition': f'attachment; filename="{out_name}"'},
    )


@app.route('/svg_raw/<job_id>/<path:filename>')
def svg_raw(job_id, filename):
    """Serve SVG file inline (not as attachment) for viewer embedding."""
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)
    work_dir = UPLOAD_DIR / job_id
    filepath = work_dir / filename
    try:
        filepath.resolve().relative_to(work_dir.resolve())
    except ValueError:
        abort(403)
    if not filename.lower().endswith('.svg'):
        abort(404)
    # ?thumb=1 -> serve a copy with normalized width/height so very tall or
    # very wide SVGs (especially L2/L3 All Areas for ~1000-device networks)
    # don't trip Chromium's 16384-px image rasterization limit and end up
    # rendered as a blank <img>. The actual file on disk and the in-memory
    # cache are never modified — only the response body is rewritten.
    thumb_mode = request.args.get('thumb') == '1'
    # Serve from in-memory cache when available (populated by run_ns_command_subprocess).
    # This avoids Windows AV file-lock issues that can occur when send_file() tries to
    # open the file after shutil.move() — even if _wait_readable() already confirmed it.
    with _svg_mem_cache_lock:
        svg_bytes = _svg_mem_cache.get(job_id, {}).get(filename)
    if svg_bytes is not None:
        if thumb_mode:
            svg_bytes = _normalize_svg_for_thumb(svg_bytes)
        return Response(svg_bytes, mimetype='image/svg+xml')
    # Fallback: serve from disk with retry (for files not yet cached, or cache miss).
    # Use continue (not break) when file is not yet visible, to handle transient
    # Windows AV / filesystem delays where the file appears shortly after move.
    for _attempt in range(3):
        if not filepath.is_file():
            if _attempt < 2:
                import time as _t_svg
                _t_svg.sleep(0.1)
                continue
            break
        try:
            if thumb_mode:
                # Read once so we can normalize before sending. send_file() can't
                # easily transform the body, so we hand back a Response instead.
                try:
                    with open(str(filepath), 'rb') as _fh:
                        raw = _fh.read()
                    return Response(_normalize_svg_for_thumb(raw),
                                    mimetype='image/svg+xml')
                except PermissionError:
                    if _attempt < 2:
                        import time as _t_svg
                        _t_svg.sleep(0.15 * (_attempt + 1))
                    continue
            return send_file(str(filepath), mimetype='image/svg+xml')
        except PermissionError:
            if _attempt < 2:
                import time as _t_svg
                _t_svg.sleep(0.15 * (_attempt + 1))
    abort(404)


@app.route('/export_nsm_to_xlsx/<job_id>', methods=['POST'])
def export_nsm_to_xlsx(job_id):
    """Convert the active .nsm master file to .xlsx for download."""
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        abort(404)

    master_filename = get_active_master(str(work_dir))
    if not master_filename or not master_filename.endswith('.nsm'):
        return jsonify({'error': 'No .nsm master file found'}), 400

    nsm_path = work_dir / master_filename
    xlsx_filename = master_filename.rsplit('.', 1)[0] + '.xlsx'
    xlsx_path = work_dir / xlsx_filename

    try:
        from ns_engine.nsm_io import nsm_to_xlsx
        nsm_to_xlsx(str(nsm_path), str(xlsx_path))
        logger.info('Exported .nsm to .xlsx: %s', xlsx_path)
        return send_file(
            str(xlsx_path),
            as_attachment=True,
            download_name=xlsx_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    except Exception as e:
        logger.error('Failed to export .nsm to .xlsx: %s', e)
        return jsonify({'error': str(e)}), 500


def _render_device_preview(job_id, tabs_data, master_basename):
    """Render interactive device data preview with 4 tabs and download buttons.

    Thin wrapper around the shared ``ns_engine.nsm_device_table_html``
    module so the Online edition and the Local MCP edition produce
    pixel-identical Device Preview HTML. ``job_id`` is retained for
    backward compatibility with callers but is no longer needed by the
    renderer itself.
    """
    from ns_engine.nsm_device_table_html import render_device_table_html
    return render_device_table_html(tabs_data, master_basename)


def _get_device_tabs_data(job_id):
    """Resolve the active master for ``job_id`` and delegate table building
    to ``ns_engine.nsm_device_table_html.build_device_tabs_data``.

    Returns ``(tabs_data, master_basename)`` or ``(None, None)`` if the
    job directory or master file is missing.
    """
    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        return None, None

    master_filename = get_active_master(str(work_dir))
    if not master_filename:
        return None, None

    master_path = str(work_dir / master_filename)
    from ns_engine.nsm_device_table_html import build_device_tabs_data
    return build_device_tabs_data(master_path)


@app.route('/device_preview_data/<job_id>')
def device_preview_data(job_id):
    """Return device tab data as JSON for the SVG grid thumbnail row."""
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    tabs_data, _basename = _get_device_tabs_data(job_id)
    if tabs_data is None:
        abort(404)

    # Return only first THUMB_ROWS rows per tab to keep the response small
    THUMB_ROWS = 4
    thumb_tabs = [
        {
            'id': t['id'],
            'label': t['label'],
            'headers': t['headers'],
            'rows': t['rows'][:THUMB_ROWS],
            'total': len(t['rows']),
        }
        for t in tabs_data
    ]
    return jsonify({'tabs': thumb_tabs})


@app.route('/device_preview/<job_id>')
def device_preview(job_id):
    """Render Device File preview (4 tabs) read directly from master file."""
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    tabs_data, master_basename = _get_device_tabs_data(job_id)
    if tabs_data is None:
        abort(404)

    html = _render_device_preview(job_id, tabs_data, master_basename)
    return html, 200, {'Content-Type': 'text/html; charset=utf-8'}


@app.route('/diagram_preview/<job_id>')
def diagram_preview(job_id):
    """Render the L1/L2/L3 tabbed live diagram preview.

    Replaces the per-SVG ``/preview/<job>/<filename>`` viewer for L1, L2 and
    L3 thumbnails: clicking any of those thumbnails opens this route, which
    presents three tabs (L1, L2, L3) for the same scope so the user can
    switch between layers in a single window. The tab content is fetched
    from ``/svg_raw/<job>/<filename>`` on first activation.

    Query parameters:
      - ``scope``: ``all`` (default) for All Areas, or ``area:<area_name>``
        for a specific area cell.
    """
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        abort(404)

    master_filename = get_active_master(str(work_dir))
    if not master_filename:
        abort(404)

    master_basename = (
        os.path.splitext(master_filename)[0].replace('[MASTER]', '')
    )

    scope = request.args.get('scope', 'all').strip() or 'all'
    if scope == 'all':
        scope_label = 'All Areas'
    elif scope.startswith('area:'):
        area_name = scope[len('area:'):].strip()
        if not area_name:
            abort(400)
        scope_label = area_name
    else:
        abort(400)

    layer_filenames = _resolve_layer_files(str(work_dir), scope)

    from ns_engine.nsm_l1l2l3_html import render_l1l2l3_html
    html = render_l1l2l3_html(
        layer_svgs={},
        master_basename=master_basename,
        scope_label=scope_label,
        mode='live',
        job_id=job_id,
        layer_filenames=layer_filenames,
    )
    return html, 200, {'Content-Type': 'text/html; charset=utf-8',
                       # /svg_raw/ caching is per-job; prevent the live viewer
                       # itself from being cached so newly generated layers
                       # appear when the user reloads.
                       'Cache-Control': 'no-store'}


@app.route('/diagram_preview_html/<job_id>')
def diagram_preview_html(job_id):
    """Build and serve the combined L1/L2/L3 standalone HTML for download.

    Backs the "↓ html(L1,L2,L3)" toolbar button in /diagram_preview/. The
    L1/L2/L3 SVGs that drive the live viewer already exist in the job
    work_dir (produced by the SVG grid SSE), so this route simply reads
    them, hands them to ``render_l1l2l3_html(mode='standalone')`` and
    streams the resulting self-contained HTML back as an attachment.

    We deliberately avoid spawning a CLI subprocess here: the engine's
    ``export combined_diagram`` is meant for headless use (Local MCP,
    ad-hoc scripting), and reusing the SVGs already on disk avoids both
    the regeneration latency and the worker's task-dir isolation that
    would otherwise hide the existing per-layer SVGs from the bundler.

    Query parameters:
      - ``scope``: ``all`` (default) for All Areas, or ``area:<area_name>``.
    """
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        abort(404)

    master_filename = get_active_master(str(work_dir))
    if not master_filename:
        abort(404)

    basename = master_filename.replace('[MASTER]', '')
    if basename.lower().endswith('.nsm'):
        basename = basename[: -len('.nsm')]
    elif basename.lower().endswith('.xlsx'):
        basename = basename[: -len('.xlsx')]

    scope = (request.args.get('scope') or 'all').strip()
    if scope == 'all':
        scope_label = 'All Areas'
        expected_filename = f'[L1L2L3_DIAGRAM]AllAreas_{basename}.html'
    elif scope.startswith('area:'):
        area_name = scope[len('area:'):].strip()
        if not area_name:
            abort(400)
        scope_label = area_name
        safe = _safe_area_for_filename(area_name)
        expected_filename = f'[L1L2L3_DIAGRAM]{safe}_{basename}.html'
    else:
        abort(400)

    layer_filenames = _resolve_layer_files(str(work_dir), scope)
    layer_svgs = {}
    for layer in ('l1', 'l2', 'l3'):
        fname = layer_filenames.get(layer)
        layer_svgs[layer] = None
        if fname and (work_dir / fname).is_file():
            try:
                with open(str(work_dir / fname), 'r', encoding='utf-8') as fh:
                    layer_svgs[layer] = fh.read()
            except Exception as exc:
                logger.warning('diagram_preview_html: could not read %s: %s',
                               fname, exc)

    if not any(layer_svgs.values()):
        # Nothing to bundle; surface a 404 rather than a confusing empty
        # download. The user must generate at least one layer SVG first
        # (initial SSE produces all-areas; per-area requires the dropdown
        # selection or "Generate Selected").
        abort(404)

    from ns_engine.nsm_l1l2l3_html import render_l1l2l3_html
    html_text = render_l1l2l3_html(
        layer_svgs=layer_svgs,
        master_basename=basename,
        scope_label=scope_label,
        mode='standalone',
    )

    target_path = work_dir / expected_filename
    try:
        with open(str(target_path), 'w', encoding='utf-8') as fh:
            fh.write(html_text)
    except Exception as exc:
        logger.error('diagram_preview_html: could not write %s: %s',
                     target_path, exc)
        abort(500)

    request._ns_download_filename = expected_filename
    return send_file(
        str(target_path),
        as_attachment=True,
        download_name=expected_filename,
        mimetype='text/html',
    )


def _render_svg_viewer(job_id, filename, work_dir, filter_prefix=''):
    """Render interactive SVG viewer with zoom, pan, page navigation, download."""
    all_svg = sorted([f for f in os.listdir(str(work_dir)) if f.lower().endswith('.svg')])
    if filter_prefix:
        svg_files = [f for f in all_svg if f.startswith(filter_prefix)]
        if not svg_files:
            svg_files = all_svg
    else:
        svg_files = all_svg
    current_idx = svg_files.index(filename) if filename in svg_files else 0
    total = len(svg_files)

    encoded = url_quote(filename, safe='')
    svg_url = f'/svg_raw/{job_id}/{encoded}'
    safe_title = filename.replace("'", "\\'").replace('"', '&quot;')

    svg_files_js = ', '.join(f'"{url_quote(f, safe="")}"' for f in svg_files)
    svg_names_js = ', '.join(f'"{f}"' for f in svg_files)

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>{safe_title}</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ background: #f0f2f5; color: #333; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
       display: flex; flex-direction: column; height: 100vh; overflow: hidden; }}
.toolbar {{
    display: flex; align-items: center; gap: 8px; padding: 10px 16px;
    background: #16213e; color: #fff; flex-shrink: 0; z-index: 10;
}}
.toolbar h1 {{ font-size: 14px; font-weight: 500; opacity: 0.9; white-space: nowrap;
               overflow: hidden; text-overflow: ellipsis; max-width: 50vw; }}
.toolbar button {{
    padding: 6px 14px; font-size: 13px; font-weight: 600;
    border: 1px solid #4A8FE7;
    border-radius: 6px; background: transparent; color: #4A8FE7; cursor: pointer;
    transition: all 0.2s; white-space: nowrap;
}}
.toolbar button:hover {{ background: #4A8FE7; color: #fff; }}
.toolbar button:disabled {{ opacity: 0.4; cursor: default; }}
.toolbar button.primary {{
    background: transparent; color: #e94560; border: 1px solid #e94560;
}}
.toolbar button.primary:hover {{ background: #e94560; color: #fff; border-color: #e94560; }}
.toolbar .sep {{ width: 1px; height: 20px; background: rgba(255,255,255,0.15); }}
.toolbar .status {{ font-size: 13px; opacity: 0.8; white-space: nowrap; }}
.toolbar .zoom-display {{ font-size: 12px; opacity: 0.6; min-width: 45px; text-align: center; }}
.toolbar .spacer {{ flex: 1; }}
.viewer {{
    flex: 1; overflow: hidden; cursor: grab; position: relative; background: #f0f2f5;
}}
.viewer.dragging {{ cursor: grabbing; }}
#svgContainer {{
    position: absolute; left: 0; top: 0;
    user-select: none; -webkit-user-select: none;
}}
#svgContainer svg {{
    display: block;
    overflow: visible;
}}
.ctx-menu {{
    position: fixed; background: #fff; border: 1px solid #ccc;
    border-radius: 6px; padding: 4px 0; z-index: 9999;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2); display: none;
    min-width: 160px;
}}
.ctx-menu li {{
    list-style: none; padding: 8px 18px; cursor: pointer;
    font-size: 13px; color: #333; white-space: nowrap;
}}
.ctx-menu li:hover {{ background: #4A8FE7; color: #fff; }}
.ctx-feedback {{
    position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%);
    background: rgba(0,0,0,0.7); color: #fff; padding: 6px 18px;
    border-radius: 20px; font-size: 13px; pointer-events: none;
    opacity: 0; transition: opacity 0.3s;
}}
</style>
</head>
<body>
<div class="toolbar">
    <h1 id="titleText">{safe_title}</h1>
    <span class="spacer"></span>
    <button id="btnPrev" title="Previous">Prev</button>
    <span class="status" id="pageInfo">Page {current_idx + 1} / {total}</span>
    <button id="btnNext" title="Next">Next</button>
    <span class="sep"></span>
    <button id="btnZoomOut" title="Zoom Out">-</button>
    <span class="zoom-display" id="zoomLevel">100%</span>
    <button id="btnZoomIn" title="Zoom In">+</button>
    <button id="btnFit" title="Fit to Window">Fit</button>
    <span class="sep"></span>
    <button class="primary" id="btnDownload" title="Download SVG">&#8681; svg</button>
    <button class="primary" id="btnDownloadVisio" title="Download SVG optimized for Visio">&#8681; svg(for visio)</button>
    <button class="primary" id="btnDownloadDrawio" title="Download as draw.io diagram">&#8681; draw.io</button>
    <button class="primary" id="btnDownloadDrawioStencil" title="Download as draw.io diagram with Cisco stencils">&#8681; draw.io(stencil)</button>
</div>
<div class="viewer" id="viewer">
    <div id="svgContainer"></div>
</div>
<ul class="ctx-menu" id="ctxMenu">
    <li id="ctxCopy">テキストをコピー</li>
</ul>
<div class="ctx-feedback" id="ctxFeedback">コピーしました</div>
<script>
(function() {{
    var files = [{svg_files_js}];
    var names = [{svg_names_js}];
    var idx = {current_idx};
    var jobId = "{job_id}";

    var viewer = document.getElementById('viewer');
    var zoomEl = document.getElementById('zoomLevel');
    var pageEl = document.getElementById('pageInfo');
    var titleEl = document.getElementById('titleText');

    var scale = 1, cx = 0, cy = 0;
    var dragging = false, dragStartX = 0, dragStartY = 0, cxStart = 0, cyStart = 0;
    var naturalW = 800, naturalH = 600;

    function updateTransform() {{
        var vw = viewer.clientWidth, vh = viewer.clientHeight;
        var iw = naturalW * scale;
        var ih = naturalH * scale;
        var left = vw / 2 - cx * scale;
        var top = vh / 2 - cy * scale;
        var container = document.getElementById('svgContainer');
        container.style.left = left + 'px';
        container.style.top = top + 'px';
        container.style.width = iw + 'px';
        container.style.height = ih + 'px';
        var svgEl = container.querySelector('svg');
        if (svgEl) {{
            svgEl.style.width = iw + 'px';
            svgEl.style.height = ih + 'px';
        }}
        zoomEl.textContent = Math.round(scale * 100) + '%';
    }}

    function fitToWindow() {{
        var vw = viewer.clientWidth, vh = viewer.clientHeight;
        scale = Math.min(vw / naturalW, vh / naturalH, 1) * 0.95;
        cx = naturalW / 2;
        cy = naturalH / 2;
        updateTransform();
    }}

    var _svgRetries = 0;
    function loadSvgInline(url) {{
        _svgRetries = 0;
        _fetchSvg(url);
    }}

    function _fetchSvg(url) {{
        fetch(url)
            .then(function(r) {{
                if (!r.ok) throw new Error('HTTP ' + r.status);
                return r.text();
            }})
            .then(function(svgText) {{
                // スコープ付き: SVG内の<style>がページ全体に漏れないようセレクタを限定する
                svgText = svgText.replace(/<style>([\s\S]*?)<\/style>/gi, function(_, css) {{
                    css = css.replace(/((?:^|[}}])\s*)((?:[^{{@/]|\/(?!\*))+)(\s*\{{)/g, function(m, pre, sel, post) {{
                        var scoped = sel.split(',').map(function(s) {{
                            return '#svgContainer ' + s.trim();
                        }}).join(', ');
                        return pre + scoped + post;
                    }});
                    return '<style>' + css + '</style>';
                }});
                var container = document.getElementById('svgContainer');
                container.innerHTML = svgText;
                var svgEl = container.querySelector('svg');
                if (!svgEl) return;
                // viewBox または width/height からサイズ取得
                var vb = svgEl.viewBox && svgEl.viewBox.baseVal;
                if (vb && vb.width && vb.height) {{
                    naturalW = vb.width;
                    naturalH = vb.height;
                }} else {{
                    naturalW = parseFloat(svgEl.getAttribute('width')) || 800;
                    naturalH = parseFloat(svgEl.getAttribute('height')) || 600;
                }}
                fitToWindow();
            }})
            .catch(function() {{
                if (_svgRetries < 3) {{
                    _svgRetries++;
                    setTimeout(function() {{ _fetchSvg(url + '?retry=' + _svgRetries); }}, 1500 * _svgRetries);
                }}
            }});
    }}

    window.addEventListener('resize', function() {{ updateTransform(); }});

    function zoomAtPoint(px, py, factor) {{
        var newScale = Math.max(0.02, Math.min(80, scale * factor));
        var vw = viewer.clientWidth, vh = viewer.clientHeight;
        var dx = px - vw / 2, dy = py - vh / 2;
        cx += dx * (1 / scale - 1 / newScale);
        cy += dy * (1 / scale - 1 / newScale);
        scale = newScale;
        updateTransform();
    }}

    function zoomCenter(factor) {{
        var newScale = Math.max(0.02, Math.min(80, scale * factor));
        scale = newScale;
        updateTransform();
    }}

    viewer.addEventListener('wheel', function(e) {{
        e.preventDefault();
        var rect = viewer.getBoundingClientRect();
        var px = e.clientX - rect.left, py = e.clientY - rect.top;
        var factor = e.deltaY < 0 ? 1.15 : 1 / 1.15;
        zoomAtPoint(px, py, factor);
    }}, {{ passive: false }});

    viewer.addEventListener('mousedown', function(e) {{
        if (e.button !== 0) return;
        dragging = true;
        dragStartX = e.clientX; dragStartY = e.clientY;
        cxStart = cx; cyStart = cy;
        viewer.classList.add('dragging');
        e.preventDefault();
    }});
    window.addEventListener('mousemove', function(e) {{
        if (!dragging) return;
        cx = cxStart - (e.clientX - dragStartX) / scale;
        cy = cyStart - (e.clientY - dragStartY) / scale;
        updateTransform();
    }});
    window.addEventListener('mouseup', function() {{
        dragging = false;
        viewer.classList.remove('dragging');
    }});

    function loadPage(i) {{
        if (i < 0 || i >= files.length) return;
        idx = i;
        loadSvgInline('/svg_raw/' + jobId + '/' + files[idx]);
        titleEl.textContent = names[idx];
        pageEl.textContent = 'Page ' + (idx + 1) + ' / ' + files.length;
        document.title = names[idx];
        document.getElementById('btnPrev').disabled = (idx === 0);
        document.getElementById('btnNext').disabled = (idx === files.length - 1);
    }}

    document.getElementById('btnPrev').onclick = function() {{ loadPage(idx - 1); }};
    document.getElementById('btnNext').onclick = function() {{ loadPage(idx + 1); }};
    document.getElementById('btnZoomIn').onclick = function() {{ zoomCenter(1.3); }};
    document.getElementById('btnZoomOut').onclick = function() {{ zoomCenter(1 / 1.3); }};
    document.getElementById('btnFit').onclick = fitToWindow;
    document.getElementById('btnDownload').onclick = function() {{
        fetch('/svg_raw/' + jobId + '/' + files[idx])
            .then(function(r) {{ return r.blob(); }})
            .then(function(blob) {{
                var url = URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = names[idx];
                a.click();
                URL.revokeObjectURL(url);
            }});
    }};

    document.getElementById('btnDownloadVisio').onclick = function() {{
        fetch('/download_visio/' + jobId + '/' + files[idx])
            .then(function(r) {{ return r.blob(); }})
            .then(function(blob) {{
                var url = URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = names[idx].replace(/\.svg$/i, '_visio.svg');
                a.click();
                URL.revokeObjectURL(url);
            }});
    }};

    document.getElementById('btnDownloadDrawio').onclick = function() {{
        fetch('/download_drawio/' + jobId + '/' + files[idx])
            .then(function(r) {{ return r.blob(); }})
            .then(function(blob) {{
                var url = URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = names[idx].replace(/\.svg$/i, '.drawio');
                a.click();
                URL.revokeObjectURL(url);
            }});
    }};

    document.getElementById('btnDownloadDrawioStencil').onclick = function() {{
        fetch('/download_drawio_stencil/' + jobId + '/' + files[idx])
            .then(function(r) {{
                if (!r.ok) {{
                    return r.text().then(function(t) {{ throw new Error(t); }});
                }}
                return r.blob();
            }})
            .then(function(blob) {{
                var url = URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = names[idx].replace(/\.svg$/i, '_stencil.drawio');
                a.click();
                URL.revokeObjectURL(url);
            }})
            .catch(function(err) {{
                alert('Stencil download failed: ' + err.message);
            }});
    }};

    document.getElementById('btnPrev').disabled = (idx === 0);
    document.getElementById('btnNext').disabled = (idx === files.length - 1);

    // 初回ロード
    loadSvgInline('/svg_raw/' + jobId + '/' + files[idx]);

    // --- 右クリック コンテキストメニュー ---
    var ctxMenu = document.getElementById('ctxMenu');
    var ctxFeedback = document.getElementById('ctxFeedback');
    var _feedbackTimer = null;

    function showCtxMenu(x, y, label) {{
        ctxMenu.style.left = x + 'px';
        ctxMenu.style.top = y + 'px';
        ctxMenu.style.display = 'block';
        document.getElementById('ctxCopy').onclick = function() {{
            navigator.clipboard.writeText(label).then(function() {{
                hideCtxMenu();
                showFeedback();
            }}).catch(function() {{
                // HTTPS以外の環境向けフォールバック
                try {{
                    var ta = document.createElement('textarea');
                    ta.value = label;
                    ta.style.position = 'fixed';
                    ta.style.opacity = '0';
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand('copy');
                    document.body.removeChild(ta);
                    hideCtxMenu();
                    showFeedback();
                }} catch(e) {{}}
            }});
        }};
    }}

    function hideCtxMenu() {{
        ctxMenu.style.display = 'none';
    }}

    function showFeedback() {{
        ctxFeedback.style.opacity = '1';
        clearTimeout(_feedbackTimer);
        _feedbackTimer = setTimeout(function() {{
            ctxFeedback.style.opacity = '0';
        }}, 1500);
    }}

    viewer.addEventListener('contextmenu', function(e) {{
        // クリック対象が <text> または <tspan> 要素か（祖先もたどる）
        var el = e.target;
        while (el && el !== viewer) {{
            var tag = el.tagName ? el.tagName.toLowerCase() : '';
            if (tag === 'text' || tag === 'tspan') break;
            el = el.parentElement;
        }}
        if (!el || el === viewer) return;
        var tag = el.tagName ? el.tagName.toLowerCase() : '';
        if (tag !== 'text' && tag !== 'tspan') return;

        // <tspan> の場合は親 <text> 全体のテキストを取得
        var textEl = (tag === 'tspan') ? el.closest('text') : el;
        var label = (textEl || el).textContent.trim();
        if (!label) return;

        e.preventDefault();
        // メニューが画面外に出ないよう位置を調整
        var menuW = 180, menuH = 42;
        var x = Math.min(e.clientX, window.innerWidth - menuW - 8);
        var y = Math.min(e.clientY, window.innerHeight - menuH - 8);
        showCtxMenu(x, y, label);
    }});

    document.addEventListener('click', function(e) {{
        if (!ctxMenu.contains(e.target)) hideCtxMenu();
    }});
    document.addEventListener('keydown', function(e) {{
        if (e.key === 'Escape') hideCtxMenu();
    }});

    window.addEventListener('keydown', function(e) {{
        if (e.key === 'ArrowLeft') loadPage(idx - 1);
        else if (e.key === 'ArrowRight') loadPage(idx + 1);
        else if (e.key === '+' || e.key === '=') document.getElementById('btnZoomIn').click();
        else if (e.key === '-') document.getElementById('btnZoomOut').click();
        else if (e.key === 'f' || e.key === 'F') fitToWindow();
    }});
}})();
</script>
</body>
</html>'''


def _render_txt_preview(safe_title, download_url):
    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Preview - {safe_title}</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        background: #f0f2f5; color: #333; height: 100vh; display: flex; flex-direction: column; }}
.toolbar {{ display: flex; align-items: center; gap: 12px; padding: 10px 20px;
            background: #16213e; color: #fff; flex-shrink: 0; }}
.toolbar h1 {{ font-size: 14px; font-weight: 500; opacity: 0.9; white-space: nowrap;
               overflow: hidden; text-overflow: ellipsis; max-width: 50vw; }}
.toolbar .spacer {{ flex: 1; }}
.toolbar a {{ padding: 6px 16px; border: 1px solid #e94560; background: transparent;
              color: #e94560; border-radius: 6px; font-size: 13px;
              text-decoration: none; transition: all 0.2s; }}
.toolbar a:hover {{ background: #e94560; color: #fff; }}
.text-wrap {{ flex: 1; overflow: auto; padding: 20px; }}
pre {{ background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.1);
       font-family: 'Consolas', 'Monaco', 'Courier New', monospace; font-size: 13px;
       line-height: 1.6; white-space: pre-wrap; word-wrap: break-word; }}
.loading {{ display: flex; align-items: center; justify-content: center; height: 100%;
            font-size: 16px; color: #666; }}
.loading .spinner {{ width: 24px; height: 24px; border: 3px solid #ddd;
                     border-top-color: #4A8FE7; border-radius: 50%;
                     animation: spin 0.8s linear infinite; margin-right: 12px; }}
@keyframes spin {{ to {{ transform: rotate(360deg); }} }}
</style>
</head>
<body>
<div class="toolbar">
    <h1>{safe_title}</h1>
    <span class="spacer"></span>
    <a href="{download_url}" rel="noopener noreferrer">Download</a>
</div>
<div id="content">
    <div class="loading" id="loadingMsg">
        <div class="spinner"></div>Loading file...
    </div>
</div>
<script>
(function() {{
    fetch('{download_url}')
        .then(function(r) {{ if (!r.ok) throw new Error('HTTP ' + r.status); return r.text(); }})
        .then(function(text) {{
            var content = document.getElementById('content');
            content.innerHTML = '<div class="text-wrap"><pre></pre></div>';
            content.querySelector('pre').textContent = text;
        }})
        .catch(function(e) {{
            document.getElementById('content').innerHTML =
                '<div style="padding:40px;text-align:center;color:#e94560">Failed to load: ' + e.message + '</div>';
        }});
}})();
</script>
<div style="text-align:center;padding:8px;font-size:11px;color:#b0b8c4;">Network Sketcher Online</div>
</body>
</html>'''


def _render_xlsx_preview(safe_title, download_url):
    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Preview - {safe_title}</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        background: #f0f2f5; color: #333; height: 100vh; display: flex; flex-direction: column; }}
.toolbar {{ display: flex; align-items: center; gap: 12px; padding: 10px 20px;
            background: #16213e; color: #fff; flex-shrink: 0; }}
.toolbar h1 {{ font-size: 14px; font-weight: 500; opacity: 0.9; white-space: nowrap;
               overflow: hidden; text-overflow: ellipsis; max-width: 50vw; }}
.toolbar .spacer {{ flex: 1; }}
.toolbar a {{ padding: 6px 16px; border: 1px solid #e94560; background: transparent;
              color: #e94560; border-radius: 6px; font-size: 13px;
              text-decoration: none; transition: all 0.2s; }}
.toolbar a:hover {{ background: #e94560; color: #fff; }}
.sheet-tabs {{ display: flex; gap: 0; background: #dee2e6; border-bottom: 2px solid #4A8FE7;
               padding: 0 16px; flex-shrink: 0; overflow-x: auto; }}
.sheet-tab {{ padding: 8px 20px; font-size: 13px; cursor: pointer; border: none;
              background: #dee2e6; color: #555; border-radius: 6px 6px 0 0;
              transition: all 0.2s; white-space: nowrap; }}
.sheet-tab:hover {{ background: #e9ecef; }}
.sheet-tab.active {{ background: #fff; color: #4A8FE7; font-weight: 600;
                     border-top: 2px solid #4A8FE7; }}
#content {{ flex: 1; display: flex; flex-direction: column; overflow: hidden; min-height: 0; }}
.table-wrap {{ flex: 1; overflow: scroll; padding: 16px; min-height: 0; }}
.table-wrap::-webkit-scrollbar {{ width: 12px; height: 12px; }}
.table-wrap::-webkit-scrollbar-track {{ background: #e9ecef; border-radius: 6px; }}
.table-wrap::-webkit-scrollbar-thumb {{ background: #a0aec0; border-radius: 6px; border: 2px solid #e9ecef; }}
.table-wrap::-webkit-scrollbar-thumb:hover {{ background: #718096; }}
.table-wrap::-webkit-scrollbar-corner {{ background: #e9ecef; }}
.table-wrap table {{ border-collapse: collapse; width: auto; min-width: 100%;
                     background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.1);
                     border-radius: 4px; overflow: hidden; font-size: 13px; }}
.table-wrap th, .table-wrap td {{ border: 1px solid #e0e0e0; padding: 6px 12px;
                                  text-align: left; white-space: nowrap; }}
.table-wrap tr.header-row th, .table-wrap tr.header-row td {{
    background: #4A8FE7; color: #fff; font-weight: 600; position: sticky; top: 0; z-index: 1; }}
.table-wrap tr.label-row {{ display: none; }}
.table-wrap tr:nth-child(even):not(.label-row):not(.header-row) td:not([style*="background"]) {{ background: #f8f9fa; }}
.table-wrap tr:hover:not(.label-row):not(.header-row) td:not([style*="background"]) {{ background: #e8f0fe; }}
.loading {{ display: flex; align-items: center; justify-content: center; height: 100%;
            font-size: 16px; color: #666; }}
.loading .spinner {{ width: 24px; height: 24px; border: 3px solid #ddd;
                     border-top-color: #4A8FE7; border-radius: 50%;
                     animation: spin 0.8s linear infinite; margin-right: 12px; }}
@keyframes spin {{ to {{ transform: rotate(360deg); }} }}
.error {{ display: none; align-items: center; justify-content: center; height: 100%;
          font-size: 16px; color: #e94560; flex-direction: column; gap: 12px; }}
.error .detail {{ font-size: 13px; opacity: 0.7; max-width: 600px; word-break: break-all; }}
</style>
</head>
<body>
<div class="toolbar">
    <h1>{safe_title}</h1>
    <span class="spacer"></span>
    <a href="{download_url}" rel="noopener noreferrer">Download</a>
</div>
<div class="sheet-tabs" id="sheetTabs"></div>
<div id="content">
    <div class="loading" id="loadingMsg">
        <div class="spinner"></div>Loading spreadsheet...
    </div>
    <div class="error" id="errorMsg">
        <div>Failed to load spreadsheet</div>
        <div class="detail" id="errorDetail"></div>
        <button onclick="location.reload()">Retry</button>
    </div>
</div>
<script src="/static/xlsx.full.min.js"></script>
<script>
(function() {{
    function showError(msg) {{
        document.getElementById('loadingMsg').style.display = 'none';
        document.getElementById('errorDetail').textContent = msg || '';
        document.getElementById('errorMsg').style.display = 'flex';
    }}
    if (typeof XLSX === 'undefined') {{
        showError('SheetJS library failed to load.');
        return;
    }}
    var sheetTabs = document.getElementById('sheetTabs');
    var contentDiv = document.getElementById('content');
    var workbook = null;

    function extractRgb(colorObj) {{
        if (!colorObj) return null;
        var rgb = colorObj.rgb;
        if (!rgb) return null;
        if (rgb.length === 8) rgb = rgb.substring(2);
        if (rgb === '000000' || rgb === 'FFFFFF') return null;
        return '#' + rgb;
    }}

    function applyCellStyles(ws, sheetName) {{
        if (sheetName !== 'Attribute') return;
        var ref = ws['!ref'];
        if (!ref) return;
        var range = XLSX.utils.decode_range(ref);
        var headerRow = 1;
        var defaultCol = -1;
        for (var hc = range.s.c; hc <= range.e.c; hc++) {{
            var hAddr = XLSX.utils.encode_cell({{ r: headerRow, c: hc }});
            var hCell = ws[hAddr];
            if (hCell && hCell.v && String(hCell.v).trim() === 'Default') {{
                defaultCol = hc;
                break;
            }}
        }}
        if (defaultCol < 0) return;
        var rows = contentDiv.querySelectorAll('.table-wrap table tr');
        for (var r = headerRow + 1; r <= range.e.r; r++) {{
            var tr = rows[r];
            if (!tr) continue;
            var cells = tr.querySelectorAll('td, th');
            for (var c = defaultCol; c <= range.e.c; c++) {{
                var addr = XLSX.utils.encode_cell({{ r: r, c: c }});
                var cell = ws[addr];
                var td = cells[c - range.s.c];
                if (!cell || !cell.s || !td) continue;
                var fill = cell.s.fgColor || (cell.s.fill && cell.s.fill.fgColor);
                var bg = extractRgb(fill);
                if (bg) td.style.backgroundColor = bg;
                var font = cell.s.color || (cell.s.font && cell.s.font.color);
                var fg = extractRgb(font);
                if (fg) td.style.color = fg;
            }}
        }}
    }}

    function showSheet(name) {{
        var tabs = sheetTabs.querySelectorAll('.sheet-tab');
        for (var i = 0; i < tabs.length; i++) {{
            tabs[i].classList.toggle('active', tabs[i].textContent === name);
        }}
        var ws = workbook.Sheets[name];
        if (!ws) return;
        var html = XLSX.utils.sheet_to_html(ws, {{ editable: false }});
        contentDiv.innerHTML = '<div class="table-wrap">' + html + '</div>';
        var rows = contentDiv.querySelectorAll('.table-wrap table tr');
        if (rows.length >= 2) {{
            rows[0].classList.add('label-row');
            rows[1].classList.add('header-row');
        }}
        applyCellStyles(ws, name);
    }}

    fetch('{download_url}')
        .then(function(r) {{ if (!r.ok) throw new Error('HTTP ' + r.status); return r.arrayBuffer(); }})
        .then(function(buf) {{
            workbook = XLSX.read(buf, {{ type: 'array', cellStyles: true }});
            document.getElementById('loadingMsg').style.display = 'none';
            sheetTabs.innerHTML = '';
            for (var i = 0; i < workbook.SheetNames.length; i++) {{
                var tab = document.createElement('button');
                tab.className = 'sheet-tab';
                tab.textContent = workbook.SheetNames[i];
                tab.addEventListener('click', function() {{ showSheet(this.textContent); }});
                sheetTabs.appendChild(tab);
            }}
            if (workbook.SheetNames.length > 0) {{
                showSheet(workbook.SheetNames[0]);
            }}
        }})
        .catch(function(e) {{ showError('Error: ' + e.message); }});
}})();
</script>
<div style="text-align:center;padding:8px;font-size:11px;color:#b0b8c4;">Network Sketcher Online</div>
</body>
</html>'''


@app.route('/preview/<job_id>/<path:filename>')
def preview_file(job_id, filename):
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    filepath = work_dir / filename

    try:
        filepath.resolve().relative_to(work_dir.resolve())
    except ValueError:
        abort(403)

    lower_name = filename.lower()
    if not filepath.is_file() or not (lower_name.endswith('.pptx') or lower_name.endswith('.xlsx') or lower_name.endswith('.txt') or lower_name.endswith('.svg')):
        abort(404)

    if lower_name.endswith('.svg'):
        filter_prefix = request.args.get('filter', '')
        return _render_svg_viewer(job_id, filename, work_dir, filter_prefix)

    encoded_filename = url_quote(filename, safe='')
    download_url = f'/download/{job_id}/{encoded_filename}'
    safe_title = filename.replace("'", "\\'").replace('"', '&quot;')

    if lower_name.endswith('.txt'):
        return _render_txt_preview(safe_title, download_url)

    if lower_name.endswith('.xlsx'):
        return _render_xlsx_preview(safe_title, download_url)

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Preview - {safe_title}</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        background: #1a1a2e; color: #fff; height: 100vh; display: flex; flex-direction: column; }}
.toolbar {{ display: flex; align-items: center; gap: 12px; padding: 10px 20px;
            background: #16213e; border-bottom: 1px solid #0f3460; flex-shrink: 0; }}
.toolbar h1 {{ font-size: 14px; font-weight: 500; opacity: 0.9; white-space: nowrap;
               overflow: hidden; text-overflow: ellipsis; max-width: 50vw; }}
.toolbar .spacer {{ flex: 1; }}
.toolbar button {{ padding: 6px 16px; border: 1px solid #4A8FE7; background: transparent;
                   color: #4A8FE7; border-radius: 6px; cursor: pointer; font-size: 13px;
                   transition: all 0.2s; white-space: nowrap; }}
.toolbar button:hover {{ background: #4A8FE7; color: #fff; }}
.toolbar button:disabled {{ opacity: 0.4; cursor: default; }}
.toolbar a {{ padding: 6px 16px; border: 1px solid #e94560; background: transparent;
              color: #e94560; border-radius: 6px; cursor: pointer; font-size: 13px;
              text-decoration: none; transition: all 0.2s; }}
.toolbar a:hover {{ background: #e94560; color: #fff; }}
.toolbar .status {{ font-size: 13px; opacity: 0.8; }}
.toolbar .zoom-info {{ font-size: 12px; opacity: 0.6; min-width: 45px; text-align: center; }}
.toolbar .sep {{ width: 1px; height: 20px; background: rgba(255,255,255,0.15); }}
#viewer {{ flex: 1; width: 100%; display: flex; }}
.canvas-wrap {{ width: 100%; height: 100%; overflow: hidden; position: relative; cursor: grab; }}
.canvas-wrap.dragging {{ cursor: grabbing; }}
canvas {{ background: white; box-shadow: 0 4px 20px rgba(0,0,0,0.5);
          position: absolute; transform-origin: 0 0; }}
.loading {{ display: flex; align-items: center; justify-content: center; height: 100%;
            font-size: 16px; opacity: 0.7; }}
.loading .spinner {{ width: 24px; height: 24px; border: 3px solid rgba(255,255,255,0.2);
                     border-top-color: #4A8FE7; border-radius: 50%;
                     animation: spin 0.8s linear infinite; margin-right: 12px; }}
@keyframes spin {{ to {{ transform: rotate(360deg); }} }}
.error {{ display: none; align-items: center; justify-content: center; height: 100%;
          font-size: 16px; color: #e94560; flex-direction: column; gap: 12px; }}
.error .detail {{ font-size: 13px; opacity: 0.7; max-width: 600px; word-break: break-all; }}
</style>
</head>
<body>
<div class="toolbar">
    <h1>{safe_title}</h1>
    <span class="spacer"></span>
    <button id="btnPrev" disabled>Prev</button>
    <span class="status" id="slideStatus"></span>
    <button id="btnNext" disabled>Next</button>
    <span class="sep"></span>
    <button id="btnZoomOut" disabled title="Zoom Out">-</button>
    <span class="zoom-info" id="zoomInfo">100%</span>
    <button id="btnZoomIn" disabled title="Zoom In">+</button>
    <button id="btnFit" disabled title="Fit to Window">Fit</button>
    <span class="sep"></span>
    <a href="{download_url}" rel="noopener noreferrer">Download</a>
</div>
<div id="viewer">
    <div class="loading" id="loadingMsg">
        <div class="spinner"></div>Loading presentation...
    </div>
    <div class="error" id="errorMsg">
        <div>Failed to load presentation</div>
        <div class="detail" id="errorDetail"></div>
        <button onclick="location.reload()">Retry</button>
    </div>
</div>
<script src="/static/chart.umd.min.js"></script>
<script src="/static/jszip.min.js"></script>
<script src="/static/PptxViewJS.min.js"></script>
<script>
(function() {{
    function showError(msg) {{
        document.getElementById('loadingMsg').style.display = 'none';
        document.getElementById('errorDetail').textContent = msg || '';
        document.getElementById('errorMsg').style.display = 'flex';
    }}
    if (typeof PptxViewJS === 'undefined') {{
        showError('PptxViewJS library failed to load.');
        return;
    }}
    if (!PptxViewJS.PPTXViewer) {{
        showError('PPTXViewer not found. Available: ' + Object.keys(PptxViewJS).join(', '));
        return;
    }}
    try {{
        var viewerDiv = document.getElementById('viewer');
        var canvas = document.createElement('canvas');
        var dpr = window.devicePixelRatio || 1;
        var renderScale = 2;
        var baseW = 1920;
        var baseH = 1080;
        var displayW = baseW * renderScale;
        var displayH = baseH * renderScale;
        canvas.width = Math.round(displayW * dpr);
        canvas.height = Math.round(displayH * dpr);
        canvas.style.width = displayW + 'px';
        canvas.style.height = displayH + 'px';
        var wrap = document.createElement('div');
        wrap.className = 'canvas-wrap';
        wrap.appendChild(canvas);

        var viewer = new PptxViewJS.PPTXViewer({{ canvas: canvas }});
        var btnPrev = document.getElementById('btnPrev');
        var btnNext = document.getElementById('btnNext');
        var slideStatus = document.getElementById('slideStatus');
        var btnZoomIn = document.getElementById('btnZoomIn');
        var btnZoomOut = document.getElementById('btnZoomOut');
        var btnFit = document.getElementById('btnFit');
        var zoomInfo = document.getElementById('zoomInfo');

        var scale = 1;
        var panX = 0;
        var panY = 0;
        var isDragging = false;
        var dragStartX = 0;
        var dragStartY = 0;
        var panStartX = 0;
        var panStartY = 0;
        var MIN_SCALE = 0.2;
        var MAX_SCALE = 5;

        function applyTransform() {{
            canvas.style.transform = 'translate(' + panX + 'px,' + panY + 'px) scale(' + scale + ')';
            zoomInfo.textContent = Math.round(scale * 100) + '%';
        }}

        function fitToWindow() {{
            var ww = wrap.clientWidth;
            var wh = wrap.clientHeight;
            var cw = parseFloat(canvas.style.width) || canvas.width;
            var ch = parseFloat(canvas.style.height) || canvas.height;
            if (cw <= 0 || ch <= 0) return;
            scale = Math.min(ww / cw, wh / ch) * 0.95;
            if (scale < MIN_SCALE) scale = MIN_SCALE;
            panX = (ww - cw * scale) / 2;
            panY = (wh - ch * scale) / 2;
            applyTransform();
        }}

        function zoomAt(delta, cx, cy) {{
            var oldScale = scale;
            scale *= delta > 0 ? 1.15 : 1 / 1.15;
            if (scale < MIN_SCALE) scale = MIN_SCALE;
            if (scale > MAX_SCALE) scale = MAX_SCALE;
            var ratio = scale / oldScale;
            panX = cx - (cx - panX) * ratio;
            panY = cy - (cy - panY) * ratio;
            applyTransform();
        }}

        wrap.addEventListener('wheel', function(e) {{
            e.preventDefault();
            var rect = wrap.getBoundingClientRect();
            zoomAt(-e.deltaY, e.clientX - rect.left, e.clientY - rect.top);
        }}, {{ passive: false }});

        wrap.addEventListener('mousedown', function(e) {{
            if (e.button !== 0) return;
            isDragging = true;
            dragStartX = e.clientX;
            dragStartY = e.clientY;
            panStartX = panX;
            panStartY = panY;
            wrap.classList.add('dragging');
        }});
        window.addEventListener('mousemove', function(e) {{
            if (!isDragging) return;
            panX = panStartX + (e.clientX - dragStartX);
            panY = panStartY + (e.clientY - dragStartY);
            applyTransform();
        }});
        window.addEventListener('mouseup', function() {{
            isDragging = false;
            wrap.classList.remove('dragging');
        }});

        btnZoomIn.addEventListener('click', function() {{
            var ww = wrap.clientWidth;
            var wh = wrap.clientHeight;
            zoomAt(1, ww / 2, wh / 2);
        }});
        btnZoomOut.addEventListener('click', function() {{
            var ww = wrap.clientWidth;
            var wh = wrap.clientHeight;
            zoomAt(-1, ww / 2, wh / 2);
        }});
        btnFit.addEventListener('click', fitToWindow);

        function updateNav() {{
            var total = viewer.getSlideCount();
            var idx = typeof viewer.getCurrentSlideIndex === 'function' ? viewer.getCurrentSlideIndex() : 0;
            slideStatus.textContent = total ? 'Slide ' + (idx + 1) + ' / ' + total : '';
            btnPrev.disabled = !(idx > 0);
            btnNext.disabled = !(idx < total - 1);
        }}

        viewer.on('renderComplete', function() {{
            updateNav();
            fitToWindow();
        }});
        viewer.on('loadComplete', function() {{
            document.getElementById('loadingMsg').style.display = 'none';
            viewerDiv.innerHTML = '';
            viewerDiv.appendChild(wrap);
            btnZoomIn.disabled = false;
            btnZoomOut.disabled = false;
            btnFit.disabled = false;
            viewer.render(canvas, {{ slideIndex: 0 }});
        }});
        viewer.on('loadError', function(e) {{
            showError('Load error: ' + (e && e.message ? e.message : String(e)));
        }});

        btnPrev.addEventListener('click', function() {{ viewer.previousSlide(canvas); }});
        btnNext.addEventListener('click', function() {{ viewer.nextSlide(canvas); }});

        window.addEventListener('resize', function() {{ fitToWindow(); }});

        viewer.loadFromUrl('{download_url}');
    }} catch(e) {{
        showError('Viewer init error: ' + e.message);
    }}
}})();
</script>
<div style="text-align:center;padding:8px;font-size:11px;color:#b0b8c4;">Network Sketcher Online</div>
</body>
</html>'''


@app.route('/download_all/<job_id>')
def download_all(job_id):
    """Build and return a ZIP of generated artifacts for the job.

    The ZIP always carries the per-layer L1/L2/L3 SVG files alongside
    the combined ``[L1L2L3_DIAGRAM]*.html`` viewer(s). Three modes are
    selected by query parameters:

      - **Lightweight (``?scope=lightweight``)**: post-upload state.
        Bundles the three all-areas SVGs
        (``[L1_DIAGRAM]AllAreasTag``, ``[L2_DIAGRAM]AllAreas``,
        ``[L3_DIAGRAM]AllAreas``), the AllAreas combined HTML viewer,
        the Device Tables HTML, AI Context txt, and the .nsm master.

      - **Area-scoped (v3.1.2a, SVG mode dropdown, ``?area=<name>``)**:
        Bundles the three all-areas SVGs *plus* the area's three
        per-area SVGs, the AllAreas + per-area combined HTML viewers,
        Device Table HTML, AI Context txt, and the .nsm master.
        Missing SVGs are (re)generated on demand. Use ``?attr=<value>``
        to align the SVG attribute with what the user picked in the
        thumbnail grid.

      - **Whole-job (legacy, no scope/area)**: bundles every artifact
        in the job directory (PPTX mode, session restore, generic
        flows). The AllAreas combined HTML is best-effort regenerated
        when the per-layer SVG triple is on disk.
    """
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        abort(404)

    master_filename = get_active_master(str(work_dir))
    if not master_filename:
        abort(404)

    area_arg = (request.args.get('area') or '').strip()
    attr_arg = request.args.get('attr', '')
    scope_arg = (request.args.get('scope') or '').strip().lower()

    basename = master_filename.replace('[MASTER]', '')
    if basename.lower().endswith('.nsm'):
        basename = basename[: -len('.nsm')]
    elif basename.lower().endswith('.xlsx'):
        basename = basename[: -len('.xlsx')]

    # ---- Helpers (shared by both ZIP paths) -------------------------------
    def _ensure_via_cli(cmd_args_extra, expected_filename):
        """Run the engine to (re)generate ``expected_filename`` if missing.

        Returns the absolute path on disk if the file exists after the call,
        else None. Failures are logged but do not abort the ZIP (so a partial
        ZIP is still useful).
        """
        target = work_dir / expected_filename
        if target.is_file():
            return target
        cli_args = list(cmd_args_extra) + [
            '--master', str(work_dir / master_filename),
            '--format', 'svg',
        ]
        if attr_arg and any(x in ('l1_diagram', 'l3_diagram') for x in cli_args):
            cli_args += ['--attribute', attr_arg]
        try:
            run_ns_command_subprocess(cli_args, work_dir, master_filename)
        except Exception as exc:
            logger.warning('download_all: regenerate %s failed: %s',
                           expected_filename, exc)
        return target if target.is_file() else None

    def _build_combined_html_for_zip(scope, scope_label):
        """Render ``[L1L2L3_DIAGRAM]<scope>_<basename>.html`` from the L1/L2/L3
        SVGs already on disk for ``scope`` and return the path on success.

        ``scope`` follows the same vocabulary as ``_resolve_layer_files``:
        ``'all'`` for All Areas, ``'area:<name>'`` for a per-area scope.
        ``scope_label`` is the human-readable label embedded in the page
        title (e.g. ``'All Areas'`` or the area name).

        Returns ``None`` (with a warning) if no per-layer SVG is available
        or if rendering fails -- the ZIP is still built without the combined
        HTML so the user receives a useful download.
        """
        if scope == 'all':
            out_name = f'[L1L2L3_DIAGRAM]AllAreas_{basename}.html'
        elif scope.startswith('area:'):
            area_name = scope[len('area:'):]
            safe = _safe_area_for_filename(area_name)
            out_name = f'[L1L2L3_DIAGRAM]{safe}_{basename}.html'
        else:
            return None

        layer_filenames = _resolve_layer_files(str(work_dir), scope)
        layer_svgs = {'l1': None, 'l2': None, 'l3': None}
        have_any = False
        for layer in ('l1', 'l2', 'l3'):
            fname = layer_filenames.get(layer)
            if fname and (work_dir / fname).is_file():
                try:
                    with open(str(work_dir / fname), 'r', encoding='utf-8') as fh:
                        layer_svgs[layer] = fh.read()
                    have_any = True
                except Exception as exc:
                    logger.warning('download_all: could not read %s: %s', fname, exc)
        if not have_any:
            return None

        try:
            try:
                from nsm_l1l2l3_html import render_l1l2l3_html
            except ImportError:
                from ns_engine.nsm_l1l2l3_html import render_l1l2l3_html
            html_text = render_l1l2l3_html(
                layer_svgs=layer_svgs,
                master_basename=basename,
                scope_label=scope_label,
                mode='standalone',
            )
            out_path = work_dir / out_name
            with open(str(out_path), 'w', encoding='utf-8') as fh:
                fh.write(html_text)
            return out_path
        except Exception as exc:
            logger.warning('download_all: combined html (%s) failed: %s', out_name, exc)
            return None

    # ---- Lightweight ZIP (v3.1.2a, SVG mode, area not yet selected) ------
    # Returned when the front-end shows the post-upload state. The ZIP
    # bundles the per-layer all-areas SVGs ([L1_DIAGRAM]AllAreasTag,
    # [L2_DIAGRAM]AllAreas, [L3_DIAGRAM]AllAreas), the combined L1/L2/L3
    # viewer ([L1L2L3_DIAGRAM]AllAreas_*.html), the Device Tables HTML,
    # AI Context txt, and the .nsm master.
    if not area_arg and scope_arg == 'lightweight':
        files_to_zip = []

        # All-areas SVGs (also inlined into the combined HTML below).
        for cli_args_extra, expected_filename in [
            (['export', 'l1_diagram', '--type', 'all_areas_tag'],
             f'[L1_DIAGRAM]AllAreasTag_{basename}.svg'),
            (['export', 'l2_diagram', '--type', 'all_areas'],
             f'[L2_DIAGRAM]AllAreas_{basename}.svg'),
            (['export', 'l3_diagram', '--type', 'all_areas'],
             f'[L3_DIAGRAM]AllAreas_{basename}.svg'),
        ]:
            svg_path = _ensure_via_cli(cli_args_extra, expected_filename)
            if svg_path:
                files_to_zip.append(svg_path)

        # Combined L1/L2/L3 standalone HTML (single artifact replacing the
        # three per-layer SVGs in the ZIP).
        combined_html = _build_combined_html_for_zip('all', 'All Areas')
        if combined_html:
            files_to_zip.append(combined_html)

        # Device Table HTML -- regenerate to guarantee freshness
        device_html_name = f'[DEVICE_TABLE]{basename}.html'
        device_html_path = work_dir / device_html_name
        try:
            try:
                from nsm_device_table_html import (
                    build_device_tabs_data, render_device_table_html,
                )
            except ImportError:
                from ns_engine.nsm_device_table_html import (
                    build_device_tabs_data, render_device_table_html,
                )
            tabs_data, master_basename = build_device_tabs_data(
                str(work_dir / master_filename))
            if tabs_data:
                html_text = render_device_table_html(
                    tabs_data, master_basename or basename)
                with open(str(device_html_path), 'w', encoding='utf-8') as df:
                    df.write(html_text)
                files_to_zip.append(device_html_path)
        except Exception as exc:
            logger.warning('download_all (lightweight): Device Table HTML failed: %s', exc)

        # AI Context (skip if not yet generated)
        ai_path = work_dir / f'[AI_Context]{basename}.txt'
        if ai_path.is_file():
            files_to_zip.append(ai_path)

        # Master .nsm itself (only when the active master is a .nsm)
        if master_filename.lower().endswith('.nsm'):
            files_to_zip.append(work_dir / master_filename)

        zip_name = f'{basename}.zip'
        zip_path = work_dir / zip_name
        with zipfile.ZipFile(str(zip_path), 'w', zipfile.ZIP_DEFLATED) as zf:
            for fpath in files_to_zip:
                if fpath and Path(fpath).is_file():
                    zf.write(str(fpath), Path(fpath).name)

        request._ns_download_filename = zip_name
        return send_file(
            str(zip_path),
            as_attachment=True,
            download_name=zip_name,
        )

    # ---- Legacy whole-job ZIP (no area filter, no scope) -----------------
    # Used by the PPTX/non-SVG mode and session-restore flows. The ZIP
    # bundles every artifact in the work_dir, including the per-layer
    # L1/L2/L3 SVGs alongside the [L1L2L3_DIAGRAM]*.html viewer(s). When
    # the all-areas SVG triple is present we (re)build the AllAreas
    # combined HTML so it is guaranteed present. Per-area combined HTMLs
    # that already exist on disk are bundled as-is.
    if not area_arg:
        # Best-effort generation of the AllAreas combined HTML; ignored if
        # the per-layer SVGs are not on disk.
        _build_combined_html_for_zip('all', 'All Areas')

        zip_name = 'network_sketcher_output.zip'
        zip_path = work_dir / zip_name
        with zipfile.ZipFile(str(zip_path), 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in sorted(os.listdir(work_dir)):
                if f == zip_name or '__TMP__' in f:
                    continue
                if f.startswith('.'):
                    continue
                if f.startswith('[MASTER]') and f.endswith('.xlsx'):
                    continue
                if f.startswith('[L2_TABLE]') or f.startswith('[L3_TABLE]'):
                    continue
                fpath = os.path.join(work_dir, f)
                if os.path.isfile(fpath):
                    zf.write(fpath, f)
        request._ns_download_filename = zip_name
        return send_file(
            str(zip_path),
            as_attachment=True,
            download_name=zip_name,
        )

    # ---- Area-scoped ZIP --------------------------------------------------
    safe_area = _safe_area_for_filename(area_arg)

    # ---- Build the file list ----------------------------------------------
    # The ZIP includes the per-layer L1/L2/L3 SVGs (3 all-areas + 3
    # per-area), the combined L1/L2/L3 viewer HTMLs (one for All Areas
    # plus one for the selected area), the Device Table HTML, AI Context
    # txt, and the .nsm master.
    files_to_zip = []

    # All-areas SVGs (also inlined into the AllAreas combined HTML).
    for cli_args_extra, expected_filename in [
        (['export', 'l1_diagram', '--type', 'all_areas_tag'],
         f'[L1_DIAGRAM]AllAreasTag_{basename}.svg'),
        (['export', 'l2_diagram', '--type', 'all_areas'],
         f'[L2_DIAGRAM]AllAreas_{basename}.svg'),
        (['export', 'l3_diagram', '--type', 'all_areas'],
         f'[L3_DIAGRAM]AllAreas_{basename}.svg'),
    ]:
        svg_path = _ensure_via_cli(cli_args_extra, expected_filename)
        if svg_path:
            files_to_zip.append(svg_path)

    # Per-area SVGs (also inlined into the per-area combined HTML).
    for cli_args_extra, expected_filename in [
        (['export', 'l1_diagram', '--type', 'per_area_tag', '--area', area_arg],
         f'[L1_DIAGRAM]PerAreaTag_{basename}_{safe_area}.svg'),
        (['export', 'l3_diagram', '--type', 'per_area', '--area', area_arg],
         f'[L3_DIAGRAM]PerArea_{basename}_{safe_area}.svg'),
        (['export', 'l2_diagram', '--area', area_arg],
         f'[L2_DIAGRAM]{area_arg}_{basename}.svg'),
    ]:
        svg_path = _ensure_via_cli(cli_args_extra, expected_filename)
        if svg_path:
            files_to_zip.append(svg_path)

    # Combined L1/L2/L3 viewer HTMLs (one for All Areas + one for the
    # selected area), bundled alongside the per-layer SVGs above.
    combined_all = _build_combined_html_for_zip('all', 'All Areas')
    if combined_all:
        files_to_zip.append(combined_all)
    combined_area = _build_combined_html_for_zip(f'area:{area_arg}', area_arg)
    if combined_area:
        files_to_zip.append(combined_area)

    # Device Table HTML -- always regenerate to guarantee freshness (cheap)
    device_html_name = f'[DEVICE_TABLE]{basename}.html'
    device_html_path = work_dir / device_html_name
    try:
        try:
            from nsm_device_table_html import (
                build_device_tabs_data, render_device_table_html,
            )
        except ImportError:
            from ns_engine.nsm_device_table_html import (
                build_device_tabs_data, render_device_table_html,
            )
        tabs_data, master_basename = build_device_tabs_data(
            str(work_dir / master_filename))
        if tabs_data:
            html_text = render_device_table_html(tabs_data, master_basename or basename)
            with open(str(device_html_path), 'w', encoding='utf-8') as df:
                df.write(html_text)
            files_to_zip.append(device_html_path)
    except Exception as exc:
        logger.warning('download_all: Device Table HTML generation failed: %s', exc)

    # AI Context (skip if not yet generated; do not block ZIP)
    ai_path = work_dir / f'[AI_Context]{basename}.txt'
    if ai_path.is_file():
        files_to_zip.append(ai_path)

    # Master .nsm itself (only when the active master is a .nsm)
    if master_filename.lower().endswith('.nsm'):
        files_to_zip.append(work_dir / master_filename)

    # ---- Assemble ZIP ------------------------------------------------------
    zip_name = f'{basename}_{safe_area}.zip'
    zip_path = work_dir / zip_name
    with zipfile.ZipFile(str(zip_path), 'w', zipfile.ZIP_DEFLATED) as zf:
        for fpath in files_to_zip:
            if fpath and Path(fpath).is_file():
                zf.write(str(fpath), Path(fpath).name)

    request._ns_download_filename = zip_name
    return send_file(
        str(zip_path),
        as_attachment=True,
        download_name=zip_name,
    )


# ---------- HTML Template ----------

HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Cache-Control" content="no-store, no-cache, must-revalidate, max-age=0">
<meta http-equiv="Pragma" content="no-cache">
<title>Network Sketcher Online</title>
<style>
:root {
    --primary: #4A8FE7;
    --primary-dark: #3570b8;
    --bg: #f0f4f8;
    --card: #ffffff;
    --text: #2c3e50;
    --text-secondary: #7f8c8d;
    --border: #e1e8ed;
    --success: #27ae60;
    --error: #e74c3c;
    --device-color: #8e44ad;
    --l1-color: #2980b9;
    --l2-color: #27ae60;
    --l3-color: #e67e22;
}
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Meiryo', Roboto, sans-serif;
    background: var(--bg);
    min-height: 100vh;
    color: var(--text);
}
.container { max-width: 860px; margin: 0 auto; padding: 32px 20px 60px; }
.header { text-align: center; margin-bottom: 32px; display: flex; align-items: center; justify-content: center; gap: 18px; }
.logo { width: 80px; height: auto; border-radius: 18px; flex-shrink: 0; }
.header-text { display: flex; flex-direction: column; align-items: center; }
h1 { font-size: 26px; color: var(--text); margin-bottom: 4px; font-weight: 800; font-style: italic; letter-spacing: 1px; }
.online-badge {
    color: #00c853;
    background: linear-gradient(135deg, #00c853, #00e676);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    font-weight: 800;
    font-style: italic;
    letter-spacing: 1px;
    position: relative;
}
.subtitle { color: var(--text-secondary); font-size: 14px; }
.card {
    background: var(--card); border-radius: 14px;
    box-shadow: 0 2px 16px rgba(0,0,0,0.06); padding: 28px; margin-bottom: 20px;
}
.card h2 {
    font-size: 16px; margin-bottom: 16px; color: var(--text);
    display: flex; align-items: center; gap: 8px;
}

/* Help icon and tooltip */
.help-icon {
    display: inline-flex; align-items: center; justify-content: center;
    width: 18px; height: 18px; border-radius: 50%;
    background: #b0b8c4; color: #fff; font-size: 12px; font-weight: 700;
    cursor: pointer; line-height: 1; flex-shrink: 0;
    transition: background 0.2s; font-style: normal;
    user-select: none;
}
.help-icon:hover { background: #8a95a5; }
.help-icon.active { background: var(--primary); }
.help-tooltip {
    background: #f0f7ff; border-left: 3px solid #4fc3f7;
    padding: 10px 14px; margin: -8px 0 12px; font-size: 13px;
    border-radius: 0 6px 6px 0; color: #4a5568; line-height: 1.6;
    display: none;
}
.help-tooltip.visible { display: block; }

/* Drop zone */
.dropzone {
    border: 2.5px dashed var(--border); border-radius: 12px;
    padding: 52px 24px; text-align: center; cursor: pointer; transition: all 0.25s ease;
}
.dropzone:hover, .dropzone.dragover { border-color: var(--primary); background: #f0f7ff; }
.dropzone.has-file { border-color: var(--success); background: #f0faf4; }
.dropzone.disabled { pointer-events: none; opacity: 0.5; }
.dropzone-icon { font-size: 42px; margin-bottom: 12px; opacity: 0.7; }
.dropzone-text { font-size: 15px; color: var(--text-secondary); line-height: 1.6; }
.dropzone-text strong { color: var(--primary); cursor: pointer; }
.dropzone-file { margin-top: 12px; font-size: 14px; color: var(--success); font-weight: 600; }
#fileInput { display: none; }

/* Analyzing indicator */
.analyzing {
    display: flex; align-items: center; justify-content: center;
    padding: 18px; gap: 10px; color: var(--text-secondary); font-size: 14px;
}
.spinner {
    width: 20px; height: 20px; border: 3px solid var(--border);
    border-top-color: var(--primary); border-radius: 50%;
    animation: spin 0.7s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }

/* Selection list */
.select-all-row {
    display: flex; align-items: center; padding: 0 0 14px;
    border-bottom: 1px solid var(--border); margin-bottom: 12px;
}
.select-all-row label {
    display: flex; align-items: center; gap: 8px;
    cursor: pointer; font-weight: 600; font-size: 14px;
}
/* Available Outputs header row with centered toggle */
.available-outputs-header {
    display: flex; align-items: center; position: relative; margin-bottom: 0;
}
.available-outputs-header h2 { flex: 1; }
/* Format toggle switch */
.format-switch {
    display: flex; align-items: center; gap: 8px;
    position: absolute; left: 50%; transform: translateX(-50%);
    background: #f0f4ff; border: 1px solid #c5d4f0;
    border-radius: 20px; padding: 4px 14px;
}
.format-switch .format-label {
    font-size: 12px; font-weight: 700; color: var(--text-secondary);
    white-space: nowrap; transition: color 0.2s;
}
.format-switch .format-label.active { color: var(--primary); }
.toggle-switch {
    position: relative; display: inline-block; width: 44px; height: 24px; flex-shrink: 0;
}
.toggle-switch input { opacity: 0; width: 0; height: 0; }
.toggle-slider {
    position: absolute; cursor: pointer; inset: 0;
    background: var(--primary); border-radius: 24px; transition: background 0.2s;
}
.toggle-slider:before {
    content: ''; position: absolute; height: 18px; width: 18px;
    left: 3px; bottom: 3px; background: #fff; border-radius: 50%; transition: transform 0.2s;
}
input:checked + .toggle-slider { background: #10b981; }
input:checked + .toggle-slider:before { transform: translateX(20px); }
.output-list { display: flex; flex-direction: column; gap: 8px; }
/* SVG mode placeholder */
.svg-mode-placeholder {
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    padding: 40px 20px; gap: 10px; text-align: center;
    border: 2px dashed var(--border); border-radius: 12px; color: var(--text-secondary);
}
.svg-mode-icon { font-size: 36px; line-height: 1; }
.svg-mode-text { font-size: 20px; font-weight: 700; color: var(--text); }
.svg-mode-sub  { font-size: 13px; opacity: 0.7; }
/* SVG grid */
.svg-grid-wrapper { display: flex; flex-direction: column; gap: 12px; padding: 4px 0; }
.svg-grid-scroll { overflow-x: auto; }
.svg-grid-table { border-collapse: collapse; min-width: 100%; font-size: 13px; }
.svg-grid-table th {
    padding: 6px 12px; background: #f0f4ff; border: 1px solid #c5d4f0;
    text-align: center; font-weight: 600; color: #2c3e50; white-space: nowrap;
}
.svg-grid-table td {
    padding: 6px; border: 1px solid #e2e8f0; text-align: center;
    vertical-align: middle; min-width: 120px;
}
.svg-grid-table td.row-header {
    padding: 6px 12px; background: #f8faff; font-weight: 600;
    color: #2c3e50; white-space: nowrap; text-align: left; min-width: 0;
}
.svg-cell-inner {
    display: flex; flex-direction: column; align-items: center;
    gap: 4px; cursor: pointer; text-decoration: none;
}
.svg-cell-inner:hover .svg-thumb { box-shadow: 0 0 0 2px #4A8FE7; border-color: #4A8FE7; }
.svg-thumb {
    width: 120px; height: 76px; object-fit: contain;
    border: 1px solid #d0d7e3; border-radius: 6px;
    background: #fff; display: block;
}
.cell-spinner {
    width: 120px; height: 76px; display: flex; align-items: center;
    justify-content: center; border: 1px dashed #c5d4f0; border-radius: 6px;
    background: #f8faff; color: #8899aa; font-size: 12px;
}
.cell-error {
    width: 120px; height: 76px; display: flex; align-items: center;
    justify-content: center; border: 1px dashed #f8b4b4; border-radius: 6px;
    background: #fff5f5; color: #c0392b; font-size: 11px; text-align: center; padding: 4px;
}
.cell-na {
    width: 120px; height: 76px; display: flex; align-items: center;
    justify-content: center; color: #c0c8d8; font-size: 12px;
}
.svg-cell-label {
    font-size: 10px; color: #6b7a8d; max-width: 120px;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
/* Device File thumbnail row — iframe scaled preview */
.device-iframe-cell {
    display: block; text-decoration: none;
    width: 120px; height: 76px; position: relative;
    margin: 0 auto;
    border: 1px solid #d0d7e3; border-radius: 6px;
    overflow: hidden; background: #fff; cursor: pointer;
}
.device-iframe-cell:hover {
    box-shadow: 0 0 0 2px #4A8FE7; border-color: #4A8FE7;
}
.device-iframe-thumb {
    width: 900px; height: 600px;
    border: none;
    transform: scale(0.1333);
    transform-origin: top left;
    pointer-events: none;
}
.device-file-row td { background: #fafbff; }
.device-file-row td.row-header {
    background: #eef2ff; font-weight: 700; color: #3a50d9; font-size: 12px;
}
.svg-grid-attr {
    display: flex; align-items: center; justify-content: flex-end; gap: 8px;
    font-size: 13px; font-weight: 500; color: #2c3e50;
}
.svg-grid-attr select {
    padding: 4px 10px; border-radius: 6px; border: 1px solid #c5d4f0;
    font-size: 13px; background: #fff; cursor: pointer;
}
.attribute-selector {
    display: flex; align-items: center; gap: 10px;
    padding: 3px 14px; background: #f0f4ff; border-radius: 10px;
    border: 1px solid #c5d4f0; font-size: 14px; font-weight: 500;
    justify-content: flex-end;
}
.attribute-selector select {
    padding: 5px 10px; border-radius: 6px; border: 1px solid var(--border);
    font-size: 14px; background: #fff; cursor: pointer; min-width: 140px;
}
.attribute-selector .attr-note {
    font-size: 11px; color: #7a8a9e; font-weight: 400;
    margin-left: -4px;
}
.output-item {
    display: flex; align-items: center; padding: 6px 14px;
    border: 1px solid var(--border); border-radius: 10px;
    cursor: pointer; transition: background 0.15s; gap: 12px;
}
.output-item:hover { background: #f8f9fa; }
.output-item input[type="checkbox"] {
    width: 18px; height: 18px; accent-color: var(--primary); cursor: pointer; flex-shrink: 0;
}
.output-icon {
    width: 38px; height: 38px; border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    color: white; font-weight: 700; font-size: 11px; flex-shrink: 0; letter-spacing: 0.5px;
}
.output-icon.device { background: var(--device-color); }
.output-icon.l1 { background: var(--l1-color); }
.output-icon.l2 { background: var(--l2-color); }
.output-icon.l3 { background: var(--l3-color); }
.output-icon.ai { background: #2d3436; }
.output-info { display: flex; flex-direction: column; flex: 1; min-width: 0; }
.output-label { font-weight: 600; font-size: 14px; }
.output-filename { font-size: 12px; color: var(--text-secondary); margin-top: 2px; word-break: break-all; }
.output-estimate {
    display: flex; align-items: center; gap: 6px;
    font-size: 12px; color: #8e6f3e; background: #fef9ef; border: 1px solid #f0e0b8;
    border-radius: 5px; padding: 2px 8px; margin-left: auto; white-space: nowrap; flex-shrink: 0;
}
.output-estimate .inline-dot {
    width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0;
    background: transparent; transition: background 0.3s; display: none;
}
.output-estimate .inline-dot.running {
    display: inline-block; background: var(--primary); animation: pulse 1s ease-in-out infinite;
}
.output-estimate .inline-dot.done {
    display: inline-block; background: var(--success);
}
.output-estimate .inline-dot.error {
    display: inline-block; background: var(--error);
}
.output-estimate.est-done { color: var(--success); background: #eafaf1; border-color: #a3d9b1; }
.output-estimate.est-error { color: var(--error); background: #fef0f0; border-color: #f0b8b8; }
.generation-total {
    display: flex; justify-content: flex-end; align-items: center;
    padding: 8px 14px; font-weight: 600; font-size: 14px; color: var(--text);
}
.svg-grid-total {
    display: flex; justify-content: flex-end; align-items: center;
    padding: 4px 14px; font-size: 11px; color: var(--text-secondary);
}
.output-dl {
    padding: 5px 14px; font-size: 12px; font-weight: 600; color: var(--primary);
    background: white; border: 1.5px solid var(--primary); border-radius: 6px;
    cursor: pointer; transition: all 0.2s; text-decoration: none; white-space: nowrap;
    margin-left: auto; flex-shrink: 0; font-family: inherit; line-height: 1.4;
}
.output-dl:hover { background: var(--primary); color: white; }
.master-dl-xlsx {
    color: var(--primary); border-color: var(--primary); background: white;
    margin-left: 0; padding: 5px 14px;
}
.master-dl-xlsx:hover { background: var(--primary) !important; color: white !important; }
.master-dl-nsm {
    color: var(--primary); border-color: var(--primary); background: white;
    margin-left: 0; padding: 5px 14px;
}
.master-dl-nsm:hover { background: var(--primary) !important; color: white !important; }
.output-preview {
    padding: 5px 14px; font-size: 12px; font-weight: 600; color: #6c5ce7;
    background: white; border: 1.5px solid #6c5ce7; border-radius: 6px;
    cursor: pointer; transition: all 0.2s; text-decoration: none; white-space: nowrap;
    flex-shrink: 0;
}
.output-preview:hover { background: #6c5ce7; color: white; }
.output-llm {
    padding: 5px 14px; font-size: 12px; font-weight: 600; color: #0984e3;
    background: white; border: 1.5px solid #0984e3; border-radius: 6px;
    cursor: pointer; transition: all 0.2s; text-decoration: none; white-space: nowrap;
    flex-shrink: 0;
}
.output-llm:hover { background: #0984e3; color: white; }
.output-copy-only {
    padding: 5px 14px; font-size: 12px; font-weight: 600; color: #0984e3;
    background: white; border: 1.5px solid #0984e3; border-radius: 6px;
    cursor: pointer; transition: all 0.2s; text-decoration: none; white-space: nowrap;
    flex-shrink: 0;
}
.output-copy-only:hover { background: #0984e3; color: white; }
.llm-prompt-area {
    display: none; width: 100%; margin-top: 4px; padding: 4px 10px;
    border: 1.5px solid #0984e3; border-radius: 8px; background: #f0f7ff;
    box-sizing: border-box; flex-basis: 100%;
}
.llm-prompt-area textarea {
    width: 100%; min-height: 96px; padding: 8px 10px; border: 1px solid #d0d8e8;
    border-radius: 6px; font-size: 13px; font-family: inherit; resize: vertical;
    box-sizing: border-box; background: white;
}
.llm-prompt-area .llm-prompt-hint {
    font-size: 11px; color: #888; margin-top: 4px;
}
.btn-create-master {
    padding: 6px 16px; font-size: 12px; font-weight: 600; color: #00b894;
    background: white; border: 1.5px solid #00b894; border-radius: 6px;
    cursor: pointer; transition: all 0.2s; white-space: nowrap;
}
.btn-create-master:hover { background: #00b894; color: white; }
.btn-create-master:disabled { opacity: 0.5; cursor: default; }
.output-gen {
    padding: 5px 14px; font-size: 12px; font-weight: 600; color: #fff;
    background: var(--primary); border: 1.5px solid var(--primary); border-radius: 6px;
    cursor: pointer; transition: all 0.2s; white-space: nowrap; flex-shrink: 0;
}
.output-gen:hover { background: var(--primary-dark); border-color: var(--primary-dark); }
.output-gen:disabled { opacity: 0.4; cursor: default; }
.output-preview-svg {
    padding: 5px 14px; font-size: 12px; font-weight: 600; color: var(--primary);
    background: #fff; border: 1.5px solid var(--primary); border-radius: 6px;
    cursor: pointer; transition: all 0.2s; white-space: nowrap; flex-shrink: 0;
}
.output-preview-svg:hover { background: #eef2ff; }
.output-preview-svg:disabled { opacity: 0.4; cursor: default; }
.output-preview-csv {
    padding: 5px 14px; font-size: 12px; font-weight: 600; color: var(--primary);
    background: #fff; border: 1.5px solid var(--primary); border-radius: 6px;
    cursor: pointer; transition: all 0.2s; white-space: nowrap; flex-shrink: 0;
}
.output-preview-csv:hover { background: #eef2ff; }
.update-desc { font-size: 13px; color: var(--text-secondary); margin-bottom: 12px; line-height: 1.5; }
.cli-command-area {
    width: 100%; padding: 2px 10px 5px; border: 1.5px solid #00b894; border-radius: 8px;
    background: #f0faf6; box-sizing: border-box;
}
/* Halve the top/bottom inner padding of the Update Master card only.
   The global .card rule uses padding: 28px; we override just the
   vertical axis on this section to make the frame more compact. */
#updateSection.card {
    padding-top: 14px;
    padding-bottom: 14px;
}
/* Halve the gap between the "Update Master" heading and the
   "Paste CLI commands..." description below it (global .card h2 uses
   margin-bottom: 16px). */
#updateSection.card h2 {
    margin-bottom: 8px;
}
.collapsible-header {
    display: flex; align-items: center; gap: 6px;
    cursor: pointer; padding: 8px 0; border-radius: 6px;
    user-select: none;
}
.collapsible-header:hover .collapsible-title { color: var(--primary); }
.collapsible-title { font-size: 13px; font-weight: 600; color: var(--text); }
.collapsible-arrow {
    margin-left: auto; font-size: 10px; color: var(--text-secondary);
    transition: transform 0.2s; display: inline-block;
}
.collapsible-arrow.open { transform: rotate(90deg); }
.collapsible-body { padding-top: 4px; }
/* ---- Command Reference drawer ---- */
.cmdref-trigger {
    display: inline-flex; align-items: center; gap: 4px;
    padding: 2px 10px; font-size: 11px; font-weight: 600;
    color: #00897b; background: #fff; border: 1px solid #00b894;
    border-radius: 12px; cursor: pointer; white-space: nowrap;
    transition: all 0.15s; line-height: 1.6;
}
.cmdref-trigger:hover { background: #00b894; color: #fff; }
/* In the Run action row, keep the reference trigger on the left while
   run-status + Run button stay right-aligned. */
.update-actions .cmdref-trigger { margin-right: auto; }
.cmdref-overlay {
    position: fixed; inset: 0; background: rgba(0,0,0,0.35);
    z-index: 10000; display: none; opacity: 0; transition: opacity 0.2s;
}
.cmdref-overlay.open { display: block; opacity: 1; }
.cmdref-drawer {
    position: fixed; top: 0; right: 0; height: 100%;
    width: min(600px, 94vw); background: #fff; z-index: 10001;
    box-shadow: -4px 0 24px rgba(0,0,0,0.18);
    transform: translateX(100%); transition: transform 0.22s ease;
    display: flex; flex-direction: column;
}
.cmdref-drawer.open { transform: translateX(0); }
.cmdref-head {
    display: flex; align-items: center; gap: 10px; padding: 14px 18px;
    background: #16213e; color: #fff; flex-shrink: 0;
}
.cmdref-head h3 { font-size: 15px; font-weight: 600; margin: 0; flex: 1; }
.cmdref-close {
    background: transparent; border: 1px solid rgba(255,255,255,0.4);
    color: #fff; border-radius: 6px; width: 28px; height: 28px;
    font-size: 16px; cursor: pointer; line-height: 1;
}
.cmdref-close:hover { background: rgba(255,255,255,0.15); }
.cmdref-searchbar { padding: 10px 18px; border-bottom: 1px solid #eee; flex-shrink: 0; }
.cmdref-searchbar input {
    width: 100%; padding: 8px 12px; font-size: 13px;
    border: 1px solid #d0d8e8; border-radius: 6px; box-sizing: border-box;
}
.cmdref-searchbar input:focus { outline: none; border-color: var(--primary); }
.cmdref-body { flex: 1; overflow-y: auto; padding: 8px 0; }
.cmdref-cat { padding: 8px 18px 2px; font-size: 12px; font-weight: 700;
    color: #00897b; text-transform: uppercase; letter-spacing: 0.04em;
    position: sticky; top: 0; background: #fff; }
.cmdref-item { padding: 8px 18px 12px; border-bottom: 1px solid #f0f0f0; }
.cmdref-item-title { font-size: 13px; font-weight: 700; color: #1a2744;
    font-family: 'Consolas', 'Monaco', 'Courier New', monospace; }
.cmdref-syntax {
    margin: 6px 0; padding: 8px 10px; background: #0f1b2d; color: #e6edf3;
    border-radius: 6px; font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
    font-size: 12px; white-space: pre-wrap; word-break: break-word;
}
.cmdref-actions { display: flex; gap: 6px; margin: 4px 0 6px; }
.cmdref-btn {
    padding: 3px 12px; font-size: 11px; font-weight: 600; border-radius: 5px;
    cursor: pointer; transition: all 0.15s; border: 1px solid #00b894;
    background: #fff; color: #00897b;
}
.cmdref-btn:hover { background: #00b894; color: #fff; }
.cmdref-btn.copy { border-color: #0984e3; color: #0984e3; }
.cmdref-btn.copy:hover { background: #0984e3; color: #fff; }
.cmdref-desc { font-size: 12px; color: #555; line-height: 1.55;
    white-space: pre-wrap; word-break: break-word; }
.cmdref-desc code { background: #eef2f7; padding: 1px 4px; border-radius: 3px;
    font-family: 'Consolas', 'Monaco', 'Courier New', monospace; font-size: 11px; }
.cmdref-empty { padding: 24px 18px; color: #888; font-size: 13px; text-align: center; }
.cmdref-feedback {
    position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%);
    background: rgba(0,0,0,0.78); color: #fff; padding: 7px 20px;
    border-radius: 20px; font-size: 13px; pointer-events: none; z-index: 10002;
    opacity: 0; transition: opacity 0.3s; max-width: 80vw; white-space: nowrap;
    overflow: hidden; text-overflow: ellipsis;
}
.cmd-input {
    width: 100%; padding: 8px 10px; font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
    font-size: 13px; border: 1px solid #d0d8e8; border-radius: 6px; resize: vertical;
    background: white; transition: border-color 0.2s; line-height: 1.5;
    box-sizing: border-box;
}
.cmd-input:focus { outline: none; border-color: var(--primary); }
.update-actions { display: flex; align-items: center; gap: 12px; margin-top: 6px; justify-content: flex-end; }
.btn-run {
    padding: 10px 28px; font-size: 14px; font-weight: 600; color: #fff;
    background: #00b894; border: none; border-radius: 8px; cursor: pointer;
    transition: all 0.2s;
}
.btn-run:hover { background: #00a381; }
.btn-run:disabled { opacity: 0.5; cursor: default; }
.run-status { font-size: 13px; }
.run-progress { display: flex; align-items: center; gap: 10px; margin-top: 12px; font-size: 14px; color: var(--text-secondary); }
.run-results {
    margin-top: 14px; padding: 12px; background: #f8f9fa; border-radius: 8px;
    font-size: 12px; font-family: 'Consolas', monospace; max-height: 300px;
    overflow-y: auto; line-height: 1.6;
}
.run-results .cmd-ok { color: var(--success); }
.run-results .cmd-err { color: #e94560; }
.run-results .cmd-skip { color: #e6a817; }
.view-log-link {
    display: inline-block; margin-top: 8px; font-size: 12px; color: var(--primary);
    cursor: pointer; text-decoration: underline;
}
.view-log-link:hover { color: var(--primary-dark); }
.copy-log-btn {
    display: inline-block; margin-top: 8px; margin-left: 6px; padding: 1px 10px;
    font-size: 11px; font-weight: 600; color: #0984e3; background: #fff;
    border: 1px solid #0984e3; border-radius: 5px; cursor: pointer;
    vertical-align: baseline; transition: all 0.15s;
}
.copy-log-btn:hover { background: #0984e3; color: #fff; }
.cmd-output {
    display: none; margin: 4px 0 8px 18px; padding: 8px 10px; background: #fff;
    border: 1px solid #e0e0e0; border-radius: 6px; white-space: pre-wrap;
    word-break: break-all; font-size: 11px; color: #444; line-height: 1.5;
}
.cmd-output.visible { display: block; }
.output-icon.master { background: #fdcb6e; color: #2d3436; }
.output-icon.master_dl { background: #059669; }
.estimate-note {
    font-size: 12px; color: var(--text-secondary); margin-top: 10px; padding: 8px 12px;
    background: #f8f9fa; border-radius: 6px; line-height: 1.5;
}
.estimate-total {
    font-size: 13px; font-weight: 600; color: #8e6f3e; margin-top: 8px;
    padding: 10px 14px; background: #fef9ef; border: 1px solid #f0e0b8; border-radius: 8px;
    display: flex; justify-content: space-between; align-items: center;
}
.selected-count { font-size: 13px; color: var(--text-secondary); margin-left: auto; margin-right: 12px; }

/* Buttons */
.btn-generate {
    display: inline-block; width: auto; padding: 8px 18px; font-size: 13px; font-weight: 600;
    color: white; background: var(--primary); border: none; border-radius: 6px;
    cursor: pointer; transition: background 0.2s; white-space: nowrap; flex-shrink: 0;
}
.btn-generate:hover { background: var(--primary-dark); }
.btn-generate:disabled { background: #bdc3c7; cursor: not-allowed; }

/* Inline progress animation */
@keyframes pulse {
    0%, 100% { opacity: 1; transform: scale(1); }
    50% { opacity: 0.5; transform: scale(0.85); }
}

/* Results */
.file-grid { display: grid; grid-template-columns: 1fr; gap: 10px; }
.file-card {
    display: flex; align-items: center; padding: 14px 16px;
    border: 1px solid var(--border); border-radius: 10px;
    background: #fafbfc; transition: background 0.2s, box-shadow 0.2s;
}
.file-card:hover { background: #f0f7ff; box-shadow: 0 2px 8px rgba(74,143,231,0.1); }
.file-icon {
    width: 42px; height: 42px; border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    color: white; font-weight: 700; font-size: 11px; margin-right: 14px;
    flex-shrink: 0; letter-spacing: 0.5px;
}
.file-icon.device { background: var(--device-color); }
.file-icon.l1 { background: var(--l1-color); }
.file-icon.l2 { background: var(--l2-color); }
.file-icon.l3 { background: var(--l3-color); }
.file-icon.unknown { background: #95a5a6; }
.file-info { flex: 1; min-width: 0; }
.file-name { font-size: 13px; font-weight: 600; word-break: break-all; }
.file-size { font-size: 12px; color: var(--text-secondary); margin-top: 2px; }
.btn-dl {
    padding: 7px 16px; font-size: 13px; font-weight: 600; color: var(--primary);
    background: white; border: 1.5px solid var(--primary); border-radius: 7px;
    cursor: pointer; transition: all 0.2s; text-decoration: none; white-space: nowrap; margin-left: 12px;
}
.btn-dl:hover { background: var(--primary); color: white; }
.btn-dl-all {
    display: block; width: 100%; padding: 13px; font-size: 15px; font-weight: 600;
    color: white; background: var(--primary); border: none; border-radius: 10px;
    cursor: pointer; text-align: center; text-decoration: none; margin-top: 16px; transition: background 0.2s;
}
.btn-dl-all:hover { background: var(--primary-dark); }
.btn-dl-all.disabled { background: #b0b8c9; cursor: not-allowed; pointer-events: none; }
.error-msg {
    background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px;
    padding: 12px 16px; color: var(--error); font-size: 14px; margin-top: 12px; display: none;
}
.btn-reset {
    display: block; width: 100%; padding: 12px; font-size: 14px; font-weight: 600;
    color: var(--text-secondary); background: transparent; border: 1.5px solid var(--border);
    border-radius: 10px; cursor: pointer; margin-top: 12px; transition: all 0.2s;
}
.btn-reset:hover { background: #f5f5f5; border-color: #ccc; }
.footer { text-align: center; margin-top: 32px; font-size: 12px; color: var(--text-secondary); }
.footer a { color: var(--primary); text-decoration: none; }
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <img src="/static/logo.png" alt="Network Sketcher" class="logo">
        <div class="header-text">
            <h1>Network Sketcher <span class="online-badge">Online</span></h1>
            <p class="subtitle">Upload master files to generate diagrams and device files or update designs</p>
        </div>
    </div>

    <!-- Restore indicator (shown only during session restore) -->
    <div id="restoreIndicator" style="display:none;text-align:center;padding:40px 20px;color:#666;">
        <div style="display:inline-block;width:28px;height:28px;border:3px solid #ddd;border-top-color:#4A8FE7;border-radius:50%;animation:spin 0.8s linear infinite;margin-bottom:12px;"></div>
        <div style="font-size:15px;">Restoring session...</div>
    </div>
    <script>
    (function(){
        try{
            var s=sessionStorage.getItem('ns_session');
            if(s&&JSON.parse(s).jobId){
                document.getElementById('restoreIndicator').style.display='block';
                document.write('<style id="hideUploadStyle">#uploadSection{display:none!important}</style>');
                window._restoreTimeout=setTimeout(function(){
                    try{sessionStorage.removeItem('ns_session');}catch(e){}
                    var ri=document.getElementById('restoreIndicator');if(ri)ri.style.display='none';
                    var hs=document.getElementById('hideUploadStyle');if(hs)hs.remove();
                    var us=document.getElementById('uploadSection');if(us)us.style.display='';
                },5000);
            }
        }catch(e){}
    })();
    </script>

    <!-- Upload Section -->
    <div class="card" id="uploadSection">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;">
            <h2 style="margin-bottom:0;">&#128194; Master File <span class="help-icon" onclick="event.stopPropagation();toggleHelp('master_file')">?</span></h2>
            <div style="display:flex;align-items:center;gap:6px;">
                <button type="button" class="btn-create-master" id="btnCreateMaster">Create New Master File</button>
                <span class="help-icon" onclick="event.stopPropagation();toggleHelp('create_new_master')">?</span>
            </div>
        </div>
        <div class="help-tooltip" id="help-create_new_master"></div>
        <div class="help-tooltip" id="help-master_file"></div>
        <div class="dropzone" id="dropzone">
            <div class="dropzone-icon">&#128196;</div>
            <div class="dropzone-text">
                Drag &amp; drop a [MASTER]*.xlsx or .nsm file here<br>
                or <strong id="browseLink">click to browse</strong>
            </div>
            <div class="dropzone-file" id="selectedFile"></div>
        </div>
        <input type="file" id="fileInput" accept=".xlsx,.nsm">
        <div class="error-msg" id="uploadError"></div>
        <div class="analyzing" id="analyzingIndicator" style="display:none;">
            <div class="spinner"></div>
            <span>Uploading and analyzing master file...</span>
            <span id="analyzingEstimate" style="margin-left:10px;color:var(--text-secondary);font-size:0.9em;"></span>
        </div>
    </div>

    <!-- Update Master Section -->
    <div class="card" id="updateSection" style="display:none;">
        <h2>&#9998; Update Master <span class="help-icon" onclick="event.stopPropagation();toggleHelp('update_master')">?</span></h2>
        <div class="help-tooltip" id="help-update_master"></div>
        <p class="update-desc">Paste CLI commands generated by an LLM from the AI Context file. Each line is executed against the current master file.</p>
        <div class="llm-prompt-area" id="llmPromptArea" style="display:none;">
            <div class="collapsible-header" id="llmCommandsHeader">
                <span class="collapsible-title">LLM Prompt</span>
                <span class="help-icon" onclick="event.stopPropagation();toggleHelp('llm_prompt')">?</span>
                <span class="collapsible-arrow">&#9654;</span>
            </div>
            <div class="help-tooltip" id="help-llm_prompt"></div>
            <div class="collapsible-body" id="llmCommandsBody" style="display:none;">
                <textarea id="llmPromptInput" placeholder="Describe what Network Sketcher commands you need (e.g., Add a router R-5 in area DC-TOP1)&#10;Multilingual support: English, Deutsch, Español, Français, Italiano, Português, Tiếng Việt, 日本語, 中文, 한국어, हिन्दी, العربية, עברית"></textarea>
                <div class="llm-prompt-hint">This text will be appended to the AI Context when copied, requesting the LLM to output only executable commands.</div>
                <div style="display:flex;gap:6px;margin-top:8px;align-items:center;flex-wrap:wrap;" id="llmBtnRow">
                    <span style="font-size:11px;font-weight:700;color:#2563eb;flex:1;min-width:200px;">&#9888; All network configuration data will be uploaded to the LLM. Ensure you fully understand the risk of information disclosure before proceeding.</span>
                    <span class="llm-gen-status" id="llmGenStatus" style="font-size:12px;color:#888;display:none;"></span>
                    <button type="button" class="output-llm" id="btnCopyOpenLlm" style="display:none;">{{AI_CONTEXT_BTN_LABEL}}</button>
                    <button type="button" class="output-copy-only" id="btnCopyOnly" style="display:none;">Copy</button>
                    <button type="button" class="output-copy-only" id="btnDlAiContext" style="display:none;">Download</button>
                </div>
            </div>
        </div>
        <div class="cli-command-area">
            <div class="collapsible-header" id="cliCommandsHeader">
                <span class="collapsible-title">CLI Commands</span>
                <span class="help-icon" onclick="event.stopPropagation();toggleHelp('cli_commands')">?</span>
                <span class="collapsible-arrow">&#9654;</span>
            </div>
            <div class="help-tooltip" id="help-cli_commands"></div>
            <div class="collapsible-body" id="cliCommandsBody" style="display:none;">
                <textarea id="cmdInput" class="cmd-input" rows="8" placeholder="show device&#10;show l1_link&#10;add device Router-A SW-1B RIGHT&#10;add l1_link_bulk &quot;[['SW-1B','WAN-1','GigabitEthernet 0/24','GigabitEthernet 0/24']]&quot;&#10;delete device OldDevice"></textarea>
                <div class="update-actions">
                    <button type="button" class="cmdref-trigger" id="btnCmdRef" onclick="openCmdRef();" title="Browse the Network Sketcher command reference">&#128214; Command Reference</button>
                    <div class="run-status" id="runStatus" style="display:none;"></div>
                    <button class="btn-run" id="btnRun">Run</button>
                </div>
            </div>
        </div>
        <div class="run-progress" id="runProgress" style="display:none;">
            <div class="spinner"></div> <span id="runProgressText">Executing commands...</span>
        </div>
        <div class="run-results" id="runResults" style="display:none;"></div>
    </div>

    <!-- Command Reference drawer -->
    <div class="cmdref-overlay" id="cmdRefOverlay" onclick="closeCmdRef()"></div>
    <div class="cmdref-drawer" id="cmdRefDrawer" role="dialog" aria-label="Command Reference">
        <div class="cmdref-head">
            <h3>&#128214; Command Reference</h3>
            <button type="button" class="cmdref-close" id="cmdRefClose" onclick="closeCmdRef()" title="Close">&#10005;</button>
        </div>
        <div class="cmdref-searchbar">
            <input type="text" id="cmdRefSearch" placeholder="Search commands (e.g., l1_link, area, ip_address)" autocomplete="off">
        </div>
        <div class="cmdref-body" id="cmdRefBody"></div>
    </div>
    <div class="cmdref-feedback" id="cmdRefFeedback"></div>

    <!-- Selection Section -->
    <div class="card" id="selectionSection" style="display:none;">
        <div class="available-outputs-header">
            <h2>&#128203; Available Outputs <span class="help-icon" onclick="event.stopPropagation();toggleHelp('available_outputs')">?</span></h2>
            <div class="format-switch">
                <span class="format-label active" id="labelSvg">SVG Mode</span>
                <label class="toggle-switch" title="Switch diagram output format between SVG/CSV and PPTX/XLSX">
                    <input type="checkbox" id="formatToggle">
                    <span class="toggle-slider"></span>
                </label>
                <span class="format-label" id="labelPptx">PPTX Mode</span>
            </div>
        </div>
        <div class="help-tooltip" id="help-available_outputs"></div>
        <div id="analysisSummary" style="font-size:13px;color:var(--text-secondary);margin-bottom:14px;padding:8px 12px;background:#f0f7ff;border-radius:6px;display:none;"></div>
        <div class="select-all-row">
            <label><input type="checkbox" id="selectAll" checked> Select / Deselect All</label>
            <span class="selected-count" id="selectedCount"></span>
            <button class="btn-generate" id="btnGenerate">Generate Selected</button>
        </div>
        <div id="fileActionsSection" style="display:none;margin-bottom:12px;"></div>
        <div class="output-list" id="outputList"></div>
        <div id="attributeContainer" style="display:none;"></div>
        <div class="estimate-total" id="estimateTotal" style="display:none;">
            <span>Estimated total time (Actual times may vary depending on your system.)</span>
            <span id="estimateTotalTime"></span>
        </div>
        <div class="generation-total" id="generationTotal" style="display:none;"></div>
        <div class="svg-grid-total" id="svgGridTotal" style="display:none;"></div>
        <a class="btn-dl-all" id="btnDownloadAll" href="#" style="display:none;">Download Generated (ZIP)</a>
        <button class="btn-reset" id="btnReset" style="display:none;">Upload New File</button>
    </div>

    <div class="footer">
        <span style="color:#2563eb;font-size:12px;font-weight:600;">All files are automatically deleted after the session ends &mdash; no data is retained on the server.</span>
        <br>
        <a href="https://github.com/cisco-open/network-sketcher" target="_blank" rel="noopener noreferrer">
            Network Sketcher
        </a> &mdash; Apache-2.0 License
        <br>
        <span style="font-size:11px;color:#b0b8c4;">&copy; 2023 Cisco Systems, Inc. and its affiliates | Created by Yusuke Ogawa - Architect, Cisco | CCIE#17583</span>
    </div>
</div>

<script type="application/json" id="helpDataJson">{{HELP_DATA_JSON}}</script>
<script>
var _helpData = {};
try { _helpData = JSON.parse(document.getElementById('helpDataJson').textContent); } catch(e) {}

function toggleHelp(key) {
    var el = document.getElementById('help-' + key);
    if (!el) return;
    if (!el.textContent && _helpData[key]) {
        el.textContent = _helpData[key];
    }
    var icon = el.closest('.card, .llm-prompt-area, div');
    var activeIcon = icon ? icon.querySelector('.help-icon.active') : null;
    var isVisible = el.classList.contains('visible');
    var allTooltips = document.querySelectorAll('.help-tooltip');
    var allIcons = document.querySelectorAll('.help-icon');
    for (var i = 0; i < allTooltips.length; i++) allTooltips[i].classList.remove('visible');
    for (var i = 0; i < allIcons.length; i++) allIcons[i].classList.remove('active');
    if (!isVisible) {
        el.classList.add('visible');
        var parentEl = el.parentElement;
        var thisIcon = parentEl ? parentEl.querySelector('.help-icon') : null;
        if (thisIcon) thisIcon.classList.add('active');
    }
}

function initCollapsible(headerId, bodyId) {
    var header = document.getElementById(headerId);
    var body = document.getElementById(bodyId);
    if (!header || !body) return;
    header.addEventListener('click', function(e) {
        if (e.target.closest('.help-icon')) return;
        var isOpen = body.style.display !== 'none';
        body.style.display = isOpen ? 'none' : '';
        var arrow = header.querySelector('.collapsible-arrow');
        if (arrow) arrow.classList.toggle('open', !isOpen);
    });
}
initCollapsible('cliCommandsHeader', 'cliCommandsBody');
initCollapsible('llmCommandsHeader', 'llmCommandsBody');

// ===== Command Reference drawer =====
var _cmdRefData = null;        // cached parsed reference (categories)
var _cmdRefLoaded = false;
var _cmdRefFbTimer = null;

function _cmdRefEscapeHtml(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

// Lightweight inline-markdown: backtick code spans only (descriptions are
// plain text otherwise). Everything is HTML-escaped first.
function _cmdRefDesc(text) {
    var esc = _cmdRefEscapeHtml(text);
    esc = esc.replace(/`([^`]+)`/g, function(m, c) { return '<code>' + c + '</code>'; });
    return esc;
}

function _cmdRefNormalizeQuotes(s) {
    // Mirror server-side _normalize_quotes so inserted text uses straight
    // quotes accepted by the Run validator.
    return String(s == null ? '' : s)
        .replace(/[\u2018\u2019]/g, "'")
        .replace(/[\u201c\u201d]/g, '"');
}

function openCmdRef() {
    var overlay = document.getElementById('cmdRefOverlay');
    var drawer = document.getElementById('cmdRefDrawer');
    overlay.classList.add('open');
    drawer.classList.add('open');
    var search = document.getElementById('cmdRefSearch');
    if (!_cmdRefLoaded) {
        var body = document.getElementById('cmdRefBody');
        body.innerHTML = '<div class="cmdref-empty">Loading...</div>';
        fetch('/command_reference').then(function(r) { return r.json(); })
            .then(function(data) {
                _cmdRefData = data || [];
                _cmdRefLoaded = true;
                _renderCmdRef('');
            })
            .catch(function() {
                body.innerHTML = '<div class="cmdref-empty">Failed to load command reference.</div>';
            });
    } else {
        _renderCmdRef(search ? search.value : '');
    }
    setTimeout(function() { if (search) search.focus(); }, 60);
}

function closeCmdRef() {
    document.getElementById('cmdRefOverlay').classList.remove('open');
    document.getElementById('cmdRefDrawer').classList.remove('open');
}

function _renderCmdRef(filter) {
    var body = document.getElementById('cmdRefBody');
    if (!_cmdRefData) { return; }
    var q = (filter || '').trim().toLowerCase();
    var html = '';
    var matchCount = 0;
    for (var ci = 0; ci < _cmdRefData.length; ci++) {
        var cat = _cmdRefData[ci];
        var items = cat.items || [];
        var rows = '';
        for (var ii = 0; ii < items.length; ii++) {
            var it = items[ii];
            if (q) {
                var hay = (it.title + ' ' + (it.body || '') + ' ' + (it.syntax || '')).toLowerCase();
                if (hay.indexOf(q) === -1) continue;
            }
            matchCount++;
            var actions = '';
            if (it.syntax) {
                actions = '<div class="cmdref-actions">'
                    + '<button type="button" class="cmdref-btn" onclick="cmdRefInsert(' + ci + ',' + ii + ')">Insert</button>'
                    + '<button type="button" class="cmdref-btn copy" onclick="cmdRefCopy(' + ci + ',' + ii + ')">Copy</button>'
                    + '</div>';
            }
            var syntaxHtml = it.syntax
                ? '<div class="cmdref-syntax">' + _cmdRefEscapeHtml(it.syntax) + '</div>'
                : '';
            rows += '<div class="cmdref-item">'
                + '<div class="cmdref-item-title">' + _cmdRefEscapeHtml(it.title) + '</div>'
                + syntaxHtml + actions
                + '<div class="cmdref-desc">' + _cmdRefDesc(it.body || '') + '</div>'
                + '</div>';
        }
        if (rows) {
            html += '<div class="cmdref-cat">' + _cmdRefEscapeHtml(cat.category) + '</div>' + rows;
        }
    }
    if (!matchCount) {
        html = '<div class="cmdref-empty">No commands match "' + _cmdRefEscapeHtml(filter) + '".</div>';
    }
    body.innerHTML = html;
}

function _cmdRefSyntaxAt(ci, ii) {
    try { return _cmdRefData[ci].items[ii].syntax || ''; } catch (e) { return ''; }
}

function cmdRefInsert(ci, ii) {
    var syntax = _cmdRefNormalizeQuotes(_cmdRefSyntaxAt(ci, ii));
    if (!syntax) return;
    var ta = document.getElementById('cmdInput');
    if (!ta) return;
    // Ensure the CLI Commands body is expanded so the user sees the result.
    var bodyEl = document.getElementById('cliCommandsBody');
    if (bodyEl && bodyEl.style.display === 'none') {
        bodyEl.style.display = '';
        var arrow = document.querySelector('#cliCommandsHeader .collapsible-arrow');
        if (arrow) arrow.classList.add('open');
    }
    var cur = ta.value;
    if (cur && !cur.endsWith('\n')) cur += '\n';
    ta.value = cur + syntax + '\n';
    ta.focus();
    ta.scrollTop = ta.scrollHeight;
    _cmdRefFeedback('Inserted into CLI Commands');
}

function cmdRefCopy(ci, ii) {
    var syntax = _cmdRefNormalizeQuotes(_cmdRefSyntaxAt(ci, ii));
    if (!syntax) return;
    var done = function() { _cmdRefFeedback('Copied: ' + syntax); };
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(syntax).then(done).catch(function() { _cmdRefCopyFallback(syntax, done); });
    } else {
        _cmdRefCopyFallback(syntax, done);
    }
}

function _cmdRefCopyFallback(text, done) {
    try {
        var t = document.createElement('textarea');
        t.value = text; t.style.position = 'fixed'; t.style.opacity = '0';
        document.body.appendChild(t); t.select();
        document.execCommand('copy'); document.body.removeChild(t);
        done();
    } catch (e) {}
}

function _cmdRefFeedback(msg) {
    var fb = document.getElementById('cmdRefFeedback');
    if (!fb) return;
    fb.textContent = msg;
    fb.style.opacity = '1';
    clearTimeout(_cmdRefFbTimer);
    _cmdRefFbTimer = setTimeout(function() { fb.style.opacity = '0'; }, 1600);
}

(function() {
    var search = document.getElementById('cmdRefSearch');
    if (search) {
        search.addEventListener('input', function() { _renderCmdRef(this.value); });
    }
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            var drawer = document.getElementById('cmdRefDrawer');
            if (drawer && drawer.classList.contains('open')) closeCmdRef();
        }
    });
})();

(function() {
    var dropzone = document.getElementById('dropzone');
    var fileInput = document.getElementById('fileInput');
    var browseLink = document.getElementById('browseLink');
    var selectedFileEl = document.getElementById('selectedFile');
    var uploadError = document.getElementById('uploadError');
    var analyzingIndicator = document.getElementById('analyzingIndicator');
    var uploadSection = document.getElementById('uploadSection');
    var selectionSection = document.getElementById('selectionSection');
    var selectAll = document.getElementById('selectAll');
    var selectedCount = document.getElementById('selectedCount');
    var outputList = document.getElementById('outputList');
    var btnGenerate = document.getElementById('btnGenerate');
    var generationTotal = document.getElementById('generationTotal');
    var btnDownloadAll = document.getElementById('btnDownloadAll');
    var btnReset = document.getElementById('btnReset');
    var updateSection = document.getElementById('updateSection');
    var cmdInput = document.getElementById('cmdInput');
    var btnRun = document.getElementById('btnRun');
    var runStatus = document.getElementById('runStatus');
    var runProgress = document.getElementById('runProgress');
    var runProgressText = document.getElementById('runProgressText');
    var runResults = document.getElementById('runResults');
    var masterVersion = 0;
    var updatedMasters = [];

    var currentJobId = null;
    var currentAreas = [];
    var currentSelectedArea = ''; // area chosen via the per-area dropdown (SVG mode only)
    var currentBasename = '';
    var currentMasterFilename = '';
    var currentDeviceCount = 0;
    var currentLinkCount = 0;
    var currentAttributeTitles = [];
    var isEmptyMaster = false;
    var heartbeatTimer = null;
    var outputMode = 'svg'; // 'pptx' or 'svg' — default SVG/CSV
    var svgGridEventSource = null; // active SSE connection for SVG grid

    // ------------------------------------------------------------------ //
    // Estimated time helpers (based on measured performance data)        //
    // ------------------------------------------------------------------ //
    // Performance reference data: [device_count, seconds]
    var _PERF_MASTER   = [[64, 4], [256, 4], [1024, 26], [4096, 492]];
    var _PERF_DIAGRAMS = [[64, 3], [256, 7], [1024, 58], [4096, 942]];

    function _estimateSec(deviceCount, perfData) {
        if (!deviceCount || deviceCount <= 0) return null;
        var d = perfData;
        if (deviceCount <= d[0][0]) return d[0][1];
        if (deviceCount >= d[d.length - 1][0]) {
            // log-log extrapolation from last two points
            var n = d.length;
            var slope = Math.log(d[n-1][1] / d[n-2][1]) / Math.log(d[n-1][0] / d[n-2][0]);
            return d[n-1][1] * Math.pow(deviceCount / d[n-1][0], slope);
        }
        // log-log interpolation between surrounding points
        for (var i = 0; i < d.length - 1; i++) {
            if (deviceCount >= d[i][0] && deviceCount < d[i+1][0]) {
                var t = Math.log(deviceCount / d[i][0]) / Math.log(d[i+1][0] / d[i][0]);
                return Math.exp(Math.log(d[i][1]) + t * Math.log(d[i+1][1] / d[i][1]));
            }
        }
        return null;
    }

    function _fmtEstimate(sec) {
        if (sec === null) return '';
        sec = Math.ceil(sec);
        if (sec < 60) return 'Estimated: ~' + sec + 's';
        return 'Estimated: ~' + Math.floor(sec / 60) + 'm ' + (sec % 60) + 's';
    }

    // ------------------------------------------------------------------ //
    // SVG Grid: hybrid dropdown model (v3.1.2a)                          //
    //                                                                    //
    // - Initial render shows three rows only: Device Tables, All Areas,  //
    //   and a Per-area dropdown row. The per-area cells start empty      //
    //   with placeholder text; nothing is generated until the user picks //
    //   an area from the dropdown.                                       //
    // - Each area selection triggers an on-demand SSE call that produces //
    //   only that area's L1 / L2 / L3 thumbnails. Server-side cache is   //
    //   keyed by (attribute, area), so re-picking the same area after    //
    //   it has been generated once is instantaneous.                     //
    // - ZIP button is hidden until the per-area generation completes,    //
    //   then becomes available with ?area=&attr= query parameters that   //
    //   restrict the ZIP to the selected area's diagrams.                //
    // ------------------------------------------------------------------ //

    // cellMap is reused by both the initial SSE and the per-area SSE, so it
    // lives outside buildSvgGrid().
    var svgGridCellMap = {};
    var svgGridSelAttr = '';
    var svgGridPerAreaEventSource = null;
    var svgGridPerAreaStartTime = 0;

    function buildSvgGrid(initAttr) {
        // Close any existing SSE streams (both initial and per-area)
        if (svgGridEventSource) { svgGridEventSource.close(); svgGridEventSource = null; }
        if (svgGridPerAreaEventSource) {
            svgGridPerAreaEventSource.close();
            svgGridPerAreaEventSource = null;
        }
        svgGridCellMap = {};

        // Determine selected attribute: prefer initAttr (passed from change handler),
        // then the existing DOM dropdown, then the first title.
        var selAttr = initAttr || '';
        if (!selAttr) {
            var svgAttrEl = document.getElementById('svgGridAttr');
            if (svgAttrEl) {
                selAttr = svgAttrEl.value || '';
            } else if (currentAttributeTitles.length > 1) {
                selAttr = currentAttributeTitles[0];
            }
        }
        svgGridSelAttr = selAttr;

        // Reset per-area selection on every (re)build (e.g. attribute toggle).
        currentSelectedArea = '';

        // Column definitions
        var cols = [
            { id: 'l1', label: 'Layer 1' },
            { id: 'l2', label: 'Layer 2' },
            { id: 'l3', label: 'Layer 3' },
        ];

        // Build HTML table
        var wrapper = document.createElement('div');
        wrapper.className = 'svg-grid-wrapper';

        // Attribute selector row for SVG mode
        if (currentAttributeTitles.length > 1) {
            var attrDiv = document.createElement('div');
            attrDiv.className = 'svg-grid-attr';
            attrDiv.innerHTML = '<label for="svgGridAttr">Attribute:</label>'
                + '<select id="svgGridAttr"></select>'
                + '<span style="font-size:11px;color:#7a8a9e;">Device color pattern for L1 / L3</span>';
            var attrSel2 = attrDiv.querySelector('#svgGridAttr');
            for (var ti = 0; ti < currentAttributeTitles.length; ti++) {
                var o = document.createElement('option');
                o.value = currentAttributeTitles[ti];
                o.textContent = currentAttributeTitles[ti];
                attrSel2.appendChild(o);
            }
            attrSel2.value = selAttr;   // restore previously selected attribute
            selAttr = attrSel2.value || currentAttributeTitles[0];
            svgGridSelAttr = selAttr;
            attrSel2.addEventListener('change', function() {
                var savedAttr = this.value;   // capture before DOM is cleared
                if (svgGridEventSource) { svgGridEventSource.close(); svgGridEventSource = null; }
                if (svgGridPerAreaEventSource) {
                    svgGridPerAreaEventSource.close();
                    svgGridPerAreaEventSource = null;
                }
                outputList.innerHTML = '';
                buildSvgGrid(savedAttr);      // pass saved value to next build
            });
            wrapper.appendChild(attrDiv);
        }

        var scrollDiv = document.createElement('div');
        scrollDiv.className = 'svg-grid-scroll';

        var table = document.createElement('table');
        table.className = 'svg-grid-table';

        // Header row
        var thead = table.createTHead();
        var hrow = thead.insertRow();
        var thCorner = document.createElement('th');
        thCorner.textContent = '';
        hrow.appendChild(thCorner);
        for (var ci = 0; ci < cols.length; ci++) {
            var th = document.createElement('th');
            th.textContent = cols[ci].label;
            hrow.appendChild(th);
        }

        var tbody = table.createTBody();

        // --- Device File row (top) ---
        var devPreviewUrl = '/device_preview/' + currentJobId;
        var devColTabMap = { l1: 'l1', l2: 'l2', l3: 'l3' };
        var devTr = tbody.insertRow();
        devTr.className = 'device-file-row';
        var devThCell = devTr.insertCell();
        devThCell.className = 'row-header';
        devThCell.innerHTML = '<a href="' + devPreviewUrl + '" target="_blank" '
            + 'style="text-decoration:none;color:inherit;">Device Tables</a>';
        var devTdByTabId = {};
        for (var dci = 0; dci < cols.length; dci++) {
            var dTd = devTr.insertCell();
            dTd.innerHTML = '<div class="cell-spinner">&#8987;</div>';
            devTdByTabId[devColTabMap[cols[dci].id]] = dTd;
        }

        // --- All Areas row ---
        var allTr = tbody.insertRow();
        var allHdr = allTr.insertCell();
        allHdr.className = 'row-header';
        allHdr.textContent = 'All Areas';
        for (var ci_all = 0; ci_all < cols.length; ci_all++) {
            var allTd = allTr.insertCell();
            var allCellId = null;
            if (cols[ci_all].id === 'l1') allCellId = 'l1_all';
            else if (cols[ci_all].id === 'l2') allCellId = 'l2_all';
            else if (cols[ci_all].id === 'l3') allCellId = 'l3_all';
            if (!allCellId) {
                allTd.innerHTML = '<div class="cell-na">N/A</div>';
            } else {
                allTd.setAttribute('data-cell', allCellId);
                allTd.innerHTML = '<div class="cell-spinner">&#8987;</div>';
                if (!svgGridCellMap[allCellId]) svgGridCellMap[allCellId] = [];
                svgGridCellMap[allCellId].push(allTd);
            }
        }

        // --- Per-area dropdown row ---
        // Only shown when the master has real (non-waypoint) areas.
        var perAreaTr = null;
        var perAreaCellsByCol = {};   // 'l1'|'l2'|'l3' -> td element
        if (currentAreas.length > 0) {
            perAreaTr = tbody.insertRow();
            perAreaTr.id = 'svgGridPerAreaRow';
            var perAreaHdr = perAreaTr.insertCell();
            perAreaHdr.className = 'row-header';
            // Build dropdown with a placeholder + every real area
            var sel = document.createElement('select');
            sel.id = 'svgGridAreaSelect';
            sel.style.cssText = 'max-width:160px;font-size:13px;padding:4px 6px;';
            var placeholder = document.createElement('option');
            placeholder.value = '';
            placeholder.textContent = '(select an area)';
            sel.appendChild(placeholder);
            for (var aii = 0; aii < currentAreas.length; aii++) {
                var opt = document.createElement('option');
                opt.value = currentAreas[aii];
                opt.textContent = currentAreas[aii];
                sel.appendChild(opt);
            }
            sel.value = '';
            perAreaHdr.appendChild(sel);

            for (var ci_pa = 0; ci_pa < cols.length; ci_pa++) {
                var paTd = perAreaTr.insertCell();
                paTd.innerHTML = '<div class="cell-na" style="color:#9aa5b1;">'
                    + 'Select an area<br>to generate</div>';
                perAreaCellsByCol[cols[ci_pa].id] = paTd;
            }

            sel.addEventListener('change', function() {
                var areaName = this.value || '';
                if (!areaName) {
                    // User cleared the selection -- reset to placeholder state
                    currentSelectedArea = '';
                    for (var k in perAreaCellsByCol) {
                        if (Object.prototype.hasOwnProperty.call(perAreaCellsByCol, k)) {
                            perAreaCellsByCol[k].innerHTML =
                                '<div class="cell-na" style="color:#9aa5b1;">'
                                + 'Select an area<br>to generate</div>';
                        }
                    }
                    if (svgGridPerAreaEventSource) {
                        svgGridPerAreaEventSource.close();
                        svgGridPerAreaEventSource = null;
                    }
                    // Re-enable the lightweight ZIP download (initial state).
                    setZipButtonLightweight();
                    return;
                }
                currentSelectedArea = areaName;
                requestAreaThumbnails(areaName, perAreaCellsByCol);
            });
        }

        scrollDiv.appendChild(table);
        wrapper.appendChild(scrollDiv);
        outputList.appendChild(wrapper);

        // --- Render Device Tables iframe thumbnails ---
        (function renderDeviceIframes() {
            Object.keys(devTdByTabId).forEach(function(tabId) {
                var tabUrl = '/device_preview/' + currentJobId + '#' + tabId;
                devTdByTabId[tabId].innerHTML = buildDeviceThumbHtml({ id: tabId }, tabUrl);
            });
        })();

        // --- Build INITIAL SSE URL: All Areas + AI Context only ---
        var sseUrl = '/svg_grid_stream/' + currentJobId + '?mode=initial';
        if (selAttr) sseUrl += '&attribute=' + encodeURIComponent(selAttr);

        var svgGridTotalEl = document.getElementById('svgGridTotal');
        if (svgGridTotalEl) {
            if (currentDeviceCount > 0) {
                // L2 All Areas runs in parallel with L1/L3 inside the
                // initial SSE pool, so the wall-clock estimate is the
                // max of the L1+L3 baseline and the L2 estimate (NOT
                // their sum).
                var initSec = _estimateSec(currentDeviceCount, _PERF_DIAGRAMS) || 0;
                var l2Sec = estimateSeconds('l2_all_areas',
                                            currentDeviceCount,
                                            currentAreas.length) || 0;
                var totalSec = Math.max(initSec, l2Sec);
                svgGridTotalEl.textContent = _fmtEstimate(totalSec);
                svgGridTotalEl.style.display = 'flex';
            } else {
                svgGridTotalEl.style.display = 'none';
            }
        }
        var svgGridStartTime = Date.now();

        // ZIP button: hidden while the initial generation is in flight;
        // the initial SSE 'done' handler below enables it with a lightweight
        // ZIP URL (Device Tables HTML + L1_all + L2_all + L3_all + AI Context
        // + .nsm), and a subsequent area pick swaps the URL to the area-scoped one.
        btnDownloadAll.classList.add('disabled');
        btnDownloadAll.style.display = 'none';

        svgGridEventSource = new EventSource(sseUrl);

        svgGridEventSource.onmessage = function(ev) {
            try {
                var msg = JSON.parse(ev.data);
                if (msg.done) {
                    svgGridEventSource.close();
                    svgGridEventSource = null;
                    var svgElapsed = (Date.now() - svgGridStartTime) / 1000;
                    if (svgGridTotalEl) {
                        svgGridTotalEl.textContent = 'Generated in ' + formatElapsed(svgElapsed);
                        svgGridTotalEl.style.display = 'flex';
                    }
                    // AI Context Download button
                    var aiRow = document.getElementById('fileActionsAiRow');
                    var aiSpinner = document.getElementById('fileActionsAiSpinner');
                    if (aiSpinner) aiSpinner.remove();
                    var aiLabelEl2 = aiRow ? aiRow.querySelector('.output-label') : null;
                    if (aiLabelEl2) aiLabelEl2.textContent = updatedMasters.length > 0 ? 'AI Context File (v' + updatedMasters.length + ')' : 'AI Context File';
                    if (msg.ai_filename) {
                        if (aiRow && !aiRow.querySelector('.output-dl')) {
                            var dlLink = document.createElement('a');
                            dlLink.className = 'output-dl';
                            dlLink.href = '/download/' + currentJobId + '/' + encodeURIComponent(msg.ai_filename);
                            dlLink.textContent = 'Download';
                            aiRow.appendChild(dlLink);
                            activateUpdateLlmButtons(dlLink.href, msg.ai_filename);
                        }
                    }
                    // Enable lightweight ZIP download now that the initial
                    // artifacts (Device Tables + L1_all + L2_all + L3_all
                    // + AI Context) are ready. If the user later picks an
                    // area, the URL is upgraded to the area-scoped ZIP in
                    // requestAreaThumbnails().
                    setZipButtonLightweight();
                    return;
                }
                renderCellFromSseMessage(msg, svgGridCellMap, svgGridSelAttr);
            } catch (e) {}
        };

        svgGridEventSource.onerror = function() {
            if (svgGridEventSource) { svgGridEventSource.close(); svgGridEventSource = null; }
        };
    }

    // Set the ZIP button to the lightweight (no area selected) download URL.
    // Used when the initial SSE completes and when the user clears their
    // area selection in the per-area dropdown.
    function setZipButtonLightweight() {
        if (!currentJobId) return;
        var href = '/download_all/' + currentJobId + '?scope=lightweight';
        if (svgGridSelAttr) href += '&attr=' + encodeURIComponent(svgGridSelAttr);
        btnDownloadAll.href = href;
        btnDownloadAll.classList.remove('disabled');
        btnDownloadAll.style.display = 'block';
    }

    // Map a cell_id to the L1/L2/L3 tabbed viewer URL (/diagram_preview/...)
    // for L1/L2/L3 cells, or null for everything else (caller falls back to
    // the per-SVG /preview/<job>/<filename> viewer).
    //
    // Cell -> scope/layer mapping (mirrors _resolve_layer_files on the server):
    //   l1_all                 -> scope=all,            layer=l1
    //   l2_all                 -> scope=all,            layer=l2
    //   l3_all                 -> scope=all,            layer=l3
    //   l1_per_area_<area>     -> scope=area:<area>,    layer=l1
    //   l2_area_<area>         -> scope=area:<area>,    layer=l2
    //   l3_per_area_<area>     -> scope=area:<area>,    layer=l3
    function buildDiagramViewerUrl(cellId, jobId) {
        if (!cellId || !jobId) return null;
        var scope = null;
        var layer = null;
        if (cellId === 'l1_all')      { scope = 'all'; layer = 'l1'; }
        else if (cellId === 'l2_all') { scope = 'all'; layer = 'l2'; }
        else if (cellId === 'l3_all') { scope = 'all'; layer = 'l3'; }
        else if (cellId.indexOf('l1_per_area_') === 0) {
            scope = 'area:' + cellId.substring('l1_per_area_'.length);
            layer = 'l1';
        } else if (cellId.indexOf('l3_per_area_') === 0) {
            scope = 'area:' + cellId.substring('l3_per_area_'.length);
            layer = 'l3';
        } else if (cellId.indexOf('l2_area_') === 0) {
            scope = 'area:' + cellId.substring('l2_area_'.length);
            layer = 'l2';
        } else {
            return null;
        }
        return '/diagram_preview/' + encodeURIComponent(jobId)
            + '?scope=' + encodeURIComponent(scope)
            + '#' + layer;
    }

    // Common SSE message renderer for a single grid cell update.
    function renderCellFromSseMessage(msg, cellMap, selAttr) {
        var cell_id = msg.cell;
        var tds = cellMap[cell_id];
        if (!tds) return;
        var inner = '';
        if (msg.error || !msg.files || msg.files.length === 0) {
            inner = '<div class="cell-error">&#10007;<br>Error</div>';
        } else {
            var firstFile = msg.files[0];
            // Prefer the L1/L2/L3 tabbed viewer for diagram cells so the
            // user can flip between layers in one window. Non-diagram cells
            // (or any unrecognised cell_id) fall back to the per-SVG
            // viewer.
            var viewerUrl = buildDiagramViewerUrl(cell_id, currentJobId);
            if (!viewerUrl) {
                viewerUrl = '/preview/' + currentJobId + '/' + encodeURIComponent(firstFile);
                if (msg.filter) viewerUrl += '?filter=' + encodeURIComponent(msg.filter);
            }
            // Always request the thumbnail variant (`?thumb=1`) so the
            // server can shrink width/height attributes that would
            // otherwise exceed Chromium's 16384-px image rasterizer
            // limit and render as a blank box. The full-size SVG remains
            // available via `?thumb=0`/no flag, used by the preview page.
            var fallbackUrl = '/svg_raw/' + currentJobId + '/' + encodeURIComponent(firstFile)
                + '?v=' + encodeURIComponent(selAttr || '_')
                + '&thumb=1';
            var thumbUrl = msg.svg_b64
                ? 'data:image/svg+xml;base64,' + msg.svg_b64
                : fallbackUrl;
            var dataFallback = msg.svg_b64
                ? ' data-fallback-src="' + escapeAttr(fallbackUrl) + '"'
                : '';
            inner = '<a class="svg-cell-inner" href="' + viewerUrl + '" target="_blank">'
                + '<img class="svg-thumb" src="' + thumbUrl + '"' + dataFallback + ' alt="' + escapeHtml(cell_id) + '" loading="eager" onerror="svgThumbRetry(this)">'
                + '<span class="svg-cell-label">' + escapeHtml(firstFile) + '</span>'
                + '</a>';
        }
        for (var ti = 0; ti < tds.length; ti++) {
            tds[ti].innerHTML = inner;
        }
    }

    // Request per-area thumbnails for the user-selected area.
    // Each call closes any previous per-area stream, repopulates the three
    // per-area cells with spinners, opens a new SSE in mode=area, and on
    // completion enables the ZIP button with area/attr query parameters.
    function requestAreaThumbnails(areaName, perAreaCellsByCol) {
        if (svgGridPerAreaEventSource) {
            svgGridPerAreaEventSource.close();
            svgGridPerAreaEventSource = null;
        }
        // Spinner placeholders + (re-)register cells in cellMap
        var l1cell = 'l1_per_area_' + areaName;
        var l2cell = 'l2_area_' + areaName;
        var l3cell = 'l3_per_area_' + areaName;
        svgGridCellMap[l1cell] = [perAreaCellsByCol.l1];
        svgGridCellMap[l2cell] = [perAreaCellsByCol.l2];
        svgGridCellMap[l3cell] = [perAreaCellsByCol.l3];
        for (var k in perAreaCellsByCol) {
            if (Object.prototype.hasOwnProperty.call(perAreaCellsByCol, k)) {
                perAreaCellsByCol[k].innerHTML = '<div class="cell-spinner">&#8987;</div>';
            }
        }

        // Keep the lightweight ZIP active while per-area generation runs;
        // the SSE 'done' handler upgrades the URL to the area-scoped ZIP.

        var sseUrl = '/svg_grid_stream/' + currentJobId
            + '?mode=area'
            + '&area=' + encodeURIComponent(areaName);
        if (svgGridSelAttr) sseUrl += '&attribute=' + encodeURIComponent(svgGridSelAttr);
        svgGridPerAreaStartTime = Date.now();

        svgGridPerAreaEventSource = new EventSource(sseUrl);
        svgGridPerAreaEventSource.onmessage = function(ev) {
            try {
                var msg = JSON.parse(ev.data);
                if (msg.done) {
                    svgGridPerAreaEventSource.close();
                    svgGridPerAreaEventSource = null;
                    // Upgrade ZIP button to the area-scoped download URL.
                    var zipHref = '/download_all/' + currentJobId
                        + '?area=' + encodeURIComponent(areaName)
                        + '&attr=' + encodeURIComponent(svgGridSelAttr || '');
                    btnDownloadAll.href = zipHref;
                    btnDownloadAll.classList.remove('disabled');
                    btnDownloadAll.style.display = 'block';
                    return;
                }
                renderCellFromSseMessage(msg, svgGridCellMap, svgGridSelAttr);
            } catch (e) {}
        };
        svgGridPerAreaEventSource.onerror = function() {
            if (svgGridPerAreaEventSource) {
                svgGridPerAreaEventSource.close();
                svgGridPerAreaEventSource = null;
            }
        };
    }

    // Build iframe thumbnail HTML for Device File row cell
    function buildDeviceThumbHtml(tab, previewUrl) {
        // Render a scaled-down iframe of the full device preview page.
        // pointer-events:none prevents iframe interaction; the wrapping <a> handles clicks.
        return '<a class="device-iframe-cell" href="' + previewUrl + '" target="_blank">'
            + '<iframe class="device-iframe-thumb" src="' + previewUrl + '"'
            + ' loading="eager" tabindex="-1" aria-hidden="true"></iframe>'
            + '</a>';
    }

    document.getElementById('formatToggle').addEventListener('change', function() {
        if (svgGridEventSource) { svgGridEventSource.close(); svgGridEventSource = null; }
        if (svgGridPerAreaEventSource) {
            svgGridPerAreaEventSource.close();
            svgGridPerAreaEventSource = null;
        }
        outputMode = this.checked ? 'pptx' : 'svg';
        document.getElementById('labelSvg').classList.toggle('active', !this.checked);
        document.getElementById('labelPptx').classList.toggle('active', this.checked);
        if (currentJobId) { buildSelectionList(); buildFileActionsSection(); }
    });

    function saveSession() {
        if (!currentJobId) return;
        try {
            sessionStorage.setItem('ns_session', JSON.stringify({
                jobId: currentJobId,
                basename: currentBasename,
                areas: currentAreas,
                deviceCount: currentDeviceCount,
                linkCount: currentLinkCount,
                masterVersion: masterVersion,
                updatedMasters: updatedMasters,
                isEmptyMaster: isEmptyMaster,
                attributeTitles: currentAttributeTitles
            }));
        } catch (e) {}
    }

    function clearSession() {
        try { sessionStorage.removeItem('ns_session'); } catch (e) {}
    }

    function startHeartbeat() {
        if (heartbeatTimer) clearInterval(heartbeatTimer);
        heartbeatTimer = setInterval(function() {
            if (!currentJobId) return;
            fetch('/heartbeat/' + currentJobId, { method: 'POST' }).catch(function() {});
        }, 30000);
        if (currentJobId) {
            fetch('/heartbeat/' + currentJobId, { method: 'POST' }).catch(function() {});
        }
    }

    function stopHeartbeat() {
        if (heartbeatTimer) { clearInterval(heartbeatTimer); heartbeatTimer = null; }
    }

    var estimateTotal = document.getElementById('estimateTotal');
    var estimateTotalTime = document.getElementById('estimateTotalTime');
    var estimateNote = document.getElementById('estimateNote'); // may be null

    var btnCreateMaster = document.getElementById('btnCreateMaster');
    btnCreateMaster.addEventListener('click', function(e) {
        e.stopPropagation();
        btnCreateMaster.disabled = true;
        btnCreateMaster.textContent = 'Creating...';
        fetch('/create_empty_master', { method: 'POST' })
            .then(function(res) {
                if (!res.ok) throw new Error('Failed to create');
                return res.blob();
            })
            .then(function(blob) {
                var file = new File([blob], '[MASTER]no_data.xlsx', {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                });
                btnCreateMaster.textContent = 'Create New Master File';
                btnCreateMaster.disabled = false;
                isEmptyMaster = true;
                handleFile(file);
            })
            .catch(function() {
                showError('Failed to create empty master file');
                btnCreateMaster.textContent = 'Create New Master File';
                btnCreateMaster.disabled = false;
            });
    });

    browseLink.addEventListener('click', function(e) { e.stopPropagation(); fileInput.click(); });
    dropzone.addEventListener('click', function() { fileInput.click(); });
    dropzone.addEventListener('dragover', function(e) { e.preventDefault(); dropzone.classList.add('dragover'); });
    dropzone.addEventListener('dragleave', function() { dropzone.classList.remove('dragover'); });
    dropzone.addEventListener('drop', function(e) {
        e.preventDefault(); dropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', function() {
        if (fileInput.files.length > 0) handleFile(fileInput.files[0]);
    });

    function handleFile(file) {
        uploadError.style.display = 'none';
        if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.nsm')) {
            showError('Only .xlsx and .nsm files are supported'); return;
        }
        if (!file.name.startsWith('[MASTER]')) { showError('Filename must start with [MASTER]'); return; }
        selectedFileEl.textContent = file.name + ' (' + formatSize(file.size) + ')';
        dropzone.classList.add('has-file');
        uploadAndAnalyze(file);
    }

    async function uploadAndAnalyze(file) {
        dropzone.classList.add('disabled');
        // Show estimated time only on re-analysis (currentDeviceCount known from previous init)
        var _aEstEl = document.getElementById('analyzingEstimate');
        if (_aEstEl) {
            _aEstEl.textContent = currentDeviceCount > 0
                ? _fmtEstimate(_estimateSec(currentDeviceCount, _PERF_MASTER)) : '';
        }
        analyzingIndicator.style.display = 'flex';
        selectionSection.style.display = 'none';
        generationTotal.style.display = 'none';

        try {
            var formData = new FormData();
            formData.append('file', file);
            var uploadRes = await fetch('/upload', { method: 'POST', body: formData });
            var uploadData = await uploadRes.json();
            if (!uploadRes.ok || uploadData.error) {
                showError(uploadData.error || 'Upload failed');
                resetDropzone();
                return;
            }
            currentJobId = uploadData.job_id;
            currentMasterFilename = uploadData.filename;
            currentBasename = file.name.replace('[MASTER]', '').replace('.xlsx', '').replace('.nsm', '');

            var res = await fetch('/generate_step/' + currentJobId + '/init', { method: 'POST' });
            var data = await res.json();
            currentAreas = (data.areas && data.areas.length > 0) ? data.areas : [];
            currentDeviceCount = data.device_count || 0;
            currentLinkCount = data.link_count || 0;
            currentAttributeTitles = data.attribute_titles || [];

            buildSelectionList();
            buildFileActionsSection();
            var summary = document.getElementById('analysisSummary');
            if (currentDeviceCount > 0 || currentLinkCount > 0) {
                summary.textContent = 'Detected: ' + currentDeviceCount + ' devices, '
                    + currentLinkCount + ' links, ' + currentAreas.length + ' areas';
                summary.style.display = 'block';
            } else {
                summary.style.display = 'none';
            }
            analyzingIndicator.style.display = 'none';
            uploadSection.style.display = 'none';
            updateSection.style.display = 'block';
            selectionSection.style.display = 'block';
            btnReset.style.display = 'block';
            saveSession();
            startHeartbeat();
            showLlmPromptArea();
        } catch (e) {
            showError('An error occurred: ' + e.message);
            resetDropzone();
        }
    }

    function interpolateTime(deviceCount, benchPoints) {
        var MIN_TIME = 5;
        if (deviceCount <= 0) return MIN_TIME;
        if (deviceCount <= benchPoints[0][0]) {
            var d0 = benchPoints[0][0], t0 = benchPoints[0][1];
            return Math.max(MIN_TIME, t0 + (t0 - benchPoints[1][1]) * (deviceCount - d0) / (benchPoints[1][0] - d0));
        }
        for (var i = 0; i < benchPoints.length - 1; i++) {
            if (deviceCount <= benchPoints[i + 1][0]) {
                var d0 = benchPoints[i][0], t0 = benchPoints[i][1];
                var d1 = benchPoints[i + 1][0], t1 = benchPoints[i + 1][1];
                return t0 + (t1 - t0) * (deviceCount - d0) / (d1 - d0);
            }
        }
        var last = benchPoints[benchPoints.length - 1];
        var prev = benchPoints[benchPoints.length - 2];
        var slope = (last[1] - prev[1]) / (last[0] - prev[0]);
        return last[1] + slope * (deviceCount - last[0]);
    }

    function estimateSeconds(type, deviceCount, numAreas) {
        var bench = {
            device_file:     [[13, 7],  [64, 19], [256, 64],  [1024, 314]],
            l1_diagram:      [[13, 5],  [64, 6],  [256, 29],  [1024, 390]],
            l2_per_area:     [[13, 7],  [64, 13], [256, 51],  [1024, 413]],
            // Empirical bench points for the L2 All Areas SVG.
            //   1024 devices on a 32x32 single-area master => ~17s end-to-end
            //   via the SSE pipeline (CLI 17s + ~2s subprocess overhead).
            // Smaller points are interpolated proportionally; replace with
            // real measurements as they become available.
            l2_all_areas:    [[13, 6],  [64, 8],  [256, 11],  [1024, 19]],
            l3_diagram:      [[13, 7],  [64, 10], [256, 56],  [1024, 863]],
            ai_context_file: [[13, 20], [64, 30], [256, 90],  [1024, 350]]
        };
        var points = (type === 'l2_diagram') ? bench.l2_per_area : bench[type];
        if (!points) return 5;
        return Math.max(5, interpolateTime(deviceCount, points));
    }

    function formatEstimate(seconds) {
        seconds = Math.round(seconds);
        if (seconds < 10) return '< 10 sec';
        if (seconds < 60) return '~' + (Math.round(seconds / 5) * 5) + ' sec';
        var min = Math.round(seconds / 60);
        if (seconds < 90) return '~1 min';
        if (seconds < 3600) return '~' + min + ' min';
        var hr = Math.floor(seconds / 3600);
        var rm = Math.round((seconds % 3600) / 60);
        return '~' + hr + 'h ' + rm + 'min';
    }

    function getSelectedAttribute() {
        var sel = document.getElementById('attributeSelect');
        return sel ? sel.value : '';
    }

    function buildSelectionList() {
        outputList.innerHTML = '';
        var attrContainer = document.getElementById('attributeContainer');
        attrContainer.innerHTML = '';
        attrContainer.style.display = 'none';
        var selectAllRow = document.querySelector('.select-all-row');

        // --- SVG / CSV mode: show SVG preview grid ---
        if (outputMode === 'svg') {
            if (selectAllRow) selectAllRow.style.display = 'none';
            if (!isEmptyMaster) {
                buildSvgGrid();
            } else {
                // Empty master: buildSvgGrid() (and its SSE stream) does not run,
                // so the SSE 'done' event will never fire to clear the AI Context
                // spinner. Generate it directly via the POST endpoint instead,
                // then replicate what the SSE 'done' handler does for the AI row.
                generateAiContextFixed().then(function() {
                    var aiRow = document.getElementById('fileActionsAiRow');
                    var aiSpinner = document.getElementById('fileActionsAiSpinner');
                    if (aiSpinner) aiSpinner.remove();
                    if (aiRow && !aiRow.querySelector('.output-dl')) {
                        var aiFilename = '[AI_Context]' + currentBasename + '.txt';
                        var dlLink = document.createElement('a');
                        dlLink.className = 'output-dl';
                        dlLink.href = '/download/' + currentJobId + '/' + encodeURIComponent(aiFilename);
                        dlLink.textContent = 'Download';
                        dlLink.addEventListener('click', function(e) { e.preventDefault(); e.stopPropagation(); window.location.href = this.href; });
                        aiRow.appendChild(dlLink);
                        activateUpdateLlmButtons(dlLink.href, aiFilename);
                    }
                });
            }
            updateSelectedCount();
            return;
        }

        // --- PPTX / XLSX mode: show full list without Preview buttons ---
        if (selectAllRow) selectAllRow.style.display = '';

        var numAreas = currentAreas.length || 1;
        var l1Est = estimateSeconds('l1_diagram', currentDeviceCount, numAreas);
        var l3Est = estimateSeconds('l3_diagram', currentDeviceCount, numAreas);

        if (currentAttributeTitles.length > 1 && !isEmptyMaster) {
            var attrRow = document.createElement('div');
            attrRow.className = 'attribute-selector';
            attrRow.innerHTML = '<div style="display:flex;flex-direction:column;align-items:flex-end;gap:2px;">'
                + '<div style="display:flex;align-items:center;gap:8px;">'
                + '<label for="attributeSelect">Attribute:</label>'
                + '<select id="attributeSelect"></select>'
                + '</div>'
                + '<span class="attr-note">Specifies device color pattern for L1 and L3 Diagrams</span>'
                + '</div>';
            attrContainer.appendChild(attrRow);
            attrContainer.style.display = 'block';
            var attrSelect = attrRow.querySelector('#attributeSelect');
            for (var a = 0; a < currentAttributeTitles.length; a++) {
                var opt = document.createElement('option');
                opt.value = currentAttributeTitles[a];
                opt.textContent = currentAttributeTitles[a];
                attrSelect.appendChild(opt);
            }
        }

        var items = [];
        if (updatedMasters.length > 0) {
            var latest = updatedMasters.length - 1;
            items.push({
                id: 'updated_master_' + latest, type: 'master',
                label: 'Updated Master File (v' + (latest + 1) + ')',
                filename: updatedMasters[latest], icon: 'MST',
                estSec: 0, generated: true
            });
        }

        if (!isEmptyMaster) {
            var isSvg = outputMode === 'svg';
            var diagExt = isSvg ? '.svg' : '.pptx';
            var devExt  = '.xlsx'; // CLI supports XLSX only for device file
            items.push(
                { id: 'device_file', type: 'device', label: 'Device Tables',
                  filename: '[DEVICE]' + currentBasename + devExt, icon: 'DEV',
                  estSec: estimateSeconds('device_file', currentDeviceCount, numAreas) },
                { id: 'l1_all_areas_tag', type: 'l1', label: 'L1 Diagram - All Areas + Tags',
                  filename: '[L1_DIAGRAM]AllAreasTag_' + currentBasename + diagExt, icon: 'L1',
                  subtype: 'all_areas_tag', estSec: l1Est },
                { id: 'l1_per_area_tag', type: 'l1', label: 'L1 Diagram - Per Area + Tags',
                  filename: '[L1_DIAGRAM]PerAreaTag_' + currentBasename + diagExt, icon: 'L1',
                  subtype: 'per_area_tag', estSec: l1Est }
            );
            for (var i = 0; i < currentAreas.length; i++) {
                items.push({
                    id: 'l2_' + currentAreas[i], type: 'l2',
                    label: 'L2 Diagram (' + currentAreas[i] + ')',
                    filename: '[L2_DIAGRAM]' + currentAreas[i] + '_' + currentBasename + diagExt,
                    icon: 'L2', area: currentAreas[i],
                    estSec: estimateSeconds('l2_diagram', currentDeviceCount, numAreas)
                });
            }
            items.push(
                { id: 'l3_all_areas', type: 'l3', label: 'L3 Diagram - All Areas',
                  filename: '[L3_DIAGRAM]AllAreas_' + currentBasename + diagExt, icon: 'L3',
                  subtype: 'all_areas', estSec: l3Est },
                { id: 'l3_per_area', type: 'l3', label: 'L3 Diagram - Per Area',
                  filename: '[L3_DIAGRAM]PerArea_' + currentBasename + diagExt, icon: 'L3',
                  subtype: 'per_area', estSec: l3Est }
            );
        }

        for (var j = 0; j < items.length; j++) {
            var it = items[j];
            var stepName = it.id.startsWith('l1_') ? 'l1_diagram'
                         : it.id.startsWith('l2_') ? 'l2_diagram'
                         : it.id.startsWith('l3_') ? 'l3_diagram'
                         : it.id;
            var row = document.createElement('label');
            row.className = 'output-item';

            if (it.isMasterDownload) {
                var masterRow = document.createElement('div');
                masterRow.className = 'output-item';
                var masterBtns = '';
                var isNsm = currentMasterFilename && currentMasterFilename.endsWith('.nsm');
                if (isNsm) {
                    masterBtns += '<button type="button" class="output-dl master-dl-xlsx">Download (.xlsx)</button>';
                    masterBtns += '<a class="output-dl master-dl-nsm" href="/download/' + currentJobId + '/' + encodeURIComponent(currentMasterFilename) + '">Download (.nsm)</a>';
                } else {
                    masterBtns += '<a class="output-dl master-dl-xlsx" href="/download/' + currentJobId + '/' + encodeURIComponent(currentMasterFilename) + '">Download (.xlsx)</a>';
                }
                masterRow.innerHTML = '<span class="output-icon master_dl">MST</span>'
                    + '<span class="output-info">'
                    + '<span class="output-label">Master File</span>'
                    + '<span class="output-filename">' + escapeHtml(currentMasterFilename) + '</span>'
                    + '</span>'
                    + '<span style="margin-left:auto;display:flex;gap:6px;flex-shrink:0;align-items:center;">'
                    + masterBtns
                    + '</span>';
                outputList.appendChild(masterRow);
                continue;
            }

            if (it.generated) {
                var encFn = encodeURIComponent(it.filename);
                var dlUrl = '/download/' + currentJobId + '/' + encFn;
                var pvUrl = '/preview/' + currentJobId + '/' + encFn;
                var btnHtml = '';
                // Preview link is shown only in PPTX mode (SVG mode uses SVG viewer via Preview button)
                if (outputMode === 'pptx' && it.type !== 'master' && it.type !== 'ai') {
                    btnHtml += '<a class="output-preview" href="' + pvUrl + '" target="_blank" rel="noopener noreferrer">Preview</a>';
                }
                btnHtml += '<a class="output-download" href="' + dlUrl + '">Download</a>';
                row.innerHTML = '<span class="output-icon ' + it.type + '">' + it.icon + '</span>'
                    + '<span class="output-info">'
                    + '<span class="output-label">' + escapeHtml(it.label) + '</span>'
                    + '<span class="output-filename">' + escapeHtml(it.filename) + '</span>'
                    + '</span>'
                    + '<span style="margin-left:auto;display:flex;gap:6px;flex-shrink:0;">'
                    + btnHtml
                    + '</span>';
                if (it.type === 'master') {
                    row.style.flexWrap = 'wrap';
                    var masterNote = document.createElement('div');
                    masterNote.style.cssText = 'flex-basis:100%;font-size:11px;color:#2563eb;font-weight:600;margin-top:2px;text-align:right;';
                    masterNote.textContent = 'Don\u0027t forget to download — this file will not be kept on the server.';
                    row.appendChild(masterNote);
                }
                outputList.appendChild(row);
                continue;
            }

            row.innerHTML = '<input type="checkbox" checked data-id="' + escapeAttr(it.id) + '"'
                + (it.area ? ' data-area="' + escapeAttr(it.area) + '"' : '')
                + (it.subtype ? ' data-subtype="' + escapeAttr(it.subtype) + '"' : '')
                + ' data-step="' + escapeAttr(stepName) + '"'
                + ' data-est="' + it.estSec + '">'
                + '<span class="output-icon ' + it.type + '">' + it.icon + '</span>'
                + '<span class="output-info">'
                + '<span class="output-label">' + escapeHtml(it.label) + '</span>'
                + '<span class="output-filename">' + escapeHtml(it.filename) + '</span>'
                + '</span>'
                + '<span class="output-estimate" id="est-' + escapeAttr(it.id) + '"><span class="inline-dot"></span><span class="est-text">&#9202; ' + formatEstimate(it.estSec) + '</span></span>'
                + '<button type="button" class="output-gen" data-idx="' + j + '">Generate</button>';
            outputList.appendChild(row);
        }

        selectAll.checked = true;
        updateSelectedCount();
        outputList.addEventListener('change', function() {
            var boxes = outputList.querySelectorAll('input[type="checkbox"]:not(:disabled)');
            var allChecked = boxes.length > 0;
            for (var k = 0; k < boxes.length; k++) { if (!boxes[k].checked) { allChecked = false; break; } }
            selectAll.checked = allChecked;
            updateSelectedCount();
        });
    }

    selectAll.addEventListener('change', function() {
        var boxes = outputList.querySelectorAll('input[type="checkbox"]:not(:disabled)');
        for (var k = 0; k < boxes.length; k++) boxes[k].checked = selectAll.checked;
        // Also toggle AI Context checkbox in fileActionsSection
        var faSection = document.getElementById('fileActionsSection');
        if (faSection) {
            var faBoxes = faSection.querySelectorAll('input[type="checkbox"]:not(:disabled)');
            for (var k2 = 0; k2 < faBoxes.length; k2++) faBoxes[k2].checked = selectAll.checked;
        }
        updateSelectedCount();
    });

    function updateSelectedCount() {
        var allBoxes = Array.from(outputList.querySelectorAll('input[type="checkbox"]'));
        var enabledBoxes = Array.from(outputList.querySelectorAll('input[type="checkbox"]:not(:disabled)'));
        // Include fileActionsSection checkboxes (AI Context in PPTX mode)
        var faSection = document.getElementById('fileActionsSection');
        if (faSection) {
            allBoxes = allBoxes.concat(Array.from(faSection.querySelectorAll('input[type="checkbox"]')));
            enabledBoxes = enabledBoxes.concat(Array.from(faSection.querySelectorAll('input[type="checkbox"]:not(:disabled)')));
        }
        var generated = allBoxes.length - enabledBoxes.length;
        var checked = 0;
        var totalEst = 0;
        for (var k = 0; k < enabledBoxes.length; k++) {
            if (enabledBoxes[k].checked) {
                checked++;
                totalEst += parseFloat(enabledBoxes[k].getAttribute('data-est') || 0);
            }
        }
        var statusText = checked + ' selected';
        if (generated > 0) statusText += ' (' + generated + ' generated)';
        selectedCount.textContent = statusText;
        btnGenerate.disabled = (checked === 0);
        btnGenerate.style.display = '';
        if (checked > 0 && currentDeviceCount > 0) {
            estimateTotalTime.textContent = formatEstimate(totalEst);
            estimateTotal.style.display = 'flex';
            if (estimateNote) estimateNote.style.display = 'block';
        } else {
            estimateTotal.style.display = 'none';
            if (estimateNote) estimateNote.style.display = 'none';
        }
    }

    var PARALLEL_LIMIT = {{PARALLEL_LIMIT}};

    function runParallel(tasks, limit) {
        return new Promise(function(resolve) {
            var index = 0;
            var active = 0;
            var results = new Array(tasks.length);
            var done = 0;

            function next() {
                while (active < limit && index < tasks.length) {
                    (function(i) {
                        active++;
                        index++;
                        tasks[i]().then(function(r) {
                            results[i] = r;
                        }).catch(function() {
                            results[i] = null;
                        }).then(function() {
                            active--;
                            done++;
                            if (done === tasks.length) {
                                resolve(results);
                            } else {
                                next();
                            }
                        });
                    })(index);
                }
            }
            if (tasks.length === 0) { resolve(results); return; }
            next();
        });
    }

    async function generateItems(selected) {
        if (selected.length === 0) return;

        btnGenerate.disabled = true;
        var genBtns = outputList.querySelectorAll('.output-gen');
        for (var g = 0; g < genBtns.length; g++) genBtns[g].disabled = true;
        // Also disable AI Context checkbox in fileActionsSection during generation
        var faSection = document.getElementById('fileActionsSection');
        var faCheckboxes = faSection ? Array.from(faSection.querySelectorAll('input[type="checkbox"]:not(:disabled)')) : [];
        for (var fc = 0; fc < faCheckboxes.length; fc++) faCheckboxes[fc].disabled = true;
        selectionSection.style.pointerEvents = 'none';
        selectionSection.style.opacity = '0.6';
        generationTotal.style.display = 'none';

        for (var i = 0; i < selected.length; i++) {
            setInlineStatus(selected[i].id, 'running');
        }

        var totalStart = Date.now();

        var selectedAttr = getSelectedAttribute();
        var tasks = selected.map(function(sel) {
            return function() {
                var url = '/generate_step/' + currentJobId + '/' + sel.step;
                var params = [];
                if (sel.area) params.push('area=' + encodeURIComponent(sel.area));
                if (sel.subtype) params.push('type=' + encodeURIComponent(sel.subtype));
                if (selectedAttr && (sel.step === 'l1_diagram' || sel.step === 'l3_diagram')) {
                    params.push('attribute=' + encodeURIComponent(selectedAttr));
                }
                if (outputMode === 'svg' && (sel.step === 'l1_diagram' || sel.step === 'l2_diagram' || sel.step === 'l3_diagram')) {
                    params.push('format=svg');
                }
                if (params.length > 0) url += '?' + params.join('&');
                return runStep(sel.id, url);
            };
        });

        await runParallel(tasks, PARALLEL_LIMIT);

        var totalElapsed = (Date.now() - totalStart) / 1000;
        generationTotal.textContent = 'Total: ' + formatElapsed(totalElapsed);
        generationTotal.style.display = 'flex';

        await addDownloadButtons();
        selectionSection.style.pointerEvents = '';
        selectionSection.style.opacity = '';
        // Re-enable AI Context checkbox only if addDownloadButtons did not permanently disable it
        // (permanent disable = Download button added to the parent row)
        for (var fc2 = 0; fc2 < faCheckboxes.length; fc2++) {
            var fcRow = faCheckboxes[fc2].closest('.output-item');
            if (fcRow && !fcRow.querySelector('.output-dl')) {
                faCheckboxes[fc2].disabled = false;
            }
        }
        for (var g2 = 0; g2 < genBtns.length; g2++) {
            var parentItem = genBtns[g2].closest('.output-item');
            var cb = parentItem ? parentItem.querySelector('input[type="checkbox"]') : null;
            genBtns[g2].disabled = !!(cb && cb.disabled);
        }
    }

    btnGenerate.addEventListener('click', async function() {
        var selected = [];
        var boxes = Array.from(outputList.querySelectorAll('input[type="checkbox"]:checked:not(:disabled)'));
        // Also include fileActionsSection checkboxes (AI Context in PPTX mode)
        var faSection = document.getElementById('fileActionsSection');
        if (faSection) {
            boxes = boxes.concat(Array.from(faSection.querySelectorAll('input[type="checkbox"]:checked:not(:disabled)')));
        }
        for (var k = 0; k < boxes.length; k++) {
            selected.push({
                id: boxes[k].getAttribute('data-id'),
                step: boxes[k].getAttribute('data-step'),
                area: boxes[k].getAttribute('data-area') || '',
                subtype: boxes[k].getAttribute('data-subtype') || ''
            });
        }
        await generateItems(selected);
    });

    function showLlmPromptArea() {
        llmPromptArea.style.display = 'block';
        btnCopyOpenLlm.style.display = '';
        btnCopyOnly.style.display = '';
        btnDlAiContext.style.display = '';
    }

    async function ensureAiContextAndDo(actionFn) {
        if (_currentAiDlUrl) {
            await actionFn();
            return;
        }
        // In PPTX mode, AI context checkbox is in fileActionsSection; in PPTX output list otherwise
        var aiItem = document.querySelector('#fileActionsSection input[data-step="ai_context_file"]')
                  || outputList.querySelector('input[data-step="ai_context_file"]');
        if (!aiItem) return;
        llmGenStatus.textContent = 'Generating AI Context...';
        llmGenStatus.style.display = 'inline';
        llmGenStatus.style.color = '#0984e3';
        btnCopyOpenLlm.disabled = true;
        btnCopyOnly.disabled = true;
        btnDlAiContext.disabled = true;
        var selected = [{
            id: aiItem.getAttribute('data-id'),
            step: 'ai_context_file',
            area: '',
            subtype: ''
        }];
        await generateItems(selected);
        llmGenStatus.style.display = 'none';
        btnCopyOpenLlm.disabled = false;
        btnCopyOnly.disabled = false;
        btnDlAiContext.disabled = false;
        if (_currentAiDlUrl) {
            await actionFn();
        }
    }

    function buildLlmClipboard(aiContextText) {
        var promptInput = document.getElementById('llmPromptInput');
        var userPrompt = promptInput ? promptInput.value.trim() : '';
        if (!userPrompt) return aiContextText;
        var suffix = '\n\n\'\'\'\nUser Request\n\'\'\'\n'
            + userPrompt + '\n\n'
            + '* Output ONLY the Network Sketcher commands needed to fulfill the request above.\n'
            + '* Do NOT include any explanation, commentary, or markdown formatting.\n'
            + '* Output one command per line so the entire output can be directly pasted into the Update Master command input.\n'
            + '* You MUST wrap all commands in a single code block (```). This is critical so the user can copy them in one click.\n';
        return aiContextText + suffix;
    }

    var _currentAiDlUrl = '';
    var _currentAiFname = '';
    var btnCopyOpenLlm = document.getElementById('btnCopyOpenLlm');
    var btnCopyOnly = document.getElementById('btnCopyOnly');
    var btnDlAiContext = document.getElementById('btnDlAiContext');
    var llmGenStatus = document.getElementById('llmGenStatus');
    var llmPromptArea = document.getElementById('llmPromptArea');

    function activateUpdateLlmButtons(dlUrl, fname) {
        _currentAiDlUrl = dlUrl;
        _currentAiFname = fname;
        btnCopyOpenLlm.style.display = '';
        btnCopyOnly.style.display = '';
        btnDlAiContext.style.display = '';
        btnDlAiContext.href = dlUrl;
        btnDlAiContext.setAttribute('data-fname', fname);
        llmGenStatus.style.display = 'none';
        llmPromptArea.style.display = 'block';
    }

    async function llmCopyAction(openUrl) {
        if (!_currentAiDlUrl) return;
        var btn = openUrl ? btnCopyOpenLlm : btnCopyOnly;
        var origText = btn.textContent;
        try {
            var r = await fetch(_currentAiDlUrl);
            var text = await r.text();
            var clipboard = buildLlmClipboard(text);
            await navigator.clipboard.writeText(clipboard);
            btn.textContent = 'Copied!';
            setTimeout(function() { btn.textContent = origText; }, 2000);
            if (openUrl) window.open('{{AI_CONTEXT_BTN_URL}}', '_blank', 'noopener,noreferrer');
        } catch (err) {
            btn.textContent = 'Copy failed';
            setTimeout(function() { btn.textContent = origText; }, 2000);
        }
    }

    btnCopyOpenLlm.addEventListener('click', async function(e) {
        e.preventDefault(); e.stopPropagation();
        await ensureAiContextAndDo(async function() { await llmCopyAction(true); });
    });
    btnCopyOnly.addEventListener('click', async function(e) {
        e.preventDefault(); e.stopPropagation();
        await ensureAiContextAndDo(async function() { await llmCopyAction(false); });
    });
    btnDlAiContext.addEventListener('click', async function(e) {
        e.preventDefault(); e.stopPropagation();
        await ensureAiContextAndDo(async function() {
            var r = await fetch(_currentAiDlUrl);
            var text = await r.text();
            var content = buildLlmClipboard(text);
            var blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url; a.download = _currentAiFname || 'AI_Context.txt';
            document.body.appendChild(a); a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        });
    });

    outputList.addEventListener('click', async function(e) {
        var csvBtn = e.target.closest('.output-preview-csv');
        if (csvBtn) {
            e.preventDefault();
            e.stopPropagation();
            window.open('/device_preview/' + currentJobId, '_blank');
            return;
        }

        var svgBtn = e.target.closest('.output-preview-svg');
        if (svgBtn && !svgBtn.disabled) {
            e.preventDefault();
            e.stopPropagation();
            svgBtn.disabled = true;
            svgBtn.textContent = 'Generating...';
            var selectedAttr = getSelectedAttribute();
            var svgStep = svgBtn.getAttribute('data-step') || 'l1_preview_svg';
            var svgArea = svgBtn.getAttribute('data-area') || '';
            var url = '/generate_step/' + currentJobId + '/' + svgStep;
            var params = [];
            if (selectedAttr) params.push('attribute=' + encodeURIComponent(selectedAttr));
            if (svgArea) params.push('area=' + encodeURIComponent(svgArea));
            if (params.length > 0) url += '?' + params.join('&');
            try {
                var startTime = Date.now();
                var res = await fetch(url, { method: 'POST' });
                var data = await res.json();
                var elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
                if (data.success) {
                    svgBtn.textContent = 'Preview (.svg) ' + elapsed + 's';
                    svgBtn.disabled = false;

                    // Per-area steps: server returns first_svg + filter_prefix directly
                    if ((svgStep === 'l1_per_area_tag_preview_svg' || svgStep === 'l3_per_area_preview_svg')
                            && data.first_svg && data.filter_prefix) {
                        var previewUrl = '/preview/' + currentJobId + '/' + encodeURIComponent(data.first_svg)
                                        + '?filter=' + encodeURIComponent(data.filter_prefix);
                        window.open(previewUrl, '_blank');
                    } else {
                        // Single-SVG steps: find the file by prefix
                        var svgExactPrefix = null;
                        var svgPrefix = '[L1_DIAGRAM]';
                        if (svgStep === 'l3_preview_svg') {
                            svgPrefix = '[L3_DIAGRAM]';
                            svgExactPrefix = '[L3_DIAGRAM]AllAreas_';
                        } else if (svgStep === 'l2_preview_svg') {
                            svgPrefix = '[L2_DIAGRAM]';
                            svgExactPrefix = svgArea ? '[L2_DIAGRAM]' + svgArea + '_' : null;
                        } else if (svgStep === 'l1_preview_svg') {
                            svgPrefix = '[L1_DIAGRAM]';
                            svgExactPrefix = '[L1_DIAGRAM]AllAreasTag_';
                        }
                        var filesRes = await fetch('/files/' + currentJobId);
                        var filesData = await filesRes.json();
                        if (filesData.files) {
                            for (var fi = 0; fi < filesData.files.length; fi++) {
                                var fn = filesData.files[fi].name;
                                if (!fn.toLowerCase().endsWith('.svg') || !fn.startsWith(svgPrefix)) continue;
                                if (svgExactPrefix && !fn.startsWith(svgExactPrefix)) continue;
                                window.open('/preview/' + currentJobId + '/' + encodeURIComponent(fn), '_blank');
                                break;
                            }
                        }
                    }
                } else {
                    svgBtn.textContent = 'Failed';
                    svgBtn.disabled = false;
                    setTimeout(function() { svgBtn.textContent = 'Preview (.svg)'; }, 3000);
                }
            } catch (err) {
                svgBtn.textContent = 'Error';
                svgBtn.disabled = false;
                setTimeout(function() { svgBtn.textContent = 'Preview (.svg)'; }, 3000);
            }
            return;
        }
        var btn = e.target.closest('.output-gen');
        if (!btn || btn.disabled) return;
        e.preventDefault();
        e.stopPropagation();
        var item = btn.closest('.output-item');
        if (!item) return;
        var cb = item.querySelector('input[type="checkbox"]');
        if (!cb || cb.disabled) return;
        var selected = [{
            id: cb.getAttribute('data-id'),
            step: cb.getAttribute('data-step'),
            area: cb.getAttribute('data-area') || '',
            subtype: cb.getAttribute('data-subtype') || ''
        }];
        await generateItems(selected);
    });

    var subtypeLabels = {
        all_areas_tag: 'All Areas + Tags',
        all_areas: 'All Areas',
        per_area_tag: 'Per Area + Tags',
        per_area: 'Per Area'
    };
    function getStepLabel(sel) {
        if (sel.step === 'device_file') return 'Generating device file...';
        if (sel.step === 'l1_diagram') return 'Generating L1 diagram (' + (subtypeLabels[sel.subtype] || '') + ')...';
        if (sel.step === 'l2_diagram') return 'Generating L2 diagram (' + sel.area + ')...';
        if (sel.step === 'l3_diagram') return 'Generating L3 diagram (' + (subtypeLabels[sel.subtype] || '') + ')...';
        if (sel.step === 'ai_context_file') return 'Generating AI context file...';
        return 'Processing...';
    }

    async function runStep(stepId, url) {
        setInlineStatus(stepId, 'running');
        var startTime = Date.now();
        try {
            var res = await fetch(url, { method: 'POST' });
            var data = await res.json();
            var elapsed = (Date.now() - startTime) / 1000;
            if (data.success) {
                setInlineStatus(stepId, 'done', elapsed);
            } else {
                setInlineStatus(stepId, 'error', elapsed);
            }
        } catch (e) {
            var elapsed = (Date.now() - startTime) / 1000;
            setInlineStatus(stepId, 'error', elapsed);
        }
    }

    function formatElapsed(sec) {
        if (sec < 60) return sec.toFixed(1) + 's';
        var m = Math.floor(sec / 60);
        var s = (sec % 60).toFixed(0);
        return m + 'm ' + s + 's';
    }

    function setInlineStatus(id, status, elapsedSec) {
        var est = document.getElementById('est-' + id);
        if (!est) return;
        var dot = est.querySelector('.inline-dot');
        var text = est.querySelector('.est-text');
        if (!dot || !text) return;

        dot.className = 'inline-dot ' + status;
        est.classList.remove('est-done', 'est-error');

        if (status === 'running') {
            text.innerHTML = 'Generating...';
            est.style.color = '';
            est.style.background = '';
            est.style.borderColor = '';
        } else if (status === 'done') {
            text.innerHTML = formatElapsed(elapsedSec);
            est.classList.add('est-done');
        } else if (status === 'error') {
            text.innerHTML = 'Failed ' + formatElapsed(elapsedSec);
            est.classList.add('est-error');
        }
    }

    function buildFileActionsSection() {
        var section = document.getElementById('fileActionsSection');
        if (!section) return;
        section.innerHTML = '';

        // --- Master File row ---
        var masterRow = document.createElement('div');
        masterRow.className = 'output-item';
        masterRow.style.flexWrap = 'wrap';
        var displayFilename, displayLabel, masterBtns;
        if (updatedMasters.length > 0) {
            var latestIdx = updatedMasters.length - 1;
            displayFilename = updatedMasters[latestIdx];
            displayLabel = 'Master File (v' + (latestIdx + 1) + ')';
            var isNsmUpdated = displayFilename && displayFilename.endsWith('.nsm');
            if (isNsmUpdated) {
                masterBtns = '<button type="button" class="output-dl master-dl-xlsx">Download (.xlsx)</button>'
                           + '<a class="output-dl master-dl-nsm" href="/download/' + currentJobId + '/' + encodeURIComponent(displayFilename) + '">Download (.nsm)</a>';
            } else {
                masterBtns = '<a class="output-dl master-dl-xlsx" href="/download/' + currentJobId + '/' + encodeURIComponent(displayFilename) + '">Download (.xlsx)</a>';
            }
        } else {
            displayFilename = currentMasterFilename;
            displayLabel = 'Master File';
            var isNsm = currentMasterFilename && currentMasterFilename.endsWith('.nsm');
            if (isNsm) {
                masterBtns = '<button type="button" class="output-dl master-dl-xlsx">Download (.xlsx)</button>'
                           + '<a class="output-dl master-dl-nsm" href="/download/' + currentJobId + '/' + encodeURIComponent(currentMasterFilename) + '">Download (.nsm)</a>';
            } else {
                masterBtns = '<a class="output-dl master-dl-xlsx" href="/download/' + currentJobId + '/' + encodeURIComponent(currentMasterFilename) + '">Download (.xlsx)</a>';
            }
        }
        masterRow.innerHTML = '<span class="output-icon master_dl">MST</span>'
            + '<span class="output-info">'
            + '<span class="output-label">' + escapeHtml(displayLabel) + '</span>'
            + '<span class="output-filename">' + escapeHtml(displayFilename) + '</span>'
            + '</span>'
            + '<span style="margin-left:auto;display:flex;gap:6px;flex-shrink:0;align-items:center;">'
            + masterBtns
            + '</span>';
        if (updatedMasters.length > 0) {
            var masterNote = document.createElement('div');
            masterNote.style.cssText = 'flex-basis:100%;font-size:11px;color:#2563eb;font-weight:600;margin-top:2px;text-align:right;';
            masterNote.textContent = 'Don\u0027t forget to download \u2014 this file will not be kept on the server.';
            masterRow.appendChild(masterNote);
        }
        section.appendChild(masterRow);

        // --- AI Context File row ---
        var aiFilename = '[AI_Context]' + currentBasename + '.txt';
        var aiLabel = updatedMasters.length > 0 ? 'AI Context File (v' + updatedMasters.length + ')' : 'AI Context File';
        var aiRow = document.createElement(outputMode === 'svg' ? 'div' : 'label');
        aiRow.className = 'output-item';
        aiRow.id = 'fileActionsAiRow';
        var aiInner = '';
        if (outputMode === 'svg') {
            // In SVG mode, AI context is auto-generated with the SVG grid.
            // Show a spinner immediately so the row is never blank during generation.
            aiInner += '<span class="output-icon ai">AI</span>'
                     + '<span class="output-info">'
                     + '<span class="output-label">' + escapeHtml(aiLabel) + '</span>'
                     + '<span class="output-filename">' + escapeHtml(aiFilename) + '</span>'
                     + '</span>'
                     + '<span class="output-estimate" id="fileActionsAiSpinner">'
                     + '<span class="inline-dot running"></span>Generating...</span>';
        } else {
            // In PPTX mode, show checkbox so AI context is part of "Generate Selected" flow
            var aiEstSec = estimateSeconds('ai_context_file', currentDeviceCount, currentAreas.length || 1);
            aiInner += '<input type="checkbox" checked data-id="ai_context_file" data-step="ai_context_file" data-est="' + aiEstSec + '">'
                     + '<span class="output-icon ai">AI</span>'
                     + '<span class="output-info">'
                     + '<span class="output-label">' + escapeHtml(aiLabel) + '</span>'
                     + '<span class="output-filename">' + escapeHtml(aiFilename) + '</span>'
                     + '</span>'
                     + '<span class="output-estimate" id="est-ai_context_file"><span class="inline-dot"></span>'
                     + '<span class="est-text">&#9202; ' + formatEstimate(aiEstSec) + '</span></span>';
        }
        aiRow.innerHTML = aiInner;
        section.appendChild(aiRow);

        // NOTE: click event delegation is registered ONCE at module init
        // (see fileActionsSection.addEventListener near outputList init).
        // Registering here would cause listener accumulation on each call,
        // leading to duplicate xlsx downloads after CLI command runs.

        section.style.display = 'block';
        // Reflect AI Context checkbox in the selected count immediately
        if (typeof updateSelectedCount === 'function') updateSelectedCount();

        // ZIP download button: in SVG mode the per-area dropdown drives it
        // (see requestAreaThumbnails); keep it hidden until the user picks
        // an area. In PPTX mode the ZIP is whole-job and is enabled by the
        // PPTX generation completion handlers below.
        if (outputMode === 'svg') {
            btnDownloadAll.classList.add('disabled');
            btnDownloadAll.style.display = 'none';
        }
    }

    async function generateAiContextFixed() {
        if (!currentJobId) return;
        // Update AI label to match current master revision before generating
        var aiRow = document.getElementById('fileActionsAiRow');
        var aiLabelEl = aiRow ? aiRow.querySelector('.output-label') : null;
        if (aiLabelEl) aiLabelEl.textContent = updatedMasters.length > 0 ? 'AI Context File (v' + updatedMasters.length + ')' : 'AI Context File';
        await runStep('ai_context_file', '/generate_step/' + currentJobId + '/ai_context_file');
        await addDownloadButtons();
    }

    async function addDownloadButtons() {
        try {
            var res = await fetch('/files/' + currentJobId);
            var data = await res.json();
            if (data.files && data.files.length > 0) {
                var fileMap = {};
                for (var i = 0; i < data.files.length; i++) {
                    fileMap[data.files[i].name] = data.files[i];
                }
                var items = outputList.querySelectorAll('.output-item');
                for (var j = 0; j < items.length; j++) {
                    var cb = items[j].querySelector('input[type="checkbox"]');
                    if (!cb) continue;
                    if (cb.disabled) continue;
                    var fname = items[j].querySelector('.output-filename');
                    if (!fname) continue;
                    var expectedName = fname.textContent;
                    var matched = fileMap[expectedName];
                    if (matched) {
                        var est = items[j].querySelector('.output-estimate');
                        if (est) {
                            var btnWrap = document.createElement('span');
                            btnWrap.style.display = 'flex';
                            btnWrap.style.gap = '6px';
                            btnWrap.style.marginLeft = 'auto';
                            btnWrap.style.flexShrink = '0';
                            btnWrap.style.alignItems = 'center';
                            if (est.classList.contains('est-done') || est.classList.contains('est-error')) {
                                est.style.marginLeft = '0';
                                btnWrap.appendChild(est);
                            } else {
                                est.remove();
                            }
                            if (matched.name.startsWith('[AI_Context]')) {
                                activateUpdateLlmButtons('/download/' + currentJobId + '/' + encodeURIComponent(matched.name), matched.name);
                            }
                            if (!matched.name.startsWith('[AI_Context]') && (matched.name.toLowerCase().endsWith('.pptx') || matched.name.toLowerCase().endsWith('.xlsx') || matched.name.toLowerCase().endsWith('.txt'))) {
                                var pv = document.createElement('a');
                                pv.className = 'output-preview';
                                pv.href = '/preview/' + currentJobId + '/' + encodeURIComponent(matched.name);
                                pv.target = '_blank';
                                pv.rel = 'noopener noreferrer';
                                pv.textContent = 'Preview';
                                pv.addEventListener('click', function(e) { e.stopPropagation(); });
                                btnWrap.appendChild(pv);
                            }
                            var dl = document.createElement('a');
                            dl.className = 'output-dl';
                            var dlHref = '/download/' + currentJobId + '/' + encodeURIComponent(matched.name);
                            dl.href = dlHref;
                            dl.textContent = 'Download';
                            dl.addEventListener('click', function(e) { e.preventDefault(); e.stopPropagation(); window.location.href = this.href; });
                            btnWrap.appendChild(dl);
                            items[j].appendChild(btnWrap);
                        }
                        var sizeEl = document.createElement('span');
                        sizeEl.className = 'output-filename';
                        sizeEl.style.color = 'var(--success)';
                        sizeEl.textContent = matched.size_human;
                        fname.parentNode.appendChild(sizeEl);
                        cb.disabled = true;
                        cb.checked = false;
                        var genBtn = items[j].querySelector('.output-gen');
                        if (genBtn) genBtn.style.display = 'none';
                    }
                }
                // Update AI Context row in fileActionsSection if file exists
                var aiFilename = '[AI_Context]' + currentBasename + '.txt';
                if (fileMap[aiFilename]) {
                    var aiRow = document.getElementById('fileActionsAiRow');
                    if (aiRow && !aiRow.querySelector('.output-dl')) {
                        activateUpdateLlmButtons('/download/' + currentJobId + '/' + encodeURIComponent(aiFilename), aiFilename);
                        var dlLink = document.createElement('a');
                        dlLink.className = 'output-dl';
                        dlLink.href = '/download/' + currentJobId + '/' + encodeURIComponent(aiFilename);
                        dlLink.textContent = 'Download';
                        dlLink.addEventListener('click', function(e) { e.preventDefault(); e.stopPropagation(); window.location.href = this.href; });
                        aiRow.appendChild(dlLink);
                        // Disable checkbox so it shows as generated
                        var aiCb = aiRow.querySelector('input[type="checkbox"]');
                        if (aiCb) { aiCb.disabled = true; aiCb.checked = false; }
                    }
                }
                if (outputMode === 'svg') {
                    // SVG mode uses area-scoped ZIP; the per-area dropdown
                    // controls visibility/enabled state of the button.
                    btnDownloadAll.classList.add('disabled');
                    btnDownloadAll.style.display = 'none';
                } else {
                    btnDownloadAll.href = '/download_all/' + currentJobId;
                    btnDownloadAll.classList.remove('disabled');
                    btnDownloadAll.style.display = 'block';
                }
                btnReset.style.display = 'block';
                updateSelectedCount();
            }
        } catch (e) {
            showError('Failed to retrieve file list');
        }
    }

    btnReset.addEventListener('click', function() { resetAll(); });

    // --- fileActionsSection click handler (register ONCE at module init) ---
    // buildFileActionsSection() is called multiple times (on upload, CLI run,
    // session restore, etc.). If the listener were registered inside that
    // function, it would accumulate and cause duplicate xlsx downloads.
    var fileActionsSectionEl = document.getElementById('fileActionsSection');
    if (fileActionsSectionEl) {
        fileActionsSectionEl.addEventListener('click', async function(ev) {
            // NSM download link
            var nsmBtn = ev.target.closest('.master-dl-nsm');
            if (nsmBtn) {
                ev.stopPropagation();
                window.location.href = nsmBtn.href;
                ev.preventDefault();
                return;
            }
            // Master XLSX download / convert
            var xlsxBtn = ev.target.closest('.master-dl-xlsx');
            if (xlsxBtn && currentJobId) {
                ev.preventDefault();
                ev.stopPropagation();
                if (xlsxBtn.tagName === 'A') {
                    window.location.href = xlsxBtn.href;
                    return;
                }
                xlsxBtn.disabled = true;
                xlsxBtn.textContent = 'Converting...';
                try {
                    var resp = await fetch('/export_nsm_to_xlsx/' + currentJobId, { method: 'POST' });
                    if (resp.ok) {
                        var blob = await resp.blob();
                        var url = window.URL.createObjectURL(blob);
                        var a = document.createElement('a');
                        a.href = url;
                        // 最新のバージョン番号付きマスタ名を採用（例: [MASTER]foo_1.nsm -> _1.xlsx）
                        var baseName = (updatedMasters.length > 0)
                            ? updatedMasters[updatedMasters.length - 1]
                            : currentMasterFilename;
                        a.download = baseName.replace(/\.nsm$/i, '.xlsx');
                        document.body.appendChild(a);
                        a.click();
                        setTimeout(function() { window.URL.revokeObjectURL(url); a.remove(); }, 1000);
                    }
                } catch (e) {}
                xlsxBtn.disabled = false;
                xlsxBtn.textContent = 'Download (.xlsx)';
                return;
            }
        });
    }

    outputList.addEventListener('click', async function(ev) {
        var nsmBtn = ev.target.closest('.master-dl-nsm');
        if (nsmBtn) {
            ev.stopPropagation();
            window.location.href = nsmBtn.href;
            ev.preventDefault();
            return;
        }
        var xlsxBtn = ev.target.closest('.master-dl-xlsx');
        if (!xlsxBtn || !currentJobId) return;
        ev.preventDefault();
        ev.stopPropagation();
        if (xlsxBtn.tagName === 'A') {
            window.location.href = xlsxBtn.href;
            return;
        }
        xlsxBtn.disabled = true;
        xlsxBtn.textContent = 'Converting...';
        try {
            var resp = await fetch('/export_nsm_to_xlsx/' + currentJobId, { method: 'POST' });
            if (resp.ok) {
                var blob = await resp.blob();
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                // 最新のバージョン番号付きマスタ名を採用（例: [MASTER]foo_1.nsm -> _1.xlsx）
                var baseName = (updatedMasters.length > 0)
                    ? updatedMasters[updatedMasters.length - 1]
                    : currentMasterFilename;
                a.download = baseName.replace(/\.nsm$/i, '.xlsx');
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
            } else {
                var errData = await resp.json();
                showError(errData.error || 'Export failed');
            }
        } catch (e) {
            showError('Export failed: ' + e.message);
        } finally {
            xlsxBtn.disabled = false;
            xlsxBtn.textContent = 'Download (.xlsx)';
        }
    });

    btnRun.addEventListener('click', async function() {
        var cmds = cmdInput.value.trim();
        if (!cmds || !currentJobId) return;
        btnRun.disabled = true;
        runProgress.style.display = 'flex';
        // Show estimated re-analysis time based on current device count
        var _runEst = _fmtEstimate(_estimateSec(currentDeviceCount, _PERF_MASTER));
        runProgressText.textContent = 'Executing commands...' + (_runEst ? '  ' + _runEst : '');
        runStatus.style.display = 'none';
        runResults.style.display = 'none';

        var nextVersion = masterVersion + 1;
        var cmdStartTime = Date.now();

        try {
            var resp = await fetch('/run_commands/' + currentJobId, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ commands: cmds, version: nextVersion })
            });
            var data = await resp.json();

            var html = '';
            if (data.results) {
                // All result rows are hidden initially inside runResultsOverflow
                var totalCount = data.results.length;
                html += '<div id="runResultsOverflow" style="display:none">';
                for (var i = 0; i < data.results.length; i++) {
                    var r = data.results[i];
                    var cls = r.skipped ? 'cmd-skip' : (r.success ? 'cmd-ok' : 'cmd-err');
                    var prefix = r.skipped ? '(skipped)' : (r.success ? '&#10003;' : '&#10007;');
                    html += '<div class="' + cls + '">' + prefix + ' ' + escapeHtml(r.command) + '</div>';
                    if (r.output && r.output.trim()) {
                        html += '<div class="cmd-output" id="cmdOut' + i + '">' + escapeHtml(r.output.trim()) + '</div>';
                    }
                }
                html += '</div>';
                html += '<span class="view-log-link" id="expandResultsLink">&#9654; Show '
                     + totalCount + ' results</span> ';
                html += '<span class="view-log-link" id="viewLogLink">&#9654; View Log</span>';
                html += '<button type="button" class="copy-log-btn" id="copyLogBtn" title="Copy the View Log to clipboard">Copy Log</button>';
            }
            runResults.innerHTML = html;
            runResults.style.display = 'block';
            var expandLink = document.getElementById('expandResultsLink');
            var overflowDiv = document.getElementById('runResultsOverflow');
            if (expandLink && overflowDiv) {
                var _totalCount = data.results.length;
                expandLink.addEventListener('click', function() {
                    var expanding = overflowDiv.style.display === 'none';
                    overflowDiv.style.display = expanding ? '' : 'none';
                    expandLink.innerHTML = expanding
                        ? '&#9660; Hide results'
                        : '&#9654; Show ' + _totalCount + ' results';
                });
            }
            var logLink = document.getElementById('viewLogLink');
            if (logLink) {
                logLink.addEventListener('click', function() {
                    // Ensure results are visible before showing log output
                    if (overflowDiv && overflowDiv.style.display === 'none') {
                        overflowDiv.style.display = '';
                        if (expandLink) expandLink.innerHTML = '&#9660; Hide results';
                    }
                    var outputs = runResults.querySelectorAll('.cmd-output');
                    var anyVisible = false;
                    for (var k = 0; k < outputs.length; k++) {
                        if (outputs[k].classList.contains('visible')) { anyVisible = true; break; }
                    }
                    for (var k = 0; k < outputs.length; k++) {
                        if (anyVisible) outputs[k].classList.remove('visible');
                        else outputs[k].classList.add('visible');
                    }
                    logLink.innerHTML = anyVisible ? '&#9654; View Log' : '&#9660; Hide Log';
                });
            }
            // Copy Log button: assemble the full run log (command + output per
            // result) as plain text and copy it to the clipboard, regardless of
            // whether the log is currently expanded in the UI.
            var copyLogBtn = document.getElementById('copyLogBtn');
            if (copyLogBtn && data.results) {
                var _logResults = data.results;
                copyLogBtn.addEventListener('click', function() {
                    var lines = [];
                    for (var li = 0; li < _logResults.length; li++) {
                        var rr = _logResults[li];
                        var mark = rr.skipped ? '(skipped)' : (rr.success ? '[OK]' : '[ERROR]');
                        lines.push(mark + ' ' + (rr.command || ''));
                        if (rr.output && rr.output.trim()) {
                            lines.push(rr.output.replace(/\s+$/, ''));
                        }
                    }
                    var logText = lines.join('\n');
                    var done = function() { _cmdRefFeedback('View Log copied'); };
                    if (navigator.clipboard && navigator.clipboard.writeText) {
                        navigator.clipboard.writeText(logText).then(done).catch(function() {
                            _cmdRefCopyFallback(logText, done);
                        });
                    } else {
                        _cmdRefCopyFallback(logText, done);
                    }
                });
            }

            var cmdElapsed = (Date.now() - cmdStartTime) / 1000;
            var cmdElapsedStr = cmdElapsed < 60
                ? cmdElapsed.toFixed(1) + 's'
                : Math.floor(cmdElapsed / 60) + 'm ' + Math.floor(cmdElapsed % 60) + 's';

            if (data.success) {
                runStatus.innerHTML = '<span style="color:var(--success)">All commands executed successfully.</span>'
                    + '<span style="color:var(--text-secondary);margin-left:12px;font-size:0.93em">' + cmdElapsedStr + '</span>';
            } else {
                runStatus.innerHTML = '<span style="color:#e94560">' + (data.errors ? data.errors.length : 0) + ' command(s) failed.</span>'
                    + '<span style="color:var(--text-secondary);margin-left:12px;font-size:0.93em">' + cmdElapsedStr + '</span>';
            }
            runStatus.style.display = 'block';

            if (data.updated_master) {
                masterVersion = nextVersion;
                updatedMasters.push(data.updated_master);
                currentBasename = data.updated_master.replace('[MASTER]', '').replace('.xlsx', '');
                isEmptyMaster = false;

                runProgressText.textContent = 'Re-analyzing master file...';
                selectionSection.style.opacity = '0.4';
                selectionSection.style.pointerEvents = 'none';
                _currentAiDlUrl = '';
                _currentAiFname = '';

                var res = await fetch('/generate_step/' + currentJobId + '/init', { method: 'POST' });
                var initData = await res.json();
                currentAreas = (initData.areas && initData.areas.length > 0) ? initData.areas : [];
                currentDeviceCount = initData.device_count || 0;
                currentLinkCount = initData.link_count || 0;
                currentAttributeTitles = initData.attribute_titles || [];

                runProgress.style.display = 'none';
                runProgressText.textContent = 'Executing commands...';

                buildSelectionList();
                buildFileActionsSection();
                // In PPTX mode, auto-regenerate AI context file after master update.
                // In SVG mode, the SSE 'done' event from buildSvgGrid() handles this.
                if (outputMode !== 'svg') {
                    generateAiContextFixed();
                }
                var summary = document.getElementById('analysisSummary');
                if (currentDeviceCount > 0 || currentLinkCount > 0) {
                    summary.textContent = 'Detected: ' + currentDeviceCount + ' devices, '
                        + currentLinkCount + ' links, ' + currentAreas.length + ' areas';
                    summary.style.display = 'block';
                }

                selectionSection.style.pointerEvents = '';
                selectionSection.style.opacity = '';
                btnGenerate.disabled = false;
                btnGenerate.style.display = '';
                if (outputMode === 'svg') {
                    // SVG mode uses area-scoped ZIP; the per-area dropdown
                    // controls visibility/enabled state of the button.
                    btnDownloadAll.classList.add('disabled');
                    btnDownloadAll.style.display = 'none';
                } else {
                    btnDownloadAll.href = '/download_all/' + currentJobId;
                    btnDownloadAll.classList.remove('disabled');
                    btnDownloadAll.style.display = 'block';
                }
                btnReset.style.display = 'block';
                generationTotal.style.display = 'none';
                saveSession();
                showLlmPromptArea();
            } else {
                runProgress.style.display = 'none';
            }
            cmdInput.value = '';

        } catch (e) {
            runProgress.style.display = 'none';
            runProgressText.textContent = 'Executing commands...';
            runStatus.innerHTML = '<span style="color:#e94560">Error: ' + escapeHtml(e.message) + '</span>';
            runStatus.style.display = 'block';
        }
        btnRun.disabled = false;
    });

    function resetDropzone() {
        dropzone.classList.remove('has-file', 'disabled');
        analyzingIndicator.style.display = 'none';
        selectedFileEl.textContent = '';
        fileInput.value = '';
    }

    function resetAll() {
        stopHeartbeat();
        clearSession();
        currentJobId = null; currentAreas = []; currentBasename = ''; currentMasterFilename = '';
        currentDeviceCount = 0; currentLinkCount = 0;
        masterVersion = 0;
        updatedMasters = [];
        isEmptyMaster = false;
        resetDropzone();
        uploadError.style.display = 'none';
        uploadSection.style.display = 'block';
        updateSection.style.display = 'none';
        cmdInput.value = '';
        runStatus.style.display = 'none';
        runProgress.style.display = 'none';
        runResults.style.display = 'none';
        selectionSection.style.display = 'none';
        selectionSection.style.pointerEvents = '';
        selectionSection.style.opacity = '';
        generationTotal.style.display = 'none';
        var _svgGridTotalReset = document.getElementById('svgGridTotal');
        if (_svgGridTotalReset) _svgGridTotalReset.style.display = 'none';
        estimateTotal.style.display = 'none';
        if (estimateNote) estimateNote.style.display = 'none';
        btnGenerate.style.display = '';
        btnGenerate.disabled = false;
        btnDownloadAll.classList.remove('disabled');
        btnDownloadAll.style.display = 'none';
        btnReset.style.display = 'none';
        selectAll.disabled = false;
        outputList.innerHTML = '';
        _currentAiDlUrl = '';
        _currentAiFname = '';
        btnCopyOpenLlm.style.display = 'none';
        btnCopyOnly.style.display = 'none';
        btnDlAiContext.style.display = 'none';
        llmGenStatus.style.display = 'none';
        llmPromptArea.style.display = 'none';
        var promptInput = document.getElementById('llmPromptInput');
        if (promptInput) promptInput.value = '';
    }

    function showError(msg) { uploadError.textContent = msg; uploadError.style.display = 'block'; }

    function formatSize(bytes) {
        var units = ['B', 'KB', 'MB', 'GB'];
        var i = 0;
        while (bytes >= 1024 && i < units.length - 1) { bytes /= 1024; i++; }
        return bytes.toFixed(1) + ' ' + units[i];
    }

    function escapeHtml(text) {
        var d = document.createElement('div'); d.appendChild(document.createTextNode(text)); return d.innerHTML;
    }
    function escapeAttr(text) {
        return text.replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    }
    function svgThumbRetry(img) {
        // If currently showing a data URI, fall back to /svg_raw/ URL immediately.
        if (img.src.startsWith('data:')) {
            var fallback = img.getAttribute('data-fallback-src');
            if (fallback) {
                img.removeAttribute('data-fallback-src');
                img.src = fallback;
            }
            return;
        }
        // URL-based retry with increasing delays.
        var retries = parseInt(img.getAttribute('data-retries') || '0');
        if (retries < 3) {
            img.setAttribute('data-retries', String(retries + 1));
            var originalSrc = img.src.replace(/&retry=\d+$/, '');
            setTimeout(function() {
                img.src = originalSrc + '&retry=' + (retries + 1);
            }, 1500 * (retries + 1));
        }
    }

    function restoreFailed() {
        if (window._restoreTimeout) { clearTimeout(window._restoreTimeout); window._restoreTimeout = null; }
        clearSession();
        var ri = document.getElementById('restoreIndicator');
        if (ri) ri.style.display = 'none';
        var hideStyle = document.getElementById('hideUploadStyle');
        if (hideStyle) hideStyle.remove();
        uploadSection.style.display = '';
    }

    (async function tryRestore() {
        var raw;
        try { raw = sessionStorage.getItem('ns_session'); } catch (e) { return; }
        if (!raw) { restoreFailed(); return; }
        var sess;
        try { sess = JSON.parse(raw); } catch (e) { restoreFailed(); return; }
        if (!sess.jobId) { restoreFailed(); return; }

        try {
            var res = await fetch('/restore/' + sess.jobId);
            if (!res.ok) { restoreFailed(); return; }
            var data = await res.json();
            if (!data.valid) { restoreFailed(); return; }

            currentJobId = sess.jobId;
            currentBasename = data.basename || sess.basename || '';
            currentMasterFilename = data.master_filename || '';
            masterVersion = sess.masterVersion || 0;
            updatedMasters = data.updated_masters || sess.updatedMasters || [];
            isEmptyMaster = sess.isEmptyMaster || false;
            currentAreas = (data.areas && data.areas.length > 0) ? data.areas : [];
            currentDeviceCount = data.device_count || 0;
            currentLinkCount = data.link_count || 0;
            currentAttributeTitles = data.attribute_titles || sess.attributeTitles || [];

            buildSelectionList();
            buildFileActionsSection();

            var summary = document.getElementById('analysisSummary');
            if (currentDeviceCount > 0 || currentLinkCount > 0) {
                summary.textContent = 'Detected: ' + currentDeviceCount + ' devices, '
                    + currentLinkCount + ' links, ' + currentAreas.length + ' areas';
                summary.style.display = 'block';
            }

            if (outputMode === 'svg') {
                // SVG mode uses area-scoped ZIP; keep the button hidden until
                // the user picks an area from the per-area dropdown.
                btnDownloadAll.classList.add('disabled');
                btnDownloadAll.style.display = 'none';
            } else {
                btnDownloadAll.href = '/download_all/' + currentJobId;
                btnDownloadAll.classList.remove('disabled');
                btnDownloadAll.style.display = 'block';
            }

            if (data.generated_files && data.generated_files.length > 0) {
                await addDownloadButtons();
            }

            if (window._restoreTimeout) { clearTimeout(window._restoreTimeout); window._restoreTimeout = null; }
            var ri = document.getElementById('restoreIndicator');
            if (ri) ri.style.display = 'none';
            var hideStyle = document.getElementById('hideUploadStyle');
            if (hideStyle) hideStyle.remove();
            uploadSection.style.display = 'none';
            updateSection.style.display = 'block';
            selectionSection.style.display = 'block';
            btnReset.style.display = 'block';
            analyzingIndicator.style.display = 'none';

            startHeartbeat();
            saveSession();
            if (!_currentAiDlUrl) showLlmPromptArea();
        } catch (e) {
            restoreFailed();
        }
    })();
})();
</script>
</body>
</html>
'''


CERT_FILE = _resolve_path(_cfg['ssl_cert_path'])
KEY_FILE = _resolve_path(_cfg['ssl_key_path'])
CERTS_DIR = CERT_FILE.parent


def _get_local_ips():
    """Return a list of (interface_description, ip_address) for local NICs."""
    ips = []
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            addr = info[4][0]
            if addr not in ('127.0.0.1', '0.0.0.0') and addr not in [i[1] for i in ips]:
                ips.append((f'IPv4 {addr}', addr))
    except socket.gaierror:
        pass
    try:
        result = subprocess.run(
            ['powershell', '-Command',
             "Get-NetIPAddress -AddressFamily IPv4 | "
             "Where-Object { $_.IPAddress -ne '127.0.0.1' -and $_.PrefixOrigin -ne 'WellKnown' } | "
             "Select-Object -Property InterfaceAlias, IPAddress | "
             "ForEach-Object { $_.InterfaceAlias + '|' + $_.IPAddress }"],
            capture_output=True, text=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0,
        )
        if result.returncode == 0:
            seen = {i[1] for i in ips}
            for line in result.stdout.strip().splitlines():
                parts = line.strip().split('|', 1)
                if len(parts) == 2 and parts[1] not in seen:
                    ips.append((parts[0].strip(), parts[1].strip()))
                    seen.add(parts[1].strip())
    except Exception:
        pass
    return ips


def _generate_self_signed_cert(cn, san_list):
    """Generate a self-signed certificate and save to Certs/."""
    from cryptography import x509
    from cryptography.x509.oid import NameOID
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    import ipaddress

    key = rsa.generate_private_key(public_exponent=65537, key_size=_cfg['ssl_key_size'])

    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, cn),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, 'Network Sketcher'),
    ])

    san_entries = []
    for entry in san_list:
        try:
            ip = ipaddress.ip_address(entry)
            san_entries.append(x509.IPAddress(ip))
        except ValueError:
            san_entries.append(x509.DNSName(entry))

    now = datetime.datetime.now(datetime.timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(now + datetime.timedelta(days=_cfg['ssl_cert_validity_days']))
        .add_extension(x509.SubjectAlternativeName(san_entries), critical=False)
        .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
        .sign(key, hashes.SHA256())
    )

    CERTS_DIR.mkdir(parents=True, exist_ok=True)

    with open(KEY_FILE, 'wb') as f:
        f.write(key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        ))

    with open(CERT_FILE, 'wb') as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))

    print(f'  Certificate saved to {CERT_FILE}')
    print(f'  Private key saved to {KEY_FILE}')
    return str(CERT_FILE), str(KEY_FILE)


def ensure_certificates():
    """Check for existing certs or auto-generate self-signed ones from config.
    Returns (cert_path, key_path, bind_host, display_host).
    """
    cfg_host = _cfg['host']
    cfg_fqdn = _cfg.get('fqdn')

    if cfg_host == '0.0.0.0':
        bind_host = '0.0.0.0'
        display_host = 'localhost'
    elif cfg_host in ('localhost', '127.0.0.1'):
        bind_host = cfg_host
        display_host = 'localhost'
    else:
        # Specific IP: bind only to that interface
        bind_host = cfg_host
        display_host = cfg_host

    if CERT_FILE.is_file() and KEY_FILE.is_file():
        print(f'  Using existing certificate: {CERT_FILE}')
        return str(CERT_FILE), str(KEY_FILE), bind_host, display_host

    print()
    print('  SSL certificate not found. Auto-generating self-signed certificate...')

    if cfg_fqdn:
        cn = cfg_fqdn
        san_list = [cfg_fqdn, '127.0.0.1', 'localhost']
        display_host = cfg_fqdn
    elif cfg_host in ('0.0.0.0', 'localhost', '127.0.0.1'):
        cn = 'localhost'
        san_list = ['localhost', '127.0.0.1']
    else:
        cn = cfg_host
        san_list = [cfg_host, '127.0.0.1', 'localhost']

    print(f'  CN={cn}, SAN={san_list}')
    print()
    cert_path, key_path = _generate_self_signed_cert(cn, san_list)
    print()
    return cert_path, key_path, bind_host, display_host


def create_ssl_context(cert_path, key_path):
    """Create a hardened TLS SSL context using settings from config."""
    _tls_versions = {
        'TLSv1_2': ssl.TLSVersion.TLSv1_2,
        'TLSv1_3': ssl.TLSVersion.TLSv1_3,
    }
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.minimum_version = _tls_versions[_cfg['ssl_min_version']]
    ctx.maximum_version = _tls_versions[_cfg['ssl_max_version']]
    ctx.set_ciphers(_cfg['ssl_ciphers'])
    ctx.options |= ssl.OP_SINGLE_ECDH_USE | ssl.OP_CIPHER_SERVER_PREFERENCE
    ctx.load_cert_chain(cert_path, key_path)
    return ctx


if __name__ == '__main__':
    print()
    print('=' * 56)
    print('  Network Sketcher Online Service')
    print('=' * 56)

    cert_path, key_path, bind_host, display_host = ensure_certificates()
    ctx = create_ssl_context(cert_path, key_path)

    port = _cfg['port']
    print()
    print(f'  URL: https://{display_host}:{port}')
    print('=' * 56)
    print()

    logging.getLogger('werkzeug').setLevel(logging.ERROR)

    app.run(host=bind_host, port=port, debug=False, ssl_context=ctx, threaded=True, use_reloader=False)
