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
    make_response, abort
)

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
NS_DIR = PROJECT_DIR / 'network-sketcher_offline'
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
SUBPROCESS_TIMEOUT = _cfg['subprocess_timeout']

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
        except Exception as e:
            logger.warning('Cleanup error: %s', e)


_cleanup_thread = threading.Thread(target=_cleanup_loop, daemon=True)
_cleanup_thread.start()


def validate_master_filename(filename):
    if not filename:
        return False, 'Filename is empty'
    if not filename.endswith('.xlsx'):
        return False, 'Only .xlsx files are supported'
    if not filename.startswith('[MASTER]'):
        return False, 'Filename must start with [MASTER]'
    return True, ''


def sanitize_job_id(job_id):
    if not re.match(r'^[a-f0-9\-]{36}$', job_id):
        return None
    return job_id


def find_master_file(work_dir):
    for f in os.listdir(work_dir):
        if f.startswith('[MASTER]') and f.endswith('.xlsx'):
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
    cmd = [sys.executable, str(NS_DIR / 'network_sketcher.py')] + args
    logger.debug('Running: %s', ' '.join(cmd))

    try:
        kwargs = {
            'capture_output': True,
            'text': True,
            'cwd': str(NS_DIR),
            'timeout': SUBPROCESS_TIMEOUT,
        }
        if sys.platform == 'win32':
            kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(cmd, **kwargs)
        logger.debug('Exit code: %d', result.returncode)
        if result.stdout:
            logger.debug('stdout: %s', result.stdout[:500])
        if result.stderr:
            logger.warning('stderr: %s', result.stderr[:500])
        return result
    except subprocess.TimeoutExpired:
        logger.error('Command timed out')
        return None
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

    basename = master_filename.replace('[MASTER]', '').replace('.xlsx', '')

    masters = sorted([
        f for f in os.listdir(str(work_dir))
        if f.startswith('[MASTER]') and f.endswith('.xlsx') and f != master_filename
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
    filename = file.filename
    valid, msg = validate_master_filename(filename)
    if not valid:
        return jsonify({'error': msg}), 400

    job_id = str(uuid.uuid4())
    work_dir = UPLOAD_DIR / job_id
    work_dir.mkdir(parents=True, exist_ok=True)

    master_path = work_dir / filename
    file.save(str(master_path))
    set_active_master(str(work_dir), filename)

    touch_heartbeat(job_id)
    request._ns_upload_filename = filename
    logger.info('Uploaded: %s -> %s (job: %s)', filename, master_path, job_id)
    return jsonify({'job_id': job_id, 'filename': filename})


def run_ns_command_isolated(args, work_dir, master_filename):
    """Run NS CLI with an isolated copy of the master file.

    Each parallel task gets its own subdirectory with a private master
    copy so that concurrent processes never contend on the same file.
    Generated outputs are moved back to work_dir after completion.
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
            dest = work_dir / f.name
            try:
                shutil.move(str(f), str(dest))
            except Exception:
                logger.warning('Could not move %s to %s', f, dest)

        return result
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

    cmd_list_path = NS_DIR / 'ns_extensions_cmd_list.txt'
    if cmd_list_path.is_file():
        with open(str(cmd_list_path), 'r', encoding='utf-8') as f:
            content += f.read()

    basename = master_filename.replace('[MASTER]', '').replace('.xlsx', '')
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
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            future_map = {
                executor.submit(run_ns_command, cmd): key
                for key, cmd in init_commands.items()
            }
            for future in concurrent.futures.as_completed(future_map):
                init_results[future_map[future]] = future.result()

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
        result = run_ns_command_isolated(cmd_args, work_dir, master_filename)
        if result and '[ERROR]' not in (result.stdout or ''):
            return jsonify({'success': True})
        msg = ''
        if result:
            msg = (result.stdout or '') + (result.stderr or '')
        return jsonify({'success': False, 'message': msg})

    elif step == 'l2_diagram':
        area = request.args.get('area', '')
        if not area:
            return jsonify({'success': False, 'message': 'Area not specified'})
        result = run_ns_command_isolated([
            'export', 'l2_diagram',
            '--master', master_path,
            '--area', area,
        ], work_dir, master_filename)
        if result and '[ERROR]' not in (result.stdout or ''):
            return jsonify({'success': True})
        msg = ''
        if result:
            msg = (result.stdout or '') + (result.stderr or '')
        return jsonify({'success': False, 'message': msg})

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
        result = run_ns_command_isolated(cmd_args, work_dir, master_filename)
        if result and '[ERROR]' not in (result.stdout or ''):
            return jsonify({'success': True})
        msg = ''
        if result:
            msg = (result.stdout or '') + (result.stderr or '')
        return jsonify({'success': False, 'message': msg})

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

    ALLOWED_VERBS = {'add', 'delete', 'rename', 'show'}
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
        new_master_name = f'[MASTER]{base_no_ver}_{version}.xlsx'
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

    FAILURE_MARKERS = ('[ERROR]', 'Input must start with', '[WARNING] No ')
    results = []
    errors = []
    for line in lines:
        try:
            parts = shlex.split(line)
        except ValueError:
            parts = line.split()
        if not parts:
            continue
        if parts[0] not in ALLOWED_VERBS:
            results.append({'command': line, 'success': True, 'skipped': True, 'output': f'Skipped (only add/delete/rename/show are supported)'})
            continue

        cmd_args = parts + ['--master', target_master_path]
        result = run_ns_command(cmd_args)
        stdout = (result.stdout or '') if result else ''
        stderr = (result.stderr or '') if result else ''
        ok = result is not None and not any(m in stdout for m in FAILURE_MARKERS)
        results.append({'command': line, 'success': ok, 'output': stdout + stderr})
        if not ok:
            errors.append(line)

    if has_mutation and new_master_name:
        set_active_master(str(work_dir), new_master_name)

    return jsonify({
        'success': len(errors) == 0,
        'results': results,
        'errors': errors,
        'updated_master': new_master_name,
    })


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
    if not filepath.is_file() or not (lower_name.endswith('.pptx') or lower_name.endswith('.xlsx') or lower_name.endswith('.txt')):
        abort(404)

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
    job_id = sanitize_job_id(job_id)
    if not job_id:
        abort(400)

    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        abort(404)

    master_filename = get_active_master(str(work_dir)) or ''

    zip_name = 'network_sketcher_output.zip'
    zip_path = work_dir / zip_name
    with zipfile.ZipFile(str(zip_path), 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(os.listdir(work_dir)):
            if f == master_filename or f == zip_name or '__TMP__' in f:
                continue
            fpath = os.path.join(work_dir, f)
            if os.path.isfile(fpath):
                zf.write(fpath, f)

    request._ns_download_filename = 'network_sketcher_output.zip'
    return send_file(
        str(zip_path),
        as_attachment=True,
        download_name='network_sketcher_output.zip',
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
.output-list { display: flex; flex-direction: column; gap: 8px; }
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
.output-dl {
    padding: 5px 14px; font-size: 12px; font-weight: 600; color: var(--primary);
    background: white; border: 1.5px solid var(--primary); border-radius: 6px;
    cursor: pointer; transition: all 0.2s; text-decoration: none; white-space: nowrap;
    margin-left: auto; flex-shrink: 0;
}
.output-dl:hover { background: var(--primary); color: white; }
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
    display: none; width: 100%; margin-top: 4px; padding: 8px 10px;
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
.update-desc { font-size: 13px; color: var(--text-secondary); margin-bottom: 12px; line-height: 1.5; }
.cli-command-area {
    width: 100%; padding: 8px 10px; border: 1.5px solid #00b894; border-radius: 8px;
    background: #f0faf6; box-sizing: border-box;
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
.cmd-output {
    display: none; margin: 4px 0 8px 18px; padding: 8px 10px; background: #fff;
    border: 1px solid #e0e0e0; border-radius: 6px; white-space: pre-wrap;
    word-break: break-all; font-size: 11px; color: #444; line-height: 1.5;
}
.cmd-output.visible { display: block; }
.output-icon.master { background: #fdcb6e; color: #2d3436; }
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
                Drag &amp; drop a [MASTER]*.xlsx file here<br>
                or <strong id="browseLink">click to browse</strong>
            </div>
            <div class="dropzone-file" id="selectedFile"></div>
        </div>
        <input type="file" id="fileInput" accept=".xlsx">
        <div class="error-msg" id="uploadError"></div>
        <div class="analyzing" id="analyzingIndicator" style="display:none;">
            <div class="spinner"></div> Uploading and analyzing master file...
        </div>
    </div>

    <!-- Update Master Section -->
    <div class="card" id="updateSection" style="display:none;">
        <h2>&#9998; Update Master <span class="help-icon" onclick="event.stopPropagation();toggleHelp('update_master')">?</span></h2>
        <div class="help-tooltip" id="help-update_master"></div>
        <p class="update-desc">Paste CLI commands generated by an LLM from the AI Context file. Each line is executed against the current master file.</p>
        <div class="cli-command-area">
            <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;">
                <span style="font-size:13px;font-weight:600;color:var(--text);">CLI Commands</span>
                <span class="help-icon" onclick="event.stopPropagation();toggleHelp('cli_commands')">?</span>
            </div>
            <div class="help-tooltip" id="help-cli_commands"></div>
            <textarea id="cmdInput" class="cmd-input" rows="8" placeholder="show device&#10;show l1_link&#10;add device Router-A SW-1B RIGHT&#10;add l1_link_bulk &quot;[['SW-1B','WAN-1','GigabitEthernet 0/24','GigabitEthernet 0/24']]&quot;&#10;delete device OldDevice"></textarea>
            <div class="update-actions">
                <div class="run-status" id="runStatus" style="display:none;"></div>
                <button class="btn-run" id="btnRun">Run</button>
            </div>
        </div>
        <div class="run-progress" id="runProgress" style="display:none;">
            <div class="spinner"></div> <span id="runProgressText">Executing commands...</span>
        </div>
        <div class="run-results" id="runResults" style="display:none;"></div>
        <div class="llm-prompt-area" id="llmPromptArea" style="display:none;">
            <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;">
                <span style="font-size:13px;font-weight:600;color:var(--text);">LLM Prompt</span>
                <span class="help-icon" onclick="event.stopPropagation();toggleHelp('llm_prompt')">?</span>
            </div>
            <div class="help-tooltip" id="help-llm_prompt"></div>
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

    <!-- Selection Section -->
    <div class="card" id="selectionSection" style="display:none;">
        <h2>&#128203; Available Outputs <span class="help-icon" onclick="event.stopPropagation();toggleHelp('available_outputs')">?</span></h2>
        <div class="help-tooltip" id="help-available_outputs"></div>
        <div id="analysisSummary" style="font-size:13px;color:var(--text-secondary);margin-bottom:14px;padding:8px 12px;background:#f0f7ff;border-radius:6px;display:none;"></div>
        <div class="select-all-row">
            <label><input type="checkbox" id="selectAll" checked> Select / Deselect All</label>
            <span class="selected-count" id="selectedCount"></span>
            <button class="btn-generate" id="btnGenerate">Generate Selected</button>
        </div>
        <div class="output-list" id="outputList"></div>
        <div id="attributeContainer" style="display:none;"></div>
        <div class="estimate-total" id="estimateTotal" style="display:none;">
            <span>Estimated total time (Actual times may vary depending on your system.)</span>
            <span id="estimateTotalTime"></span>
        </div>
        <div class="generation-total" id="generationTotal" style="display:none;"></div>
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
        <span style="font-size:11px;color:#b0b8c4;">Author : Yusuke Ogawa - Architect, Cisco | CCIE#17583</span>
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
    var currentBasename = '';
    var currentDeviceCount = 0;
    var currentLinkCount = 0;
    var currentAttributeTitles = [];
    var isEmptyMaster = false;
    var heartbeatTimer = null;

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
        if (!file.name.endsWith('.xlsx')) { showError('Only .xlsx files are supported'); return; }
        if (!file.name.startsWith('[MASTER]')) { showError('Filename must start with [MASTER]'); return; }
        selectedFileEl.textContent = file.name + ' (' + formatSize(file.size) + ')';
        dropzone.classList.add('has-file');
        uploadAndAnalyze(file);
    }

    async function uploadAndAnalyze(file) {
        dropzone.classList.add('disabled');
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
            currentBasename = file.name.replace('[MASTER]', '').replace('.xlsx', '');

            var res = await fetch('/generate_step/' + currentJobId + '/init', { method: 'POST' });
            var data = await res.json();
            currentAreas = (data.areas && data.areas.length > 0) ? data.areas : [];
            currentDeviceCount = data.device_count || 0;
            currentLinkCount = data.link_count || 0;
            currentAttributeTitles = data.attribute_titles || [];

            buildSelectionList();
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
            l2_per_area:     [[13, 7],  [64, 8],  [256, 17],  [1024, 100]],
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

        items.push(
            { id: 'ai_context_file', type: 'ai', label: 'AI Context File',
              filename: '[AI_Context]' + currentBasename + '.txt', icon: 'AI',
              estSec: estimateSeconds('ai_context_file', currentDeviceCount, numAreas) }
        );
        if (!isEmptyMaster) {
            items.push(
                { id: 'device_file', type: 'device', label: 'Device File',
                  filename: '[DEVICE]' + currentBasename + '.xlsx', icon: 'DEV',
                  estSec: estimateSeconds('device_file', currentDeviceCount, numAreas) },
                { id: 'l1_all_areas_tag', type: 'l1', label: 'L1 Diagram - All Areas + Tags',
                  filename: '[L1_DIAGRAM]AllAreasTag_' + currentBasename + '.pptx', icon: 'L1',
                  subtype: 'all_areas_tag', estSec: l1Est },
                { id: 'l1_per_area_tag', type: 'l1', label: 'L1 Diagram - Per Area + Tags',
                  filename: '[L1_DIAGRAM]PerAreaTag_' + currentBasename + '.pptx', icon: 'L1',
                  subtype: 'per_area_tag', estSec: l1Est }
            );
            for (var i = 0; i < currentAreas.length; i++) {
                items.push({
                    id: 'l2_' + currentAreas[i], type: 'l2',
                    label: 'L2 Diagram (' + currentAreas[i] + ')',
                    filename: '[L2_DIAGRAM]' + currentAreas[i] + '_' + currentBasename + '.pptx',
                    icon: 'L2', area: currentAreas[i],
                    estSec: estimateSeconds('l2_diagram', currentDeviceCount, numAreas)
                });
            }
            items.push(
                { id: 'l3_all_areas', type: 'l3', label: 'L3 Diagram - All Areas',
                  filename: '[L3_DIAGRAM]AllAreas_' + currentBasename + '.pptx', icon: 'L3',
                  subtype: 'all_areas', estSec: l3Est },
                { id: 'l3_per_area', type: 'l3', label: 'L3 Diagram - Per Area',
                  filename: '[L3_DIAGRAM]PerArea_' + currentBasename + '.pptx', icon: 'L3',
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

            if (it.generated) {
                var encFn = encodeURIComponent(it.filename);
                var dlUrl = '/download/' + currentJobId + '/' + encFn;
                var pvUrl = '/preview/' + currentJobId + '/' + encFn;
                var btnHtml = '';
                
                if (it.type !== 'master' && it.type !== 'ai') {
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
        updateSelectedCount();
    });

    function updateSelectedCount() {
        var allBoxes = outputList.querySelectorAll('input[type="checkbox"]');
        var enabledBoxes = outputList.querySelectorAll('input[type="checkbox"]:not(:disabled)');
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
        for (var g2 = 0; g2 < genBtns.length; g2++) {
            var parentItem = genBtns[g2].closest('.output-item');
            var cb = parentItem ? parentItem.querySelector('input[type="checkbox"]') : null;
            genBtns[g2].disabled = !!(cb && cb.disabled);
        }
    }

    btnGenerate.addEventListener('click', async function() {
        var selected = [];
        var boxes = outputList.querySelectorAll('input[type="checkbox"]:checked:not(:disabled)');
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
        var aiItem = outputList.querySelector('input[data-step="ai_context_file"]');
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
                btnDownloadAll.href = '/download_all/' + currentJobId;
                btnDownloadAll.style.display = 'block';
                btnReset.style.display = 'block';
                updateSelectedCount();
            }
        } catch (e) {
            showError('Failed to retrieve file list');
        }
    }

    btnReset.addEventListener('click', function() { resetAll(); });

    btnRun.addEventListener('click', async function() {
        var cmds = cmdInput.value.trim();
        if (!cmds || !currentJobId) return;
        btnRun.disabled = true;
        runProgress.style.display = 'flex';
        runStatus.style.display = 'none';
        runResults.style.display = 'none';

        var nextVersion = masterVersion + 1;

        try {
            var resp = await fetch('/run_commands/' + currentJobId, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ commands: cmds, version: nextVersion })
            });
            var data = await resp.json();

            var html = '';
            if (data.results) {
                for (var i = 0; i < data.results.length; i++) {
                    var r = data.results[i];
                    var cls = r.skipped ? 'cmd-skip' : (r.success ? 'cmd-ok' : 'cmd-err');
                    var prefix = r.skipped ? '(skipped)' : (r.success ? '&#10003;' : '&#10007;');
                    html += '<div class="' + cls + '">' + prefix + ' ' + escapeHtml(r.command) + '</div>';
                    if (r.output && r.output.trim()) {
                        html += '<div class="cmd-output" id="cmdOut' + i + '">' + escapeHtml(r.output.trim()) + '</div>';
                    }
                }
                html += '<span class="view-log-link" id="viewLogLink">&#9654; View Log</span>';
            }
            runResults.innerHTML = html;
            runResults.style.display = 'block';
            var logLink = document.getElementById('viewLogLink');
            if (logLink) {
                logLink.addEventListener('click', function() {
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

            if (data.success) {
                runStatus.innerHTML = '<span style="color:var(--success)">All commands executed successfully.</span>';
            } else {
                runStatus.innerHTML = '<span style="color:#e94560">' + (data.errors ? data.errors.length : 0) + ' command(s) failed.</span>';
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

                runProgress.style.display = 'none';
                runProgressText.textContent = 'Executing commands...';

                buildSelectionList();
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
                btnDownloadAll.style.display = 'none';
                btnReset.style.display = 'none';
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
        currentJobId = null; currentAreas = []; currentBasename = '';
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
        estimateTotal.style.display = 'none';
        if (estimateNote) estimateNote.style.display = 'none';
        btnGenerate.style.display = '';
        btnGenerate.disabled = false;
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
            masterVersion = sess.masterVersion || 0;
            updatedMasters = data.updated_masters || sess.updatedMasters || [];
            isEmptyMaster = sess.isEmptyMaster || false;
            currentAreas = (data.areas && data.areas.length > 0) ? data.areas : [];
            currentDeviceCount = data.device_count || 0;
            currentLinkCount = data.link_count || 0;
            currentAttributeTitles = data.attribute_titles || sess.attributeTitles || [];

            buildSelectionList();

            var summary = document.getElementById('analysisSummary');
            if (currentDeviceCount > 0 || currentLinkCount > 0) {
                summary.textContent = 'Detected: ' + currentDeviceCount + ' devices, '
                    + currentLinkCount + ' links, ' + currentAreas.length + ' areas';
                summary.style.display = 'block';
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

    if cfg_host in ('0.0.0.0', 'localhost', '127.0.0.1'):
        bind_host = cfg_host
        display_host = 'localhost'
    else:
        bind_host = '0.0.0.0'
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

    app.run(host=bind_host, port=port, debug=False, ssl_context=ctx, threaded=True)
