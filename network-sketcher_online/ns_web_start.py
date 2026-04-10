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

# Tracks the last attribute value used when SVG thumbnails were generated for a job.
# Used by svg_grid_stream to skip CLI re-execution when attribute unchanged and cache is warm.
# {job_id: attr_string}  (attr_string may be '' if no attribute was selected)
_svg_grid_last_attr: dict = {}
_svg_grid_last_attr_lock = threading.Lock()


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

    FAILURE_MARKERS = ('[ERROR]', 'Input must start with', '[WARNING] No ')
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
        ok = result is not None and not any(m in stdout for m in FAILURE_MARKERS)
        results.append({'command': line, 'success': ok, 'output': stdout + stderr})
        if not ok:
            errors.append(line)

        if not skip_sync and pending_sync_master and current_syncable:
            pending_sync_master = None

    if pending_sync_master:
        _run_deferred_sync(pending_sync_master)

    if has_mutation and new_master_name:
        set_active_master(str(work_dir), new_master_name)
        # Invalidate SVG grid cache so next buildSvgGrid() regenerates from updated master
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
    if cell_id.startswith('l2_area_'):
        area = cell_id[len('l2_area_'):]
        return sorted([f for f in files if f.startswith(f'[L2_DIAGRAM]{area}_')])
    return []


@app.route('/svg_grid_stream/<job_id>')
def svg_grid_stream(job_id):
    """SSE endpoint: generate all SVGs in parallel and stream completion events."""
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
    areas = request.args.getlist('area')
    limit = _cfg.get('parallel_limit', 5)

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

    tasks = [
        _task('l1_all', make_args(['export', 'l1_diagram', '--type', 'all_areas_tag'])),
        _task([f'l1_per_area_{a}' for a in areas],
              make_args(['export', 'l1_diagram', '--type', 'per_area_tag'])),
        _task('l3_all', make_args(['export', 'l3_diagram', '--type', 'all_areas'])),
        _task([f'l3_per_area_{a}' for a in areas],
              make_args(['export', 'l3_diagram', '--type', 'per_area'])),
    ] + [
        _task(f'l2_area_{area}', make_args(['export', 'l2_diagram', '--area', area]))
        for area in areas
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
                per_cell_b64[fname] = _b64_svg.b64encode(svg_bytes).decode('ascii')
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
                                    per_cell_b64[found[0]] = _b64_svg.b64encode(_f.read()).decode('ascii')
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
            basename = master_filename.replace('[MASTER]', '').replace('.xlsx', '').replace('.nsm', '')
            ai_filename_expected = f'[AI_Context]{basename}.txt'

            # --- Cache fast-path ---
            # If the same attribute was used last time and SVGs exist (memory or disk),
            # serve everything from cache without re-running any CLI command.
            with _svg_grid_last_attr_lock:
                last_attr = _svg_grid_last_attr.get(job_id)
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
                    # AI context: reuse existing file if already generated
                    ai_filename = ai_filename_expected if (work_dir / ai_filename_expected).exists() else None
                    if not ai_filename:
                        try:
                            ai_ok = generate_ai_context_parallel(work_dir, master_filename)
                            ai_filename = ai_filename_expected if ai_ok else None
                        except Exception as exc:
                            logger.error('AI context error (cache path): %s', exc)
                    yield 'data: ' + json.dumps({'done': True, 'ai_filename': ai_filename}) + '\n\n'
                    return

            # --- Full generation path ---
            # Start AI context generation in a background thread immediately so it
            # runs concurrently with SVG generation instead of sequentially after it.
            ai_result = [None]
            ai_done_event = threading.Event()

            def _run_ai_context():
                try:
                    ai_ok = generate_ai_context_parallel(work_dir, master_filename)
                    ai_result[0] = ai_filename_expected if ai_ok else None
                except Exception as exc:
                    logger.error('AI context generation error in svg_grid_stream: %s', exc)
                finally:
                    ai_done_event.set()

            threading.Thread(target=_run_ai_context, daemon=True).start()

            with concurrent.futures.ThreadPoolExecutor(max_workers=limit) as executor:
                futures = {
                    executor.submit(_run_task, cell_ids, cmd_args): cell_ids
                    for cell_ids, cmd_args in tasks
                }
                for future in concurrent.futures.as_completed(futures):
                    per_cell, success, per_cell_b64 = future.result()
                    yield from _stream_events(per_cell, per_cell_b64, success)

            # Record the attribute so next call can use cache fast-path
            with _svg_grid_last_attr_lock:
                _svg_grid_last_attr[job_id] = attr

            # Wait for AI context thread (usually already finished by now)
            ai_done_event.wait()
            yield 'data: ' + json.dumps({'done': True, 'ai_filename': ai_result[0]}) + '\n\n'
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
    # Serve from in-memory cache when available (populated by run_ns_command_subprocess).
    # This avoids Windows AV file-lock issues that can occur when send_file() tries to
    # open the file after shutil.move() — even if _wait_readable() already confirmed it.
    with _svg_mem_cache_lock:
        svg_bytes = _svg_mem_cache.get(job_id, {}).get(filename)
    if svg_bytes is not None:
        return Response(svg_bytes, mimetype='image/svg+xml')
    # Fallback: serve from disk with retry (for files not yet cached, or cache miss).
    for _attempt in range(3):
        if not filepath.is_file():
            break
        try:
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

    tabs_data: list of dicts with keys 'id', 'label', 'headers', 'rows'
    """
    safe_title = f'Device Preview - {master_basename}'

    tabs_js_list = []
    for tab in tabs_data:
        escaped_headers = [h.replace('\\', '\\\\').replace("'", "\\'") for h in tab['headers']]
        headers_js = '[' + ', '.join(f"'{h}'" for h in escaped_headers) + ']'
        rows_js_parts = []
        for row in tab['rows']:
            cells = []
            for cell in row:
                s = str(cell) if cell is not None else ''
                s = s.replace('\\', '\\\\').replace("'", "\\'").replace('\n', ' ').replace('\r', '')
                cells.append(f"'{s}'")
            rows_js_parts.append('[' + ', '.join(cells) + ']')
        rows_js = '[' + ', '.join(rows_js_parts) + ']'
        tab_id = tab['id'].replace('\\', '\\\\').replace("'", "\\'")
        tab_label = tab['label'].replace('\\', '\\\\').replace("'", "\\'")
        # Optionally serialize per-cell background colors (Attribute tab only)
        row_colors_js = ''
        if tab.get('row_colors'):
            color_parts = []
            for crow in tab['row_colors']:
                cells = ['null' if c is None else "'" + str(c) + "'" for c in crow]
                color_parts.append('[' + ','.join(cells) + ']')
            row_colors_js = ',row_colors:[' + ','.join(color_parts) + ']'
        tabs_js_list.append(
            '{id:\'' + tab_id + '\', label:\'' + tab_label + '\', headers:' + headers_js
            + ', rows:' + rows_js + row_colors_js + '}'
        )
    tabs_js = '[' + ', '.join(tabs_js_list) + ']'

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
.toolbar .row-count {{ font-size: 12px; opacity: 0.6; white-space: nowrap; }}
.toolbar button {{ padding: 6px 16px; border: 1px solid rgba(255,255,255,0.4);
                   background: transparent; color: #fff; border-radius: 6px; font-size: 13px;
                   cursor: pointer; transition: all 0.2s; white-space: nowrap; }}
.toolbar button:hover {{ background: rgba(255,255,255,0.15); border-color: rgba(255,255,255,0.7); }}
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
.table-wrap tr.header-row th {{
    background: #4A8FE7; color: #fff; font-weight: 600; position: sticky; top: 0; z-index: 1; }}
.table-wrap tr:nth-child(even):not(.header-row) td {{ background: #f8f9fa; }}
.table-wrap tr:hover:not(.header-row) td {{ background: #e8f0fe; }}
</style>
</head>
<body>
<div class="toolbar">
    <h1>{safe_title}</h1>
    <span class="spacer"></span>
    <span class="row-count" id="rowCount"></span>
    <button id="btnDlCsv" title="Download current tab as CSV">&#8681; Download CSV</button>
    <button id="btnDlHtml" title="Download current tab as HTML">&#8681; Download HTML</button>
</div>
<div class="sheet-tabs" id="sheetTabs"></div>
<div id="content">
    <div class="table-wrap" id="tableWrap"></div>
</div>
<script>
(function() {{
    var TABS = {tabs_js};
    var currentTab = 0;
    var masterBase = '{master_basename.replace("'", "\\'")}';

    function escHtml(s) {{
        return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }}

    function buildTable(tab) {{
        var h = '<table><thead><tr class="header-row">';
        for (var i = 0; i < tab.headers.length; i++) {{
            h += '<th>' + escHtml(tab.headers[i]) + '</th>';
        }}
        h += '</tr></thead><tbody>';
        for (var r = 0; r < tab.rows.length; r++) {{
            h += '<tr>';
            var row = tab.rows[r];
            var colors = tab.row_colors ? tab.row_colors[r] : null;
            var colCount = Math.max(tab.headers.length, row.length);
            for (var c = 0; c < colCount; c++) {{
                var v = c < row.length ? row[c] : '';
                var bg = colors && c < colors.length && colors[c] ? colors[c] : null;
                var style = bg ? ' style="background-color:' + bg + '"' : '';
                h += '<td' + style + '>' + (v !== '' ? escHtml(v) : '') + '</td>';
            }}
            h += '</tr>';
        }}
        h += '</tbody></table>';
        return h;
    }}

    function showTab(idx) {{
        currentTab = idx;
        document.getElementById('tableWrap').innerHTML = buildTable(TABS[idx]);
        document.getElementById('rowCount').textContent = TABS[idx].rows.length + ' rows';
        var btns = document.querySelectorAll('.sheet-tab');
        for (var i = 0; i < btns.length; i++) {{
            btns[i].classList.toggle('active', i === idx);
        }}
    }}

    function initTabs() {{
        var container = document.getElementById('sheetTabs');
        for (var i = 0; i < TABS.length; i++) {{
            (function(idx) {{
                var b = document.createElement('button');
                b.className = 'sheet-tab';
                b.textContent = TABS[idx].label;
                b.onclick = function() {{ showTab(idx); }};
                container.appendChild(b);
            }})(i);
        }}
    }}

    function buildCsv(tab) {{
        var lines = [];
        var h = tab.headers.map(function(c) {{ return '"' + String(c).replace(/"/g,'""') + '"'; }}).join(',');
        lines.push(h);
        for (var r = 0; r < tab.rows.length; r++) {{
            var row = tab.rows[r];
            var cells = [];
            for (var c = 0; c < tab.headers.length; c++) {{
                var v = c < row.length ? row[c] : '';
                cells.push('"' + String(v).replace(/"/g,'""') + '"');
            }}
            lines.push(cells.join(','));
        }}
        return lines.join('\\r\\n');
    }}

    function buildHtml(tab) {{
        var rows = '';
        for (var r = 0; r < tab.rows.length; r++) {{
            rows += '<tr>';
            var colors = tab.row_colors ? tab.row_colors[r] : null;
            for (var c = 0; c < tab.headers.length; c++) {{
                var v = c < tab.rows[r].length ? tab.rows[r][c] : '';
                var bg = colors && c < colors.length && colors[c] ? colors[c] : null;
                var style = bg ? ' style="background-color:' + bg + '"' : '';
                rows += '<td' + style + '>' + escHtml(v) + '</td>';
            }}
            rows += '</tr>';
        }}
        var ths = tab.headers.map(function(h) {{ return '<th>' + escHtml(h) + '</th>'; }}).join('');
        return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + escHtml(tab.label) + ' - ' + escHtml(masterBase) +
            '</title><style>body{{font-family:sans-serif;font-size:13px;background:#f0f2f5;color:#333;margin:0}}' +
            '.toolbar{{background:#16213e;color:#fff;padding:10px 20px;font-size:14px}}' +
            '.wrap{{padding:16px}}.wrap table{{border-collapse:collapse;width:auto;min-width:100%;background:#fff;' +
            'box-shadow:0 1px 4px rgba(0,0,0,0.1);font-size:13px}}' +
            'th{{background:#4A8FE7;color:#fff;font-weight:600;padding:6px 12px;text-align:left;border:1px solid #e0e0e0}}' +
            'td{{padding:6px 12px;border:1px solid #e0e0e0}}' +
            'tr:nth-child(even) td{{background:#f8f9fa}}tr:hover td{{background:#e8f0fe}}' +
            '</style></head><body>' +
            '<div class="toolbar">' + escHtml(tab.label) + ' - ' + escHtml(masterBase) +
            ' (' + tab.rows.length + ' rows)</div>' +
            '<div class="wrap"><table><thead><tr>' + ths + '</tr></thead><tbody>' + rows + '</tbody></table></div>' +
            '</body></html>';
    }}

    function download(content, filename, mime) {{
        var bom = mime === 'text/csv' ? '\\uFEFF' : '';
        var blob = new Blob([bom + content], {{type: mime}});
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(a.href);
    }}

    document.getElementById('btnDlCsv').onclick = function() {{
        var tab = TABS[currentTab];
        var safeName = tab.label.replace(/[^a-zA-Z0-9_\\-]/g, '_');
        download(buildCsv(tab), masterBase + '_' + safeName + '.csv', 'text/csv');
    }};

    document.getElementById('btnDlHtml').onclick = function() {{
        var tab = TABS[currentTab];
        var safeName = tab.label.replace(/[^a-zA-Z0-9_\\-]/g, '_');
        download(buildHtml(tab), masterBase + '_' + safeName + '.html', 'text/html');
    }};

    initTabs();
    // Open the tab specified by the URL hash (e.g. #l1, #l2, #l3)
    var hashTabId = window.location.hash.replace('#', '');
    var initIdx = 0;
    if (hashTabId) {{
        for (var hi = 0; hi < TABS.length; hi++) {{
            if (TABS[hi].id === hashTabId) {{ initIdx = hi; break; }}
        }}
    }}
    showTab(initIdx);
}})();
</script>
</body>
</html>'''


def _get_device_tabs_data(job_id):
    """Shared helper: fetch all 4 device tabs data from master file.

    Returns (tabs_data list, master_basename str) or (None, None) on error.
    Uses nsm_def.convert_master_to_array directly (in-process, no subprocess,
    no XLSX generation) for L1/L2/L3. Device list uses run_ns_command.
    """
    import ast as _ast
    import sys as _sys

    work_dir = UPLOAD_DIR / job_id
    if not work_dir.is_dir():
        return None, None

    master_filename = get_active_master(str(work_dir))
    if not master_filename:
        return None, None

    master_path = str(work_dir / master_filename)
    master_basename = master_filename
    for ext in ('.nsm', '.xlsx'):
        if master_basename.lower().endswith(ext):
            master_basename = master_basename[:-len(ext)]
            break
    master_basename = master_basename.replace('[MASTER]', '').strip()

    # Add ns_engine dir to path so nsm_def can be imported directly
    _engine_dir = str(BASE_DIR / 'ns_engine')
    if _engine_dir not in _sys.path:
        _sys.path.insert(0, _engine_dir)
    try:
        import nsm_def as _nsm_def
    except ImportError:
        logger.warning('nsm_def import failed; falling back to empty tabs')
        return None, None

    def _read_section(ws_name, section):
        try:
            return _nsm_def.convert_master_to_array(ws_name, master_path, section)
        except Exception as exc:
            logger.warning('convert_master_to_array %s/%s failed: %s', ws_name, section, exc)
            return []

    # --- POSITION_SHAPE: build device -> area lookup (used by L1 and Attribute) ---
    ps_raw = _read_section('Master_Data', '<<POSITION_SHAPE>>')
    _device_area_map = {}
    _cur_folder = None
    for _item in ps_raw:
        if not isinstance(_item, list) or len(_item) < 2:
            continue
        _row = _item[1]
        if not isinstance(_row, list):
            continue
        if _row and _row[0] and _row[0] not in ('', '<END>', '<<POSITION_SHAPE>>', '_AIR_'):
            _cur_folder = _row[0]
        if _cur_folder:
            _area_label = '_N/A_' if '_wp_' in _cur_folder else _cur_folder
            for _val in _row:
                if _val and _val not in ('', '<END>', '_AIR_', '<<POSITION_SHAPE>>', _cur_folder) \
                        and not str(_val).startswith('_AIR_'):
                    _device_area_map[str(_val)] = _area_label
        if isinstance(_row, list) and len(_row) == 1 and _row[0] == '<END>':
            _cur_folder = None

    # --- L2: Master_Data_L2 / <<L2_TABLE>> ---
    # Cols: 0=Area, 1=Device Name, 2=Port Mode(formula/empty), 3=Port Name,
    #       4=Virtual Port Mode(formula/empty), 5=Virtual Port Name,
    #       6=Connected L2 Segment, 7=L2 (L3 Virtual Port)
    # Skip formula cols 2 and 4.
    l2_raw = _read_section('Master_Data_L2', '<<L2_TABLE>>')
    l2_data = [r[1] for r in l2_raw if isinstance(r, list) and r[0] > 2]
    USE_L2 = [0, 1, 3, 5, 6, 7]
    l2_headers = ['Area', 'Device Name', 'Port Name', 'Virtual Port Name',
                  'Connected L2 Segment', 'L2 (L3 Virtual Port)']
    l2_rows = [
        [str(row[i]) if i < len(row) and row[i] is not None and row[i] != '' else ''
         for i in USE_L2]
        for row in l2_data
    ]

    # --- L3: Master_Data_L3 / <<L3_TABLE>> ---
    # Cols: 0=Area, 1=Device Name, 2=L3 IF Name, 3=L3 Instance Name,
    #       4=IP Address / Subnet mask, 5=[VPN] Target Device Name,
    #       6=[VPN] Target L3 Port Name
    l3_raw = _read_section('Master_Data_L3', '<<L3_TABLE>>')
    l3_hdr_row = next((r[1] for r in l3_raw if isinstance(r, list) and r[0] == 2), [])
    l3_headers = [h for h in l3_hdr_row if h is not None]
    l3_data = [r[1] for r in l3_raw if isinstance(r, list) and r[0] > 2]
    l3_rows = [
        [str(c) if c is not None else '' for c in row[:len(l3_headers)]]
        for row in l3_data
    ]

    # --- L1: Excel-compatible format (one row per device×port, 2 rows per link) ---
    # POSITION_LINE raw col indices:
    #   0=From_Name, 1=To_Name, 2=From_Tag_raw, 3=To_Tag_raw,
    #   12=From_Port_prefix, 13=From_Speed, 14=From_Duplex, 15=From_Port_Type,
    #   16=To_Port_prefix, 17=To_Speed, 18=To_Duplex, 19=To_Port_Type
    pl_raw = _read_section('Master_Data', '<<POSITION_LINE>>')
    pl_data = [r[1] for r in pl_raw if isinstance(r, list) and r[0] > 2]
    l1_headers = [
        'Area', 'Device Name', 'Port Name', 'Abbreviation(Diagram)',
        'Speed', 'Duplex', 'Port Type',
        '[src] Device Name', '[src] Port Name', '[dst] Device Name', '[dst] Port Name',
    ]

    def _make_port(raw_tag, prefix):
        """Return (full_port_name, abbreviation) from raw tag and prefix."""
        if ' ' in raw_tag:
            parts = raw_tag.split(' ')
            return (prefix + ' ' + parts[-1]).strip(), parts[0]
        return (prefix or raw_tag).strip(), raw_tag

    l1_rows = []
    for row in pl_data:
        if len(row) < 20:
            continue
        from_dev = row[0] or ''
        to_dev   = row[1] or ''
        from_full, from_abbr = _make_port(row[2] or '', row[12] or '')
        to_full,   to_abbr   = _make_port(row[3] or '', row[16] or '')
        from_raw = row[2] or ''
        to_raw   = row[3] or ''
        # Row A: From デバイス視点
        l1_rows.append([
            _device_area_map.get(from_dev, ''), from_dev,
            from_full, from_abbr,
            row[13] or '', row[14] or '', row[15] or '',
            from_dev, from_raw, to_dev, to_raw,
        ])
        # Row B: To デバイス視点
        l1_rows.append([
            _device_area_map.get(to_dev, ''), to_dev,
            to_full, to_abbr,
            row[17] or '', row[18] or '', row[19] or '',
            from_dev, from_raw, to_dev, to_raw,
        ])

    # Sort: Area昇順 → Device Name昇順 → ポート番号昇順（数値ソート）
    l1_rows.sort(key=lambda x: (
        x[0], x[1],
        _nsm_def.get_if_value(x[2]),
        x[2],
    ))

    # --- Attribute: show attribute --one_msg ---

    # Each non-device-name cell is "['value', [R, G, B]]". Extract text + color.
    attr_r = run_ns_command(['show', 'attribute', '--master', master_path, '--one_msg'])
    attr_headers = ['Area', 'Device Name']
    attr_rows = []
    attr_row_colors = []   # same shape as attr_rows; None or 'rgb(R,G,B)' per cell
    if attr_r and attr_r.returncode == 0:
        try:
            raw_attr = _ast.literal_eval(attr_r.stdout.strip())
            if raw_attr and isinstance(raw_attr[0], list):
                attr_headers = ['Area'] + raw_attr[0]
                for row in raw_attr[1:]:
                    vals = []
                    cols = []
                    for i, cell in enumerate(row):
                        if i == 0:
                            dev = str(cell) if cell is not None else ''
                            area = _device_area_map.get(dev, '')
                            vals = [area, dev]
                            cols = [None, None]
                        else:
                            try:
                                cell_list = _ast.literal_eval(str(cell))
                                text = cell_list[0] if cell_list else ''
                                text = '' if text in ('<EMPTY>', None) else str(text)
                                rgb = cell_list[1] if len(cell_list) > 1 else None
                                if rgb and isinstance(rgb, list) and len(rgb) == 3 \
                                        and tuple(rgb) != (255, 255, 255):
                                    color = f'rgb({rgb[0]},{rgb[1]},{rgb[2]})'
                                else:
                                    color = None
                            except Exception:
                                text = str(cell) if cell is not None else ''
                                color = None
                            vals.append(text)
                            cols.append(color)
                    attr_rows.append(vals)
                    attr_row_colors.append(cols)
        except Exception:
            pass

    # Sort Attribute rows: Area descending, then Device Name descending within each Area
    if attr_rows:
        _combined = list(zip(attr_rows, attr_row_colors))
        _combined.sort(
            key=lambda x: (x[0][0] if x[0] else '', x[0][1] if len(x[0]) > 1 else ''),
            reverse=False,
        )
        attr_rows, attr_row_colors = map(list, zip(*_combined))

    tabs_data = [
        {
            'id': 'l1',
            'label': 'L1 Table',
            'headers': l1_headers,
            'rows': l1_rows,
        },
        {
            'id': 'l2',
            'label': 'L2 Table',
            'headers': l2_headers,
            'rows': l2_rows,
        },
        {
            'id': 'l3',
            'label': 'L3 Table',
            'headers': l3_headers,
            'rows': l3_rows,
        },
        {
            'id': 'attribute',
            'label': 'Attribute',
            'headers': attr_headers,
            'rows': attr_rows,
            'row_colors': attr_row_colors,
        },
    ]
    return tabs_data, master_basename


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
    transition: all 0.2s;
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
.viewer img {{
    position: absolute; left: 0; top: 0;
    max-width: none; max-height: none;
    user-select: none; -webkit-user-drag: none;
    image-rendering: optimizeQuality;
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
    <button class="primary" id="btnDownload" title="Download SVG">&#8681; Download</button>
    <button class="primary" id="btnDownloadVisio" title="Download SVG optimized for Visio">&#8681; Download for Visio</button>
</div>
<div class="viewer" id="viewer">
    <img id="svgImg" src="{svg_url}" draggable="false" />
</div>
<script>
(function() {{
    var files = [{svg_files_js}];
    var names = [{svg_names_js}];
    var idx = {current_idx};
    var jobId = "{job_id}";

    var viewer = document.getElementById('viewer');
    var img = document.getElementById('svgImg');
    var zoomEl = document.getElementById('zoomLevel');
    var pageEl = document.getElementById('pageInfo');
    var titleEl = document.getElementById('titleText');

    var scale = 1, cx = 0, cy = 0;
    var dragging = false, dragStartX = 0, dragStartY = 0, cxStart = 0, cyStart = 0;

    function updateTransform() {{
        var vw = viewer.clientWidth, vh = viewer.clientHeight;
        var iw = (img.naturalWidth || 1) * scale;
        var ih = (img.naturalHeight || 1) * scale;
        var left = vw / 2 - cx * scale;
        var top = vh / 2 - cy * scale;
        img.style.left = left + 'px';
        img.style.top = top + 'px';
        img.style.width = iw + 'px';
        img.style.height = ih + 'px';
        img.style.transform = '';
        zoomEl.textContent = Math.round(scale * 100) + '%';
    }}

    function fitToWindow() {{
        var vw = viewer.clientWidth, vh = viewer.clientHeight;
        var iw = img.naturalWidth || 1, ih = img.naturalHeight || 1;
        scale = Math.min(vw / iw, vh / ih, 1) * 0.95;
        cx = iw / 2;
        cy = ih / 2;
        updateTransform();
    }}

    img.onload = function() {{ fitToWindow(); }};
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
        img.src = '/svg_raw/' + jobId + '/' + files[idx];
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

    document.getElementById('btnPrev').disabled = (idx === 0);
    document.getElementById('btnNext').disabled = (idx === files.length - 1);

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
    width: 100%; padding: 4px 10px 10px; border: 1.5px solid #00b894; border-radius: 8px;
    background: #f0faf6; box-sizing: border-box;
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
                    <div class="run-status" id="runStatus" style="display:none;"></div>
                    <button class="btn-run" id="btnRun">Run</button>
                </div>
            </div>
        </div>
        <div class="run-progress" id="runProgress" style="display:none;">
            <div class="spinner"></div> <span id="runProgressText">Executing commands...</span>
        </div>
        <div class="run-results" id="runResults" style="display:none;"></div>
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
    </div>

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
    // SVG Grid: build the preview grid and start streaming SVG generation //
    // ------------------------------------------------------------------ //
    function buildSvgGrid(initAttr) {
        // Close any existing SSE stream
        if (svgGridEventSource) { svgGridEventSource.close(); svgGridEventSource = null; }

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

        // Row definitions: [row_id, row_label]
        var rows = [
            ['all_areas', 'All Areas'],
        ];
        for (var ai = 0; ai < currentAreas.length; ai++) {
            rows.push(['area_' + currentAreas[ai], currentAreas[ai]]);
        }

        // Column definitions: [col_id, col_label, cell_for_row_fn]
        // cell_for_row_fn(row_id) => grid cell_id for this col+row combo
        var cols = [
            { id: 'l1', label: 'Layer 1', cellId: function(row_id) {
                if (row_id === 'all_areas') return 'l1_all';
                return 'l1_per_area_' + row_id.replace(/^area_/, ''); } },
            { id: 'l2', label: 'Layer 2', cellId: function(row_id) {
                if (row_id === 'all_areas') return null;
                return 'l2_area_' + row_id.replace(/^area_/, ''); } },
            { id: 'l3', label: 'Layer 3', cellId: function(row_id) {
                if (row_id === 'all_areas') return 'l3_all';
                return 'l3_per_area_' + row_id.replace(/^area_/, ''); } },
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
            attrSel2.addEventListener('change', function() {
                var savedAttr = this.value;   // capture before DOM is cleared
                if (svgGridEventSource) { svgGridEventSource.close(); svgGridEventSource = null; }
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

        // Body rows — collect all unique cell_ids to stream
        var tbody = table.createTBody();
        var cellMap = {}; // cell_id => td element (may map to multiple tds)

        // --- Device File row (top, before All Areas) ---
        var devPreviewUrl = '/device_preview/' + currentJobId;
        // col id -> tab id mapping: L1->l1 interface, L2->l2 interface, L3->l3 interface
        var devColTabMap = { l1: 'l1', l2: 'l2', l3: 'l3' };
        var devTr = tbody.insertRow();
        devTr.className = 'device-file-row';
        var devThCell = devTr.insertCell();
        devThCell.className = 'row-header';
        devThCell.innerHTML = '<a href="' + devPreviewUrl + '" target="_blank" '
            + 'style="text-decoration:none;color:inherit;">Device Tables</a>';
        var devTdByTabId = {};  // tabId -> td element
        for (var dci = 0; dci < cols.length; dci++) {
            var dTd = devTr.insertCell();
            dTd.innerHTML = '<div class="cell-spinner">&#8987;</div>';
            devTdByTabId[devColTabMap[cols[dci].id]] = dTd;
        }

        // --- Diagram rows (All Areas + per-area) ---
        for (var ri = 0; ri < rows.length; ri++) {
            var row_id = rows[ri][0];
            var row_label = rows[ri][1];
            var tr = tbody.insertRow();

            var tdHeader = tr.insertCell();
            tdHeader.className = 'row-header';
            tdHeader.textContent = row_label;

            for (var ci2 = 0; ci2 < cols.length; ci2++) {
                var cell_id = cols[ci2].cellId(row_id);
                var td = tr.insertCell();

                if (!cell_id) {
                    // N/A cell (e.g., L2 for All Areas)
                    td.innerHTML = '<div class="cell-na">N/A</div>';
                    continue;
                }

                // Spinner placeholder
                td.setAttribute('data-cell', cell_id);
                td.innerHTML = '<div class="cell-spinner">&#8987;</div>';

                if (!cellMap[cell_id]) cellMap[cell_id] = [];
                cellMap[cell_id].push(td);
            }
        }

        scrollDiv.appendChild(table);
        wrapper.appendChild(scrollDiv);
        outputList.appendChild(wrapper);

        // --- Fetch Device File thumbnail data in parallel with SSE ---
        // Render iframe thumbnails directly — no fetch needed
        (function renderDeviceIframes() {
            Object.keys(devTdByTabId).forEach(function(tabId) {
                var tabUrl = '/device_preview/' + currentJobId + '#' + tabId;
                devTdByTabId[tabId].innerHTML = buildDeviceThumbHtml({ id: tabId }, tabUrl);
            });
        })();

        // Build SSE URL
        var sseUrl = '/svg_grid_stream/' + currentJobId;
        var sseParams = [];
        if (selAttr) sseParams.push('attribute=' + encodeURIComponent(selAttr));
        for (var ai2 = 0; ai2 < currentAreas.length; ai2++) {
            sseParams.push('area=' + encodeURIComponent(currentAreas[ai2]));
        }
        if (sseParams.length) sseUrl += '?' + sseParams.join('&');

        // Start SSE
        var svgGridTotalEl = document.getElementById('svgGridTotal');
        if (svgGridTotalEl) {
            // Show estimated generation time until SSE done replaces it with actual time
            if (currentDeviceCount > 0) {
                svgGridTotalEl.textContent = _fmtEstimate(_estimateSec(currentDeviceCount, _PERF_DIAGRAMS));
                svgGridTotalEl.style.display = 'flex';
            } else {
                svgGridTotalEl.style.display = 'none';
            }
        }
        var svgGridStartTime = Date.now();

        // Re-disable ZIP button while SVG generation is running (covers attribute toggle)
        if (currentJobId) {
            btnDownloadAll.href = '/download_all/' + currentJobId;
            btnDownloadAll.classList.add('disabled');
            btnDownloadAll.style.display = 'block';
        }

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
                    // Enable ZIP download button now that all SVGs are ready
                    btnDownloadAll.classList.remove('disabled');
                    // Remove spinner and add Download button
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
                    return;
                }
                var cell_id = msg.cell;
                var tds = cellMap[cell_id];
                if (!tds) return;

                var inner = '';
                if (msg.error || !msg.files || msg.files.length === 0) {
                    inner = '<div class="cell-error">&#10007;<br>Error</div>';
                } else {
                    // Build viewer URL: use first SVG file, pass filter prefix if multiple
                    var firstFile = msg.files[0];
                    var viewerUrl = '/preview/' + currentJobId + '/' + encodeURIComponent(firstFile);
                    if (msg.filter) viewerUrl += '?filter=' + encodeURIComponent(msg.filter);
                    // Thumbnail: use base64 data URI from SSE payload if available (avoids
                    // Windows file-lock races between shutil.move and browser HTTP requests).
                    // Fall back to /svg_raw/ URL for large files not embedded in the SSE event.
                    var fallbackUrl = '/svg_raw/' + currentJobId + '/' + encodeURIComponent(firstFile)
                        + '?v=' + encodeURIComponent(selAttr || '_');
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
                for (var ti2 = 0; ti2 < tds.length; ti2++) {
                    tds[ti2].innerHTML = inner;
                }

            } catch (e) {}
        };

        svgGridEventSource.onerror = function() {
            if (svgGridEventSource) { svgGridEventSource.close(); svgGridEventSource = null; }
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

        // --- Event delegation for this section ---
        section.addEventListener('click', async function(ev) {
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
                        a.download = currentMasterFilename.replace('.nsm', '.xlsx');
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

        section.style.display = 'block';
        // Reflect AI Context checkbox in the selected count immediately
        if (typeof updateSelectedCount === 'function') updateSelectedCount();

        // ZIP download button: show disabled in SVG mode immediately after upload;
        // hide again when switching away from SVG mode before generation completes.
        if (outputMode === 'svg' && currentJobId) {
            btnDownloadAll.href = '/download_all/' + currentJobId;
            btnDownloadAll.classList.add('disabled');
            btnDownloadAll.style.display = 'block';
        } else if (btnDownloadAll.classList.contains('disabled')) {
            // Switched away from SVG mode while generation was still pending — hide button
            btnDownloadAll.classList.remove('disabled');
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
                btnDownloadAll.href = '/download_all/' + currentJobId;
                btnDownloadAll.classList.remove('disabled');
                btnDownloadAll.style.display = 'block';
                btnReset.style.display = 'block';
                updateSelectedCount();
            }
        } catch (e) {
            showError('Failed to retrieve file list');
        }
    }

    btnReset.addEventListener('click', function() { resetAll(); });

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
                a.download = currentMasterFilename.replace('.nsm', '.xlsx');
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
                btnDownloadAll.href = '/download_all/' + currentJobId;
                btnDownloadAll.classList.remove('disabled');
                btnDownloadAll.style.display = 'block';
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

            btnDownloadAll.href = '/download_all/' + currentJobId;
            btnDownloadAll.classList.remove('disabled');
            btnDownloadAll.style.display = 'block';

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
