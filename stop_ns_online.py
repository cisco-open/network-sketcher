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
Stop the Network Sketcher Online web server.

Usage:
    python stop_ns_online.py

Works on Windows, Mac OS, and Linux.
"""

import subprocess
import sys
import os
import platform
import signal
import time
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ONLINE_DIR = BASE_DIR / 'network-sketcher_online'
PID_FILE = ONLINE_DIR / 'ns_online.pid'


def _is_process_running(pid):
    """Check if a process with the given PID is still running."""
    try:
        if platform.system() == 'Windows':
            result = subprocess.run(
                ['tasklist', '/FI', f'PID eq {pid}', '/NH'],
                capture_output=True, text=True,
            )
            return str(pid) in result.stdout
        else:
            os.kill(pid, 0)
            return True
    except (OSError, ProcessLookupError):
        return False


def _is_ns_server(pid):
    """Verify that the PID belongs to a Network Sketcher Online server process."""
    try:
        if platform.system() == 'Windows':
            result = subprocess.run(
                ['wmic', 'process', 'where', f'ProcessId={pid}',
                 'get', 'CommandLine', '/VALUE'],
                capture_output=True, text=True,
            )
            return 'ns_web_start.py' in result.stdout
        else:
            result = subprocess.run(
                ['ps', '-p', str(pid), '-o', 'args='],
                capture_output=True, text=True,
            )
            return 'ns_web_start.py' in result.stdout
    except Exception:
        return False


def main():
    print('=' * 56)
    print('  Network Sketcher Online — Stop')
    print('=' * 56)

    if not PID_FILE.exists():
        print('  Server is not running (no PID file found).')
        sys.exit(0)

    try:
        pid = int(PID_FILE.read_text().strip())
    except (ValueError, OSError):
        print('  Invalid PID file. Removing it.')
        PID_FILE.unlink(missing_ok=True)
        sys.exit(1)

    if not _is_process_running(pid):
        print(f'  Process {pid} is not running. Removing stale PID file.')
        PID_FILE.unlink(missing_ok=True)
        sys.exit(0)

    if not _is_ns_server(pid):
        print(f'  PID {pid} is not a Network Sketcher process. Removing stale PID file.')
        PID_FILE.unlink(missing_ok=True)
        sys.exit(1)

    print(f'  Stopping server (PID: {pid})...')
    try:
        if platform.system() == 'Windows':
            subprocess.run(
                ['taskkill', '/F', '/T', '/PID', str(pid)],
                capture_output=True,
            )
        else:
            os.kill(pid, signal.SIGTERM)

        for _ in range(20):
            time.sleep(0.5)
            if not _is_process_running(pid):
                break
        else:
            print('  Warning: Process may still be running.')

    except (OSError, ProcessLookupError):
        pass

    PID_FILE.unlink(missing_ok=True)
    print('  Server stopped.')
    print('=' * 56)


if __name__ == '__main__':
    main()
