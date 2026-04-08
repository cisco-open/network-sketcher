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
Stop all running Network Sketcher Online web server processes.

Usage:
    python stop_ns_online.py

All running ns_web_start.py processes are stopped (including any that were
started outside of this script), and the PID file is removed.
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


def _find_all_ns_server_pids():
    """Return a list of PIDs for all running ns_web_start.py processes.

    Searches the OS process list directly so that processes started outside
    of this script (e.g. from a terminal or IDE) are also found.
    """
    own_pid = os.getpid()
    pids = []
    try:
        if platform.system() == 'Windows':
            result = subprocess.run(
                ['wmic', 'process', 'where',
                 '(CommandLine like "%ns_web_start.py%") AND (Name like "%python%")',
                 'get', 'ProcessId', '/VALUE'],
                capture_output=True, text=True,
            )
            for line in result.stdout.splitlines():
                line = line.strip()
                if line.startswith('ProcessId='):
                    try:
                        pid = int(line.split('=', 1)[1])
                        if pid and pid != own_pid:
                            pids.append(pid)
                    except ValueError:
                        pass
        else:
            result = subprocess.run(
                ['pgrep', '-f', 'ns_web_start.py'],
                capture_output=True, text=True,
            )
            for line in result.stdout.splitlines():
                try:
                    pid = int(line.strip())
                    if pid != own_pid:
                        pids.append(pid)
                except ValueError:
                    pass
    except Exception:
        pass

    # Also include any PID recorded in the PID file that was missed above
    if PID_FILE.exists():
        try:
            pid = int(PID_FILE.read_text().strip())
            if pid and pid != own_pid and pid not in pids:
                pids.append(pid)
        except (ValueError, OSError):
            pass

    return pids


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


def main():
    print('=' * 56)
    print('  Network Sketcher Online — Stop')
    print('=' * 56)

    pids = _find_all_ns_server_pids()
    if not pids:
        print('  No running server processes found.')
        PID_FILE.unlink(missing_ok=True)
        print('=' * 56)
        sys.exit(0)

    print(f'  Found {len(pids)} running server process(es): {pids}')
    for pid in pids:
        print(f'  Stopping PID {pid}...')
        try:
            if platform.system() == 'Windows':
                subprocess.run(
                    ['taskkill', '/F', '/T', '/PID', str(pid)],
                    capture_output=True,
                )
            else:
                os.kill(pid, signal.SIGTERM)
        except (OSError, ProcessLookupError):
            pass

    # Wait up to 10 seconds for all processes to exit
    deadline = time.time() + 10.0
    while time.time() < deadline:
        time.sleep(0.5)
        still_running = [pid for pid in pids if _is_process_running(pid)]
        if not still_running:
            break
    else:
        still_running = [pid for pid in pids if _is_process_running(pid)]
        if still_running:
            print(f'  Warning: Process(es) {still_running} may still be running.')

    PID_FILE.unlink(missing_ok=True)
    print(f'  Stopped {len(pids)} process(es).')
    print('=' * 56)


if __name__ == '__main__':
    main()
