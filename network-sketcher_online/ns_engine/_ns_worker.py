#!/usr/bin/env python3
"""
Subprocess worker for parallel SVG/diagram generation.

Called by run_ns_command_subprocess() in ns_web_start.py.
Each invocation runs in its own Python interpreter, bypassing
the in-process _cli_lock and enabling true parallelism.

Usage (internal only):
    python _ns_worker.py <ns_cli_args...>
"""
import sys
import os
import traceback

# Ensure ns_engine directory is on sys.path (for direct imports: import nsm_cli, etc.)
_ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))
if _ENGINE_DIR not in sys.path:
    sys.path.insert(0, _ENGINE_DIR)

# Also add the PARENT of ns_engine so that 'from ns_engine.xxx import yyy' style
# imports inside nsm_adapter (e.g. 'from ns_engine.nsm_io import nsm_to_xlsx')
# resolve correctly when loading .nsm master files via openpyxl.
_ONLINE_DIR = os.path.dirname(_ENGINE_DIR)
if _ONLINE_DIR not in sys.path:
    sys.path.insert(0, _ONLINE_DIR)

from nsm_adapter import bootstrap, run_cli  # noqa: E402


if __name__ == '__main__':
    try:
        bootstrap()
        cli_args = sys.argv[1:]
        result = run_cli(cli_args, cwd=_ENGINE_DIR)
        if result.returncode != 0:
            # Write captured CLI output to actual stderr so the parent process
            # can capture it via subprocess.run(capture_output=True).
            if result.stderr:
                sys.stderr.write('[WORKER STDERR]\n' + result.stderr + '\n')
            if result.stdout:
                sys.stderr.write('[WORKER STDOUT]\n' + result.stdout + '\n')
        sys.exit(result.returncode)
    except Exception:
        sys.stderr.write('[WORKER EXCEPTION]\n' + traceback.format_exc() + '\n')
        sys.exit(1)
