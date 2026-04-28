"""
Network Sketcher CLI Worker – executed as a subprocess by ns_mcp_server.py.

Accepts a JSON payload on stdin, runs one or more CLI commands in-process,
and writes the results as a JSON array to stdout.  stderr is left open so
any Python tracebacks reach the parent process.
"""

import json
import sys


def main() -> None:
    payload = json.loads(sys.stdin.buffer.read())
    engine_dir: str = payload['engine_dir']
    online_dir: str = payload['online_dir']
    commands: list = payload['commands']  # list of list[str]

    sys.path.insert(0, engine_dir)
    sys.path.insert(0, online_dir)

    from ns_engine.nsm_adapter import bootstrap, run_cli  # noqa: E402
    bootstrap()

    results = []
    for args in commands:
        result = run_cli(list(args), cwd=engine_dir)
        results.append({
            'returncode': result.returncode,
            'stdout': result.stdout,
            'stderr': result.stderr,
        })

    # Write JSON result to the REAL stdout (not sys.stdout which run_cli may
    # have temporarily redirected – it's always restored, but writing to the
    # raw buffer is safer here).
    output = json.dumps(results, ensure_ascii=False)
    sys.stdout.buffer.write(output.encode('utf-8'))
    sys.stdout.buffer.flush()


if __name__ == '__main__':
    main()
