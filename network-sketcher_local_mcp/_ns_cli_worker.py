"""
Network Sketcher CLI Worker – executed as a subprocess by ns_mcp_server.py.

Two modes, selected by argv:

Persistent mode (``_ns_cli_worker.py <engine_dir> <online_dir>``)
    Imports the engine once at startup, announces readiness, then serves
    length-prefixed JSON requests from stdin until stdin closes or a
    shutdown request arrives.  This keeps the ~1.2 s interpreter + import
    cost out of every MCP tool call.

One-shot mode (no argv)
    Reads a single JSON payload from stdin, runs the commands, writes a
    JSON array to stdout and exits.  Kept as the fallback path used by
    ns_mcp_server._run_batch_oneshot.

Frame format (persistent mode, both directions):
    a decimal byte count followed by '\\n', then that many bytes of UTF-8
    JSON.  Requests carry a monotonically increasing ``id`` which is
    echoed in the response so the parent can detect a desynchronised
    stream.

Request types:
    {"id": 1, "type": "cli", "commands": [[...], ...]}
    {"id": 2, "type": "ping"}
    {"id": 3, "type": "shutdown"}

stderr is left open so Python tracebacks reach the parent process.
"""

import json
import os
import sys
import traceback


# ---------------------------------------------------------------------------
# Framing helpers
# ---------------------------------------------------------------------------

# Upper bound on a single frame, so a desynchronised stream cannot turn a
# garbage length into an unbounded allocation. Well above any real request.
_MAX_FRAME_BYTES = 64 * 1024 * 1024


def _write_frame(out, obj) -> None:
    raw = json.dumps(obj, ensure_ascii=False).encode('utf-8')
    out.write(str(len(raw)).encode('ascii') + b'\n')
    out.write(raw)
    out.flush()


def _read_exact(fh, n: int):
    """Read exactly n bytes, or return None if the stream ends early."""
    chunks = []
    remaining = n
    while remaining > 0:
        chunk = fh.read(remaining)
        if not chunk:
            return None
        chunks.append(chunk)
        remaining -= len(chunk)
    return b''.join(chunks)


def _read_frame(fh):
    """Read one frame, or return None at end of stream."""
    header = fh.readline()
    if not header:
        return None
    header = header.strip()
    if not header:
        return None
    try:
        length = int(header)
    except ValueError:
        raise ValueError(f'invalid frame header: {header!r}')
    if not 0 <= length <= _MAX_FRAME_BYTES:
        raise ValueError(f'frame length {length} outside 0..{_MAX_FRAME_BYTES}')
    body = _read_exact(fh, length)
    if body is None:
        return None
    return json.loads(body.decode('utf-8'))


# ---------------------------------------------------------------------------
# Engine helpers
# ---------------------------------------------------------------------------

def _clear_xlsx_cache() -> None:
    """Drop every cached .nsm -> .xlsx conversion.

    nsm_adapter._nsm_xlsx_cache is keyed by path alone and is only
    invalidated when *this* process writes the master.  A long-lived
    worker would therefore keep serving a stale conversion after the
    master is replaced from the outside (import_master and
    create_empty_master both write masters from the parent process).
    Clearing per request removes the whole class of bug; the cache only
    feeds the pptx / xlsx export paths, so rebuilding it is rare.
    """
    try:
        from ns_engine import nsm_adapter
    except ImportError:
        return
    cache = getattr(nsm_adapter, '_nsm_xlsx_cache', None)
    if not cache:
        return
    for path in list(cache.keys()):
        try:
            nsm_adapter.invalidate_nsm_cache(path)
        except Exception:
            cache.pop(path, None)


def _warm_up(run_cli, engine_dir: str) -> None:
    """Pay the engine's lazy imports at startup instead of on first use.

    ``import nsm_cli`` alone is not enough: pandas, pyarrow, nsm_def and
    pptx (~0.5 s together) are only pulled in once a command actually
    opens a master.  Pointing a read-only command at a path that cannot
    exist reaches that code and then fails on open, so nothing is read
    or written.
    """
    import tempfile
    missing = os.path.join(tempfile.gettempdir(),
                           f'__ns_warmup_{os.getpid()}__.nsm')
    try:
        run_cli(['show', 'area', '--master', missing], cwd=engine_dir)
    except Exception:
        pass


def _run_commands(run_cli, engine_dir, commands):
    results = []
    for args in commands:
        result = run_cli(list(args), cwd=engine_dir)
        results.append({
            'returncode': result.returncode,
            'stdout': result.stdout,
            'stderr': result.stderr,
        })
    return results


# ---------------------------------------------------------------------------
# Persistent mode
# ---------------------------------------------------------------------------

def serve(engine_dir: str, online_dir: str) -> None:
    # Move both ends of the framing channel onto private descriptors before
    # any engine code runs.  A one-shot worker could tolerate the engine
    # touching stdio (it corrupts a single response); a persistent one
    # cannot, because the stream would stay broken for the rest of the
    # session.  Concretely: run_cli closes sys.stdin while executing a
    # command, and anything printing to fd 1 would desynchronise the
    # response stream.
    result_fd = os.dup(1)
    os.dup2(2, 1)                    # stray writes to fd 1 land on stderr
    out = os.fdopen(result_fd, 'wb')

    request_fd = os.dup(0)
    _devnull = os.open(os.devnull, os.O_RDONLY)
    os.dup2(_devnull, 0)             # engine sees an empty stdin
    os.close(_devnull)
    inp = os.fdopen(request_fd, 'rb')

    sys.path.insert(0, engine_dir)
    sys.path.insert(0, online_dir)

    try:
        from ns_engine.nsm_adapter import bootstrap, run_cli
        bootstrap()
        _warm_up(run_cli, engine_dir)
    except Exception as e:
        _write_frame(out, {
            'ready': False,
            'error': f'{type(e).__name__}: {e}',
            'traceback': traceback.format_exc(),
        })
        sys.exit(1)

    _write_frame(out, {'ready': True, 'pid': os.getpid()})

    while True:
        try:
            request = _read_frame(inp)
        except (ValueError, json.JSONDecodeError) as e:
            sys.stderr.write(f'[WORKER] malformed request: {e}\n')
            return
        if request is None:
            # End of the request pipe. This is the safety net against
            # orphaning: when the MCP server is force-killed, atexit never
            # runs in the parent, but the pipe closes and we exit here.
            return

        req_id = request.get('id')
        req_type = request.get('type', 'cli')

        if req_type == 'shutdown':
            _write_frame(out, {'id': req_id, 'results': []})
            return
        if req_type == 'ping':
            _write_frame(out, {'id': req_id, 'results': []})
            continue

        try:
            _clear_xlsx_cache()
            results = _run_commands(run_cli, engine_dir,
                                    request.get('commands') or [])
            _write_frame(out, {'id': req_id, 'results': results})
        except Exception as e:
            _write_frame(out, {
                'id': req_id,
                'error': f'{type(e).__name__}: {e}',
                'traceback': traceback.format_exc(),
            })


# ---------------------------------------------------------------------------
# One-shot mode (fallback)
# ---------------------------------------------------------------------------

def run_once() -> None:
    payload = json.loads(sys.stdin.buffer.read())
    engine_dir: str = payload['engine_dir']
    online_dir: str = payload['online_dir']
    commands: list = payload['commands']  # list of list[str]

    sys.path.insert(0, engine_dir)
    sys.path.insert(0, online_dir)

    from ns_engine.nsm_adapter import bootstrap, run_cli  # noqa: E402
    bootstrap()

    results = _run_commands(run_cli, engine_dir, commands)

    # Write JSON result to the REAL stdout (not sys.stdout which run_cli may
    # have temporarily redirected – it's always restored, but writing to the
    # raw buffer is safer here).
    output = json.dumps(results, ensure_ascii=False)
    sys.stdout.buffer.write(output.encode('utf-8'))
    sys.stdout.buffer.flush()


def main() -> None:
    if len(sys.argv) >= 3:
        serve(sys.argv[1], sys.argv[2])
    else:
        run_once()


if __name__ == '__main__':
    main()
