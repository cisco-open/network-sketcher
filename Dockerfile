# Network Sketcher Local MCP — container image for Glama introspection and stdio runs.
#
# Build:
#   docker build -t network-sketcher-mcp .
#
# Run (stdio MCP transport):
#   docker run --rm -i network-sketcher-mcp
#
# The image bundles network-sketcher_local_mcp plus network-sketcher_online/ns_engine.
# No credentials are embedded; workspace paths are configured at runtime by the MCP host.

FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

COPY network-sketcher_local_mcp/requirements_mcp.txt network-sketcher_local_mcp/requirements_mcp.txt
RUN python -m pip install -r network-sketcher_local_mcp/requirements_mcp.txt

COPY network-sketcher_online/ network-sketcher_online/
COPY network-sketcher_local_mcp/ network-sketcher_local_mcp/

RUN useradd --create-home --shell /bin/bash mcp
USER mcp
WORKDIR /app

ENTRYPOINT ["python", "network-sketcher_local_mcp/ns_mcp_server.py"]
