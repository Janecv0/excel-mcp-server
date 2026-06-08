# syntax=docker/dockerfile:1
FROM ghcr.io/astral-sh/uv:python3.11-bookworm-slim

WORKDIR /app

# Install dependencies first (better layer caching), then the project itself.
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-install-project --no-dev

COPY . .
RUN uv sync --frozen --no-dev

# Serve over streamable HTTP on the platform's MCP port.
ENV FASTMCP_HOST=0.0.0.0 \
    FASTMCP_PORT=8008 \
    PATH="/app/.venv/bin:$PATH"

EXPOSE 8008

ENTRYPOINT ["excel-mcp-server"]
CMD ["streamable-http"]
