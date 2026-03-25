<p align="center">
  <img src="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo.png" alt="Excel MCP Server Logo" width="300"/>
</p>

[![PyPI version](https://img.shields.io/pypi/v/excel-mcp-server.svg)](https://pypi.org/project/excel-mcp-server/)
[![Total Downloads](https://static.pepy.tech/badge/excel-mcp-server)](https://pepy.tech/project/excel-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![smithery badge](https://smithery.ai/badge/@haris-musa/excel-mcp-server)](https://smithery.ai/server/@haris-musa/excel-mcp-server)
[![Install MCP Server](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=excel-mcp-server&config=eyJjb21tYW5kIjoidXZ4IGV4Y2VsLW1jcC1zZXJ2ZXIgc3RkaW8ifQ%3D%3D)

A Model Context Protocol (MCP) server that lets you manipulate Excel files without needing Microsoft Excel installed. Create, read, and modify Excel workbooks with your AI agent.

## Features

- 📊 **Excel Operations**: Create, read, update workbooks and worksheets
- 📈 **Data Manipulation**: Formulas, formatting, charts, pivot tables, and Excel tables
- 🔍 **Data Validation**: Built-in validation for ranges, formulas, and data integrity
- 🎨 **Formatting**: Font styling, colors, borders, alignment, and conditional formatting
- 📋 **Table Operations**: Create and manage Excel tables with custom styling
- 📊 **Chart Creation**: Generate various chart types (line, bar, pie, scatter, etc.)
- 🔄 **Pivot Tables**: Create dynamic pivot tables for data analysis
- 🔧 **Sheet Management**: Copy, rename, delete worksheets with ease
- 🔌 **Triple transport support**: stdio, SSE (deprecated), and streamable HTTP
- 🌐 **Remote & Local**: Works both locally and as a remote service

## Usage

The server supports three transport methods:

### 1. Stdio Transport (for local use)

```bash
uvx excel-mcp-server stdio
```

```json
{
   "mcpServers": {
      "excel": {
         "command": "uvx",
         "args": ["excel-mcp-server", "stdio"]
      }
   }
}
```

### 2. SSE Transport (Server-Sent Events - Deprecated)

```bash
uvx excel-mcp-server sse
```

**SSE transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/sse",
      }
   }
}
```

### 3. Streamable HTTP Transport (Recommended for remote connections)

```bash
uvx excel-mcp-server streamable-http
```

**Streamable HTTP transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/mcp",
      }
   }
}
```

## Environment Variables & File Path Handling

### SSE and Streamable HTTP Transports

When running the server with the **SSE or Streamable HTTP protocols**, you should set an output directory env var on the server side:
- Primary: `DOC_OUTPUT_DIR`
- Backward-compatible fallbacks: `EXCEL_FILES_PATH`, then `MCP_OUTPUT_DIR`

The resolved directory is used for reading/writing generated Excel files.
- If not set, it defaults to `./excel_files`.

You can also set the `FASTMCP_PORT` environment variable to control the port the server listens on (default is `8017` if not set).
Optionally, in **streamable HTTP mode**, you can protect all HTTP endpoints (`/mcp` and `/files`) with an API key:
- `EXCEL_MCP_API_KEY`: Required API key value.
- `EXCEL_MCP_API_KEY_HEADER`: Header name to read (default: `x-api-key`).

Optional download URL env vars used in tool responses:
- `DOC_DOWNLOAD_BASE_URL` (primary)
- `MCP_DOWNLOAD_BASE_URL` (fallback)

- Example (Windows PowerShell):
  ```powershell
  $env:DOC_OUTPUT_DIR="E:\MyExcelFiles"
  $env:FASTMCP_PORT="8007"
  $env:EXCEL_MCP_API_KEY="replace-with-secret"
  $env:DOC_DOWNLOAD_BASE_URL="https://your-server-domain"
  uvx excel-mcp-server streamable-http
  ```
- Example (Linux/macOS):
  ```bash
  DOC_OUTPUT_DIR=/path/to/excel_files FASTMCP_PORT=8007 EXCEL_MCP_API_KEY=replace-with-secret DOC_DOWNLOAD_BASE_URL=https://your-server-domain uvx excel-mcp-server streamable-http
  ```

### HTTP File Endpoints (SSE and Streamable HTTP)

When using SSE or Streamable HTTP, the server now exposes:

- `GET /files` - list files currently available under the resolved output directory
- `GET /files/{filename}` - download a specific `.xlsx` file by basename (for example: `/files/report.xlsx`)
- `GET /healthz` - simple health check endpoint (always public)

Example:

```bash
curl -s https://your-server-domain/files
curl -L -o report.xlsx https://your-server-domain/files/report.xlsx
```

If `EXCEL_MCP_API_KEY` is set, include the key header:

```bash
curl -s -H "x-api-key: replace-with-secret" https://your-server-domain/files
curl -L -H "x-api-key: replace-with-secret" -o report.xlsx https://your-server-domain/files/report.xlsx
```

LibreChat streamable-http example with API key:

```yaml
mcpServers:
  excel:
    type: streamable-http
    url: https://your-server-domain/mcp
    headers:
      x-api-key: "${EXCEL_MCP_API_KEY}"
```

### Save and Listing Tools

- `save_excel_file(file_path, source_filename)` copies/saves an existing workbook and returns:
  - `message`
  - `file_path`
  - `file_size_bytes`
  - optional `download_url`
- `list_excel_files(directory="")` lists `.xlsx` files.
  - If directory is empty or `"."` and output dir is configured, it lists the configured output dir.

### Stdio Transport

When using the **stdio protocol**, the file path is provided with each tool call, so you do **not** need to set `EXCEL_FILES_PATH` on the server. The server will use the path sent by the client for each operation.

## Available Tools

The server provides a comprehensive set of Excel manipulation tools. See [TOOLS.md](TOOLS.md) for complete documentation of all available tools.

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=haris-musa/excel-mcp-server&type=Date)](https://www.star-history.com/#haris-musa/excel-mcp-server&Date)

## License

MIT License - see [LICENSE](LICENSE) for details.
