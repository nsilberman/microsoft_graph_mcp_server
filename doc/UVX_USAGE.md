# UVX Usage Guide

This guide explains how to use UVX (from the `uv` package manager) to run the Microsoft Graph MCP Server.

## What is UVX?

UVX is a tool that allows you to run Python applications directly from PyPI or local packages without manual installation. It automatically manages dependencies and provides a fast, isolated execution environment.

## Installation

First, install the `uv` package manager:

```bash
# Using pip
pip install uv

# Or using the official installer (recommended for best performance)
# Visit https://github.com/astral-sh/uv for platform-specific instructions
```

## Running with UVX

### From Local Development

Run the MCP server directly from the local project directory:

```bash
uvx --from . microsoft-graph-mcp-server
```

This will:
- Automatically install all dependencies from [pyproject.toml](../pyproject.toml)
- Run the server using the entry point defined in the project
- Use an isolated environment to avoid conflicts with your system Python

**Important**: When running the MCP server directly in a terminal, you may see JSON parsing errors. This is **normal behavior** - MCP servers communicate via stdio using JSON-RPC protocol and expect to be run by an MCP client like Claude Desktop. The server is working correctly; these errors occur because there's no MCP client connected to send proper JSON messages. To properly test the server, use it with Claude Desktop or another MCP client.

### From PyPI (when published)

Once the package is published to PyPI, you can run it directly:

```bash
uvx microsoft-graph-mcp-server
```

### With Development Dependencies

To run with development dependencies (for testing or debugging):

```bash
uvx --with pytest --with pytest-asyncio --from . microsoft-graph-mcp-server
```

### With Specific Python Version

To use a specific Python version:

```bash
uvx --python 3.11 --from . microsoft-graph-mcp-server
```

## Configuration in Claude Desktop

Update your Claude Desktop configuration file to use UVX:

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["--from", "C:/Project/microsoft_graph_mcp_server", "microsoft-graph-mcp-server"]
    }
  }
}
```

**Important**: 
- Use an absolute path to your project directory
- Use forward slashes (`/`) in paths, even on Windows
- Replace `C:/Project/microsoft_graph_mcp_server` with your actual project path
- Do NOT use relative paths like `"."` as they won't work with Claude Desktop

## Advantages of UVX

1. **Fast Dependency Resolution**: UVX uses the `uv` resolver, which is significantly faster than pip
2. **Isolated Environments**: Each run uses an isolated environment, preventing dependency conflicts
3. **No Manual Installation**: No need to run `pip install -e .` before running
4. **Automatic Caching**: Dependencies are cached locally for faster subsequent runs
5. **Version Pinning**: Easily specify exact versions of dependencies
6. **Cross-Platform**: Works consistently across Windows, macOS, and Linux

## Common UVX Commands

### Run with specific dependency versions

```bash
uvx --from "mcp>=1.0.0,<2.0.0" --from . microsoft-graph-mcp-server
```

### Run with additional packages

```bash
uvx --with rich --with click --from . microsoft-graph-mcp-server
```

### Run in editable mode (development)

```bash
uvx --editable --from . microsoft-graph-mcp-server
```

### Run with environment variables

```bash
USER_TIMEZONE=America/New_York uvx --from . microsoft-graph-mcp-server
```

### Clear cache and reinstall

```bash
uvx --clear-cache --from . microsoft-graph-mcp-server
```

## Troubleshooting

### JSON Parsing Errors When Running in Terminal

If you see errors like this when running `uvx --from . microsoft-graph-mcp-server`:

```
Received exception from stream: 1 validation error for JSONRPCMessage
Invalid JSON: EOF while parsing a value at line 2 column 0
```

**This is normal behavior!** MCP servers communicate via stdio using JSON-RPC protocol and expect to be run by an MCP client like Claude Desktop. When run directly in a terminal, there's no client sending proper JSON messages, so the server reports JSON parsing errors.

**Solution**: Use the server with Claude Desktop or another MCP client. See the "Configuration in Claude Desktop" section above.

### Dependencies not found

If UVX can't find dependencies, ensure your [pyproject.toml](../pyproject.toml) is properly configured:

```toml
[project]
dependencies = [
    "mcp>=1.0.0",
    "msal>=1.20.0",
    # ... other dependencies
]
```

### Entry point not found

If UVX can't find the entry point, verify the console_scripts section in [pyproject.toml](../pyproject.toml):

```toml
[project.scripts]
microsoft-graph-mcp-server = "microsoft_graph_mcp_server.main:main"
```

### Permission issues

On some systems, you may need to add the `--no-cache` flag:

```bash
uvx --no-cache --from . microsoft-graph-mcp-server
```

### Python version issues

Ensure you have Python 3.8+ installed. UVX will automatically use the appropriate Python version:

```bash
# Check available Python versions
uv python list

# Install a specific Python version
uv python install 3.11

# Use a specific version
uvx --python 3.11 --from . microsoft-graph-mcp-server
```

## Development Workflow with UVX

### 1. Install dependencies for development

```bash
uv sync --dev
```

This installs all dependencies including development tools (pytest, black, isort, mypy).

### 2. Run tests

```bash
uv run pytest
```

### 3. Format code

```bash
uv run black .
uv run isort .
```

### 4. Type check

```bash
uv run mypy .
```

### 5. Run the server

```bash
uv run python -m microsoft_graph_mcp_server.main
# or
uv run microsoft-graph-mcp-server
```

## Comparison with Traditional Setup

### Traditional (pip)

```bash
pip install -e .
python -m microsoft_graph_mcp_server.main
```

### With UVX

```bash
uvx --from . microsoft-graph-mcp-server
```

**Benefits:**
- No need to install the package globally
- Faster dependency resolution
- Isolated environment prevents conflicts
- Automatic caching for faster subsequent runs

## Performance

UVX is significantly faster than traditional pip:

- **Dependency resolution**: 10-100x faster
- **Installation**: 2-10x faster
- **Cold start**: ~1-2 seconds (after first run)
- **Warm start**: ~0.1-0.5 seconds (with cached dependencies)

## Additional Resources

- [UV Documentation](https://github.com/astral-sh/uv)
- [UVX Usage Guide](https://github.com/astral-sh/uv/blob/main/README.md#uvx)
- [Project Configuration](../pyproject.toml)
