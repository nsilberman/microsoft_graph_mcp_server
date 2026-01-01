# Example Claude Desktop Configuration for UVX

This file shows how to configure Claude Desktop to use UVX with the Microsoft Graph MCP Server.

## Configuration File Location

- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

## Configuration Examples

### Example 1: UVX with Absolute Path (Recommended for Local Development)

Use an absolute path to your project directory:

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

**Important**: Use forward slashes (`/`) in paths, even on Windows. Do NOT use relative paths like `"."` as they won't work with Claude Desktop.

### Example 2: UVX with Development Dependencies

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["--with", "pytest", "--from", "C:/Project/microsoft_graph_mcp_server", "microsoft-graph-mcp-server"]
    }
  }
}
```

### Example 3: UVX with Specific Python Version

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["--python", "3.11", "--from", "C:/Project/microsoft_graph_mcp_server", "microsoft-graph-mcp-server"]
    }
  }
}
```

### Example 5: Traditional Python Configuration (Fallback)

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "python",
      "args": ["-m", "microsoft_graph_mcp_server.main"]
    }
  }
}
```

### Example 4: Multiple MCP Servers

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["--from", "C:/Project/microsoft_graph_mcp_server", "microsoft-graph-mcp-server"]
    },
    "filesystem": {
      "command": "npx",
      "args": ["-y", "@modelcontextprotocol/server-filesystem", "/path/to/allowed/files"]
    },
    "brave-search": {
      "command": "uvx",
      "args": ["--from", "mcp-brave-search", "mcp-brave-search"]
    }
  }
}
```

### Example 5: UVX with Environment Variables

Note: Environment variables should be set in your system environment, not in the config file.

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["--from", "C:/Project/microsoft_graph_mcp_server", "microsoft-graph-mcp-server"],
      "env": {
        "USER_TIMEZONE": "America/New_York",
        "DEFAULT_SEARCH_DAYS": "90"
      }
    }
  }
}
```

## UVX Command Options Reference

### Basic Options

- `--from <path>`: Specify the package source (local path or PyPI package name)
- `--python <version>`: Use specific Python version (e.g., 3.11)
- `--with <package>`: Include additional packages
- `--no-cache`: Disable caching
- `--clear-cache`: Clear the cache before running

### Development Options

- `--with pytest`: Include pytest for testing
- `--with black`: Include black for code formatting
- `--with mypy`: Include mypy for type checking
- `--editable`: Run in editable mode (development)

### Performance Options

- `--no-sync`: Skip dependency synchronization
- `--frozen`: Use frozen lockfile
- `--no-build-isolation`: Disable build isolation

## Troubleshooting

### JSON Parsing Errors When Running in Terminal

If you see errors like this when running `uvx --from . microsoft-graph-mcp-server` in a terminal:

```
Received exception from stream: 1 validation error for JSONRPCMessage
Invalid JSON: EOF while parsing a value at line 2 column 0
```

**This is normal behavior!** MCP servers communicate via stdio using JSON-RPC protocol and expect to be run by an MCP client like Claude Desktop. When run directly in a terminal, there's no client sending proper JSON messages, so the server reports JSON parsing errors.

**Solution**: Use the server with Claude Desktop (see examples above) or another MCP client. The server is working correctly - it just needs to be run by an MCP client to function properly.

### UVX Command Not Found

If UVX is not found, ensure `uv` is installed:

```bash
pip install uv
```

### Permission Issues

On Windows, you may need to add the UVX path to your PATH or use the full path:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "C:\\Users\\<username>\\.local\\bin\\uvx.exe",
      "args": ["--from", ".", "microsoft-graph-mcp-server"]
    }
  }
}
```

### Dependencies Not Found

If dependencies are not found, ensure your [pyproject.toml](../pyproject.toml) is properly configured with all dependencies listed.

### Path Issues

Use forward slashes in paths, even on Windows:

```json
{
  "args": ["--from", "C:/Project/microsoft_graph_mcp_server", "microsoft-graph-mcp-server"]
}
```

## Additional Resources

- [UVX Usage Guide](UVX_USAGE.md)
- [Main README](../README.md)
- [Installation Guide](INSTALLATION.md)
