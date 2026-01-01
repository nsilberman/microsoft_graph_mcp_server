# Quick Start Guide

This guide will help you get the Microsoft Graph MCP Server up and running with UVX in minutes.

## Prerequisites

- Python 3.8 or higher
- Microsoft account (personal or work/school)
- Claude Desktop (for using the MCP server)

## Installation Steps

### 1. Install UVX

```bash
pip install uv
```

### 2. Verify the Setup

```bash
python verify_setup.py
```

You should see all checks pass:
```
✅ Package imported successfully
✅ Server class imported successfully
✅ Settings loaded: microsoft-graph-mcp-server
✅ Server instance created successfully
✅ Main entry point imported successfully
```

### 3. Configure Claude Desktop

Edit your Claude Desktop configuration file:

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Linux**: `~/.config/Claude/claude_desktop_config.json`

Add this configuration:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["--from", "C:/Path/To/microsoft_graph_mcp_server", "microsoft-graph-mcp-server"]
    }
  }
}
```

**Important**: Replace `C:/Path/To/microsoft_graph_mcp_server` with your actual project path. Use forward slashes even on Windows.

### 4. Restart Claude Desktop

Close and reopen Claude Desktop. The MCP server will start automatically.

### 5. Authenticate

In Claude Desktop, ask to use the `login` tool:

```
Please use the login tool to authenticate with Microsoft Graph
```

Follow the prompts:
1. You'll see a device code
2. Visit `https://microsoft.com/devicelogin`
3. Enter the code
4. Sign in with your Microsoft account
5. Grant permissions

### 6. Start Using the Server

Once authenticated, you can use all Microsoft Graph tools:

- **Email**: Read, send, search, and manage emails
- **Calendar**: View, create, and manage events
- **Files**: Access OneDrive files and folders
- **Teams**: View teams and channels
- **Contacts**: Search contacts and people

## Common Issues

### Issue: JSON Parsing Errors in Terminal

**Symptom**: When running `uvx --from . microsoft-graph-mcp-server` in a terminal, you see JSON parsing errors.

**Solution**: This is **normal behavior**. MCP servers communicate via stdio and expect to be run by an MCP client like Claude Desktop. The server is working correctly. Use it with Claude Desktop instead of running it directly in a terminal.

### Issue: UVX Command Not Found

**Symptom**: `uvx` command is not recognized.

**Solution**: Make sure you've installed `uv`:
```bash
pip install uv
```

### Issue: Dependencies Not Found

**Symptom**: UVX reports missing dependencies.

**Solution**: Ensure your [pyproject.toml](../pyproject.toml) has the `dependencies` field properly configured. Run the verification script:
```bash
python verify_setup.py
```

## Verification

To verify everything is working:

1. Run the verification script:
   ```bash
   python verify_setup.py
   ```

2. Check that all tests pass

3. Start Claude Desktop and try using a Microsoft Graph tool

## Additional Resources

- [UVX Usage Guide](UVX_USAGE.md) - Detailed UVX documentation
- [Configuration Examples](UVX_CONFIG_EXAMPLES.md) - More configuration options
- [Installation Guide](INSTALLATION.md) - Detailed installation instructions
- [Main README](../README.md) - Full project documentation

## Next Steps

1. ✅ Install UVX
2. ✅ Verify setup with `python verify_setup.py`
3. ✅ Configure Claude Desktop
4. ✅ Restart Claude Desktop
5. ✅ Authenticate with `login` tool
6. ✅ Start using Microsoft Graph tools!

Enjoy using the Microsoft Graph MCP Server with UVX!
