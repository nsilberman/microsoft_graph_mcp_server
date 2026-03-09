# Setup and Installation Guide

Complete guide to install, configure, and run the Microsoft Graph MCP Server.

## Prerequisites

- Python 3.8 or higher
- Microsoft account (personal or work/school)
- Claude Desktop (for using the MCP server)

---

## Quick Start

### 1. Install UVX (Recommended)

```bash
pip install uv
```

### 2. Configure Claude Desktop

Edit your Claude Desktop configuration file:

- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

Add this configuration:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["C:/Path/To/microsoft_graph_mcp_server"]
    }
  }
}
```

**Important**: 
- Replace `C:/Path/To/microsoft_graph_mcp_server` with your actual project path
- Use forward slashes (`/`) even on Windows

### 3. Restart Claude Desktop

Close and reopen Claude Desktop. The MCP server will start automatically.

### 4. Authenticate

In Claude Desktop, use the `login` tool:

```
Please use the login tool to authenticate with Microsoft Graph
```

Follow the prompts:
1. You'll see a device code
2. Visit `https://microsoft.com/devicelogin`
3. Enter the code
4. Sign in with your Microsoft account
5. Grant permissions

---

## Installation Methods

### Method 1: UVX (Recommended)

UVX automatically manages dependencies and provides an isolated execution environment.

```bash
# Clone the project
git clone <repository-url>
cd microsoft_graph_mcp_server

# Run the server directly
uvx .
```

**Advantages:**
- No manual dependency installation
- Fast dependency resolution (10-100x faster than pip)
- Isolated environments prevent conflicts
- Automatic caching for faster subsequent runs

### Method 2: Traditional pip

```bash
# Clone the project
git clone <repository-url>
cd microsoft_graph_mcp_server

# Create and activate virtual environment
python -m venv venv
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
pip install -e .
```

---

## Claude Desktop Configuration Examples

### Basic UVX Configuration

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["C:/Project/microsoft_graph_mcp_server"]
    }
  }
}
```

### With Environment Variables

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["C:/Project/microsoft_graph_mcp_server"],
      "env": {
        "USER_TIMEZONE": "America/New_York",
        "DEFAULT_SEARCH_DAYS": "90"
      }
    }
  }
}
```

### With Specific Python Version

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["--python", "3.11", "C:/Project/microsoft_graph_mcp_server"]
    }
  }
}
```

### Traditional Python (Fallback)

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

### Multiple MCP Servers

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["C:/Project/microsoft_graph_mcp_server"]
    },
    "filesystem": {
      "command": "npx",
      "args": ["-y", "@modelcontextprotocol/server-filesystem", "/path/to/files"]
    }
  }
}
```

---

## Environment Configuration

### Optional: Custom Azure App Registration

Create a `.env` file in the project root:

```env
CLIENT_ID=your_custom_client_id
TENANT_ID=organizations
USER_TIMEZONE=Asia/Shanghai
DEFAULT_SEARCH_DAYS=7
MAX_SEARCH_DAYS=90
PAGE_SIZE=5
LLM_PAGE_SIZE=20
CONTACT_SEARCH_LIMIT=10
MAX_BCC_BATCH_SIZE=500
```

**Configuration Options:**

| Variable | Description | Default |
|----------|-------------|---------|
| `CLIENT_ID` | Custom Azure app client ID | Microsoft public client |
| `TENANT_ID` | Azure tenant ID | `organizations` |
| `USER_TIMEZONE` | User timezone (IANA format) | Auto-detected |
| `DEFAULT_SEARCH_DAYS` | Default email search range | 7 |
| `MAX_SEARCH_DAYS` | Maximum search range | 90 |
| `PAGE_SIZE` | Emails per page (user mode) | 5 |
| `LLM_PAGE_SIZE` | Emails per page (LLM mode) | 20 |
| `CONTACT_SEARCH_LIMIT` | Max contacts in search | 10 |
| `MAX_BCC_BATCH_SIZE` | Max BCC recipients per batch | 500 |

---

## Running the Server

### With UVX

```bash
# Run from local directory
uvx .

# Run from PyPI (when published)
uvx microsoft-graph-mcp-server

# Run with development dependencies
uvx --with pytest .

# Run with specific Python version
uvx --python 3.11 .
```

### With Python

```bash
# Run using Python module
python -m microsoft_graph_mcp_server.main

# Or use the installed command
microsoft-graph-mcp-server
```

**Note**: When running directly in a terminal, you may see JSON parsing errors. This is **normal** - MCP servers communicate via stdio and expect to be run by an MCP client like Claude Desktop.

---

## Development Workflow

### Using UV (Recommended)

```bash
# Install dependencies for development
uv sync --dev

# Run tests
uv run pytest

# Format code
uv run black .
uv run isort .

# Type check
uv run mypy .

# Run the server
uv run python -m microsoft_graph_mcp_server.main
```

---

## Troubleshooting

### JSON Parsing Errors in Terminal

**Symptom**: When running `uvx .` in a terminal, you see:
```
Invalid JSON: EOF while parsing a value at line 2 column 0
```

**Solution**: This is normal. MCP servers communicate via stdio and expect to be run by Claude Desktop. Use with Claude Desktop instead.

### UVX Command Not Found

**Solution**: Install uv:
```bash
pip install uv
```

### Dependencies Not Found

**Solution**: Ensure your `pyproject.toml` has dependencies properly configured. Run:
```bash
python verify_setup.py
```

### Authentication Failure

**Solutions**:
- Ensure the device code is entered correctly
- Check network connection
- Confirm your Microsoft account has permission to access required resources

### Permission Issues on Windows

**Solution**: Add UVX path to PATH or use full path:
```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "C:\\Users\\<username>\\.local\\bin\\uvx.exe",
      "args": ["C:/Project/microsoft_graph_mcp_server"]
    }
  }
}
```

---

## Post-Installation

After successful setup, you can use all Microsoft Graph tools:

- **Email**: Read, send, search, and manage emails
- **Calendar**: View, create, and manage events
- **Files**: Access OneDrive files and folders
- **Teams**: View teams and channels
- **Contacts**: Search contacts and people

---

## Related Documentation

- [README.md](../README.md) - Main project documentation
- [EMAIL_FUNCTIONS.md](EMAIL_FUNCTIONS.md) - Email functionality
- [CALENDAR.md](CALENDAR.md) - Calendar functionality
- [ERROR_CODES.md](ERROR_CODES.md) - Error reference
