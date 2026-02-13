# Installation and Configuration Guide

## Prerequisites

- Python 3.8 or higher
- Microsoft account (personal account or work/school account)

## Step 1: Install the Project

### Method 1: Using UVX (Recommended)

UVX is a fast Python package manager that automatically manages dependencies and provides an isolated execution environment.

```bash
# Clone the project
git clone <repository-url>
cd microsoft_graph_mcp_server

# Install the uv package manager
pip install uv

# Run the server directly (UVX will automatically install dependencies)
uvx .
```

### Method 2: Using Traditional pip Installation

```bash
# Clone the project
git clone <repository-url>
cd microsoft_graph_mcp_server

# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
pip install -e .
```

For detailed UVX usage instructions, please refer to [UVX_USAGE.md](UVX_USAGE.md).

## Step 2: Configure Claude Desktop

Edit the Claude Desktop configuration file:
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

### Option 1: Using UVX (Recommended)

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

**Important**: Use forward slashes (`/`) in paths, even on Windows. Replace `C:/Project/microsoft_graph_mcp_server` with your actual project path.

### Option 2: Using Python (Traditional)

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

## Step 3: First Use - Interactive Authentication

### Authentication Using the login Tool

1. In Claude Desktop, first call the **`login`** tool
2. The server will display a device code authentication prompt:

```
======================================================================
MICROSOFT GRAPH AUTHENTICATION
======================================================================

To sign in, use a web browser to open the page:

https://microsoft.com/devicelogin

And enter the code:

ABCD-EFGH-IJKL

======================================================================
Waiting for authentication...
======================================================================
```

3. Open a browser and visit `https://microsoft.com/devicelogin`
4. Enter the device code displayed
5. Sign in with your Microsoft account
6. Grant the necessary permissions
7. After successful authentication, the login tool will return a success message

### Using Other Tools

After successful authentication, you can use the following tools:
- **get_user_info** - Get current user information
- **list_users** - List users in the organization
- **get_messages** - Get emails
- **get_events** - Get calendar events
- **create_event** - Create calendar events
- **list_files** - List OneDrive files and folders
- **get_teams** - Get Microsoft Teams list
- **get_team_channels** - Get channels for a specific team

**Important**: If you call other tools without authentication, they will automatically trigger the authentication flow, but it's recommended to use the login tool first to ensure successful authentication.

## Optional Configuration

### Using Custom Azure App Registration

If you have your own Azure app registration, you can create a `.env` file:

```bash
cp .env.example .env
```

Edit the `.env` file:

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

**Configuration Options**:
- `CLIENT_ID`: Custom Azure application client ID (optional)
- `TENANT_ID`: Azure tenant ID (default: "organizations")
- `USER_TIMEZONE`: User timezone in IANA format (e.g., "Asia/Shanghai", "America/New_York")
- `DEFAULT_SEARCH_DAYS`: Default number of days for email search when not specified (default: 7)
- `MAX_SEARCH_DAYS`: Maximum allowed search range in days (default: 90)
- `PAGE_SIZE`: Number of items to display per page when users browse emails (default: 5)
- `LLM_PAGE_SIZE`: Number of items to display per page when LLM browses emails (default: 20)
- `CONTACT_SEARCH_LIMIT`: Maximum number of contacts to return in search results (default: 10)
- `MAX_BCC_BATCH_SIZE`: Maximum BCC recipients per batch when forwarding emails (default: 500)

**Note**: By default, Microsoft's public client ID is used, no configuration required.

### Environment Variable Configuration

You can also configure through environment variables:

```bash
# Windows PowerShell
$env:CLIENT_ID="your_client_id"
$env:TENANT_ID="organizations"

# Linux/Mac
export CLIENT_ID="your_client_id"
export TENANT_ID="organizations"
```

## Running the Server

### Using UVX (Recommended)

```bash
# Run from local directory
uvx .

# Run from PyPI (when package is published)
uvx microsoft-graph-mcp-server

# Run with development dependencies
uvx --with pytest .
```

**Important Note**: When running the MCP server directly in a terminal, you may see JSON parsing errors. This is **normal behavior** - MCP servers communicate via stdio using the JSON-RPC protocol and expect to be run by an MCP client (like Claude Desktop). The server is working correctly; these errors occur because there is no MCP client connected to send proper JSON messages. To properly test the server, use it with Claude Desktop or another MCP client.

### Using Python (Traditional)

```bash
# Run using Python module
python -m microsoft_graph_mcp_server.main

# Or use the installed command
microsoft-graph-mcp-server
```

## Troubleshooting

### Common Issues

1. **Authentication Failure**:
   - Ensure the device code is entered correctly
   - Check network connection
   - Confirm your Microsoft account has permission to access required resources

2. **Insufficient Permissions**:
   - Ensure all necessary permissions were granted during authentication
   - Some features may require administrator approval

3. **Connection Timeout**:
   - Check network connection and firewall settings
   - Device codes are valid for 15 minutes, re-authentication required after timeout

4. **JSON Parsing Errors When Running in Terminal**:
   
   If you see the following error when running `uvx .` in a terminal:
   
   ```
   Received exception from stream: 1 validation error for JSONRPCMessage
   Invalid JSON: EOF while parsing a value at line 2 column 0
   ```
   
   **This is normal behavior!** MCP servers communicate via stdio using the JSON-RPC protocol and expect to be run by an MCP client (like Claude Desktop). When running directly in a terminal, there is no client to send proper JSON messages, so the server reports JSON parsing errors.
   
   **Solution**: Use the server with Claude Desktop or another MCP client. Please refer to "Step 2: Configure Claude Desktop" section.

### Debug Mode

Enable verbose logging:

```bash
python -m microsoft_graph_mcp_server.main --verbose
```

### Re-authentication

To switch accounts or re-authenticate, simply delete the cached tokens and restart the server.

## Security Notes

- When using device code flow, there is no need to store a client secret
- Access tokens are automatically refreshed, no need to re-authenticate
- Tokens are cached locally, ensure your computer is secure
- To clear authentication information, delete the local cache file
