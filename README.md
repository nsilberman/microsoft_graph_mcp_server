# Microsoft Graph MCP Server

A Model Context Protocol (MCP) Server based on Microsoft Graph API, providing comprehensive access to the Microsoft 365 ecosystem.

## Features

### User and Organization Management
- Get basic user information such as profile, photo, and email address
- Query user lists and group information in the organization
- Manage user permissions and group memberships

### Email and Calendar Operations
- Read and send Outlook emails
- Manage calendar events, create, update, and delete meetings
- Query contacts and personal contact information

### File and Document Management
- Access files and folders in OneDrive and SharePoint
- Upload, download, move, and delete files
- Manage file permissions and sharing settings

### Team Collaboration Features
- Access Teams teams and channels
- Manage team members and channel messages
- Create and manage Planner tasks and to-do items

### Data Analysis and Intelligent Insights
- Get user activity data and common file trends
- Analyze meeting requests and collaboration patterns
- Generate personalized data insights and decision support

## Installation

```bash
pip install -r requirements.txt
```

## Configuration

### Interactive Authentication (Recommended)

This server uses device code flow for interactive authentication, without the need for Azure app registration or client secrets.

On first run, you will see a device code that requires:
1. Open your browser and visit `https://microsoft.com/devicelogin`
2. Enter the displayed device code
3. Sign in with your Microsoft account

### Optional Configuration

If you want to use a custom Azure app registration, you can create a `.env` file:

```env
CLIENT_ID=your_custom_client_id
TENANT_ID=organizations
```

**Note**: By default, Microsoft's public client ID is used, and no configuration is required.

## Usage

### Configure in Claude Desktop

Edit the Claude Desktop configuration file:
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

Add the following configuration:

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

**Note**: By default, Microsoft's public client ID and "organizations" tenant are used, no additional configuration is required.

### Usage Steps

1. **First, run the `login` tool** - This will trigger device code flow authentication
2. **Complete authentication as prompted** - Open your browser and enter the device code
3. **Use other tools** - After successful authentication, all tools can be used normally

### Available Tools

- **login** - Authenticate with Microsoft Graph using device code flow (run this tool first)
- **check_login_status** - Check current authentication status
- **logout** - Logout from Microsoft Graph and clear authentication state
- **get_user_info** - Get current user information
- **search_contacts** - Search contacts and people relevant to you
- **browse_emails** - Browse emails in a folder with pagination
- **get_email** - Get full email content by ID
- **search_emails** - Search emails by keywords
- **send_message** - Send an email message
- **get_events** - Get calendar events
- **create_event** - Create a calendar event
- **list_files** - List files and folders from OneDrive
- **get_teams** - Get list of Microsoft Teams
- **get_team_channels** - Get channels for a specific Team

### Direct Run

```bash
# Start MCP Server
python -m microsoft_graph_mcp_server.main

# Or use the installed command
microsoft-graph-mcp-server
```

## Development

```bash
# Install development dependencies
pip install -e .

# Run tests
pytest

# Code formatting
black .
isort .
```

## Documentation

Additional documentation is available in the `doc/` folder:
- [INSTALLATION.md](doc/INSTALLATION.md) - Detailed installation instructions
- [LOGIN_DOCUMENTATION.md](doc/LOGIN_DOCUMENTATION.md) - Authentication guide
- [TEST_README.md](doc/TEST_README.md) - Testing guide
- [CONTRIBUTING.md](doc/CONTRIBUTING.md) - Contribution guidelines

## License

MIT License
