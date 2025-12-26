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

#### Authentication Tools
- **login** - Authenticate with Microsoft Graph using device code flow (run this tool first)
- **check_login_status** - Check current authentication status and token expiry
- **logout** - Logout from Microsoft Graph and clear authentication state

#### User and Contact Management
- **get_user_info** - Get current user information from Microsoft Graph
- **search_contacts** - Search contacts and people relevant to you

#### Email Management
- **list_mail_folders** - List all mail folders with their paths (e.g., 'Inbox', 'Inbox/Projects', 'Archive/2024')
- **list_recent_emails** - List recent emails from Inbox with optional days parameter (default: 1 day, maximum: 7 days)
- **load_emails_by_folder** - Load emails from a folder into cache with filtering options (by days or top number)
- **browse_email_cache** - Browse emails in the cache with pagination (returns current_page and total_pages)
- **search_emails** - Search emails by sender, recipient, subject, or body text
- **get_email_content** - Get full email content by ID (with optional text-only mode)
- **send_message** - Send an email message
- **clear_email_cache** - Clear the email browsing cache

#### Calendar Management
- **browse_events** - Browse calendar events with pagination
- **get_event** - Get full calendar event by ID
- **search_events** - Search calendar events by keywords
- **create_event** - Create a calendar event

#### File and Team Management
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

## Recent Improvements

### Timezone Support
Email timestamps are now automatically converted to the user's local timezone for better readability. The system:
- Retrieves the user's mailbox timezone from Microsoft Graph
- Falls back to the `USER_TIMEZONE` environment variable if unavailable
- Defaults to UTC if no timezone information is available
- Displays timestamps in a user-friendly format (e.g., "Fri 12/26/2025 09:27 PM")

### Email Sorting and Filtering
Emails are now properly sorted by received date (newest first) with:
- Original ISO datetime preserved for accurate sorting
- Formatted display timestamps for readability
- Local-time-aware filtering when using the `days` parameter
- Consistent ordering across both API responses and cached data

### Enhanced Pagination
The `browse_email_cache` tool now provides clear pagination information:
- `current_page` - Current page number being viewed
- `total_pages` - Total number of pages available
- `count` - Number of emails on current page
- `total_count` - Total number of emails in cache

## Documentation

Additional documentation is available in the `doc/` folder:
- [INSTALLATION.md](doc/INSTALLATION.md) - Detailed installation instructions
- [LOGIN_DOCUMENTATION.md](doc/LOGIN_DOCUMENTATION.md) - Authentication guide
- [TEST_README.md](doc/TEST_README.md) - Testing guide
- [CONTRIBUTING.md](doc/CONTRIBUTING.md) - Contribution guidelines

## License

MIT License
