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

### Using pip (Traditional)

```bash
pip install -r requirements.txt
```

### Using UVX (Recommended)

Install the `uv` package manager:

```bash
pip install uv
```

Then run the server directly with UVX:

```bash
uvx .
```

For detailed UVX usage instructions, see [UVX_USAGE.md](doc/UVX_USAGE.md).

### Verify Installation

To verify that the MCP server is properly configured, run:

```bash
python verify_setup.py
```

This will check that all dependencies are installed and the server can be instantiated correctly.

## Configuration

### Interactive Authentication (Recommended)

This server uses device code flow for interactive authentication, without the need for Azure app registration or client secrets.

**Authentication Process:**
1. Call the `login` tool to initiate authentication
2. Open your browser and visit the provided verification URL
3. Enter the displayed user code
4. Sign in with your Microsoft account
5. **IMPORTANT**: Call `complete_login` to verify your authentication status and complete the login process
6. (Optional) Call `check_status` anytime to check your authentication state and token expiry

**Notes:**
- The device_code is automatically saved during login and loaded during complete_login - you don't need to manually handle it
- Previous authentication tokens are cleared when you initiate a new login
- You must call `complete_login` after completing browser authentication to verify your status
- Use `check_status` to monitor token expiry and authentication state without triggering actions

### Optional Configuration

If you want to use a custom Azure app registration, you can create a `.env` file:

```env
CLIENT_ID=your_custom_client_id
TENANT_ID=organizations
USER_TIMEZONE=Asia/Shanghai
DEFAULT_SEARCH_DAYS=90
PAGE_SIZE=5
LLM_PAGE_SIZE=20
CONTACT_SEARCH_LIMIT=10
```

**Configuration Options**:
- `CLIENT_ID`: Custom Azure application client ID (optional)
- `TENANT_ID`: Azure tenant ID (default: "organizations")
- `USER_TIMEZONE`: User's timezone in IANA format (e.g., "Asia/Shanghai", "America/New_York")
- `DEFAULT_SEARCH_DAYS`: Default search range in days for email searches (default: 90)
- `PAGE_SIZE`: Number of items per page for browsing emails (default: 5)
- `LLM_PAGE_SIZE`: Number of items per page for LLM browsing (default: 20)
- `CONTACT_SEARCH_LIMIT`: Maximum contacts to return in search results (default: 10)

**Note**: By default, Microsoft's public client ID is used, and no configuration is required.

## Usage

### Configure in Claude Desktop

Edit the Claude Desktop configuration file:
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

#### Option 1: Using UVX (Recommended)

Use an absolute path to your project directory:

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

#### Option 2: Using Python (Traditional)

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

1. **Call the `login` tool** - This will trigger device code flow authentication and provide a verification URL and user code
2. **Complete authentication in your browser** - Open the provided URL and enter the user code, then sign in with your Microsoft account
3. **Call `complete_login`** - **Mandatory step** to verify your authentication status and complete the login process
4. **Use other tools** - After successful authentication, all tools can be used normally
5. **(Optional) Call `check_status`** - Check your authentication state and token expiry anytime without triggering actions

### Available Tools

#### Authentication Tools
- **auth** - Manage authentication with Microsoft Graph. Supports four actions:
  - `login`: Initiates device code flow authentication. Returns a verification URL and user code to complete authentication in your browser. Previous tokens are cleared on new login.
  - `complete_login`: Completes the login process after browser authentication. **Mandatory step** - must be called after login to verify authentication status and finalize the login. The device_code is automatically loaded from the latest login session.
  - `check_status`: Checks current authentication state and token expiry without triggering any actions (read-only). Returns authentication status, token expiry time, remaining time, and refresh token availability. Useful for debugging and monitoring.
  - `logout`: Clears authentication tokens and signs out from Microsoft Graph.

#### User and Contact Management
- **get_user_info** - Get current user information from Microsoft Graph
- **search_contacts** - Search contacts and people relevant to you. Returns people you interact with most, including organization users and personal contacts. Results are limited (default: 10). Response includes: contacts array, count (number of contacts returned), limit_reached (boolean), and message. If more results exist, limit_reached will be true - use more specific search terms to narrow results.

#### Email Management
- **manage_mail_folder** - Manage mail folders. Supports list, create, delete, rename, get_details, and move operations
- **manage_emails** - Manage emails with multiple actions. Supports moving, deleting, archiving, flagging, and categorizing emails. Actions include: move_single, move_all, delete_single, delete_multiple, delete_all, archive_single, archive_multiple, flag_single, flag_multiple, categorize_single, categorize_multiple
- **search_emails** - Search emails by sender, recipient, subject, or body text. If no search criteria provided, lists recent emails from Inbox (default: 1 day, maximum: 7 days)
- **browse_email_cache** - Browse emails in the cache with pagination (returns current_page and total_pages)
- **get_email_content** - Get full email content by ID (with optional text-only mode)
- **send_email** - Compose, reply to, or forward emails. Supports multiple recipients, CC, and BCC. The htmlbody parameter accepts HTML format for rich email content.

#### Calendar Management
- **browse_events** - Browse calendar events with pagination
- **get_event** - Get full calendar event by ID
- **search_events** - Search calendar events by keywords
- **create_event** - Create a calendar event
- **check_attendee_availability** - Check availability of attendees for a given date. Automatically includes the organizer (you) in the availability check to ensure overlap-free time slots. Automatically calculates time range based on all attendees' working hours. Returns availability view string and schedule items for each attendee. Useful for finding optimal meeting times when creating or updating events. Availability view string uses single-character codes for each time interval: 0=Free, 1=Tentative, 2=Busy, 3=Out of office (OOF), 4=Working elsewhere, ?=Unknown. Timezone defaults to user's mailbox settings, but can be explicitly specified.

#### File and Team Management
- **list_files** - List files and folders from OneDrive
- **get_teams** - Get list of Microsoft Teams
- **get_team_channels** - Get channels for a specific Team

### Direct Run

#### Using UVX (Recommended)

```bash
# Run from local directory
uvx .

# Run from PyPI (when published)
uvx microsoft-graph-mcp-server

# Run with development dependencies
uvx --with pytest .
```

**Important Note**: When running the MCP server directly in a terminal, you may see JSON parsing errors. This is **normal behavior** - MCP servers communicate via stdio using JSON-RPC protocol and expect to be run by an MCP client like Claude Desktop. The server is working correctly; these errors occur because there's no MCP client connected to send proper JSON messages.

#### Using Python (Traditional)

```bash
# Start MCP Server
python -m microsoft_graph_mcp_server.main

# Or use the installed command
microsoft-graph-mcp-server
```

## Development

### Using UVX (Recommended)

```bash
# Install development dependencies
uv sync --dev

# Run tests
uv run pytest

# Code formatting
uv run black .
uv run isort .

# Type checking
uv run mypy .

# Run the server
uv run python -m microsoft_graph_mcp_server.main
# or
uv run microsoft-graph-mcp-server
```

### Using pip (Traditional)

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
- Retrieves the user's timezone from the server's local timezone
- Supports Windows timezone names with automatic conversion to IANA format
- Falls back to the `USER_TIMEZONE` environment variable if needed
- Defaults to UTC if no timezone information is available
- Displays timestamps in a user-friendly format (e.g., "Fri 12/26/2025 09:27 PM")

### Date Range Information
Email list and search tools now return date range information to help users understand the time span of results:
- **date_range**: Shows the actual date range of emails returned (from most recent to oldest)
- **filter_date_range**: Shows the filter applied to the search (e.g., "last 90 days")
- **timezone**: Shows the user's timezone for reference

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

### Configurable Search Range
Email search functions now support a configurable default search range:
- Default search range: 90 days (configurable via `DEFAULT_SEARCH_DAYS` environment variable)
- Override default by specifying `days` parameter in search functions
- Makes searches more efficient and predictable by limiting the time range

### Inline Attachment Support
When replying to emails, the system now properly handles inline attachments (embedded images):
- **Automatic Detection**: Inline attachments are identified by the `isInline` property
- **Content Preservation**: Original attachment content (`contentBytes`) and content ID (`contentId`) are preserved
- **HTML Rendering**: Embedded images referenced with `cid:` (Content-ID) are properly rendered in email replies
- **Thread Continuity**: Original email content including inline images is maintained in reply threads
- **Error Handling**: Failed attachment fetches are logged without breaking the reply process

The implementation uses Microsoft Graph API's separate endpoint calls to retrieve complete attachment data:
1. Fetch email with basic attachment metadata using `$expand=attachments($select=id,name,contentType,isInline)`
2. For each inline attachment, call `/me/messages/{message_id}/attachments/{attachment_id}` to get `contentBytes` and `contentId`
3. Re-attach inline attachments to the reply email with preserved metadata

This ensures that emails with embedded images, logos, or other inline content maintain their visual integrity when replied to.

### Folder and Email Deletion
The system now provides robust folder and email deletion functionality following standard email client conventions:

- **Soft Deletion**: Both folders and emails are moved to the Deleted Items folder rather than being permanently deleted
- **Recoverable**: Deleted items can be recovered from the Deleted Items folder if needed
- **Conflict Handling**: Folder deletion automatically handles naming conflicts by removing existing folders with the same name in Deleted Items before moving
- **Consistent Behavior**: Email deletion follows the same pattern as folder deletion for a unified user experience

The implementation ensures that:
1. Folders are moved to Deleted Items using the Microsoft Graph API's move endpoint
2. Emails are moved to Deleted Items using the same approach
3. The email cache is properly updated when emails are deleted
4. API propagation delays are handled with appropriate wait times

### Folder Management and Response Format Standardization
All folder operations now return consistent, user-friendly response formats:

- **Path Field**: All folder operations include a `path` field showing the full folder path (e.g., 'Inbox/Projects')
- **Standardized Metadata**: Responses include `displayName`, `totalItemCount`, `unreadItemCount`, and `childFolderCount`
- **No Internal Fields**: Internal implementation details like `id` and `parentFolderId` are excluded from responses
- **Clear Messages**: Each operation returns a clear success message describing what was accomplished

Affected operations:
- **create_folder**: Returns folder path and metadata
- **get_folder_details**: Returns folder path and detailed information
- **rename_folder**: Returns new folder path and updated metadata
- **move_folder**: Returns new folder path after moving
- **delete_folder**: Returns confirmation of folder move to Deleted Items

This standardization provides a consistent API for folder operations and makes responses easier to parse and display.

### Bulk Email Movement Performance

The `move_all_emails_from_folder` tool provides highly optimized bulk email movement with the following performance characteristics:

**Performance Scaling**:
- **10-20 emails**: ~2-3 seconds
- **50-100 emails**: ~2-3 seconds (5-12x faster per email)
- **200-400 emails**: ~4-6 seconds
- **1000 emails**: ~8-10 seconds

**Key Optimizations**:
- **Batch Operations**: Uses Microsoft Graph API $batch endpoint with 20 emails per batch (API limit)
- **Concurrent Processing**: Up to 20 batches run in parallel using asyncio.gather()
- **Folder ID Caching**: Eliminates repeated folder lookups (0.63s → 0.00s)
- **Connection Pooling**: Reuses HTTP client connections to reduce overhead
- **Efficient Error Handling**: Tracks moved/failed counts and collects errors without stopping

**Performance Benefits**:
- Time per email decreases as volume increases (amortized fixed overhead)
- Consistent ~2-3 second execution time for most practical use cases (up to ~200-400 emails)
- Scales efficiently to thousands of emails with predictable performance

**Example Usage**:
```
Move all emails from 'Inbox/Projects' to 'Archive/2024':
- 50 emails: ~3 seconds (0.06s per email)
- 100 emails: ~3 seconds (0.03s per email)
- 500 emails: ~6 seconds (0.01s per email)
```

### Email Search Performance Optimizations

The email search functionality has been optimized for handling large batches of emails with improved performance and safety:

**Performance Results**:
- **100 emails**: 2.22 seconds (45.1 emails/second)
- **500 emails**: 3.50 seconds (142.7 emails/second)
- **1000 emails**: 5.39 seconds (185.6 emails/second)

**Key Optimizations**:
- **Hard Limit**: Maximum MAX_EMAIL_SEARCH_LIMIT emails per search to prevent excessive resource usage
- **Reduced API Response**: Only essential fields requested (10 fields instead of 18), reducing response size by ~40%
- **Efficient Processing**: Direct list comprehension for email summary creation (no thread overhead)
- **Cached Timezone Objects**: ZoneInfo objects cached to avoid redundant timezone conversions
- **Scalable Performance**: Processing rate improves with larger batches due to amortized overhead

**Safety Features**:
- All email search methods enforce the MAX_EMAIL_SEARCH_LIMIT email limit
- Clear error messages when limit is exceeded
- Consistent validation across all search functions (search_emails, search_emails_by_sender, search_emails_by_recipient, search_emails_by_subject, search_emails_by_body, load_emails_by_folder)

**Example Usage**:
```
Search with different batch sizes:
- 100 emails: 2.22s (45.1 emails/sec)
- 500 emails: 3.50s (142.7 emails/sec)
- 1000 emails: 5.39s (185.6 emails/sec)
```

The optimizations ensure efficient handling of large email batches while maintaining system stability and predictable performance.

## Documentation

Additional documentation is available in the `doc/` folder:
- [QUICK_START.md](doc/QUICK_START.md) - Quick start guide for UVX setup
- [INSTALLATION.md](doc/INSTALLATION.md) - Detailed installation instructions
- [UVX_USAGE.md](doc/UVX_USAGE.md) - UVX usage guide and configuration
- [UVX_CONFIG_EXAMPLES.md](doc/UVX_CONFIG_EXAMPLES.md) - Claude Desktop configuration examples
- [LOGIN_DOCUMENTATION.md](doc/LOGIN_DOCUMENTATION.md) - Authentication guide
- [TEST_README.md](doc/TEST_README.md) - Testing guide
- [INLINE_ATTACHMENTS.md](doc/INLINE_ATTACHMENTS.md) - Inline attachment handling documentation
- [CONTRIBUTING.md](doc/CONTRIBUTING.md) - Contribution guidelines

## License

MIT License
