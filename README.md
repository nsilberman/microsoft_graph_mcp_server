# Microsoft Graph MCP Server

<div align="center">

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Microsoft Graph API](https://img.shields.io/badge/Microsoft%20Graph-API-0078D4.svg)](https://developer.microsoft.com/en-us/graph)
<br>**🚀 A powerful Model Context Protocol (MCP) Server for Microsoft 365 integration**
<br><br>
[![GitHub stars](https://img.shields.io/github/stars/marlonluo2018/microsoft_graph_mcp_server)](https://github.com/marlonluo2018/microsoft_graph_mcp_server/stargazers)<br>
*Seamlessly connect your AI assistants with Microsoft Graph API - Email, Calendar, Teams, OneDrive, and more!*

[Features](#-features) • [Quick Start](doc/QUICK_START.md) • [Documentation](#-documentation) • [Examples](examples/)

</div>

---

## 🌟 Why This Project?

This MCP server brings the full power of Microsoft 365 to your AI workflows:

- ✅ **Zero Configuration** - Device flow authentication, no Azure app registration required
- 🔒 **Secure** - Token management with automatic refresh and secure storage
- ⚡ **Fast** - Optimized batch operations and intelligent caching
- 📦 **Comprehensive** - 30+ tools covering email, calendar, contacts, files, and Teams
- 📖 **Well-Documented** - Extensive guides and examples for every feature
- 🛡️ **Production-Ready** - Robust error handling, rate limiting, and validation

---

## 📚 Features


## Features

### 🔐 Authentication and User Settings
- Interactive device code flow authentication with Microsoft Graph (no Azure app registration required)
- User settings management including timezone, search days, and page sizes
- Authentication status checking and token management
- Token refresh with `extend_token` action to extend sessions without re-login
- Support for custom Azure app registration via `.env` configuration

### 📧 Email Management
- Search and browse emails with advanced filtering (by sender, subject, body)
- Configurable search range (default: 90 days, adjustable via `DEFAULT_SEARCH_DAYS`)
- Compose, reply, and forward emails with HTML support
- Inline attachment support for embedded images and content
- Manage email folders (create, delete, rename, move)
- Bulk email operations (move, delete, archive, flag, categorize)
- Email caching with pagination for efficient browsing
- Timezone-aware email timestamps and filtering

### 📝 Template Management *(Experimental)*
> **Note:** Template management is still experimental. We're actively improving this feature based on user feedback.

- Create email templates from existing emails stored in a Templates folder
- Browse templates with pagination support
- View templates in simple text or full HTML format
- Update templates while preserving HTML formatting and structure
- Send templates while preserving the original (creates a copy to send)
- Soft delete functionality (moves to Deleted Items folder for recovery)
- Smart update workflow:
  - Users view simple text content for easy reading
  - Users provide natural language update instructions to LLM
  - LLM retrieves full HTML and applies changes while maintaining formatting
  - Users verify changes in simple text before sending
- Ideal for recurring emails like newsletters, meeting reminders, and status reports

### 📅 Calendar Management
- Search and browse calendar events with pagination
- Create, update, and cancel your own events
- Create recurring events with flexible patterns (daily, weekly, monthly, yearly)
- Respond to events organized by others (accept, decline, tentatively accept, propose new time)
- Accept/decline entire recurring series with `series` parameter
- Delete cancelled events from your calendar
- Check attendee availability for scheduling
- Forward events and reply to event attendees
- Timezone-aware event scheduling
- Smart all-day event filtering (events ending at midnight are correctly excluded from next day)

### 👥 Contact Management
- Search organization directory for people by name or email address
- Returns contact information (name, email, etc.) from your organization
- Configurable search limit (default: 10)
- Automatic rate limiting with exponential backoff and retry logic
- Clear error messages with retry-after information when rate limits are exceeded

### 📁 File and Team Management *(Workflows in Development)*
- 🚀 **list_files** - List OneDrive files and folders
- 🚀 **get_teams** - Get Teams you're a member of *(No workflow example yet)*
- 🚀 **get_team_channels** - Get channels for a Team *(No workflow example yet)*

> **Note:** File and Teams tools are functional but don't have complete workflow examples yet. Contributions welcome!


### Performance Optimizations
- Efficient bulk email operations with batch processing
- Optimized email search with configurable limits
- Timezone-aware date and time handling
- Disk-based caching for persistence
- API propagation delay handling for reliable operations

### Architecture Improvements
- **Input Validation Layer**: Comprehensive validation system for all tool inputs with clear error messages
  - Validates email addresses, cache numbers, and required/optional strings
  - Early error detection prevents processing invalid inputs
  - Gradual rollout to critical handlers first, then expanded to all handlers
  - See [VALIDATION.md](doc/VALIDATION.md) for complete documentation
- **Optimized Tool Routing**: O(1) dispatch table instead of O(n) if/elif chain
  - Faster tool execution with dictionary-based lookup
  - Easier to maintain when adding/removing tools
  - 85% reduction in routing code (65 lines ?10 lines)
- **Code Quality Improvements**:
  - Removed dead `create_event` tool code (legacy duplicate functionality)
  - Added comprehensive unit tests for dispatch table and validation
  - Better separation of concerns with centralized validators



---

## 📊 Performance Benchmarks

| Operation | Volume | Time | Rate |
|-----------|--------|------|------|
| Email Search | 100 emails | 2.2s | 45/s |
| Email Search | 500 emails | 3.5s | 143/s |
| Email Search | 1000 emails | 5.4s | 186/s |
| Bulk Move | 50 emails | ~3s | 17/s |
| Bulk Move | 100 emails | ~3s | 33/s |
| Bulk Move | 1000 emails | ~10s | 100/s |

---


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

### Advanced Configuration: Token Lifetime

By default, Microsoft Entra ID issues access tokens with a **1-hour lifetime**. You have two options to extend your session:

#### Option 1: Use `extend_token` Action (Recommended for Most Users)

The easiest way to refresh your session is to use the `extend_token` action:

```bash
# Refresh the access token
auth action="extend_token"
```

This action:
- Uses the refresh token to obtain a new access token
- Gives you a fresh token with a new 1-hour lifetime starting from the time you call it (does NOT extend the old token's expiry time)
- Works without requiring user login or admin privileges
- Can be called multiple times to refresh further

#### Option 2: Configure Microsoft Entra ID Token Lifetime Policy (Requires Admin Access)

If you have Microsoft Entra ID admin privileges, you can configure a **Token Lifetime Policy** to extend the default access token lifetime beyond 1 hour (up to maximum 24 hours).

**Steps to Configure:**

1. **Visit Microsoft Entra Admin Center:**
   - Go to: https://entra.microsoft.com/
   - Sign in with your admin account

2. **Create a Token Lifetime Policy:**
   - Navigate to **Identity** ?**Applications** ?**App registrations**
   - Find your application
   - Configure token lifetime policies through the portal interface

3. **Documentation:**
   - Full configuration guide: https://docs.azure.cn/zh-cn/entra/identity-platform/configurable-token-lifetimes
   - Alternative link: https://learn.microsoft.com/zh-cn/entra/identity-platform/configurable-token-lifetimes

**Important Notes:**
- This requires **Microsoft Entra ID admin privileges**
- This is a **one-time configuration** - no code changes needed
- Maximum configurable access token lifetime is **24 hours**
- After configuration, new authentication flows will automatically use the extended lifetime
- The `extend_token` action still works for further extensions if needed

**Recommendation:** Most users should use the `extend_token` action (Option 1) as it doesn't require admin access and provides the same result - longer token lifetime.

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
5. **(Optional) Extend your session** - Access tokens expire after 1 hour by default. Use `auth action="extend_token"` to refresh your access token without requiring login again. This provides a fresh 1-hour token starting from the time you call it.
6. **(Optional) Call `check_status`** - Check your authentication state and token expiry anytime without triggering actions

### Available Tools

#### Authentication Tools
- **auth** - Manage authentication with Microsoft Graph. Supports five actions:
  - `login`: Initiates device code flow authentication. Returns a verification URL and user code to complete authentication in your browser. Previous tokens are cleared on new login.
  - `complete_login`: Completes the login process after browser authentication. **Mandatory step** - must be called after login to verify authentication status and finalize the login. The device_code is automatically loaded from the latest login session.
  - `check_status`: Checks current authentication state and token expiry without triggering any actions (read-only). Returns authentication status, token expiry time, remaining time, refresh token availability, and user timezone. Useful for debugging and monitoring.
  - `extend_token`: Refreshes the access token using the refresh token without requiring user login. This provides a fresh access token with a new 1-hour lifetime starting from the time you call it (does NOT extend the old token's expiry time). This is useful when your token is about to expire and you want to refresh your session without going through the login process again. By default, access tokens expire after 1 hour. Example: `auth action="extend_token"` to refresh your token.
  - `logout`: Clears authentication tokens and signs out from Microsoft Graph.

#### User and Contact Management
- **user_settings** - Manage user settings with two actions: 'init' to sync USER_TIMEZONE and set default values (DEFAULT_SEARCH_DAYS=90, PAGE_SIZE=5, LLM_PAGE_SIZE=20), or 'update' to allow user to update USER_TIMEZONE, DEFAULT_SEARCH_DAYS, PAGE_SIZE, and LLM_PAGE_SIZE. Note: Both actions require login - user_info and LLM settings will only be returned when authenticated.
- **search_contacts** - Search for people by name or email address in organization directory. Returns contact information (name, email, etc.). Use this when you need to find information about a person, such as 'who is John Smith' or 'find contact with email john@company.com'. This is NOT for searching email messages - use search_emails for that. Uses smart detection to automatically choose the optimal search method: email addresses use fast `$filter` with exact match (~10x faster), while name searches use `$search` with tokenization for contains-like behavior. Results are limited (default: 10). Response includes: contacts array, count (number of contacts returned), limit_reached (boolean), and message. If more results exist, limit_reached will be true - use more specific search terms to narrow results. Note: If you encounter a rate limit error (429), the response will include a 'retry_after' field indicating how many seconds to wait before retrying.

#### Email Management
- **manage_mail_folder** - Manage mail folders. Supports list, create, delete, rename, get_details, and move operations
- **manage_emails** - Manage emails with multiple actions. Supports moving, deleting, archiving, flagging, and categorizing emails. Actions include: move_single, move_all, delete_single, delete_multiple, delete_all, archive_single, archive_multiple, flag_single, flag_multiple, categorize_single, categorize_multiple
- **search_emails** - Search or list emails by keywords, sender, subject, or body. Returns matching emails with summary information. If no search_type and query are provided, lists emails within the specified time range. All time parameters use your local timezone. When using time_range, the response includes a user-friendly display string (e.g., 'Today', 'This Week', 'This Month'). Note: Subject and body searches use exact substring matching (contains) for precise results, while sender searches use fuzzy matching. The search_type parameter supports: "sender" (search by sender name/email), "subject" (search by subject text), and "body" (search by body content).
- **browse_email_cache** - Browse emails in the cache with pagination. Returns summary information with number column indicating position in cache. Use page_number to navigate. Automatically manages browsing state with disk cache for persistence.
- **get_email_content** - Get full email content by cache number. Use the cache number from browse_email_cache (e.g., 1, 2, 3) to retrieve complete email with body, attachments, and all details.
- **send_email** - Unified tool for composing, replying to, and forwarding emails. Supports multiple recipients, CC, and BCC. The htmlbody parameter accepts HTML format for rich email content.
- **manage_templates** - Manage email templates stored as drafts in a Templates folder. Actions include: create_from_email (copy an email as a template), list (browse templates), get (view template details with simple text or full HTML), update (edit template content), delete (remove a template - soft delete, moves to Deleted Items folder), and send (send a template while preserving the original). The template update workflow allows users to view simple text, provide update instructions to an LLM, and have the LLM apply changes to the full HTML while preserving formatting.

#### Calendar Management
- **browse_events** - Browse calendar events in the cache with pagination. Returns summary information with number column indicating position in cache. Use page_number to navigate. Automatically manages browsing state with disk cache for persistence. Use search_events to load events into the cache first.
- **get_event_detail** - Get detailed information for a specific calendar event by its cache number. Use the cache number from browse_events (e.g., 1, 2, 3) to retrieve complete event details.
- **search_events** - Search or list calendar events by keywords. Returns matching events with summary information. If no query is provided, lists events within the specified time range. All time parameters use your local timezone. When using time_range, the response includes a user-friendly display string (e.g., 'Today', 'This Week', 'This Month'). The search_type parameter supports: "subject" (search by event subject) and "organizer" (search by organizer name with fuzzy matching).
- **check_attendee_availability** - Check availability of attendees for a given date. Automatically includes the organizer (you) in the availability check to ensure overlap-free time slots. Automatically calculates time range based on all attendees' working hours. Returns availability view string and schedule items for each attendee. Useful for finding optimal meeting times when creating or updating events. Availability view string uses single-character codes for each time interval: 0=Free, 1=Tentative, 2=Busy, 3=Out of office (OOF), 4=Working elsewhere, ?=Unknown. Timezone defaults to user's mailbox settings, but can be explicitly specified.
- **respond_to_event** - Respond to calendar events organized by others with multiple actions: accept, decline, tentatively accept, propose new time, and delete cancelled events. This tool is for responding to events that you are invited to, not events you organized yourself. Supports accepting/declining/tentatively accepting entire recurring series using the `series` parameter. The `delete` action is specifically for removing cancelled events from your calendar that were organized by others. Use the cache number from browse_events or search_events results.
- **manage_my_event** - Manage your own calendar events with multiple actions: create, update, cancel, forward, and reply. This tool is for events that you organized yourself, not events you were invited to. Use this tool to create new events, update or cancel events you created, forward events to others, or reply to event attendees. Use the cache number from browse_events or search_events results.

#### File and Team Management
- **list_files** - List files and folders in OneDrive
- **get_teams** - Get list of Teams that you are a member of
- **get_team_channels** - Get channels for a specific Team

### Template Management Workflow

The template management system allows you to create, edit, and send email templates stored as drafts in a Templates folder. Templates are useful for emails you send regularly with similar content.

#### Template Actions

1. **Create from Email** - Copy an existing email as a template
2. **List** - Browse all templates with pagination
3. **Get** - View template details (simple text or full HTML)
4. **Update** - Edit template content
5. **Delete** - Remove a template (soft delete - moves to Deleted Items folder)
6. **Send** - Send a template while preserving the original

#### Template Update Workflow

The smart update workflow allows users to view simple text content and provide natural language update instructions, while the LLM handles the full HTML formatting:

**Step 1: User views template in simple text**
```json
{
  "action": "get",
  "template_number": 1,
  "text_only": true
}
```

**Step 2: User provides update instructions to LLM**
Example: "Change the meeting time to 2:00 PM and update the agenda to include the new project discussion"

**Step 3: LLM retrieves full HTML**
```json
{
  "action": "get",
  "template_number": 1,
  "text_only": false
}
```

**Step 4: LLM updates and saves complete HTML**
```json
{
  "action": "update",
  "template_number": 1,
  "htmlbody": "<html>...</html>"
}
```

**Step 5: User verifies changes in simple text**
```json
{
  "action": "get",
  "template_number": 1,
  "text_only": true
}
```

**Step 6: User gives command to send**
Example: "Send this template to john@example.com"

**Step 7: LLM sends template and saves a copy**
```json
{
  "action": "send",
  "template_number": 1,
  "to": ["john@example.com"]
}
```

#### Key Features

- **HTML Formatting Preservation**: The LLM updates the full HTML while maintaining the original formatting, styles, and structure
- **Simple Text View**: Users can view templates in simple text format for easy reading and editing instructions
- **Original Preservation**: Sending a template creates a copy and sends it, leaving the original template unchanged
- **Soft Delete**: Deleted templates are moved to the Deleted Items folder and can be recovered

#### Example: Creating a Weekly Newsletter Template

1. **Create from an existing email:**
```json
{
  "action": "create_from_email",
  "cache_number": 5
}
```

2. **View the template:**
```json
{
  "action": "get",
  "template_number": 1,
  "text_only": true
}
```

3. **Update the template:**
```json
{
  "action": "update",
  "template_number": 1,
  "htmlbody": "<html><body><h1>Weekly Newsletter - Week 52</h1><p>Hello Team,</p>...</body></html>"
}
```

4. **Send the template:**
```json
{
  "action": "send",
  "template_number": 1,
  "to": ["team@example.com"]
}
```

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
- **Folder ID Caching**: Eliminates repeated folder lookups (0.63s ?0.00s)
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

The email search functionality has been significantly optimized with server-side filtering and query optimization for dramatic performance improvements:

**Performance Improvements**:
- **Date-filtered searches**: ~90% faster (server-side filtering vs client-side)
- **Sender/recipient searches**: ~70% faster (targeted $filter vs full-text $search)
- **Common folder searches**: ~50% faster (well-known folder cache)
- **General searches**: ~30-50% faster (combined filters)

**Benchmark Results**:
- **100 emails**: 2.22 seconds (45.1 emails/second)
- **500 emails**: 3.50 seconds (142.7 emails/second)
- **1000 emails**: 5.39 seconds (185.6 emails/second)

**Key Optimizations**:

1. **Server-Side Date Filtering**
   - **Before**: Fetched all emails, then filtered by date in Python
   - **After**: Added date filters to API `$filter` parameter
   - **Impact**: Reduces data transfer and processing time by ~90% for date-filtered searches

2. **Replaced $search with $filter**
   - **Before**: Used slow `$search` for full-text search across all fields
   - **After**:
     - Sender search: `contains(from/emailAddress/address, '{sender}')` - targeted field filtering (eq doesn't work for sender email)
     - Recipient search: `toRecipients/any(r: r/emailAddress/address eq '{recipient}')`
     - Subject/Body: `contains(subject, '{query}')` or `contains(body, '{query}')`
   - **Impact**: ~70% faster for field-specific searches

3. **Well-Known Folder Cache**
   - **Before**: Always called `_get_folder_id_by_path()` requiring API calls
   - **After**: Added cache for common folders (Inbox, Sent, Drafts, Deleted, Archive, Junk)
   - **Impact**: Eliminates API calls for standard folders (~50% faster)

4. **Combined Filter Expressions**
   - **Before**: Date filtering done separately after API call
   - **After**: Combined all filters in single `$filter` expression
   - **Impact**: Reduces network round trips

5. **Other Optimizations**:
   - **Hard Limit**: Maximum MAX_EMAIL_SEARCH_LIMIT (1000) emails per search to prevent excessive resource usage
   - **Reduced API Response**: Only essential fields requested (10 fields instead of 18), reducing response size by ~40%
   - **Efficient Processing**: Direct list comprehension for email summary creation (no thread overhead)
   - **Cached Timezone Objects**: ZoneInfo objects cached to avoid redundant timezone conversions
   - **Scalable Performance**: Processing rate improves with larger batches due to amortized overhead

**Safety Features**:
- All email search methods enforce MAX_EMAIL_SEARCH_LIMIT email limit
- Clear error messages when limit is exceeded
- Consistent validation across all search functions (search_emails, search_emails_by_sender, search_emails_by_recipient, search_emails_by_subject, search_emails_by_body, load_emails_by_folder)

### Template Management System

A comprehensive template management system has been implemented for creating, editing, and sending email templates:

**Features**:
- **Template Creation**: Copy any email as a template from the email cache
- **Template Storage**: Templates stored as draft emails in a Templates folder (auto-created if missing)
- **Template Browsing**: List and browse templates with pagination support
- **Template Viewing**: View templates with simple text (user-friendly) or full HTML (for LLM processing)
- **Template Editing**: Update template content while preserving HTML formatting
- **Template Sending**: Send templates while preserving the original (creates a copy and sends it)

**Template Update Workflow**:
The 7-step workflow allows users to easily update templates while maintaining HTML formatting:
1. User calls get with text_only=true - views simple text for easy reading
2. User provides update instructions to LLM - describes desired changes
3. LLM calls get with text_only=false - retrieves full HTML body
4. LLM calls update with htmlbody - applies changes and provides complete updated HTML
5. User calls get with text_only=true - verifies changes in simple text format
6. User gives command to send - approves the template for sending
7. LLM calls send - sends the template and saves a copy in the Templates folder

**Key Benefits**:
- Users can view simple, readable text without HTML complexity
- LLMs can work with full HTML to preserve formatting and structure
- Original templates are never modified when sending (creates copies)
- Automatic folder creation ensures Templates folder always exists
- Disk-based cache provides persistence across sessions

**Example JSON Requests**:
```json
// User views template (simple text)
{"action": "get", "template_number": 1, "text_only": true}

// LLM retrieves full HTML for editing
{"action": "get", "template_number": 1, "text_only": false}

// LLM updates template with complete HTML
{"action": "update", "template_number": 1, "htmlbody": "<html><body>Updated content</body></html>"}

// Send template (preserves original)
{"action": "send", "template_number": 1}
```

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
- [VALIDATION.md](doc/VALIDATION.md) - Input validation system documentation
- [CONTRIBUTING.md](doc/CONTRIBUTING.md) - Contribution guidelines

## License

MIT License
