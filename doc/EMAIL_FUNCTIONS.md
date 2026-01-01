# Email Functions Documentation

This document describes the email-related functions available in the Microsoft Graph MCP Server.

## Table of Contents

1. [List Recent Emails](#list-recent-emails)
2. [Load Emails by Folder](#load-emails-by-folder)
3. [Search Emails](#search-emails)
4. [Browse Email Cache](#browse-email-cache)
5. [Clear Email Cache](#clear-email-cache)
6. [Get Email Content](#get-email-content)
7. [Mail Folder Management](#mail-folder-management)
8. [Move Email](#move-email)

---

## List Recent Emails

### Description
Lists recent emails from the Inbox folder with a configurable time range. This is a convenience function that loads emails from the last N days (default: 1 day, maximum: 7 days) into the cache for browsing.

### Parameters
- `days` (optional, integer): Number of days to look back
  - Default: 1
  - Minimum: 1
  - Maximum: 7

### Returns
```json
{
  "message": "Loaded X recent emails from Inbox (last N day(s))",
  "folder": "Inbox",
  "days": N,
  "count": X,
  "timezone": "Asia/Shanghai",
  "date_range": {
    "from": "Fri 12/27/2025 09:36 PM",
    "to": "Wed 12/24/2025 10:28 PM"
  },
  "hint": "Use browse_email_cache to view the loaded emails"
}
```

### Return Fields
- `message`: Confirmation message
- `folder`: Folder loaded (always "Inbox")
- `days`: Number of days looked back
- `count`: Number of emails loaded
- `timezone`: User's timezone for reference
- `date_range`: Actual date range of emails loaded (from most recent to oldest)
- `hint`: Instructions for viewing results

### Example Usage
```python
# List emails from the last day (default)
result = await list_recent_emails()

# List emails from the last 3 days
result = await list_recent_emails(days=3)

# List emails from the last 7 days (maximum)
result = await list_recent_emails(days=7)
```

### Notes
- Automatically clears the cache before loading new emails
- Loads emails into cache for subsequent browsing
- Use `browse_email_cache` to view the loaded emails

---

## Search Emails

### Description
Unified search tool for emails. Search by sender, recipient, subject, or body text with configurable date range filtering.

### Parameters
- `search_type` (required, string): Type of search to perform
  - Values: "sender", "recipient", "subject", "body"
- `query` (required, string): Search query
  - For sender: sender name or email address
  - For recipient: recipient name or email address
  - For subject: subject text
  - For body: body text content
- `folder` (optional, string): Folder to search (default: all folders)
- `days` (optional, integer): Number of days to search back
  - Default: 90 (configurable via `DEFAULT_SEARCH_DAYS` environment variable)
  - Set to `null` to search all emails (not recommended for large mailboxes)

### Returns
```json
{
  "search_type": "sender",
  "query": "search query",
  "folder": "Inbox",
  "count": X,
  "timezone": "Asia/Shanghai",
  "date_range": {
    "from": "Fri 12/27/2025 09:36 PM",
    "to": "Wed 12/24/2025 10:28 PM"
  },
  "filter_date_range": "last 90 days",
  "hint": "Found X emails. Use browse_email_cache to view the results."
}
```

### Return Fields
- `search_type`: Type of search performed
- `query`: Search query used
- `folder`: Folder searched (or null for all folders)
- `count`: Number of emails found
- `timezone`: User's timezone for reference
- `date_range`: Actual date range of emails returned (from most recent to oldest)
- `filter_date_range`: Filter applied to the search (e.g., "last 90 days")
- `hint`: Instructions for viewing results

### Example Usage
```python
# Search by sender in Inbox (default 90 days)
result = await search_emails(search_type="sender", query="john@example.com", folder="Inbox")

# Search by subject with custom date range
result = await search_emails(search_type="subject", query="meeting", days=30)

# Search by body text in Sent Items with custom date range
result = await search_emails(search_type="body", query="project update", folder="Sent Items", days=60)

# Search all emails (not recommended for large mailboxes)
result = await search_emails(search_type="sender", query="john@example.com", days=None)
```

### Configuration
The default search range can be configured via the `DEFAULT_SEARCH_DAYS` environment variable:

```env
# .env file
DEFAULT_SEARCH_DAYS=90
```

### Notes
- Automatically clears the cache before performing search
- Loads search results into cache for browsing
- Use `browse_email_cache` to view the search results
- Date range filtering makes searches more efficient and predictable
- Setting `days` to `null` will search all emails (may be slow for large mailboxes)

---

## Browse Email Cache

### Description
Browse emails in the cache with pagination. Returns summary information with number column indicating position in cache. The page size is configured via the `PAGE_SIZE` environment variable (default: 5).

### Parameters
- `page_number` (required, integer): Page number to view
  - Minimum: 1

### Configuration
The number of emails per page is controlled by the `PAGE_SIZE` environment variable:

```env
# .env file
PAGE_SIZE=5
```

Default page size is 5 emails per page. This ensures consistent pagination across all browsing operations.

### Returns
```json
{
  "emails": [...],
  "count": X,
  "total_count": Y,
  "page": N,
  "pages": Z
}
```

### Email Structure
Each email in the cache contains:
```json
{
  "number": 1,
  "subject": "Email Subject",
  "from": {
    "name": "Sender Name",
    "email": "sender@example.com"
  },
  "to": [
    {
      "name": "Recipient Name",
      "email": "recipient@example.com"
    }
  ],
  "cc": [],
  "receivedDateTime": "Fri 12/27/2025 09:36 PM",
  "receivedDateTimeOriginal": "2025-12-27T13:36:27Z",
  "isRead": false,
  "hasAttachments": false,
  "hasEmbeddedImages": false,
  "embeddedImageCount": 0,
  "importance": "normal"
}
```

### Return Fields
- `emails`: Array of email summaries
- `count`: Number of emails on current page
- `total_count`: Total emails in cache
- `current_page`: Current page number
- `total_pages`: Total pages available
- `timezone`: User's timezone for reference

### Example Usage
```python
# View first page (default 20 emails per page)
result = await browse_email_cache(page_number=1)

# View second page (uses configured page_size)
result = await browse_email_cache(page_number=2)
```

### Notes
- Requires emails to be loaded into cache first (via list_recent_emails or search_emails)
- Automatically manages browsing state with disk cache for persistence
- Use the `number` field to reference specific emails

---

## Get Email Content

### Description
Get full email content by cache number. Use the email number from browse_email_cache (e.g., 1, 2, 3) to retrieve complete email with body, attachments, and all details.

### Parameters
- `emailNumber` (required, integer): Email number from browse_email_cache (e.g., 1, 2, 3)
- `text_only` (optional, boolean): If true, return only text content without embedded images and attachments. If false, return full content including embedded images and attachments.
  - Default: true

### Returns
User-facing email content:
```json
{
  "emailNumber": 1,
  "subject": "Email Subject",
  "from": {
    "name": "Sender Name",
    "email": "sender@example.com"
  },
  "to": [
    {
      "name": "Recipient Name",
      "email": "recipient@example.com"
    }
  ],
  "cc": [],
  "receivedDateTimeDisplay": "Fri 12/27/2025 09:36 PM",
  "importance": "normal",
  "isRead": false,
  "hasAttachments": false,
  "body": {
    "contentType": "html",
    "content": "<html>...</html>"
  }
}
```

### Return Fields
- `emailNumber`: Email number in cache
- `subject`: Email subject
- `from`: Sender information (name, email)
- `to`: Array of to recipients (name, email)
- `cc`: Array of CC recipients (name, email)
- `receivedDateTimeDisplay`: Formatted local time
- `importance`: Email importance level
- `isRead`: Read status
- `hasAttachments`: Has attachments flag
- `body`: Email body content (contentType, content)
- `attachments`: Array of attachments (only when text_only=false)
  - `name`: Attachment filename
  - `size`: Attachment size in bytes
  - `contentType`: MIME type of the attachment
  - `isInline`: Whether the attachment is inline (embedded in the email body)

### Example Usage
```python
# Get email content with text only (default)
result = await get_email_content(emailNumber=1)

# Get email content with full attachments and images
result = await get_email_content(emailNumber=1, text_only=false)
```

### Notes
- Requires valid email number from cache
- Returns only user-facing content (system metadata is excluded)
- Use `text_only=false` to include attachments and embedded images

---

## Mail Folder Management

### Description
Manage mail folders with support for list, create, delete, rename, get_details, and move operations.

### Parameters
- `action` (required, string): Action to perform
  - Values: "list", "create", "delete", "rename", "get_details", "move"
- `folder_path` (optional, string): Path of the folder (e.g., 'Inbox', 'Archive/2024')
  - Required for: delete, rename, get_details, move actions
- `folder_name` (optional, string): Name of the folder to create
  - Required for: create action
- `parent_folder` (optional, string): Optional parent folder path for create action (e.g., 'Inbox', 'Archive/2024')
  - If not provided, creates a top-level folder
- `new_name` (optional, string): New name for the folder
  - Required for: rename action
- `destination_parent` (optional, string): Path of the destination parent folder (e.g., 'Archive', 'Sent Items')
  - Required for: move action

### Actions

#### List Folders (action="list")
Lists all mail folders with hierarchy including child folders, total item count, and unread item count.

**Parameters:**
- `action`: "list"

**Returns:**
```json
[
  {
    "id": "folder-id",
    "displayName": "Inbox",
    "parentFolderId": null,
    "childFolderCount": 0,
    "totalItemCount": 150,
    "unreadItemCount": 5
  }
]
```

**Example Usage:**
```python
result = await mail_folder(action="list")
```

#### Create Folder (action="create")
Creates a new mail folder.

**Parameters:**
- `action`: "create"
- `folder_name`: Name of the folder to create
- `parent_folder` (optional): Parent folder path

**Returns:**
```json
{
  "id": "new-folder-id",
  "displayName": "New Folder",
  "parentFolderId": "parent-folder-id"
}
```

**Example Usage:**
```python
# Create top-level folder
result = await mail_folder(action="create", folder_name="Projects")

# Create folder under Inbox
result = await mail_folder(action="create", folder_name="2024", parent_folder="Inbox")
```

#### Delete Folder (action="delete")
Deletes a mail folder.

**Parameters:**
- `action`: "delete"
- `folder_path`: Path of the folder to delete

**Returns:**
```json
{
  "message": "Folder deleted successfully"
}
```

**Example Usage:**
```python
result = await mail_folder(action="delete", folder_path="Archive/2023")
```

#### Rename Folder (action="rename")
Renames a mail folder.

**Parameters:**
- `action`: "rename"
- `folder_path`: Path of the folder to rename
- `new_name`: New name for the folder

**Returns:**
```json
{
  "id": "folder-id",
  "displayName": "New Folder Name"
}
```

**Example Usage:**
```python
result = await mail_folder(action="rename", folder_path="Archive/OldName", new_name="NewName")
```

#### Get Folder Details (action="get_details")
Gets detailed information about a mail folder.

**Parameters:**
- `action`: "get_details"
- `folder_path`: Path of the folder

**Returns:**
```json
{
  "id": "folder-id",
  "displayName": "Inbox",
  "parentFolderId": null,
  "childFolderCount": 2,
  "totalItemCount": 150,
  "unreadItemCount": 5
}
```

**Example Usage:**
```python
result = await mail_folder(action="get_details", folder_path="Inbox")
```

#### Move Folder (action="move")
Moves a mail folder to a different parent folder.

**Parameters:**
- `action`: "move"
- `folder_path`: Path of the folder to move
- `destination_parent`: Path of the destination parent folder

**Returns:**
```json
{
  "id": "folder-id",
  "displayName": "Moved Folder",
  "parentFolderId": "new-parent-folder-id"
}
```

**Example Usage:**
```python
result = await mail_folder(action="move", folder_path="Inbox/Projects", destination_parent="Archive")
```

### Notes
- Returns all folders in the mailbox for list action
- Includes folder hierarchy information
- Shows total and unread item counts
- Folder paths use '/' separator for nested folders

---

## Move Email

### Description
Move emails to a different folder. Supports moving a single email or all emails from a folder.

### Parameters
- `action` (required, string): Action to perform
  - Values: "single", "all"
- `email_number` (optional, integer): Email number from browse_email_cache (e.g., 1, 2, 3)
  - Required for: "single" action
- `source_folder` (optional, string): Source folder path (e.g., 'Inbox', 'Archive/2024')
  - Required for: "all" action
- `destination_folder` (required, string): Destination folder path (e.g., 'Archive/2024', 'Inbox/Projects')

### Actions

#### Move Single Email (action="single")
Moves a single email to a different folder.

**Parameters:**
- `action`: "single"
- `email_number`: Email number from browse_email_cache
- `destination_folder`: Destination folder path

**Returns:**
```json
{
  "message": "Email moved successfully",
  "email_id": "email-id"
}
```

**Example Usage:**
```python
result = await move_email(
    action="single",
    email_number=1,
    destination_folder="Archive/2024"
)
```

#### Move All Emails from Folder (action="all")
Moves all emails from a source folder to a destination folder.

**Parameters:**
- `action`: "all"
- `source_folder`: Source folder path
- `destination_folder`: Destination folder path

**Returns:**
```json
{
  "message": "Moved X emails from SourceFolder to DestinationFolder",
  "count": X
}
```

**Example Usage:**
```python
result = await move_email(
    action="all",
    source_folder="Inbox",
    destination_folder="Archive/2024"
)
```

### Notes
- Requires valid email number from cache for "single" action
- Use `browse_email_cache` to get email numbers
- "all" action moves all emails from source folder
- Destination folder must exist (use mail_folder to create if needed)

---

## Cache Management

### Automatic Cache Clearing
The cache is automatically cleared before the following operations:
- `list_recent_emails`
- `search_emails`

### Async Cache Performance
Cache saves to disk are performed asynchronously using threading for optimal performance. This ensures non-blocking I/O operations when saving large amounts of email data.

### Cache Persistence
The cache is persisted to disk at `~/.microsoft_graph_mcp_browsing.json` with the following characteristics:
- Version: 2.0
- Expiry: 1 hour
- Max age: 24 hours
- Automatic cleanup of expired caches

---

## Best Practices

1. **Use list_recent_emails for quick access**: For viewing recent emails, use `list_recent_emails` with the default 1-day parameter.

2. **Specify folder for targeted searches**: Use the `folder` parameter in `search_emails` to narrow down search results.

3. **Use pagination for large result sets**: When browsing emails, use pagination with the configured `page_size` to manage memory usage. Configure `PAGE_SIZE` in your environment variables to adjust the number of items per page.

5. **Respect parameter limits**: 
   - `days` parameter: maximum 29 (not 30)
   - `top` parameter: maximum 99 (not 100)
   - `list_recent_emails` days: maximum 7

---

## Error Handling

### Common Errors

**"Days parameter must be 7 or less"**
- Occurs when `days` > 7 is provided to `list_recent_emails`
- Solution: Use a value between 1 and 7

**"No emails in cache"**
- Occurs when trying to browse cache without loading emails first
- Solution: Use `list_recent_emails` or `search_emails` first

---

## Performance Considerations

1. **Cache Size**: The cache is automatically cleared before each load/search operation to prevent memory bloat.

2. **Async Operations**: Cache saves are performed asynchronously using `asyncio.to_thread` for non-blocking I/O.

3. **Pagination**: Use pagination when browsing large email lists to optimize memory usage.

4. **Parameter Limits**: The limits on `days` and `top` parameters are designed to prevent excessive API calls and memory usage.

---

## Testing

Run the test suite to verify email functionality:

```bash
python tests/test_new_email_functions.py
```

This will test all new email functions including:
- List recent emails with various day parameters
- Search emails with folder parameter
- Cache clearing functionality
- Async cache save performance
- Clear email cache tool
- List mail folders

All tests should pass successfully.
