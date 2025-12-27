# Email Functions Documentation

This document describes the email-related functions available in the Microsoft Graph MCP Server.

## Table of Contents

1. [List Recent Emails](#list-recent-emails)
2. [Load Emails by Folder](#load-emails-by-folder)
3. [Search Emails](#search-emails)
4. [Browse Email Cache](#browse-email-cache)
5. [Clear Email Cache](#clear-email-cache)
6. [Get Email](#get-email)
7. [List Mail Folders](#list-mail-folders)

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

## Load Emails by Folder

### Description
Loads emails from a specific folder into the cache. Can filter by days or limit by page size (mutually exclusive parameters).

### Parameters
- `folder` (optional, string): Mail folder name
  - Default: "Inbox"
- `days` (optional, integer): Number of days to look back
  - Minimum: 1
  - Maximum: 29 (cannot be 30 or more)
- `top` (optional, integer): Maximum number of emails to load
  - Minimum: 1
  - Maximum: 99 (cannot be 100 or more)

### Configuration
The number of emails per page is controlled by the `PAGE_SIZE` environment variable:

```env
# .env file
PAGE_SIZE=5
```

Default page size is 5 emails per page. This ensures consistent pagination across all browsing operations.

### Constraints
- Cannot specify both `days` and `top` parameters simultaneously
- `days` must be less than 30
- `top` must be less than 100

### Returns
```json
{
  "message": "Loaded X emails from FolderName",
  "folder": "FolderName",
  "filter_days": N,
  "limit_top": M,
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
- `folder`: Folder loaded
- `filter_days`: Days filter applied (if days parameter used)
- `limit_top`: Top limit applied (if top parameter used)
- `count`: Number of emails loaded
- `timezone`: User's timezone for reference
- `date_range`: Actual date range of emails loaded (from most recent to oldest)
- `hint`: Instructions for viewing results

### Example Usage
```python
# Load emails from last 7 days in Inbox
result = await load_emails_by_folder(folder="Inbox", days=7)

# Load most recent 50 emails from Sent Items
result = await load_emails_by_folder(folder="Sent Items", top=50)

# Load emails from last 14 days in Archive
result = await load_emails_by_folder(folder="Archive", days=14)
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
  "id": "email-id",
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
  "receivedDateTime": "2025-12-26T11:01:26Z",
  "isRead": false,
  "hasAttachments": false,
  "importance": "normal"
}
```

### Example Usage
```python
# View first page (default 20 emails per page)
result = await browse_email_cache(page_number=1)

# View second page (uses configured page_size)
result = await browse_email_cache(page_number=2)
```

### Notes
- Requires emails to be loaded into cache first (via list_recent_emails, load_emails_by_folder, or search_emails)
- Automatically manages browsing state with disk cache for persistence
- Use the `number` field to reference specific emails

---

## Clear Email Cache

### Description
Clears the email browsing cache. This removes all cached emails from memory and disk.

### Parameters
None

### Returns
```json
{
  "message": "Email cache cleared successfully",
  "status": "success"
}
```

### Example Usage
```python
result = await clear_email_cache()
```

### Notes
- Removes all cached emails from memory
- Deletes cache file from disk
- Useful for freeing up memory or starting fresh

---

## Get Email

### Description
Get full email content by ID. Use the email ID from browse_email_cache or search_emails to retrieve complete email with body, attachments, and all details.

### Parameters
- `email_id` (required, string): Email ID from browse_email_cache or search_emails

### Returns
Full email object including:
- Subject, from, to, cc, bcc
- Body content (HTML and text)
- Attachments
- Headers
- All other email properties

### Example Usage
```python
# Get email by ID
result = await get_email(email_id="AAMkADc4MDRiMTA2LWM5MTctNGY0Yy1iNGZhLTk5N2FkZDYzNTdlMABGAAAAAADvtPkv_HNvQ5BD_qICmjfLBwAS9HxkbFcHSLwsZGppet3QAAAAAAEMAAAS9HxkbFcHSLwsZGppet3QAAdhgdQfAAA=")
```

### Notes
- Requires valid email ID from cache
- Returns complete email with all details

---

## List Mail Folders

### Description
Lists all mail folders with hierarchy including child folders, total item count, and unread item count.

### Parameters
None

### Returns
Array of folder objects:
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

### Example Usage
```python
result = await list_mail_folders()
```

### Notes
- Returns all folders in the mailbox
- Includes folder hierarchy information
- Shows total and unread item counts

---

## Cache Management

### Automatic Cache Clearing
The cache is automatically cleared before the following operations:
- `list_recent_emails`
- `load_emails_by_folder`
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

3. **Clear cache when needed**: Use `clear_email_cache` to free up memory or start fresh.

4. **Use pagination for large result sets**: When browsing emails, use pagination with the configured `page_size` to manage memory usage. Configure `PAGE_SIZE` in your environment variables to adjust the number of items per page.

5. **Respect parameter limits**: 
   - `days` parameter: maximum 29 (not 30)
   - `top` parameter: maximum 99 (not 100)
   - `list_recent_emails` days: maximum 7

---

## Error Handling

### Common Errors

**"Cannot specify both 'days' and 'top' parameters simultaneously"**
- Occurs when both parameters are provided to `load_emails_by_folder`
- Solution: Use only one parameter at a time

**"Days parameter must be less than 30"**
- Occurs when `days` >= 30 is provided
- Solution: Use a value less than 30

**"Top parameter must be less than 100"**
- Occurs when `top` >= 100 is provided
- Solution: Use a value less than 100

**"Days parameter must be 7 or less"**
- Occurs when `days` > 7 is provided to `list_recent_emails`
- Solution: Use a value between 1 and 7

**"No emails in cache"**
- Occurs when trying to browse cache without loading emails first
- Solution: Use `list_recent_emails`, `load_emails_by_folder`, or `search_emails` first

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
