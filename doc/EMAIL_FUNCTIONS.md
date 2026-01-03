# Email Functions Documentation

This document describes the email-related functions available in the Microsoft Graph MCP Server.

## Table of Contents

1. [Search Emails](#search-emails)
2. [Browse Email Cache](#browse-email-cache)
3. [Get Email Content](#get-email-content)
4. [Compose Reply Forward Email](#compose-reply-forward-email)
5. [Manage Mail Folder](#manage-mail-folder)
6. [Move Delete Emails](#move-delete-emails)

---

## Search Emails

### Description
Unified search tool for emails. Search by sender, recipient, subject, or body text with configurable date range filtering. If no search_type and query are provided, lists recent emails from Inbox (default: 1 day, maximum: 7 days).

### Parameters
- `search_type` (optional, string): Type of search to perform
  - Values: "sender", "recipient", "subject", "body"
  - If not provided, lists recent emails from Inbox
- `query` (optional, string): Search query
  - For sender: sender name or email address
  - For recipient: recipient name or email address
  - For subject: subject text
  - For body: body text content
  - Required when search_type is provided
- `folder` (optional, string): Folder to search (default: "Inbox" for recent emails, null for searches)
- `days` (optional, integer): Number of days to search back
  - Default: 1 for recent emails, 90 for searches (configurable via `DEFAULT_SEARCH_DAYS` environment variable)
  - Maximum: 7 for recent emails
  - Set to `null` to search all emails (not recommended for large mailboxes)

### Returns

#### For Recent Emails (no search_type and query):
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

#### For Search Results:
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
- `message`: Confirmation message (for recent emails)
- `search_type`: Type of search performed (for searches)
- `query`: Search query used (for searches)
- `folder`: Folder loaded or searched
- `count`: Number of emails found
- `timezone`: User's timezone for reference
- `date_range`: Actual date range of emails returned (from most recent to oldest)
- `filter_date_range`: Filter applied to the search (for searches)
- `hint`: Instructions for viewing results

### Example Usage

#### List Recent Emails:
```python
# List emails from the last day (default)
result = await search_emails()

# List emails from the last 3 days
result = await search_emails(days=3)

# List emails from the last 7 days (maximum)
result = await search_emails(days=7)
```

#### Search Emails:
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
- Automatically clears the cache before performing search or loading recent emails
- Loads search results or recent emails into cache for browsing
- Use `browse_email_cache` to view the results
- Date range filtering makes searches more efficient and predictable
- Setting `days` to `null` will search all emails (may be slow for large mailboxes)
- When no search_type and query are provided, lists recent emails from Inbox with a maximum of 7 days

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

## Compose Reply Forward Email

### Description
Unified tool for composing, replying to, and forwarding emails. Supports multiple recipients, CC, and BCC. The htmlbody parameter accepts HTML format for rich email content.

### Parameters
- `action` (required, string): Action to perform
  - Values: "compose", "reply", "forward"
- `to` (required, array of strings): List of recipient email addresses
- `htmlbody` (required, string): Email body content in HTML format. Use HTML tags like <p>, <br>, <strong>, <em>, <ul>, <li>, etc. Example: '<p>Hello,</p><p>This is <strong>important</strong>.</p><br><p>Best regards</p>'
- `subject` (optional, string): Email subject
  - Required for: "compose" action
  - Optional for: "reply" and "forward" actions
- `emailNumber` (optional, integer): Email number from browse_email_cache (e.g., 1, 2, 3)
  - Required for: "reply" and "forward" actions
- `cc` (optional, array of strings): List of CC recipient email addresses
- `bcc` (optional, array of strings): List of BCC recipient email addresses
- `bcc_csv_file` (optional, string): Path to CSV file containing BCC recipients
  - CSV must have a single column with header 'Email' or 'email'
  - Only available for: "forward" action

### Actions

#### Compose Email (action="compose")
Composes and sends a new email.

**Parameters:**
- `action`: "compose"
- `to`: List of recipient email addresses
- `subject`: Email subject
- `htmlbody`: Email body content (HTML format)
- `cc` (optional): List of CC recipient email addresses
- `bcc` (optional): List of BCC recipient email addresses

**Returns:**
```json
{
  "message": "Email composed and sent successfully: {result}"
}
```

**Example Usage:**
```python
result = await send_email(
    action="compose",
    to=["recipient@example.com"],
    subject="Meeting Tomorrow",
    htmlbody="<p>Hi, let's meet tomorrow at 2 PM.</p>",
    cc=["manager@example.com"]
)
```

#### Reply to Email (action="reply")
Replies to an existing email. The reply will be linked to the original email thread and will include inline attachments from the original email.

**Parameters:**
- `action`: "reply"
- `emailNumber`: Email number from browse_email_cache
- `to`: List of recipient email addresses
- `htmlbody`: Email body content (HTML format)
- `subject` (optional): Email subject
- `cc` (optional): List of CC recipient email addresses
- `bcc` (optional): List of BCC recipient email addresses

**Returns:**
```json
{
  "message": "Reply email sent successfully: {result}"
}
```

**Example Usage:**
```python
result = await send_email(
    action="reply",
    emailNumber=1,
    to=["original_sender@example.com"],
    htmlbody="<p>Thank you for your email. I'll review it and get back to you.</p>",
    subject="Re: Meeting Tomorrow"
)
```

#### Forward Email (action="forward")
Forwards an email to recipients. The original email will be included in the forwarded message with 'FW:' prefix on the subject. Supports BCC recipients via CSV file with automatic batching for large recipient lists.

**Parameters:**
- `action`: "forward"
- `emailNumber`: Email number from browse_email_cache
- `to`: List of recipient email addresses
- `htmlbody`: Email body content (HTML format)
- `subject` (optional): Email subject (defaults to 'FW: ' + original subject)
- `cc` (optional): List of CC recipient email addresses
- `bcc` (optional): List of BCC recipient email addresses
- `bcc_csv_file` (optional): Path to CSV file containing BCC recipients

**Returns:**
```json
{
  "message": "Email forwarded successfully: {result}"
}
```

**For Large BCC Lists (with batching):**
```json
{
  "message": "Email forwarded successfully in N batches (total X BCC recipients): {results}"
}
```

**Example Usage:**
```python
# Basic forward
result = await send_email(
    action="forward",
    emailNumber=1,
    to=["new_recipient@example.com"],
    body="<p>Please review the forwarded email.</p>"
)

# Forward with custom subject and CC
result = await send_email(
    action="forward",
    emailNumber=1,
    to=["new_recipient@example.com"],
    subject="FW: Important - Please Review",
    body="<p>Please review the forwarded email and provide feedback.</p>",
    cc=["manager@example.com"]
)

# Forward with BCC recipients from CSV file
result = await send_email(
    action="forward",
    emailNumber=1,
    to=["main_recipient@example.com"],
    body="<p>FYI - please review.</p>",
    bcc_csv_file="recipients.csv"
)

# Forward with BCC recipients from array
result = await send_email(
    action="forward",
    emailNumber=1,
    to=["main_recipient@example.com"],
    body="<p>FYI - please review.</p>",
    bcc=["bcc1@example.com", "bcc2@example.com"]
)
```

### BCC CSV File Format
The CSV file must have a single column with header 'Email' or 'email':

```csv
Email
recipient1@example.com
recipient2@example.com
recipient3@example.com
```

### BCC Batching
When forwarding emails with BCC recipients that exceed the maximum batch size (default: 500, configurable via `MAX_BCC_BATCH_SIZE`), the system automatically batches the recipients:

- Splits BCC recipients into batches of the maximum size
- Sends the forward email multiple times, once per batch
- Returns a summary of all batches sent

**Configuration:**
```env
# .env file
MAX_BCC_BATCH_SIZE=500
```

### Notes
- The body parameter must be in HTML format for all actions
- For reply and forward actions, use `browse_email_cache` to get the email number
- The original email is included in the forwarded message
- BCC recipients can be provided via array or CSV file
- Large BCC lists are automatically batched to stay within API limits
- The subject for forward actions defaults to 'FW: ' + original subject if not provided

---

## Manage Mail Folder

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
result = await manage_mail_folder(action="list")
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
result = await manage_mail_folder(action="move", folder_path="Inbox/Projects", destination_parent="Archive")
```

### Notes
- Returns all folders in the mailbox for list action
- Includes folder hierarchy information
- Shows total and unread item counts
- Folder paths use '/' separator for nested folders

---

## Move Delete Emails

### Description
Move or delete emails. Supports moving a single email, moving all emails from a folder, deleting a single email, deleting multiple emails, or deleting all emails from a folder.

### Parameters
- `action` (required, string): Action to perform
  - Values: "move_single", "move_all", "delete_single", "delete_multiple", "delete_all"
- `email_number` (optional, integer): Email number from browse_email_cache (e.g., 1, 2, 3)
  - Required for: "move_single" and "delete_single" actions
- `email_numbers` (optional, array of integers): List of email numbers from browse_email_cache (e.g., [1, 2, 3])
  - Required for: "delete_multiple" action
- `source_folder` (optional, string): Source folder path (e.g., 'Inbox', 'Archive/2024')
  - Required for: "move_all" and "delete_all" actions
- `destination_folder` (optional, string): Destination folder path (e.g., 'Archive/2024', 'Inbox/Projects')
  - Required for: "move_single" and "move_all" actions

### Actions

#### Move Single Email (action="move_single")
Moves a single email to a different folder.

**Parameters:**
- `action`: "move_single"
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
result = await move_delete_emails(
    action="move_single",
    email_number=1,
    destination_folder="Archive/2024"
)
```

#### Move All Emails from Folder (action="move_all")
Moves all emails from a source folder to a destination folder.

**Parameters:**
- `action`: "move_all"
- `source_folder`: Source folder path
- `destination_folder`: Destination folder path

**Returns:**
```json
{
  "status": "success",
  "message": "Moved X emails from SourceFolder to DestinationFolder",
  "moved_count": X,
  "failed_count": 0,
  "errors": null
}
```

**Example Usage:**
```python
result = await move_delete_emails(
    action="move_all",
    source_folder="Inbox",
    destination_folder="Archive/2024"
)
```

#### Delete Single Email (action="delete_single")
Deletes a single email by moving it to Deleted Items (recoverable).

**Parameters:**
- `action`: "delete_single"
- `email_number`: Email number from browse_email_cache

**Returns:**
```json
{
  "status": "success",
  "message": "Email moved to Deleted Items"
}
```

**Example Usage:**
```python
result = await move_delete_emails(
    action="delete_single",
    email_number=1
)
```

#### Delete Multiple Emails (action="delete_multiple")
Deletes multiple emails by moving them to Deleted Items (recoverable).

**Parameters:**
- `action`: "delete_multiple"
- `email_numbers`: List of email numbers from browse_email_cache

**Returns:**
```json
{
  "status": "success",
  "message": "Deleted X emails",
  "deleted_count": X,
  "failed_count": 0,
  "errors": null
}
```

**Example Usage:**
```python
result = await move_delete_emails(
    action="delete_multiple",
    email_numbers=[1, 2, 3]
)
```

#### Delete All Emails from Folder (action="delete_all")
Deletes all emails from a folder by moving them to Deleted Items (recoverable).

**Parameters:**
- `action`: "delete_all"
- `source_folder`: Source folder path

**Returns:**
```json
{
  "status": "success",
  "message": "Deleted X emails from SourceFolder",
  "deleted_count": X,
  "failed_count": 0,
  "errors": null
}
```

**Example Usage:**
```python
result = await move_delete_emails(
    action="delete_all",
    source_folder="Inbox"
)
```

### Notes
- Requires valid email number(s) from cache for "move_single", "delete_single", and "delete_multiple" actions
- Use `browse_email_cache` to get email numbers
- "move_all" and "delete_all" actions operate on all emails in the source folder
- Destination folder must exist for move actions (use manage_mail_folder to create if needed)
- Deleted emails are moved to Deleted Items folder and can be recovered
- Batch operations use batch processing for efficiency (20 emails per batch)
- All operations return status, count, and error information for tracking

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

1. **Use search_emails for quick access**: For viewing recent emails, use `search_emails` with the default 1-day parameter (no search_type and query).

2. **Specify folder for targeted searches**: Use the `folder` parameter in `search_emails` to narrow down search results.

3. **Use pagination for large result sets**: When browsing emails, use pagination with the configured `page_size` to manage memory usage. Configure `PAGE_SIZE` in your environment variables to adjust the number of items per page.

4. **Use HTML format for email bodies**: When using `send_email`, ensure the htmlbody parameter is in HTML format for all actions (compose, reply, forward).

5. **Respect parameter limits**: 
   - `days` parameter for recent emails: maximum 7
   - `days` parameter for searches: configurable via `DEFAULT_SEARCH_DAYS` environment variable
   - `top` parameter: maximum 99 (not 100)

---

## Error Handling

### Common Errors

**"Days parameter must be 7 or less"**
- Occurs when `days` > 7 is provided to `search_emails` for recent emails
- Solution: Use a value between 1 and 7 for recent emails

**"No emails in cache"**
- Occurs when trying to browse cache without loading emails first
- Solution: Use `search_emails` first (with or without search_type and query)

**"Invalid action: X. Must be 'compose', 'reply', or 'forward'."**
- Occurs when an invalid action is provided to `send_email`
- Solution: Use one of the valid actions: "compose", "reply", or "forward"

**"Error: Email number X is out of range"**
- Occurs when an invalid email number is provided to `send_email` for reply or forward actions
- Solution: Use a valid email number from `browse_email_cache` (between 1 and total count)

**"Error: No valid email ID found"**
- Occurs when the email cache doesn't contain a valid email ID
- Solution: Check the cache with `browse_email_cache` and try again

**"Error reading BCC CSV file: X"**
- Occurs when there's an error reading the BCC CSV file
- Solution: Ensure the CSV file exists and has the correct format (single column with header 'Email' or 'email')

---

## Performance Considerations

1. **Hard Limit**: All email search methods have a maximum limit defined by `MAX_EMAIL_SEARCH_LIMIT` (50 emails) per search to prevent excessive API calls and memory usage. This limit is implemented as a constant in the codebase rather than a magic number, making it easier to maintain and modify in the future.

2. **Performance Metrics**:
   - 100 emails: 2.22 seconds (45.1 emails/second)
   - 500 emails: 3.50 seconds (142.7 emails/second)
   - 1000 emails: 5.39 seconds (185.6 emails/second)

3. **Key Optimizations**:
   - **List Comprehension**: Email summaries are generated using list comprehension instead of parallel processing for optimal performance
   - **Reduced API Response Size**: Only essential fields are requested from the API (id, subject, from, toRecipients, ccRecipients, receivedDateTime, sentDateTime, isRead, hasAttachments, importance, bodyPreview)
   - **Timezone Object Caching**: ZoneInfo objects are cached to avoid redundant timezone conversions
   - **User Timezone Caching**: User timezone information is cached to avoid repeated calculations
   - **Optimized Field Selection**: API requests retrieve only needed fields, reducing response size by ~40%
   - **Regex Pattern Optimization**: Multiple regex operations consolidated into fewer, more efficient patterns

4. **Cache Size**: The cache is automatically cleared before each load/search operation to prevent memory bloat.

5. **Async Operations**: Cache saves are performed asynchronously using `asyncio.to_thread` for non-blocking I/O.

6. **Email Content Processing**:
   - HTML extraction removes all HTML tags, styles, classes, IDs, and images
   - Configurable maximum email body length (default: 5000 characters) to reduce token usage
   - Content truncation with clear indication when content is truncated

7. **Pagination**: Use pagination when browsing large email lists to optimize memory usage.

8. **Parameter Limits**: The limits on `days` and `top` parameters are designed to prevent excessive API calls and memory usage.

---

## Testing

Run the test suite to verify email functionality:

```bash
python tests/test_new_email_functions.py
```

This will test all email functions including:
- Search emails with various parameters (recent emails and searches)
- Cache clearing functionality
- Async cache save performance
- Browse email cache
- Get email content
- Compose, reply, and forward emails
- Manage mail folders
- Move emails

All tests should pass successfully.
