# Email Functions Documentation

This document describes the email-related functions available in the Microsoft Graph MCP Server.

## Table of Contents

1. [Search Emails](#search-emails)
2. [Browse Email Cache](#browse-email-cache)
3. [Get Email Content](#get-email-content)
4. [Send Email](#send-email)
5. [Manage Mail Folder](#manage-mail-folder)
6. [Manage Emails](#manage-emails)
7. [Manage Templates](#manage-templates)

---

## Search Emails

### Description
Unified search tool for emails. Search by sender, recipient, subject, or body text with configurable date range filtering. If no search_type and query are provided, lists recent emails from Inbox (default: 1 day, maximum: 7 days).

### Server-Side Filtering Performance

The search function uses optimized server-side filtering with Microsoft Graph API for maximum performance:

| Search Type | API Method | Date Filtering | Performance |
|-------------|------------|----------------|-------------|
| **Sender fuzzy name** (e.g., "beng") | `$search="from:beng received>=YYYY-MM-DD"` | **Server-side (KQL)** | **Best** |
| **Sender exact email** (e.g., "john@example.com") | `$filter="from/emailAddress/address eq '...'"` | **Server-side** | **Best** |
| **Subject search** | `$search="subject:keywords received>=YYYY-MM-DD"` | **Server-side (KQL)** | **Best** |
| **Body search** | `$search="keywords received>=YYYY-MM-DD"` | **Server-side (KQL)** | **Best** |

**Key Features:**
- **All searches now use server-side date filtering** for optimal performance
- **KQL date syntax** is embedded in `$search` queries (e.g., `received:2026-02-01..2026-02-26`)
- **Exact email searches** use `$filter` for precise matching
- **Fuzzy name, subject, and body searches** use `$search` with KQL date filters
- Maximum 100 emails returned per search to balance performance and completeness

**KQL Date Syntax:**
- Date range: `received:2026-02-01..2026-02-26`
- From date: `received>=2026-02-01`
- Until date: `received<=2026-02-26`

### Parameters
- `search_type` (optional, string): Type of search to perform
  - Values: "sender", "subject", "body"
  - If not provided, lists recent emails from Inbox
- `query` (optional, string): Search query
  - For sender: sender name or email address (fuzzy matching for names, exact match for emails)
  - For subject: subject text (supports multiple keywords with AND logic)
  - For body: body text content (supports multiple keywords with AND logic)
  - Required when search_type is provided
- `folder` (optional, string): Folder to search (default: "Inbox" for recent emails, null for searches)
- `days` (optional, integer): Number of days to search back
  - Default: 90 (configurable via `DEFAULT_SEARCH_DAYS` environment variable)
  - Maximum: 90 (configurable via `MAX_SEARCH_DAYS` environment variable)
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
The search range can be configured via environment variables:

```env
# .env file
DEFAULT_SEARCH_DAYS=7    # Default search range when not specified
MAX_SEARCH_DAYS=90       # Maximum allowed search range
```

### Notes
- Automatically clears the cache before performing search or loading recent emails
- Loads search results or recent emails into cache for browsing
- Use `browse_email_cache` to view the results
- Date range filtering makes searches more efficient and predictable
- Setting `days` to `null` will search all emails (may be slow for large mailboxes)
- When no search_type and query are provided, lists recent emails from Inbox with a maximum of 7 days
- Subject and body searches use exact substring matching (contains) for precise results, while sender searches use fuzzy matching

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
Get full email content by cache number. Use the cache number from browse_email_cache (e.g., 1, 2, 3) to retrieve complete email with body, attachments, and all details.

### Parameters
- `cache_number` (required, integer): Cache number from browse_email_cache (e.g., 1, 2, 3)
- `text_only` (optional, boolean): If true, return only text content without embedded images and attachments. If false, return full content including embedded images and attachments.
  - Default: true

### Returns
User-facing email content:
```json
{
  "cache_number": 1,
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
- `cache_number`: Cache number in cache
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
result = await get_email_content(cache_number=1)

# Get email content with full attachments and images
result = await get_email_content(cache_number=1, text_only=false)
```

### Notes
- Requires valid cache number from cache
- Returns only user-facing content (system metadata is excluded)
- Use `text_only=false` to include attachments and embedded images

---

## Send Email

### Description
Send emails directly without creating drafts. Supports three actions: 'send_new' to send a new email, 'reply' to reply to an existing email, and 'forward' to forward an existing email. All actions send emails immediately - no drafts are created. Supports multiple recipients, CC, and BCC. The htmlbody parameter accepts HTML format for rich email content.

**⭐ RECOMMENDED FOR BCC**: Use the `bcc_csv_file` parameter to provide BCC recipients from a CSV file. This is the **preferred method** for handling BCC recipients, especially for large lists. The system automatically batches large BCC lists (up to 500 recipients per batch by default) and sends multiple emails as needed. This approach is more efficient and handles API limits gracefully compared to manually managing BCC arrays.

### Parameters
- `action` (required, string): Action to perform
  - Values: "send_new", "reply", "forward"
- `to` (required, array of strings): List of recipient email addresses
- `htmlbody` (required, string): Email body content in HTML format. Use HTML tags like <p>, <br>, <strong>, <em>, <ul>, <li>, etc. Example: '<p>Hello,</p><p>This is <strong>important</strong>.</p><br><p>Best regards</p>'
- `subject` (optional, string): Email subject
  - Required for: "send_new" action
  - Optional for: "reply" and "forward" actions
- `cache_number` (optional, integer): Cache number from browse_email_cache (e.g., 1, 2, 3)
  - Required for: "reply" and "forward" actions
- `cc` (optional, array of strings): List of CC recipient email addresses
- `bcc` (optional, array of strings): List of BCC recipient email addresses
- `bcc_csv_file` (optional, string): **PREFERRED METHOD FOR BCC** - Path to CSV file containing BCC recipients
  - CSV must have a single column with header 'Email' or 'email'
  - Only available for: "forward" action
  - **Recommended over manual `bcc` array** for large recipient lists
  - Automatically batches recipients (up to 500 per batch by default)
  - When both `bcc` and `bcc_csv_file` are provided, they are combined

### Actions

#### Send New Email (action="send_new")
Sends a new email immediately without creating a draft.

**Parameters:**
- `action`: "send_new"
- `to`: List of recipient email addresses
- `subject`: Email subject
- `htmlbody`: Email body content (HTML format)
- `cc` (optional): List of CC recipient email addresses
- `bcc` (optional): List of BCC recipient email addresses

**Returns:**
```json
{
  "message": "Email sent successfully: {result}"
}
```

**Example Usage:**
```python
result = await send_email(
    action="send_new",
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
- `cache_number`: Cache number from browse_email_cache
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
    cache_number=1,
    to=["original_sender@example.com"],
    htmlbody="<p>Thank you for your email. I'll review it and get back to you.</p>",
    subject="Re: Meeting Tomorrow"
)
```

#### Forward Email (action="forward")
Forwards an email to recipients. The original email will be included in the forwarded message with 'FW:' prefix on the subject. **⭐ RECOMMENDED FOR BCC**: Use `bcc_csv_file` parameter for BCC recipients - this is the preferred method with automatic batching support for large recipient lists. **NOTE**: The `to` parameter is optional when using `bcc_csv_file` or `bcc` - you can forward emails using only BCC recipients.

**Parameters:**
- `action`: "forward"
- `cache_number`: Cache number from browse_email_cache
- `to` (optional): List of recipient email addresses (optional when using `bcc_csv_file` or `bcc`)
- `htmlbody`: Email body content (HTML format)
- `subject` (optional): Email subject (defaults to 'FW: ' + original subject)
- `cc` (optional): List of CC recipient email addresses
- `bcc` (optional): List of BCC recipient email addresses
- `bcc_csv_file` (optional): **PREFERRED METHOD** - Path to CSV file containing BCC recipients

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
# ⭐ RECOMMENDED: Forward with BCC recipients from CSV file (no TO recipients)
# This is the preferred approach for handling BCC recipients
result = await send_email(
    action="forward",
    cache_number=1,
    htmlbody="<p>FYI - please review.</p>",
    bcc_csv_file="recipients.csv"
)

# Basic forward
result = await send_email(
    action="forward",
    cache_number=1,
    to=["new_recipient@example.com"],
    htmlbody="<p>Please review the forwarded email.</p>"
)

# Forward with custom subject and CC
result = await send_email(
    action="forward",
    cache_number=1,
    to=["new_recipient@example.com"],
    subject="FW: Important - Please Review",
    htmlbody="<p>Please review the forwarded email and provide feedback.</p>",
    cc=["manager@example.com"]
)

# Forward with BCC recipients from array (small lists only)
# Note: For large BCC lists, use bcc_csv_file instead
result = await send_email(
    action="forward",
    cache_number=1,
    to=["main_recipient@example.com"],
    htmlbody="<p>FYI - please review.</p>",
    bcc=["bcc1@example.com", "bcc2@example.com"]
)

# Forward with only BCC recipients (no TO recipients)
result = await send_email(
    action="forward",
    cache_number=1,
    htmlbody="<p>FYI - please review.</p>",
    bcc=["bcc1@example.com", "bcc2@example.com", "bcc3@example.com"]
)

# Forward combining both CSV and array BCC recipients
result = await send_email(
    action="forward",
    cache_number=1,
    to=["main_recipient@example.com"],
    htmlbody="<p>FYI - please review.</p>",
    bcc=["manual_bcc@example.com"],
    bcc_csv_file="additional_recipients.csv"
)
```

### BCC CSV File Format
**⭐ RECOMMENDED APPROACH**: Using a CSV file for BCC recipients is the preferred method for handling BCC lists, especially for large recipient lists. The system automatically handles batching and API limits.

The CSV file must have a single column with header 'Email' or 'email':

```csv
Email
recipient1@example.com
recipient2@example.com
recipient3@example.com
```

**Why use CSV file instead of manual BCC array?**
- ✅ **Automatic batching**: Splits large lists into batches (up to 500 recipients per batch)
- ✅ **Error handling**: Gracefully handles API limits and retries
- ✅ **Scalability**: Works with any number of recipients
- ✅ **Maintainability**: Easy to update recipient lists without code changes
- ✅ **Flexibility**: Can combine CSV file with manual BCC array if needed

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

**Example Response with Batching:**
```json
{
  "message": "Email forwarded successfully in 3 batches (total 1250 BCC recipients): {results}"
}
```

### Notes
- The htmlbody parameter must be in HTML format for all actions
- For reply and forward actions, use `browse_email_cache` to get the cache number
- The original email is included in the forwarded message
- **⭐ RECOMMENDED**: Use `bcc_csv_file` for BCC recipients instead of manual `bcc` array
- Large BCC lists are automatically batched to stay within API limits
- The subject for forward actions defaults to 'FW: ' + original subject if not provided
- When both `bcc` and `bcc_csv_file` are provided, they are combined
- **For forward action**: The `to` parameter is optional when using `bcc_csv_file` or `bcc` - you can forward emails using only BCC recipients
- At least one of `to`, `bcc`, or `bcc_csv_file` must be provided for forward action

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

## Manage Emails

### Description
Manage emails with multiple actions. Supports moving, deleting, archiving, flagging, and categorizing emails. Actions include: move_single, move_all, delete_single, delete_multiple, delete_all, archive_single, archive_multiple, flag_single, flag_multiple, categorize_single, categorize_multiple.

### Parameters
- `action` (required, string): Action to perform
  - Values: "move_single", "move_all", "delete_single", "delete_multiple", "delete_all", "archive_single", "archive_multiple", "flag_single", "flag_multiple", "categorize_single", "categorize_multiple"
- `cache_number` (optional, integer): Cache number from browse_email_cache (e.g., 1, 2, 3)
  - Required for: "move_single", "delete_single", "archive_single", "flag_single", and "categorize_single" actions
- `cache_numbers` (optional, array of integers): List of cache numbers from browse_email_cache (e.g., [1, 2, 3])
  - Required for: "delete_multiple", "archive_multiple", "flag_multiple", and "categorize_multiple" actions
- `source_folder` (optional, string): Source folder path (e.g., 'Inbox', 'Archive/2024')
  - Required for: "move_all" and "delete_all" actions
- `destination_folder` (optional, string): Destination folder path (e.g., 'Archive/2024', 'Inbox/Projects')
  - Required for: "move_single" and "move_all" actions
- `flag_status` (optional, string): Flag status
  - Values: "flagged", "complete"
  - Required for: "flag_single" and "flag_multiple" actions
- `categories` (optional, array of strings): List of category names to apply (e.g., ['Important', 'Work'])
  - Required for: "categorize_single" and "categorize_multiple" actions

### Actions

#### Move Single Email (action="move_single")
Moves a single email to a different folder.

**Parameters:**
- `action`: "move_single"
- `cache_number`: Cache number from browse_email_cache
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
result = await manage_emails(
    action="move_single",
    cache_number=1,
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
result = await manage_emails(
    action="move_all",
    source_folder="Inbox",
    destination_folder="Archive/2024"
)
```

#### Delete Single Email (action="delete_single")
Deletes a single email by moving it to Deleted Items (recoverable).

**Parameters:**
- `action`: "delete_single"
- `cache_number`: Cache number from browse_email_cache

**Returns:**
```json
{
  "status": "success",
  "message": "Email moved to Deleted Items"
}
```

**Example Usage:**
```python
result = await manage_emails(
    action="delete_single",
    cache_number=1
)
```

#### Delete Multiple Emails (action="delete_multiple")
Deletes multiple emails by moving them to Deleted Items (recoverable).

**Parameters:**
- `action`: "delete_multiple"
- `cache_numbers`: List of cache numbers from browse_email_cache

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
result = await manage_emails(
    action="delete_multiple",
    cache_numbers=[1, 2, 3]
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
result = await manage_emails(
    action="delete_all",
    source_folder="Inbox"
)
```

#### Archive Single Email (action="archive_single")
Archives a single email by moving it to the Archive folder.

**Parameters:**
- `action`: "archive_single"
- `cache_number`: Cache number from browse_email_cache

**Returns:**
```json
{
  "status": "success",
  "message": "Email archived"
}
```

**Example Usage:**
```python
result = await manage_emails(
    action="archive_single",
    cache_number=1
)
```

#### Archive Multiple Emails (action="archive_multiple")
Archives multiple emails by moving them to the Archive folder.

**Parameters:**
- `action`: "archive_multiple"
- `cache_numbers`: List of cache numbers from browse_email_cache

**Returns:**
```json
{
  "status": "success",
  "message": "Archived X emails",
  "archived_count": X,
  "failed_count": 0,
  "errors": null
}
```

**Example Usage:**
```python
result = await manage_emails(
    action="archive_multiple",
    cache_numbers=[1, 2, 3]
)
```

#### Flag Single Email (action="flag_single")
Flags or unflags a single email.

**Parameters:**
- `action`: "flag_single"
- `cache_number`: Cache number from browse_email_cache
- `flag_status`: Flag status ("flagged" or "complete")

**Returns:**
```json
{
  "status": "success",
  "message": "Email flagged"
}
```

**Example Usage:**
```python
# Flag an email
result = await manage_emails(
    action="flag_single",
    cache_number=1,
    flag_status="flagged"
)

# Mark as complete
result = await manage_emails(
    action="flag_single",
    cache_number=1,
    flag_status="complete"
)
```

#### Flag Multiple Emails (action="flag_multiple")
Flags or unflags multiple emails.

**Parameters:**
- `action`: "flag_multiple"
- `cache_numbers`: List of cache numbers from browse_email_cache
- `flag_status`: Flag status ("flagged" or "complete")

**Returns:**
```json
{
  "status": "success",
  "message": "Flagged X emails",
  "flagged_count": X,
  "failed_count": 0,
  "errors": null
}
```

**Example Usage:**
```python
result = await manage_emails(
    action="flag_multiple",
    cache_numbers=[1, 2, 3],
    flag_status="flagged"
)
```

#### Categorize Single Email (action="categorize_single")
Adds categories to a single email.

**Parameters:**
- `action`: "categorize_single"
- `cache_number`: Cache number from browse_email_cache
- `categories`: List of category names to apply (e.g., ['Important', 'Work'])

**Returns:**
```json
{
  "status": "success",
  "message": "Email categorized with: Important, Work"
}
```

**Example Usage:**
```python
result = await manage_emails(
    action="categorize_single",
    cache_number=1,
    categories=["Important", "Work"]
)
```

#### Categorize Multiple Emails (action="categorize_multiple")
Adds categories to multiple emails.

**Parameters:**
- `action`: "categorize_multiple"
- `cache_numbers`: List of cache numbers from browse_email_cache
- `categories`: List of category names to apply (e.g., ['Important', 'Work'])

**Returns:**
```json
{
  "status": "success",
  "message": "Categorized X emails",
  "categorized_count": X,
  "failed_count": 0,
  "errors": null
}
```

**Example Usage:**
```python
result = await manage_emails(
    action="categorize_multiple",
    cache_numbers=[1, 2, 3],
    categories=["Important", "Work"]
)
```

### Notes
- Requires valid cache number(s) from cache for single and multiple email actions
- Use `browse_email_cache` to get cache numbers
- "move_all" and "delete_all" actions operate on all emails in the source folder
- Destination folder must exist for move actions (use manage_mail_folder to create if needed)
- Deleted emails are moved to Deleted Items folder and can be recovered
- Archived emails are moved to the Archive folder
- Flag status can be "flagged" (mark for follow-up) or "complete" (mark as done)
- Categories are user-defined labels that help organize emails
- Batch operations use batch processing for efficiency (20 emails per batch)
- All operations return status, count, and error information for tracking

---

## Manage Templates

### Description
Manage email templates stored as drafts in a Templates folder. Templates are draft emails that can be edited and sent. Actions include: create_from_email, list, get, update, delete, and send.

### Parameters
- `action` (required, string): Action to perform
  - Values: "create_from_email", "list", "get", "update", "delete", "send"
- `cache_number` (optional, integer): Cache number from browse_email_cache to copy as template
  - Required for: "create_from_email" action
- `template_number` (optional, integer): Template cache number
  - Required for: "get", "update", "delete", and "send" actions
- `subject` (optional, string): Email subject (title)
  - Optional for: "update" action
- `to` (optional, array of strings): List of recipient email addresses
  - Optional for: "update" action - if not provided, keeps existing recipients
- `cc` (optional, array of strings): List of CC recipient email addresses
  - Optional for: "update" action
- `bcc` (optional, array of strings): List of BCC recipient email addresses
  - Optional for: "update" action
- `htmlbody` (optional, string): Email body content in HTML format
  - Optional for: "update" action - if not provided, keeps existing body
  - Note: When updating body, you should first call get with text_only=false to get the full HTML, then provide the complete updated HTML here
- `text_only` (optional, boolean): For "get" action: if true, returns simple text body (default). If false, returns full HTML body.
  - Default: true

### Actions

#### Create Template from Email (action="create_from_email")
Copies an existing email as a template in the Templates folder.

**Parameters:**
- `action`: "create_from_email"
- `cache_number`: Cache number from browse_email_cache

**Returns:**
```json
{
  "message": "Template created successfully",
  "template_id": "template-id",
  "subject": "Email Subject",
  "folder": "Templates"
}
```

**Example Usage:**
```python
result = await manage_templates(
    action="create_from_email",
    cache_number=1
)
```

#### List Templates (action="list")
Lists all templates in the Templates folder with pagination.

**Parameters:**
- `action`: "list"

**Returns:**
```json
{
  "templates": [
    {
      "number": 1,
      "id": "template-id",
      "subject": "Weekly Newsletter",
      "folder": "Templates",
      "createdDateTime": "Fri 01/03/2026 10:00 AM",
      "lastModifiedDateTime": "Fri 01/03/2026 10:00 AM"
    }
  ],
  "count": 1,
  "total_count": 1,
  "page": 1,
  "pages": 1
}
```

**Example Usage:**
```python
result = await manage_templates(action="list")
```

#### Get Template (action="get")
Retrieves template details. Returns simple text body by default, or full HTML when text_only=false.

**Parameters:**
- `action`: "get"
- `template_number`: Template cache number
- `text_only` (optional): If true, returns simple text body. If false, returns full HTML body.

**Returns (text_only=true):**
```json
{
  "template_number": 1,
  "id": "template-id",
  "subject": "Weekly Newsletter",
  "to": ["recipient@example.com"],
  "cc": [],
  "body": "Hello, this is a simple text version of the email body.",
  "createdDateTime": "Fri 01/03/2026 10:00 AM",
  "lastModifiedDateTime": "Fri 01/03/2026 10:00 AM"
}
```

**Returns (text_only=false):**
```json
{
  "template_number": 1,
  "id": "template-id",
  "subject": "Weekly Newsletter",
  "to": ["recipient@example.com"],
  "cc": [],
  "body": {
    "contentType": "html",
    "content": "<html><body><p>Hello, this is the full HTML version.</p></body></html>"
  },
  "createdDateTime": "Fri 01/03/2026 10:00 AM",
  "lastModifiedDateTime": "Fri 01/03/2026 10:00 AM"
}
```

**Example Usage:**
```python
# Get simple text body (default)
result = await manage_templates(
    action="get",
    template_number=1
)

# Get full HTML body
result = await manage_templates(
    action="get",
    template_number=1,
    text_only=false
)
```

#### Update Template (action="update")
Updates template content. Requires the complete updated HTML body for any body changes.

**Parameters:**
- `action`: "update"
- `template_number`: Template cache number
- `subject` (optional): New email subject
- `to` (optional): New list of recipient email addresses
- `cc` (optional): New list of CC recipient email addresses
- `bcc` (optional): New list of BCC recipient email addresses
- `htmlbody` (optional): Complete updated HTML body content

**Returns:**
```json
{
  "message": "Template updated successfully",
  "template_id": "template-id",
  "subject": "Updated Subject",
  "folder": "Templates"
}
```

**Example Usage:**
```python
result = await manage_templates(
    action="update",
    template_number=1,
    subject="Updated Subject",
    htmlbody="<html><body><p>Updated content here.</p></body></html>"
)
```

#### Delete Template (action="delete")
Deletes a template by moving it to the Deleted Items folder (soft delete).

**Parameters:**
- `action`: "delete"
- `template_number`: Template cache number

**Returns:**
```json
{
  "message": "Template moved to Deleted Items",
  "template_id": "template-id",
  "deleted": true
}
```

**Example Usage:**
```python
result = await manage_templates(
    action="delete",
    template_number=1
)
```

**Note:** This is a soft delete - the template is moved to the Deleted Items folder and can be recovered if needed.

#### Send Template (action="send")
Sends a template (creates a copy and sends it, preserving the original template).

**Parameters:**
- `action`: "send"
- `template_number`: Template cache number
- `to` (optional): Override recipient email addresses
- `cc` (optional): Override CC recipient email addresses
- `bcc` (optional): Override BCC recipient email addresses

**Returns:**
```json
{
  "message": "Template sent successfully",
  "subject": "Weekly Newsletter",
  "to": ["recipient@example.com"],
  "sent": true,
  "savedCopyId": "saved-copy-id"
}
```

**Example Usage:**
```python
result = await manage_templates(
    action="send",
    template_number=1
)
```

### Template Update Workflow

The recommended workflow for updating templates involves a 7-step process:

1. **User calls get with text_only=true** - Gets simple text body for easy reading
2. **User provides update instructions to LLM** - Describes what changes they want
3. **LLM calls get with text_only=false** - Retrieves the full HTML body
4. **LLM calls update with htmlbody** - Applies user's changes and provides complete updated HTML
5. **User calls get with text_only=true** - Verifies the changes in simple text format
6. **User gives command to send** - Approves the template for sending
7. **LLM calls send** - Sends the template and saves a copy in the Templates folder

**JSON Examples:**

**User calls get action (simple text):**
```json
{
  "action": "get",
  "template_number": 1,
  "text_only": true
}
```

**LLM calls get action (full HTML):**
```json
{
  "action": "get",
  "template_number": 1,
  "text_only": false
}
```

**LLM calls update action (complete HTML):**
```json
{
  "action": "update",
  "template_number": 1,
  "htmlbody": "<html><body><p>Updated content with all HTML preserved.</p></body></html>"
}
```

### Notes
- Templates are stored as draft emails in the Templates folder
- The Templates folder is automatically created if it doesn't exist
- Use `text_only=true` for user-friendly simple text viewing
- Use `text_only=false` when LLM needs to work with full HTML
- When updating body, always provide the complete HTML (not partial updates)
- Sending a template creates a copy and sends it, preserving the original template
- The saved copy is also stored in the Templates folder for reference

---

## Email Sorting

### Sorting Order

Emails are always sorted by received date in **descending order** (newest first). This applies to:
- API responses from `search_emails`
- Cached emails retrieved via `browse_email_cache`

### Timestamp Fields

Each email includes two timestamp fields:

| Field | Purpose | Format | Timezone |
|-------|---------|--------|----------|
| `receivedDateTime` | Display | "Fri 12/26/2025 09:27 PM" | User's local timezone |
| `receivedDateTimeOriginal` | Sorting | "2025-12-26T13:27:44Z" | UTC |

### Best Practices

1. **Always use `receivedDateTimeOriginal` for sorting** - This ensures accurate sorting regardless of timezone
2. **Check pagination bounds** - Verify `page_number <= total_pages` before requesting
3. **Configure appropriate page sizes**:
   - Small (5-10): Slow connections, limited display space
   - Medium (20-50): Good balance for most use cases
   - Large (100+): Batch processing or data export

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

1. **Use search_emails for quick access**: For viewing recent emails, use `search_emails` with the default 1-day parameter.

2. **Specify folder for targeted searches**: Use the `folder` parameter to narrow down results.

3. **Use pagination for large result sets**: Configure `PAGE_SIZE` in environment variables.

4. **Use HTML format for email bodies**: The `htmlbody` parameter must be in HTML format.

5. **Respect parameter limits**:
   - `days` parameter for recent emails: maximum 7
   - `days` parameter for searches: configurable via `DEFAULT_SEARCH_DAYS`

---

## Error Handling

### Common Errors

**"Days parameter must be 7 or less"**
- Occurs when `days` > 7 is provided to `search_emails` for recent emails
- Solution: Use a value between 1 and 7 for recent emails

**"No emails in cache"**
- Occurs when trying to browse cache without loading emails first
- Solution: Use `search_emails` first (with or without search_type and query)

**"Invalid action: X. Must be 'send_new', 'reply', or 'forward'."**
- Occurs when an invalid action is provided to `send_email`
- Solution: Use one of the valid actions: "send_new", "reply", or "forward"

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

The email search functionality has been significantly optimized with server-side filtering and query optimization for dramatic performance improvements:

### Recent Optimizations (v2.0)

1. **KQL Date Filtering (NEW)**
   - **Before**: `$search` queries couldn't combine with `$filter` for date filtering
   - **After**: KQL date syntax embedded directly in `$search` queries
   - **Impact**: All searches now have server-side date filtering
   - **Example**: `"subject:meeting received:2026-02-01..2026-02-26"`
   - **Syntax**:
     - Date range: `received:2026-02-01..2026-02-26`
     - From date: `received>=2026-02-01`
     - Until date: `received<=2026-02-26`

2. **Server-Side Date Filtering**
   - **Before**: Fetched all emails, then filtered by date in Python
   - **After**: All date filters applied at server-side (via `$filter` or KQL)
   - **Impact**: ~90% faster for date-filtered searches
   - **Example**: Searching last 7 days now returns only matching emails from the server, reducing data transfer and processing

3. **Targeted $filter vs Full-Text $search**
   - **Before**: Used slow `$search` for full-text search across all fields
   - **After**:
     - Sender (exact email): `from/emailAddress/address eq '{sender}'` with `$filter`
     - Sender (fuzzy name): `"from:{name} received>=YYYY-MM-DD"` with KQL
     - Subject: `"subject:{query} received>=YYYY-MM-DD"` with KQL
     - Body: `"{query} received>=YYYY-MM-DD"` with KQL
   - **Impact**: ~70% faster for field-specific searches

4. **Well-Known Folder Cache**
   - **Before**: Always called `_get_folder_id_by_path()` requiring API calls
   - **After**: Added cache for common folders (Inbox, Sent, Drafts, Deleted, Archive, Junk)
   - **Impact**: Eliminates API calls for standard folders (~50% faster)
   - **Example**: Searching "Inbox" uses cached ID instead of API lookup

5. **Combined Filter Expressions**
   - **Before**: Date filtering done separately after API call
   - **After**: Combined all filters in single `$filter` expression or KQL query
   - **Impact**: Reduces network round trips
   - **Example**: `from/emailAddress/address eq 'john@example.com' and receivedDateTime ge 2024-01-01T00:00:00Z`

### Performance Metrics

After implementing the above optimizations:

1. **Date-Filtered Searches**: ~90% faster
2. **Sender/Recipient Searches**: ~70% faster
3. **Common Folder Searches**: ~50% faster
4. **General Searches**: ~30-50% faster

**Benchmark Results**:
- **100 emails**: 2.22 seconds (45.1 emails/second)
- **500 emails**: 3.50 seconds (142.7 emails/second)
- **1000 emails**: 5.39 seconds (185.6 emails/second)

### Legacy Optimizations

5. **Hard Limit**: All email search methods have a maximum limit defined by `MAX_EMAIL_SEARCH_LIMIT` (1000 emails) per search to prevent excessive API calls and memory usage. This limit is implemented as a constant in the codebase rather than a magic number, making it easier to maintain and modify in the future.

6. **Key Legacy Optimizations**:
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
