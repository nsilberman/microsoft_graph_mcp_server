# Error Codes Reference

This document provides comprehensive error messages, their meanings, and solutions for the Microsoft Graph MCP Server.

## Table of Contents

- [Authentication Errors](#authentication-errors)
- [Cache Errors](#cache-errors)
- [Parameter Errors](#parameter-errors)
- [Folder Errors](#folder-errors)
- [API Errors](#api-errors)
- [Email Errors](#email-errors)
- [Calendar Errors](#calendar-errors)
- [Contact Errors](#contact-errors)
- [Template Errors](#template-errors)
- [Generic Errors](#generic-errors)

---

## Authentication Errors

### "Not authenticated"

**Meaning:** User hasn't logged in or authentication token has expired.

**When this happens:**
- Calling any tool (except `auth`) before authenticating
- Access token expired and not refreshed
- Authentication cleared via logout

**Solution:**
```python
# 1. Start login
auth(action="login")
# User completes browser authentication with verification_url and user_code

# 2. Complete login (REQUIRED!)
auth(action="complete_login")

# 3. Retry the failed operation
search_emails(days=7)  # Now works
```

**Prevention:**
- Always call `auth(action="complete_login") after `auth(action="login")`
- Use `auth(action="extend_token")` to refresh before token expires (1 hour lifetime)
- Check authentication status with `auth(action="check_status")` if unsure

---

### "Token expired"

**Meaning:** Access token expired (default 1 hour lifetime).

**When this happens:**
- More than 1 hour since last authentication
- Calling any tool with expired token

**Solution:**
```python
# Preferred: Extend token without user interaction
auth(action="extend_token")
# Returns fresh access token with new 1-hour lifetime

# Alternative: Full re-login
auth(action="logout")
auth(action="login")
# User completes browser auth
auth(action="complete_login")
```

**Prevention:**
- Extend token before it expires using `auth(action="extend_token")`
- Refresh tokens are long-lived (up to 90 days with regular use)
- Check token expiry with `auth(action="check_status")`

---

### "Failed to complete login"

**Meaning:** User didn't complete authentication in browser within timeout period.

**When this happens:**
- Calling `auth(action="complete_login")` before user completes browser auth
- User didn't visit verification_url
- User didn't enter user_code

**Solution:**
```python
# 1. Start new login
auth(action="login")

# 2. User completes authentication in browser:
#    - Visit verification_url
#    - Enter user_code

# 3. Complete login (wait for browser auth)
auth(action="complete_login")
```

**Prevention:**
- Ensure user visits verification_url and enters user_code
- Wait for user to complete browser auth before calling `complete_login`
- Timeout is 60 seconds for user to complete authentication

---

## Cache Errors

### "No emails in cache"

**Meaning:** Email cache is empty.

**When this happens:**
- Calling `browse_email_cache` without first calling `search_emails`
- Calling `get_email_content` with cache_number from empty cache
- Cache expired and cleared

**Solution:**
```python
# 1. Search emails to load into cache
search_emails(days=7, folder="Inbox")

# 2. Browse cache
browse_email_cache(page_number=1, mode="llm")

# 3. Now get_email_content works
get_email_content(cache_number=1)
```

**Prevention:**
- Always call `search_emails` before `browse_email_cache`
- `search_emails` clears cache and loads new emails
- Cache persists across tool calls

---

### "No events in cache"

**Meaning:** Event cache is empty.

**When this happens:**
- Calling `browse_events` without first calling `search_events`
- Calling `get_event_detail` with cache_number from empty cache
- Cache expired and cleared

**Solution:**
```python
# 1. Search events to load into cache
search_events(time_range="this_week")

# 2. Browse cache
browse_events(page_number=1, mode="llm")

# 3. Now get_event_detail works
get_event_detail(cache_number="1")
```

**Prevention:**
- Always call `search_events` before `browse_events`
- `search_events` clears cache and loads new events
- Cache persists across tool calls

---

### "Invalid cache_number"

**Meaning:** Cache number doesn't exist or is out of range.

**When this happens:**
- Using array index instead of cache number (e.g., cache_number=1 but should be 21)
- Cache number exceeds current cache size
- Cache cleared and trying to use old cache number

**Solution:**
```python
# 1. Browse cache to see valid numbers
browse_email_cache(page_number=1, mode="llm")
# Or browse_events(page_number=1, mode="llm")

# 2. Use a cache_number from the results
# Look at the "number" field in browse results, not array index
# Example: If browse returns emails[0].number=21, use cache_number=21
get_email_content(cache_number=21)  # Correct
```

**Prevention:**
- Always use the `number` field from browse results, not array index
- Cache numbers are 1-based (start at 1, not 0)
- Cache numbers persist across pagination

---

## Parameter Errors

### "Days parameter must be X or less"

**Meaning:** Exceeded maximum search days limit.

**When this happens:**
- Using `days` parameter greater than `max_search_days` (default: 90)

**Solution:**
```python
# Use smaller days value
search_emails(days=90)  # Maximum
# Or use specific date ranges
search_emails(
    start_date="2024-01-01",
    end_date="2024-03-31"
)
# Or use time_range
search_emails(time_range="this_month")
```

**Prevention:**
- Maximum `days` is 90 (configurable via MAX_SEARCH_DAYS)
- Use `time_range` for broader ranges without hitting limit
- Use `start_date`/`end_date` for specific ranges

---

### "Invalid action"

**Meaning:** Action parameter not recognized.

**When this happens:**
- Typo in action parameter
- Using action not in supported list
- Case-sensitive action name

**Solution:**
```python
# Check tool description for valid actions
# Examples:
auth(action="login")  # Valid
auth(action="Login")   # INVALID - case-sensitive

manage_emails(action="move_single")  # Valid
manage_emails(action="MOVE_SINGLE")  # INVALID - case-sensitive
```

**Prevention:**
- Use exact action strings from tool descriptions
- Actions are case-sensitive
- Refer to tool documentation for valid action values

---

### "Required parameter missing"

**Meaning:** Required parameter not provided.

**When this happens:**
- Missing required parameter in tool call
- Providing None or empty string for required parameter

**Solution:**
```python
# Check tool description for required parameters
# Example - send_email requires action and htmlbody:
send_email(
    action="send_new",
    to=["recipient@example.com"],
    subject="Test",
    htmlbody="<p>Hello!</p>"  # Required!
)
```

**Prevention:**
- Check tool description for required parameters
- All required parameters are listed
- Optional parameters have default values

---

## Folder Errors

### "Folder not found"

**Meaning:** Folder path doesn't exist.

**When this happens:**
- Typo in folder path
- Folder doesn't exist in user's mailbox
- Invalid folder hierarchy

**Solution:**
```python
# 1. List all folders to see what exists
manage_mail_folder(action="list")

# 2. Use correct folder path from list results
search_emails(folder="Inbox")  # Correct
search_emails(folder="Inbox/Projects/2024")  # Correct
```

**Prevention:**
- Use `manage_mail_folder(action="list")` to see all folders
- Folder paths use forward slashes: "Inbox/Projects/2024"
- Well-known folders: Inbox, SentItems, Drafts, DeletedItems, Archive, JunkEmail

---

### "Folder already exists"

**Meaning:** Trying to create folder with name that already exists.

**When this happens:**
- Creating folder with duplicate name
- Not checking if folder exists first

**Solution:**
```python
# Option A: Use existing folder
search_emails(folder="Projects")  # Use existing folder

# Option B: Choose different name
manage_mail_folder(
    action="create",
    folder_name="Projects2024",  # Different name
    parent_folder="Inbox"
)

# Option C: Delete existing first
manage_mail_folder(action="delete", folder_path="Inbox/Projects")
manage_mail_folder(
    action="create",
    folder_name="Projects",
    parent_folder="Inbox"
)
```

**Prevention:**
- List folders first to check if name exists
- Choose unique folder names
- Include date or version in folder names

---

## API Errors

### "Rate limit exceeded (HTTP 429)"

**Meaning:** Too many API calls in short time period.

**When this happens:**
- Exceeding Microsoft Graph API rate limits
- Making too many rapid API calls
- Batch operations exceeding limits

**Solution:**
```python
# 1. Check response for retry_after field
# Response includes:
# {
#   "retry_after": 60,
#   "message": "Rate limit exceeded. Please wait 60 seconds."
# }

# 2. Wait the specified time
import time
time.sleep(60)  # Wait 60 seconds (or retry_after value)

# 3. Retry the operation
search_emails(days=7)  # Retry
```

**Prevention:**
- Use pagination instead of fetching all at once
- Use caching (email_cache, event_cache) to reduce API calls
- Wait `retry_after` seconds before retrying
- Batch operations when possible (manage_emails with cache_numbers)

---

### "Service unavailable"

**Meaning:** Microsoft Graph API is down.

**When this happens:**
- Microsoft Graph API service outage
- Temporary server issues

**Solution:**
```python
# 1. Wait a few minutes
import time
time.sleep(60)  # Wait 1 minute

# 2. Retry the operation
search_emails(days=7)

# 3. Check service status if persists
# Visit: https://status.microsoft.com
```

**Prevention:**
- Check https://status.microsoft.com for service status
- Wait and retry on temporary failures
- No code-level prevention (service issue)

---

## Email Errors

### "Email not found"

**Meaning:** Email ID doesn't exist or was deleted.

**When this happens:**
- Trying to access deleted email
- Email moved to different folder
- Using invalid cache number

**Solution:**
```python
# 1. Search to find email again
search_emails(query="meeting", days=7)

# 2. Browse cache
browse_email_cache(page_number=1, mode="llm")

# 3. Use new cache number
get_email_content(cache_number=5)  # Use current cache number
```

**Prevention:**
- Use recent cache numbers (cache expires after 24 hours)
- Search again if email might have been moved/deleted
- Use `search_emails` to refresh cache

---

### "Recipient limit exceeded"

**Meaning:** Too many recipients (TO + CC + BCC).

**When this happens:**
- Total recipients exceed 500 limit
- Too many recipients in single email

**Solution:**
```python
# Option A: Use BCC CSV file for large lists
send_email(
    action="send_new",
    to=["recipient@example.com"],
    subject="Newsletter",
    htmlbody="<p>...</p>",
    bcc_csv_file="/path/to/recipients.csv"  # Auto-batches to 500
)

# Option B: Reduce recipient count
send_email(
    action="send_new",
    to=["r1@example.com"],  # Smaller list
    subject="Update",
    htmlbody="<p>...</p>"
)
```

**Prevention:**
- Maximum total recipients: 500 (TO + CC + BCC)
- Use BCC CSV file for large lists (auto-batches)
- Check recipient count before sending

---

## Calendar Errors

### "Event not found"

**Meaning:** Event doesn't exist or was cancelled.

**When this happens:**
- Trying to access deleted/cancelled event
- Event ID invalid
- Using invalid cache number

**Solution:**
```python
# 1. Search to find event again
search_events(query="Project Review", time_range="this_week")

# 2. Browse cache
browse_events(page_number=1, mode="llm")

# 3. Use new cache number
get_event_detail(cache_number="3")  # Use current cache number
```

**Prevention:**
- Use recent cache numbers (cache expires after 24 hours)
- Search again if event might have been deleted
- Use `search_events` to refresh cache

---

### "Event time conflicts"

**Meaning:** New event overlaps with existing event.

**When this happens:**
- Creating event at time when already busy
- Updating event to time with conflict

**Solution:**
```python
# 1. Check availability first
check_attendee_availability(
    attendees=["user@example.com"],
    date="2024-01-15"
)
# Returns: availability_view, schedule_items, top_slots

# 2. Create event at free slot from top_slots
manage_my_event(
    action="create",
    subject="Meeting",
    start="2024-01-15T14:00",  # Free slot
    end="2024-01-15T15:00",
    attendees=["user@example.com"]
)
```

**Prevention:**
- Always use `check_attendee_availability` before creating events
- Choose free time slots from `top_slots` in availability check
- Review existing events with `search_events` first

---

### "Already responded to event"

**Meaning:** Trying to respond to an event that was already responded to.

**When this happens:**
- Responding to event after accepting/declining it
- Trying to change response after it was sent

**Solution:**
```python
# 1. Check current response status
get_event_detail(cache_number="5")
# Returns event with responseStatus field

# 2. Cannot change response
# Once response is sent, it cannot be changed
# Contact organizer to cancel/reschedule if needed
```

**Prevention:**
- Check event details before responding
- Response status: accepted, declined, tentative, none
- Only respond if status is "none" or "tentative"

---

## Contact Errors

### "No contacts found"

**Meaning:** Search returned no results.

**When this happens:**
- Search query too specific
- Person doesn't exist in organization directory
- Typo in name or email

**Solution:**
```python
# Option A: Try broader search terms
search_contacts(query="John")  # Broader than "John Smith"

# Option B: Use partial name match
search_contacts(query="Smith")  # Match last name

# Option C: Use email address
search_contacts(query="john.smith@company.com")  # Exact match
```

**Prevention:**
- Use partial names for broader search
- Try multiple search terms
- Check spelling

---

### "Contact search limit reached"

**Meaning:** More contacts exist than returned.

**When this happens:**
- Search returns max contacts (default: 10)
- More results available

**Solution:**
```python
# Response includes:
# {
#   "contacts": [...],
#   "count": 10,
#   "limit_reached": true,  # Indicates more results available
#   "message": "Showing 10 contacts (limit reached). More results available - use more specific search terms to narrow results."
# }

# Use more specific search terms to narrow results
search_contacts(query="John Smith Engineering")  # More specific
```

**Prevention:**
- Use specific search terms to narrow results
- Maximum contacts returned: 10 (configurable via CONTACT_SEARCH_LIMIT)
- Use department or role in search query

---

## Template Errors

### "Template not found"

**Meaning:** Template doesn't exist or was deleted.

**When this happens:**
- Using invalid template_number
- Template deleted

**Solution:**
```python
# 1. List templates to see valid numbers
manage_templates(action="list", page_number=1)

# 2. Use valid template_number from list
manage_templates(action="get", template_number=1)
```

**Prevention:**
- List templates first to see valid template numbers
- Template numbers are 1-based
- Cache persists across tool calls

---

## Generic Errors

### "Internal server error"

**Meaning:** Unexpected error on server side.

**When this happens:**
- Server-side exception
- Unknown error condition

**Solution:**
```python
# 1. Retry the operation
search_emails(days=7)

# 2. If persists, report error
# Check logs for more details

# 3. Restart server if needed
# Restart MCP server
```

**Prevention:**
- Retry on transient errors
- Check logs for details
- Report persistent errors

---

### "Validation error"

**Meaning:** Input validation failed.

**When this happens:**
- Invalid parameter type
- Invalid parameter value
- Parameter constraint violated

**Solution:**
```python
# Check error message for specific validation failure
# Example error:
# "cache_number must be 1 or greater, got 0"

# Fix the validation issue
get_email_content(cache_number=1)  # Use valid cache number (>= 1)
```

**Prevention:**
- Check tool descriptions for parameter requirements
- Use correct parameter types (string, integer, array, etc.)
- Respect parameter constraints (minimum, maximum, enum values)

---

### "Error during operation"

**Meaning:** General error during tool execution.

**When this happens:**
- Network error
- API error
- Unexpected exception

**Solution:**
```python
# Check error message for specific issue
# Usually includes context about what failed

# Retry if transient
search_emails(days=7)

# Check authentication if auth-related error
auth(action="check_status")
```

**Prevention:**
- Ensure authentication is valid
- Check network connectivity
- Verify parameter values

---

## Quick Troubleshooting Guide

### Before calling tools:

1. **Check authentication:**
   ```python
   auth(action="check_status")
   ```
   If `authenticated: false`, authenticate first.

2. **Load cache for email/event tools:**
   - For emails: Call `search_emails` before `browse_email_cache`
   - For events: Call `search_events` before `browse_events`

3. **Check folder paths:**
   ```python
   manage_mail_folder(action="list")
   ```
   Verify folder path exists before using.

### Common mistakes:

❌ **Don't forget complete_login:**
```python
auth(action="login")
# User completes browser auth
# MISSING: auth(action="complete_login")  ← FORGOT!
search_emails(days=7)  # Will fail!
```

✅ **Correct:**
```python
auth(action="login")
# User completes browser auth
auth(action="complete_login")  # ← REMEMBER THIS!
search_emails(days=7)  # Now works
```

❌ **Don't use array indices:**
```python
browse_email_cache(page_number=1, mode="llm")
# Returns emails[0].number=21
get_email_content(cache_number=0)  # Wrong! Array index
```

✅ **Correct:**
```python
browse_email_cache(page_number=1, mode="llm")
# Returns emails[0].number=21
get_email_content(cache_number=21)  # Correct! Use number field
```

❌ **Don't search wrong tool:**
```python
search_contacts(query="meeting@example.com")  # Wrong! Searches people
```

✅ **Correct:**
```python
search_emails(query="meeting@example.com", search_type="sender")  # Correct! Searches emails
```

---

## Getting Help

If you encounter errors not covered here:

1. **Check tool descriptions:** Each tool has detailed parameter descriptions
2. **Check TOOL_DEPENDENCIES.md:** For tool call sequences and workflows
3. **Check logs:** Review server logs for detailed error information
4. **Consult documentation:** See `doc/` directory for detailed guides

---

**Document Version:** 1.0
**Last Updated:** 2025-01-09
