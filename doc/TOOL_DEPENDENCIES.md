# Tool Dependencies

This document describes tool call dependencies and workflow patterns to help LLMs understand how to use tools effectively.

## Table of Contents

- [Authentication Required](#authentication-required)
- [Cache-Dependent Tools](#cache-dependent-tools)
  - [Email Cache](#email-cache)
  - [Event Cache](#event-cache)
  - [Template Cache](#template-cache)
- [Sequence Dependencies](#sequence-dependencies)
  - [Authentication Workflow](#authentication-workflow)
  - [Email Workflow](#email-workflow)
  - [Calendar Workflow](#calendar-workflow)
  - [Template Workflow](#template-workflow)
- [Error Recovery](#error-recovery)

---

## Authentication Required

**All tools require authentication EXCEPT:**
- `auth` (obviously - it handles authentication)

### How to Authenticate

**When authentication needed, always try `refresh` FIRST:**

```python
result = auth(action="refresh")
# Check result["authenticated"]
```

**If `authenticated: true`:** Already logged in, proceed with your task.

**If `authenticated: false`:** Need to login, follow steps below:

1. **Start login process:**
   ```python
   auth(action="start")
   ```
   Returns:
   - `verification_uri`: URL for user to visit
   - `user_code`: Code for user to enter
   - `status`: "pending"

2. **User completes authentication in browser:**
   - User visits `verification_uri`
   - User enters `user_code`
   - User signs in to Microsoft 365

3. **Complete login (CRITICAL - REQUIRED!):**
   ```python
   auth(action="complete")
   ```
   Returns:
   - `status`: "success" if successful
   - `authenticated`: true
   - `message`: Confirmation message

   **IMPORTANT:** Without calling `complete`, authentication fails!

4. **Use any other tools:**
   Now you're authenticated and can use all other tools.

### Logout

```python
auth(action="logout")
```
Clears all authentication tokens and requires new login.

---

## Cache-Dependent Tools

Some tools require data to be loaded into cache first.

### Email Cache

These tools require emails to be loaded into cache first:

**Cache-Dependent Tools:**
- `browse_email_cache` - Requires `search_emails` to be called first
- `get_email_content` - Requires cache_number from `browse_email_cache`
- `manage_emails` (single actions) - Requires cache_number from `browse_email_cache`
- `send_email` (reply/forward) - Requires cache_number from `browse_email_cache`

**Workflow:**
```
search_emails → browse_email_cache → get_email_content / manage_emails / send_email
```

**Example - Complete email workflow:**
```python
# Step 1: Search emails (loads into cache)
search_emails(query="meeting", folder="Inbox", days=7)

# Step 2: Browse cache
browse_email_cache(page_number=1, mode="llm")  # Returns 20 emails

# Step 3: Get full content
get_email_content(cache_number=5)  # Get full email #5

# Step 4: Act on email
send_email(action="reply", cache_number=5, htmlbody="<p>Thanks!</p>")
```

**Important Notes:**
- `search_emails` clears existing cache and loads new emails
- Cache numbers use the 'number' field from browse results
- Cache numbers persist across pagination (email with number 5 is always number 5)
- Browse modes:
  - `mode="user"`: 5 emails per page (human-friendly)
  - `mode="llm"`: 20 emails per page (LLM-friendly)

**When cache is empty:**
If you see "No emails in cache" error:
1. Call `search_emails` to load emails into cache
2. Retry `browse_email_cache`

**Invalid cache number:**
If you see "Invalid cache_number" error:
1. Call `browse_email_cache` or `browse_events` to see valid cache numbers
2. Use a cache_number from the results

### Event Cache

These tools require events to be loaded into cache first:

**Cache-Dependent Tools:**
- `browse_events` - Requires `search_events` to be called first
- `get_event_detail` - Requires cache_number from `browse_events` or `search_events`
- `manage_event_as_attendee` - Requires cache_number from `browse_events` or `search_events`
- `manage_event_as_organizer` (update/cancel/forward/email_attendees) - Requires cache_number

**Workflow:**
```
search_events → browse_events → get_event_detail / manage_event_as_attendee / manage_event_as_organizer
```

**Example - Complete event workflow:**
```python
# Step 1: Search events (loads into cache)
search_events(time_range="this_week")

# Step 2: Browse cache
browse_events(page_number=1, mode="llm")  # Returns 20 events

# Step 3: Get full details
get_event_detail(cache_number="3")  # Get full event #3

# Step 4: Act on event
manage_event_as_attendee(action="accept", cache_number=3, comment="I'll be there!")
```

**Important Notes:**
- `search_events` clears existing cache and loads new events
- Cache numbers use the 'number' field from browse results
- Cache numbers persist across pagination (event with number 5 is always number 5)
- Browse modes:
  - `mode="user"`: 5 events per page (human-friendly)
  - `mode="llm"`: 20 events per page (LLM-friendly)

### Template Cache

These tools use template cache internally:

**Cache-Dependent Tools:**
- `manage_templates` (list action) - Uses template cache
- `manage_templates` (get/update/send/delete) - Requires template_number from list

**Workflow:**
```
manage_templates(action="list") → manage_templates(action="get"/"update"/"send"/"delete")
```

**Example - Complete template workflow:**
```python
# Step 1: Find email to template
search_emails(query="newsletter", folder="Inbox")
browse_email_cache(page_number=1, mode="llm")

# Step 2: Create template from email
manage_templates(action="create_from_email", cache_number=5)

# Step 3: List templates
manage_templates(action="list", page_number=1)

# Step 4: View as simple text (user view)
manage_templates(action="get", template_number=1, return_html=false)

# Step 5: LLM gets full HTML for editing
manage_templates(action="get", template_number=1, return_html=true)

# Step 6: LLM updates template
manage_templates(action="update", template_number=1, htmlbody="<html>...</html>")

# Step 7: User verifies changes
manage_templates(action="get", template_number=1, return_html=false)

# Step 8: Send template
manage_templates(action="send", template_number=1, to=["recipient@example.com"])
```

**Important Notes:**
- Templates are draft emails in Templates folder
- Templates have subject prefix "Template:" for identification
- Template numbers use the 'number' field from browse results
- Sending a template creates a copy and sends it (preserves original)

---

## Sequence Dependencies

### Authentication Workflow

**When authentication needed:**

```python
# 1. Try refresh FIRST
result = auth(action="refresh")

if result["authenticated"]:
    # Already logged in, proceed with your task
    pass
else:
    # 2. Need to login
    result1 = auth(action="start")
    # User completes browser authentication with result1['verification_uri'] and result1['user_code']

    # 3. Complete login (REQUIRED!)
    result2 = auth(action="complete")
    # Now authenticated

# 4. Use any tools
# Note: Tokens auto-refresh when expired, no manual action needed
```

**Common mistake:**
❌ **Incorrect:**
```python
auth(action="start")
# User completes in browser
# MISSING: auth(action="complete")
search_emails(days=7)  # This will fail!
```

✅ **Correct:**
```python
# Try refresh first
result = auth(action="refresh")
if not result["authenticated"]:
    auth(action="start")
    # User completes in browser
    auth(action="complete")  # MUST call this!
search_emails(days=7)  # Now works
```

### Email Workflow

**Complete email workflow from search to action:**

```python
# 1. Search emails
search_emails(
    query="project update",
    search_type="subject",
    folder="Inbox",
    days=7
)

# 2. Browse cache
browse_email_cache(page_number=1, mode="llm")

# 3. Get full content
get_email_content(cache_number=3)

# 4. Act on email (choose one):

# Option A: Reply
send_email(
    action="reply",
    cache_number=3,
    htmlbody="<p>Thank you for the update!</p>"
)

# Option B: Forward
send_email(
    action="forward",
    cache_number=3,
    to=["manager@example.com"],
    htmlbody="<p>FYI</p>"
)

# Option C: Move
manage_emails(
    action="move_single",
    cache_number=3,
    destination_folder="Archive/2024"
)

# Option D: Flag
manage_emails(
    action="flag_single",
    cache_number=3,
    flag_status="flagged"
)

# Option E: Delete
manage_emails(
    action="delete_single",
    cache_number=3
)
```

### Calendar Workflow

**Complete calendar workflow from availability to event management:**

```python
# 1. Check availability first (recommended)
check_attendee_availability(
    attendees=["user1@example.com", "user2@example.com"],
    date="2024-01-15"
)
# Returns: availability_view, schedule_items, top_slots

# 2. Create meeting at free slot
manage_event_as_organizer(
    action="create",
    subject="Project Review",
    start="2024-01-15T10:00",
    end="2024-01-15T11:00",
    timezone="America/New_York",
    attendees=["user1@example.com", "user2@example.com"],
    isOnlineMeeting=True,
    onlineMeetingProvider="teamsForBusiness"
)

# 3. Verify it was created
search_events(query="Project Review", time_range="this_week")

# 4. Browse events
browse_events(page_number=1, mode="llm")

# 5. Get event details
get_event_detail(cache_number="1")

# 6. Optional: Respond to someone else's event
manage_event_as_attendee(
    action="accept",
    cache_number="5",
    comment="I'll attend!"
)
```

### Template Workflow

**Complete template workflow for email templates:**

```python
# 1. Find email to use as template
search_emails(query="newsletter", folder="Inbox")

# 2. Browse emails
browse_email_cache(page_number=1, mode="llm")

# 3. Create template from email
manage_templates(action="create_from_email", cache_number=5)

# 4. List templates
manage_templates(action="list", page_number=1)

# 5. User views as simple text
manage_templates(action="get", template_number=1, return_html=false)

# 6. User provides update instructions to LLM

# 7. LLM gets full HTML for editing
manage_templates(action="get", template_number=1, return_html=true)

# 8. LLM applies updates and saves
manage_templates(
    action="update",
    template_number=1,
    subject="Updated Newsletter",
    htmlbody="<html>...updated content...</html>"
)

# 9. User verifies changes
manage_templates(action="get", template_number=1, return_html=false)

# 10. Send template
manage_templates(
    action="send",
    template_number=1,
    to=["recipient1@example.com", "recipient2@example.com"]
)
```

---

## Error Recovery

### Authentication Errors

**"Not authenticated" error**

**Meaning:** User hasn't logged in or token expired

**Solution:**
```python
# 1. Try refresh first
result = auth(action="refresh")

if result["authenticated"]:
    # Already authenticated, retry the operation
    search_emails(days=7)
else:
    # 2. Need to login
    auth(action="start")
    # User completes browser auth

    # 3. Complete login (REQUIRED!)
    auth(action="complete")

    # 4. Retry the failed operation
    search_emails(days=7)  # Now works
```

### Token Expired

**"Token expired" error**

**Meaning:** Access token expired (default 1 hour lifetime)

**Solution:**
Tokens auto-refresh, so just retry your operation. If that fails:

```python
# Try refresh
result = auth(action="refresh")
if result["authenticated"]:
    # Retry operation
    search_emails(days=7)
else:
    # Need to re-login
    auth(action="start")
    # User completes browser auth
    auth(action="complete")
```

### Cache Errors

**"No emails in cache" error**

**Meaning:** Email cache is empty

**Solution:**
```python
# 1. Search emails to load into cache
search_emails(days=7, folder="Inbox")

# 2. Browse cache
browse_email_cache(page_number=1, mode="llm")
```

**"No events in cache" error**

**Meaning:** Event cache is empty

**Solution:**
```python
# 1. Search events to load into cache
search_events(time_range="this_week")

# 2. Browse cache
browse_events(page_number=1, mode="llm")
```

**"Invalid cache_number" error**

**Meaning:** Cache number doesn't exist or out of range

**Solution:**
```python
# 1. Browse cache to see valid numbers
browse_email_cache(page_number=1, mode="llm")
# Or browse_events(page_number=1, mode="llm")

# 2. Use a cache_number from the 'number' field in browse results
get_email_content(cache_number=5)  # Must be a number from the 'number' field in browse results
```

---

## Quick Reference

### Email Tools
- `search_emails`: Load emails into cache
- `browse_email_cache`: Browse cached emails (pagination)
- `get_email_content`: Get full email content
- `send_email`: Send new email, reply, forward
- `manage_emails`: Move, delete, archive, flag, categorize
- `manage_mail_folder`: Create, rename, move, delete, get_details, list folders

### Calendar Tools
- `search_events`: Load events into cache
- `browse_events`: Browse cached events (pagination)
- `get_event_detail`: Get full event details
- `check_attendee_availability`: Check when people are free
- `manage_event_as_organizer`: Create, update, cancel, forward, email_attendees
- `manage_event_as_attendee`: Accept, decline, tentatively accept, propose_new_time, delete_cancelled

### User Tools
- `search_contacts`: Find people in organization directory
- `user_settings`: Initialize or update user settings

### Other Tools
- `manage_templates`: Create, edit, send email templates
- `auth`: Authentication management

---

## Best Practices

1. **Always call `complete` after `start`**
   - Without this, authentication fails

2. **Load cache before browsing**
   - Use `search_emails` or `search_events` first
   - Then use `browse_*_cache` tools

3. **Use cache numbers from browse results**
   - Use the `number` field from browse results (e.g., 21, 22, 23)
   - Cache numbers are displayed in the browse results and persist across pagination

4. **Check availability before creating meetings**
   - Use `check_attendee_availability` first
   - Create meeting at free slot

5. **Use appropriate browsing mode**
   - `mode="user"`: 5 items (for user review)
   - `mode="llm"`: 20 items (for LLM processing)

6. **Handle rate limits gracefully**
   - Wait `retry_after` seconds on HTTP 429
   - Don't immediately retry

7. **Always try `refresh` first when auth needed**
   - `auth(action="refresh")` checks auth status
   - If authenticated, proceed; if not, use `start` flow

---

**Document Version:** 3.0
**Last Updated:** 2026-03-23
