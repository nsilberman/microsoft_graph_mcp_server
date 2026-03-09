# Error Codes Reference

Quick reference for common errors and their solutions.

---

## Authentication Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Not authenticated" | No login or token expired | Call `login` then `complete_login` |
| "Token expired" | Access token expired (1hr) | Call `extend_token` to refresh |
| "Failed to complete login" | Browser auth not completed | Wait for user to complete browser auth |

**Quick Fix:**
```python
auth(action="login")       # Start login
# User completes browser auth...
auth(action="complete_login")  # REQUIRED!
```

---

## Cache Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "No emails in cache" | Cache empty | Call `search_emails` first |
| "No events in cache" | Cache empty | Call `search_events` first |
| "Invalid cache_number" | Number out of range | Use `number` field from browse results |

**Quick Fix:**
```python
search_emails(days=7)           # Load cache
browse_email_cache(page_number=1)  # See valid numbers
get_email_content(cache_number=21)  # Use number from browse
```

---

## Parameter Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Days parameter must be X or less" | Exceeded max days | Use smaller value or `time_range` |
| "Invalid action" | Typo or wrong case | Use exact action string from docs |
| "Required parameter missing" | Missing required param | Check tool description |

---

## Folder Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Folder not found" | Invalid path | List folders first: `manage_mail_folder(action="list")` |
| "Folder already exists" | Duplicate name | Use different name or delete existing |

---

## API Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Rate limit exceeded (HTTP 429)" | Too many calls | Wait `retry_after` seconds |
| "Service unavailable" | API down | Wait and retry |

---

## Email Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Email not found" | Deleted or moved | Search again to refresh cache |
| "Recipient limit exceeded" | >500 recipients | Use `bcc_csv_file` for large lists |

---

## Calendar Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Event not found" | Deleted or cancelled | Search again to refresh cache |
| "Event time conflicts" | Overlapping event | Check availability first |
| "Already responded to event" | Response already sent | Cannot change response |

**Prevention:**
```python
check_attendee_availability(attendees=["user@example.com"], date="2024-01-15")
# Use free slots from top_slots for event creation
```

---

## Contact Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "No contacts found" | No matches | Use broader search terms |
| "Contact search limit reached" | Max 10 results | Use more specific search |

---

## Template Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Template not found" | Invalid number or deleted | List templates first |

---

## Quick Troubleshooting Checklist

### Before calling tools:

1. **Check authentication:**
   ```python
   auth(action="check_status")
   ```

2. **Load cache first:**
   - Emails: `search_emails(days=7)` before `browse_email_cache`
   - Events: `search_events(time_range="this_week")` before `browse_events`

3. **Verify folders:**
   ```python
   manage_mail_folder(action="list")
   ```

---

## Common Mistakes

### ❌ Forgot complete_login
```python
auth(action="login")
# MISSING: auth(action="complete_login")
search_emails(days=7)  # Will fail!
```

### ✅ Correct
```python
auth(action="login")
# User completes browser auth
auth(action="complete_login")
search_emails(days=7)  # Works!
```

### ❌ Using array index
```python
browse_email_cache(page_number=1)
# Returns emails[0].number=21
get_email_content(cache_number=0)  # Wrong!
```

### ✅ Using number field
```python
browse_email_cache(page_number=1)
# Returns emails[0].number=21
get_email_content(cache_number=21)  # Correct!
```

---

## Getting Help

1. **Check tool descriptions** - Each tool has detailed parameter docs
2. **Check TOOL_DEPENDENCIES.md** - For tool call sequences
3. **Check logs** - Review server logs for details
4. **Consult documentation** - See `doc/` directory

---

**Document Version:** 2.0
**Last Updated:** 2026-03-09
