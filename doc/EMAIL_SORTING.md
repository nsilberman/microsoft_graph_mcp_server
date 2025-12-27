# Email Sorting and Pagination Documentation

## Overview

The Microsoft Graph MCP Server provides robust email sorting and pagination features to ensure efficient browsing of large email collections.

## Email Sorting

### Sorting Order

Emails are always sorted by received date in **descending order** (newest first). This applies to:

- API responses from `load_emails_by_folder`
- Cached emails retrieved via `email_cache.get_cached_emails()`
- Browsing results from `browse_email_cache`

### Timestamp Fields

Each email includes two timestamp fields to support both display and sorting:

#### receivedDateTime (Display)
- **Purpose**: Human-readable timestamp for display
- **Format**: "Fri 12/26/2025 09:27 PM"
- **Timezone**: User's local timezone
- **Usage**: Display purposes only

#### receivedDateTimeOriginal (Sorting)
- **Purpose**: Original ISO datetime for accurate sorting
- **Format**: "2025-12-26T13:27:44Z"
- **Timezone**: UTC
- **Usage**: Sorting and filtering

### Sorting Implementation

#### Graph Client (graph_client.py)

```python
sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTimeOriginal", ""), reverse=True)
```

#### Email Cache (email_cache.py)

```python
def get_cached_emails(self) -> List[Dict[str, Any]]:
    emails = self.cache["list_state"]["emails"].copy()
    return sorted(emails, key=lambda x: x.get("receivedDateTimeOriginal", ""), reverse=True)
```

## Pagination

### browse_email_cache Tool

The `browse_email_cache` tool provides paginated access to cached emails.

### Parameters

- **page_number** (required): Page number to view (starts at 1)

### Configuration

The number of emails per page is controlled by the `PAGE_SIZE` environment variable:

```env
# .env file
PAGE_SIZE=5
```

Default page size is 5 emails per page. This ensures consistent pagination across all browsing operations.

### Response Fields

The tool returns the following pagination metadata:

- **current_page**: Current page number being viewed
- **total_pages**: Total number of pages available
- **count**: Number of emails on current page
- **total_count**: Total number of emails in cache

### Pagination Calculation

```python
page_size = settings.page_size
total_pages = (total_count + page_size - 1) // page_size
start_idx = (page_number - 1) * page_size
end_idx = start_idx + page_size
```

## Usage Examples

### Example 1: Browse First Page

```python
{
  "page_number": 1
}
```

Response:
```json
{
  "emails": [...],
  "count": 5,
  "total_count": 10,
  "current_page": 1,
  "total_pages": 2
}
```

### Example 2: Browse Second Page

```python
{
  "page_number": 2
}
```

Response:
```json
{
  "emails": [...],
  "count": 5,
  "total_count": 10,
  "current_page": 2,
  "total_pages": 2
}
```

### Example 3: Adjust Page Size

To change the number of emails per page, modify the `PAGE_SIZE` environment variable:

```env
# .env file
PAGE_SIZE=10
```

Then browse:
```python
{
  "page_number": 1
}
```

Response:
```json
{
  "emails": [...],
  "count": 10,
  "total_count": 25,
  "current_page": 1,
  "total_pages": 3
}
```

## Email Filtering

### Days Parameter

The `load_emails_by_folder` tool supports filtering by days using the user's local timezone:

```python
# Load emails from last 1 day
result = await client.load_emails_by_folder(folder="Inbox", days=1)

# Load emails from last 7 days
result = await client.load_emails_by_folder(folder="Inbox", days=7)
```

### Local Time Filtering

When using the `days` parameter:

1. System calculates the cutoff time in user's local timezone
2. Converts cutoff to UTC for Microsoft Graph API query
3. Returns emails received after the cutoff

Example:
- User timezone: Asia/Shanghai (UTC+8)
- Current time: 2025-12-26 21:00:00 CST
- days=1
- Cutoff (local): 2025-12-25 21:00:00 CST
- Cutoff (UTC): 2025-12-25 13:00:00 UTC

## Testing

### Test Scripts

Two test scripts are available to verify sorting and pagination:

#### test_load_emails.py

Comprehensive testing including:
- Loading emails with `days` parameter
- Verifying sorting order (newest first)
- Testing cache operations
- Timezone conversion verification
- Pagination with configured page_size

Run with:
```bash
python test_load_emails.py
```

#### test_last_day.py

Focused testing for "last 1 day" email loading:
- Verifies local time filtering
- Checks sorting order
- Displays first and last 5 emails

Run with:
```bash
python test_last_day.py
```

## Best Practices

### 1. Always Use receivedDateTimeOriginal for Sorting

When implementing custom sorting logic, always use `receivedDateTimeOriginal`:

```python
sorted_emails = sorted(emails, key=lambda x: x['receivedDateTimeOriginal'], reverse=True)
```

### 2. Check Pagination Bounds

Before requesting a page, verify it's within bounds:

```python
if page_number > total_pages:
    print(f"Page {page_number} is out of range. Maximum page: {total_pages}")
```

### 3. Use Appropriate Page Sizes

Configure the `PAGE_SIZE` environment variable based on your needs:
- Small pages (5-10): Better for slow connections or limited display space
- Medium pages (20-50): Good balance for most use cases
- Large pages (100+): Only for batch processing or data export

Example:
```env
# .env file
PAGE_SIZE=20
```

### 4. Cache Results When Possible

The email cache persists to disk, so you can:
- Load emails once
- Browse multiple pages without re-fetching
- Clear cache when needed with `clear_email_cache`

## Troubleshooting

### Emails Not Sorted Correctly

**Symptom**: Newest emails don't appear first

**Solutions**:
1. Verify `receivedDateTimeOriginal` field is present
2. Check sorting logic uses `receivedDateTimeOriginal`, not `receivedDateTime`
3. Ensure `reverse=True` is set in sorting

### Pagination Shows Wrong Total

**Symptom**: `total_pages` doesn't match expected value

**Solutions**:
1. Verify `total_count` is correct
2. Check page size calculation: `(total_count + page_size - 1) // page_size`
3. Ensure cache is up to date (reload if needed)
4. Verify `PAGE_SIZE` environment variable is set correctly

### Local Time Filtering Not Working

**Symptom**: `days` parameter returns wrong emails

**Solutions**:
1. Verify user timezone is set correctly
2. Check `USER_TIMEZONE` in `.env` file
3. Review timezone handling documentation

## Related Documentation

- [README.md](../README.md) - Main project documentation
- [TIMEZONE_HANDLING.md](TIMEZONE_HANDLING.md) - Timezone conversion documentation
- [TEST_README.md](TEST_README.md) - Testing guide
