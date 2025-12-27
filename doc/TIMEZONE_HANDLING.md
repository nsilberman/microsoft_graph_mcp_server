# Timezone Handling Documentation

## Overview

The Microsoft Graph MCP Server automatically converts email timestamps to the user's local timezone for better readability and accuracy.

## How It Works

### Timezone Detection Priority

The system uses a three-tier fallback mechanism to determine the user's timezone:

1. **Microsoft Graph Mailbox Settings** (Primary)
   - Retrieves the user's mailbox timezone from Microsoft Graph API
   - Endpoint: `GET /me/mailboxSettings`
   - Field: `timeZone`

2. **Environment Variable** (Secondary)
   - Falls back to `USER_TIMEZONE` if Graph API fails
   - Set in `.env` file: `USER_TIMEZONE=Asia/Shanghai`
   - Supports IANA timezone names (e.g., `America/New_York`, `Europe/London`)

3. **UTC** (Default)
   - Used if both methods fail
   - Ensures the system always has a valid timezone

### Implementation Details

#### Graph Client Methods

```python
async def get_mailbox_settings(self) -> Dict[str, Any]:
    """Get user mailbox settings including timezone."""
    # Returns mailbox settings or empty dict on error

async def get_user_timezone(self) -> str:
    """Get user timezone with fallback."""
    # Tries mailbox settings, then env var, then UTC
```

#### Email Timestamp Conversion

```python
def _create_email_summary(self, email: Dict[str, Any], index: int, timezone_str: str = "UTC") -> Dict[str, Any]:
    # Converts UTC timestamp to user's local timezone
    # Preserves original ISO datetime for sorting
    # Returns formatted display timestamp
```

### Timestamp Fields

Each email summary includes two timestamp fields:

- **receivedDateTime**: Formatted display timestamp in user's local timezone
  - Format: "Fri 12/26/2025 09:27 PM"
  - Used for display purposes only

- **receivedDateTimeOriginal**: Original ISO 8601 datetime in UTC
  - Format: "2025-12-26T13:27:44Z"
  - Used for accurate sorting and filtering

## Configuration

### Setting User Timezone via Environment Variable

Create or update the `.env` file:

```env
USER_TIMEZONE=Asia/Shanghai
```

### Supported Timezone Formats

The system uses Python's `zoneinfo` module and supports all IANA timezone names:

- `Asia/Shanghai`
- `America/New_York`
- `Europe/London`
- `Asia/Tokyo`
- `Australia/Sydney`
- And many more...

See the [IANA Time Zone Database](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones) for a complete list.

## Error Handling

### Mailbox Settings Access Denied

If the Microsoft Graph API returns a 403 Forbidden error when accessing mailbox settings:

- The system gracefully falls back to the environment variable
- No error is raised to the user
- Email loading continues normally with the fallback timezone

### Invalid Timezone

If an invalid timezone string is provided:

- The system falls back to UTC
- A warning is logged for debugging purposes
- Email loading continues normally

## Examples

### Example 1: User in Shanghai

Input (UTC): `2025-12-26T13:27:44Z`

Output (Asia/Shanghai): `Fri 12/26/2025 09:27 PM`

### Example 2: User in New York

Input (UTC): `2025-12-26T13:27:44Z`

Output (America/New_York): `Fri 12/26/2025 08:27 AM`

### Example 3: User in London

Input (UTC): `2025-12-26T13:27:44Z`

Output (Europe/London): `Fri 12/26/2025 01:27 PM`

## Testing

Run the unit tests to verify timezone functionality:

```bash
# Test email loading with timezone conversion
python -m unittest tests/test_new_email_functions.py

# Test all email functions including timezone handling
python -m unittest tests/test_email_functions_unit.py
```

## Troubleshooting

### Timestamps showing UTC instead of local time

1. Check if `USER_TIMEZONE` is set in `.env` file
2. Verify the timezone name is valid (IANA format)
3. Check MCP server logs for mailbox settings errors
4. Ensure you have permission to access mailbox settings

### Incorrect timezone conversion

1. Verify your system timezone matches your location
2. Check the `USER_TIMEZONE` environment variable
3. Test with a known timezone (e.g., `UTC`, `America/New_York`)

## Related Documentation

- [README.md](../README.md) - Main project documentation
- [EMAIL_SORTING.md](EMAIL_SORTING.md) - Email sorting and filtering documentation
- [TEST_README.md](TEST_README.md) - Testing guide
