# Timezone Handling Documentation

## Overview

The Microsoft Graph MCP Server automatically converts email and calendar timestamps to the user's local timezone for better readability and accuracy.

## How It Works

### Timezone Detection Priority

The system uses a three-tier fallback mechanism to determine the user's timezone:

1. **Server Local Timezone** (Primary)
   - Uses server's local timezone detected from the system
   - Automatically converts Windows timezone names to IANA format
   - Code location: `email_client.py:60-81`

2. **Environment Variable** (Secondary)
   - Falls back to `USER_TIMEZONE` if system timezone detection fails
   - Set in `.env` file: `USER_TIMEZONE=Asia/Shanghai`
   - Supports IANA timezone names (e.g., `America/New_York`, `Europe/London`)

3. **UTC** (Default)
   - Used if both methods fail
   - Ensures system always has a valid timezone

### Implementation Details

#### EmailClient.get_user_timezone()

```python
async def get_user_timezone(self) -> str:
    """Get user's timezone identifier. Uses server local timezone with caching."""
    if self._user_timezone_cache is not None:
        return self._user_timezone_cache

    try:
        # Get server's local timezone from the system
        local_tz = datetime.now().astimezone().tzinfo
        if local_tz:
            tz_str = str(local_tz)
            if tz_str and tz_str != "UTC":
                # Validate it's a valid IANA timezone
                self._user_timezone_cache = date_handler.convert_to_iana_timezone(tz_str)
                return self._user_timezone_cache
    except Exception:
        pass

    # Fall back to environment variable
    self._user_timezone_cache = date_handler.convert_to_iana_timezone(
        settings.user_timezone
    )
    return self._user_timezone_cache
```

#### Email Timestamp Conversion

```python
def _create_email_summary(self, email: Dict[str, Any], index: int, user_tz: ZoneInfo) -> Dict[str, Any]:
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

Create or update `.env` file:

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

See [IANA Time Zone Database](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones) for a complete list.

### Windows Timezone Conversion

The system includes automatic conversion from Windows timezone names to IANA format via `date_handler.convert_to_iana_timezone()`:

- Windows: "China Standard Time" → IANA: "Asia/Shanghai"
- Windows: "Eastern Standard Time" → IANA: "America/New_York"
- Windows: "GMT Standard Time" → IANA: "Europe/London"
- Windows: "India Standard Time" → IANA: "Asia/Kolkata"
- Windows: "FLE Standard Time" → IANA: "Europe/Kiev" (Finland, Russia, Estonia, Latvia, Lithuania)
- Windows: "W. Europe Standard Time" → IANA: "Europe/Paris"
- Windows: "Romance Standard Time" → IANA: "Europe/Paris"
- Windows: "E. Europe Standard Time" → IANA: "Europe/Chisinau"
- Windows: "Central Europe Standard Time" → IANA: "Europe/Budapest"
- Windows: "Central European Standard Time" → IANA: "Europe/Warsaw"
- Windows: "Tokyo Standard Time" → IANA: "Asia/Tokyo"
- Windows: "Singapore Standard Time" → IANA: "Asia/Singapore"
- Windows: "Pacific Standard Time" → IANA: "America/Los_Angeles"
- Windows: "Central Standard Time" → IANA: "America/Chicago"
- Windows: "AUS Eastern Standard Time" → IANA: "Australia/Sydney"

This conversion is applied when the `USER_TIMEZONE` environment variable contains a Windows timezone name.

## Error Handling

### Timezone Detection Failure

If timezone detection fails at all levels:

- The system gracefully falls back to UTC
- No error is raised to the user
- Email/calendar loading continues normally with UTC timestamps
- A warning is logged for debugging purposes

### Invalid Timezone

If an invalid timezone string is provided:

- The system attempts to convert Windows timezone names to IANA format
- If conversion still fails, falls back to UTC
- A warning is logged for debugging purposes
- Email/calendar loading continues normally

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

Run unit tests to verify timezone functionality:

```bash
# Test email loading with timezone conversion
python -m pytest tests/test_email_handlers.py -k timezone

# Test date handler timezone conversion
python -m pytest tests/test_date_handler.py
```

## Troubleshooting

### Timestamps showing UTC instead of local time

1. Check if `USER_TIMEZONE` is set in `.env` file
2. Verify timezone name is valid (IANA format)
3. Check your system timezone settings
4. Ensure the MCP server is running with access to system timezone information

### Incorrect timezone conversion

1. Verify your system timezone matches your location
2. Check `USER_TIMEZONE` environment variable
3. Test with a known timezone (e.g., `UTC`, `America/New_York`)
4. Check MCP server logs for timezone conversion warnings

### Windows timezone not converting

If using Windows timezone names:

1. Verify the Windows timezone name is correct
2. Check if the conversion is supported in `date_handler.convert_to_iana_timezone()`
3. Try using the IANA timezone name directly

## Related Documentation

- [README.md](../README.md) - Main project documentation
- [EMAIL_SORTING.md](EMAIL_SORTING.md) - Email sorting and filtering documentation
- [CALENDAR.md](CALENDAR.md) - Calendar timezone handling
