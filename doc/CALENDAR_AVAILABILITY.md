# Calendar Availability Documentation

## Overview

The Microsoft Graph MCP Server provides comprehensive calendar availability checking functionality, including working hours retrieval, availability time slot display, and timezone-aware scheduling.

## Working Hours Retrieval

### How It Works

The system retrieves working hours from Microsoft Graph mailbox settings to determine when attendees are available for meetings.

### Implementation

The working hours are retrieved in the `check_attendee_availability` function in [calendar_handlers.py](file:///c:/Project/microsoft_graph_mcp_server/microsoft_graph_mcp_server/handlers/calendar_handlers.py):

```python
async def check_attendee_availability(
    attendees: List[str],
    date: str,
    availability_view_interval: Optional[int] = None,
    time_zone: Optional[str] = None
) -> Dict[str, Any]:
    """Check availability of attendees for a given date."""
    # Get user timezone
    timezone_str = time_zone if time_zone else await get_user_timezone()
    user_tz = DateHandler.get_user_timezone_object(timezone_str)
    
    # Get working hours for each attendee
    working_hours_map = {}
    for attendee in attendees:
        try:
            attendee_tz_name = await get_attendee_timezone(attendee)
            attendee_tz = DateHandler.get_user_timezone_object(attendee_tz_name)
            
            working_hours = await calendar_client.get_working_hours(attendee)
            working_hours_map[attendee] = {
                'timezone': attendee_tz_name,
                'working_hours': working_hours,
                'timezone_obj': attendee_tz
            }
        except Exception as e:
            working_hours_map[attendee] = {
                'timezone': 'UTC',
                'working_hours': None,
                'timezone_obj': ZoneInfo("UTC"),
                'error': str(e)
            }
```

### Working Hours Data Structure

Working hours are retrieved from Microsoft Graph API endpoint `GET /users/{id}/mailboxSettings` and include:

- **daysOfWeek**: Array of working days (e.g., `['monday', 'tuesday', 'wednesday', 'thursday', 'friday']`)
- **startTime**: Start time in attendee's timezone (e.g., `"09:00:00"`)
- **endTime**: End time in attendee's timezone (e.g., `"17:00:00"`)
- **timeZone**: Timezone information for the working hours

### Working Hours Calculation

The system calculates working hours for the specified date:

```python
# Parse working hours
working_hours = working_hours_map[attendee]['working_hours']
if working_hours:
    days_of_week = working_hours.get('daysOfWeek', [])
    day_name = date_obj.strftime("%A").lower()
    
    if day_name in days_of_week:
        start_time = working_hours.get('startTime', '09:00:00')
        end_time = working_hours.get('endTime', '17:00:00')
        
        working_start_dt = datetime.combine(date_obj, datetime.strptime(start_time, "%H:%M:%S").time(), tzinfo=attendee_tz)
        working_end_dt = datetime.combine(date_obj, datetime.strptime(end_time, "%H:%M:%S").time(), tzinfo=attendee_tz)
```

## Availability by Time Slot Display

### How It Works

The system displays availability in 30-minute time slots throughout the day, showing the status of each slot for each attendee.

### Availability Status Codes

The Microsoft Graph API returns an availability view string where each character represents a 30-minute time slot:

- **0**: Free
- **1**: Tentative
- **2**: Busy
- **3**: Out of office
- **4**: Working elsewhere
- **?**: Unknown

### Time Slot Generation

The system generates time slots starting from UTC midnight and converts them to both attendee and user timezones:

```python
# Calculate UTC midnight as reference point
utc_midnight = datetime.combine(today, datetime.strptime("00:00:00", "%H:%M:%S").time(), tzinfo=ZoneInfo("UTC"))
utc_midnight_attendee = utc_midnight.astimezone(attendee_tz)

# Calculate start slot index based on working hours
interval_minutes = availability_view_interval or 30
minutes_from_utc_midnight = int((working_start_dt - utc_midnight_attendee).total_seconds() / 60)
start_slot_index = minutes_from_utc_midnight // interval_minutes

if start_slot_index < 0:
    start_slot_index = 0

# Generate time slots
total_slots = len(availability_view)

for i in range(start_slot_index, total_slots):
    status_code = availability_view[i] if i < total_slots else '?'
    
    # Calculate slot boundaries in UTC
    slot_start_utc = utc_midnight + timedelta(minutes=i * interval_minutes)
    slot_end_utc = utc_midnight + timedelta(minutes=(i + 1) * interval_minutes)
    
    # Convert to attendee timezone
    slot_start_attendee = slot_start_utc.astimezone(attendee_tz)
    slot_end_attendee = slot_end_utc.astimezone(attendee_tz)
    
    # Stop if slot is beyond working hours
    if slot_start_attendee >= working_end_dt:
        break
    
    # Convert to user timezone for display
    slot_start_user = slot_start_attendee.astimezone(user_tz)
    slot_end_user = slot_end_attendee.astimezone(user_tz)
    
    # Format time strings
    slot_start_attendee_str = slot_start_attendee.strftime("%H:%M")
    slot_end_attendee_str = slot_end_attendee.strftime("%H:%M")
    slot_start_user_str = slot_start_user.strftime("%H:%M")
    slot_end_user_str = slot_end_user.strftime("%H:%M")

    # Map status code to text
    status_map = {
        '0': 'Free',
        '1': 'Tentative',
        '2': 'Busy',
        '3': 'Out of office',
        '4': 'Working elsewhere',
        '?': 'Unknown'
    }
    status_text = status_map.get(status_code, 'Unknown')

    formatted_results.append(f"  {slot_start_attendee_str}-{slot_end_attendee_str} ({attendee_timezone}) / {slot_start_user_str}-{slot_end_user_str} ({timezone_str}): {status_code} ({status_text})")
```

### Example Output

```
Availability for 2026-01-05:

Attendee: ajay.kulkarni@in.ibm.com
Timezone: Asia/Kolkata
Working Hours: 09:00-18:00
  09:00-09:30 (Asia/Kolkata) / 11:30-12:00 (Asia/Shanghai): 0 (Free)
  09:30-10:00 (Asia/Kolkata) / 12:00-12:30 (Asia/Shanghai): 0 (Free)
  10:00-10:30 (Asia/Kolkata) / 12:30-13:00 (Asia/Shanghai): 2 (Busy)
  10:30-11:00 (Asia/Kolkata) / 13:00-13:30 (Asia/Shanghai): 2 (Busy)
  ...
```

## Timezone Conversion

### How It Works

The system handles timezone conversion at multiple levels to ensure accurate availability display across different timezones.

### Timezone Detection

The system uses a three-tier fallback mechanism to determine the user's timezone:

1. **Microsoft Graph Mailbox Settings** (Primary)
   - Retrieves the user's mailbox timezone from Microsoft Graph API
   - Endpoint: `GET /me/mailboxSettings`
   - Field: `timeZone`

2. **Environment Variable** (Secondary)
   - Falls back to `USER_TIMEZONE` if Graph API fails
   - Set in `.env` file: `USER_TIMEZONE=Asia/Shanghai`

3. **UTC** (Default)
   - Used if both methods fail

### Windows to IANA Timezone Conversion

Microsoft Graph API returns Windows timezone names (e.g., `"India Standard Time"`), which need to be converted to IANA timezone names (e.g., `"Asia/Kolkata"`). This conversion is handled by the `DateHandler` class in [date_handler.py](file:///c:/Project/microsoft_graph_mcp_server/microsoft_graph_mcp_server/date_handler.py):

```python
@staticmethod
def convert_to_iana_timezone(windows_tz: str) -> str:
    """Convert Windows timezone name to IANA timezone name."""
    if not windows_tz:
        return "UTC"

    iana_tz = DateHandler.WINDOWS_TO_IANA_TIMEZONES.get(windows_tz)
    if iana_tz:
        return iana_tz

    try:
        ZoneInfo(windows_tz)
        return windows_tz
    except Exception:
        return "UTC"
```

### UTC to User Timezone Conversion

The system converts UTC datetime strings to the user's local timezone for display:

```python
@staticmethod
def convert_utc_to_user_timezone(
    utc_datetime: str,
    timezone_str: str = "UTC",
    format_str: str = "%a %m/%d/%Y %I:%M %p",
) -> str:
    """Convert UTC datetime string to user timezone."""
    if not utc_datetime:
        return ""

    try:
        dt = datetime.fromisoformat(utc_datetime.replace("Z", "+00:00"))
        user_tz = DateHandler.get_user_timezone_object(timezone_str)
        dt_converted = dt.astimezone(user_tz)
        return dt_converted.strftime(format_str)
    except Exception as e:
        return utc_datetime
```

### Timezone Object Retrieval

The system retrieves timezone objects for both attendees and users:

```python
@staticmethod
def get_user_timezone_object(timezone_str: str = "UTC") -> ZoneInfo:
    """Get user timezone object with fallback to UTC."""
    try:
        return ZoneInfo(timezone_str)
    except Exception:
        return ZoneInfo("UTC")
```

## Configuration

### Setting User Timezone via Environment Variable

Create or update the `.env` file:

```env
USER_TIMEZONE=Asia/Shanghai
```

### Supported Timezone Formats

The system uses Python's `zoneinfo` module and supports all IANA timezone names:

- `Asia/Shanghai` (CST, UTC+8)
- `Asia/Kolkata` (IST, UTC+5:30)
- `America/New_York` (EST, UTC-5)
- `Europe/London` (GMT, UTC+0)
- `Asia/Tokyo` (JST, UTC+9)
- And many more...

## API Reference

### check_attendee_availability

Check availability of attendees for a given date.

**Parameters:**
- `attendees` (List[str]): List of attendee email addresses
- `date` (str): Date in ISO format (e.g., "2026-01-05")
- `availability_view_interval` (Optional[int]): Time interval in minutes for availability slots (default: 30)
- `time_zone` (Optional[str]): User timezone for display (default: auto-detected)

**Returns:**
- Dictionary containing availability information for each attendee

**Example:**
```python
result = await check_attendee_availability(
    attendees=["ajay.kulkarni@in.ibm.com"],
    date="2026-01-05",
    availability_view_interval=30,
    time_zone="Asia/Shanghai"
)
```

### get_working_hours

Get working hours for a specific user.

**Parameters:**
- `user_id` (str): User email address or ID

**Returns:**
- Dictionary containing working hours information

**Example:**
```python
working_hours = await calendar_client.get_working_hours("ajay.kulkarni@in.ibm.com")
# Returns: {
#     'daysOfWeek': ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
#     'startTime': '09:00:00',
#     'endTime': '18:00:00',
#     'timeZone': {'@odata.type': '#microsoft.graph.customTimeZone', ...}
# }
```

## Error Handling

### Working Hours Not Available

If working hours cannot be retrieved for an attendee:

- The system falls back to default working hours (09:00-17:00)
- An error message is included in the response
- Availability checking continues for other attendees

### Invalid Timezone

If an invalid timezone string is provided:

- The system falls back to UTC
- A warning is logged for debugging purposes
- Availability checking continues normally

### API Access Denied

If the Microsoft Graph API returns a 403 Forbidden error:

- The system gracefully falls back to default settings
- No error is raised to the user
- Availability checking continues with fallback values

## Examples

### Example 1: Check Availability for Single Attendee

```python
result = await check_attendee_availability(
    attendees=["ajay.kulkarni@in.ibm.com"],
    date="2026-01-05"
)
```

Output:
```
Availability for 2026-01-05:

Attendee: ajay.kulkarni@in.ibm.com
Timezone: Asia/Kolkata
Working Hours: 09:00-18:00
  09:00-09:30 (Asia/Kolkata) / 11:30-12:00 (Asia/Shanghai): 0 (Free)
  09:30-10:00 (Asia/Kolkata) / 12:00-12:30 (Asia/Shanghai): 0 (Free)
  10:00-10:30 (Asia/Kolkata) / 12:30-13:00 (Asia/Shanghai): 2 (Busy)
  ...
```

### Example 2: Check Availability for Multiple Attendees

```python
result = await check_attendee_availability(
    attendees=["ajay.kulkarni@in.ibm.com", "john.doe@example.com"],
    date="2026-01-05",
    availability_view_interval=15
)
```

### Example 3: Custom Timezone Display

```python
result = await check_attendee_availability(
    attendees=["ajay.kulkarni@in.ibm.com"],
    date="2026-01-05",
    time_zone="America/New_York"
)
```

## Troubleshooting

### All Time Slots Showing as Free

1. Verify the attendee has calendar events on the specified date
2. Check if you have permission to access the attendee's calendar
3. Verify the date format is correct (YYYY-MM-DD)
4. Check MCP server logs for API errors

### Incorrect Timezone Conversion

1. Verify your system timezone matches your location
2. Check the `USER_TIMEZONE` environment variable
3. Test with a known timezone (e.g., `UTC`, `America/New_York`)
4. Verify Windows to IANA timezone conversion is working correctly

### Working Hours Not Displaying

1. Check if you have permission to access mailbox settings
2. Verify the attendee's mailbox settings are configured
3. Check MCP server logs for API errors
4. Ensure the attendee's timezone is correctly configured

## Related Documentation

- [README.md](../README.md) - Main project documentation
- [CALENDAR.md](CALENDAR.md) - General calendar functionality documentation
- [TIMEZONE_HANDLING.md](TIMEZONE_HANDLING.md) - Timezone handling for email functions
- [TEST_README.md](TEST_README.md) - Testing guide
