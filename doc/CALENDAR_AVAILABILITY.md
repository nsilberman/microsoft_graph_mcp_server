# Calendar Availability Documentation

## Overview

The Microsoft Graph MCP Server provides comprehensive calendar availability checking functionality, including working hours retrieval, availability time slot display, and timezone-aware scheduling.

---

## API Reference

### check_attendee_availability

Check availability of attendees for a given date.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `attendees` | List[str] | Yes | List of mandatory attendee email addresses |
| `optional_attendees` | List[str] | No | List of optional attendee email addresses |
| `date` | str | Yes | Date in ISO format (e.g., "2026-01-05") |
| `availability_view_interval` | int | No | Time interval in minutes (5, 6, 10, 15, 30, 60). Default: 30 |
| `time_zone` | str | No | User timezone for display. Default: auto-detected |
| `top_slots` | int | No | Number of top time slots to display. Default: 5 |
| `meeting_duration` | int | No | Required meeting duration in minutes (15, 30, 45, 60, 90, 120, etc.). Default: 30 |

**Returns:**

```json
{
  "date": "2026-01-05",
  "interval_minutes": 30,
  "timezone": "Asia/Shanghai",
  "attendees": [
    {
      "email": "colleague1@example.com",
      "type": "Mandatory",
      "working_hours": {
        "start": "09:00",
        "end": "18:00",
        "timezone": "Asia/Kolkata"
      },
      "free_time_slots": [...],
      "scheduled_items": [...],
      "timezone": "Asia/Kolkata"
    }
  ],
  "summary": {
    "top_time_slots": [
      {
        "rank": 1,
        "start_time": "2026-01-05 11:30",
        "end_time": "2026-01-05 12:00",
        "timezone": "Asia/Shanghai",
        "free_attendees": 3,
        "total_attendees": 3,
        "percentage_free": 100.0,
        "unavailable_attendees": []
      }
    ],
    "total_attendees": 3
  }
}
```

### get_working_hours

Get working hours for a specific user.

**Parameters:**
- `user_id` (str): User email address or ID

**Returns:**
```json
{
  "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"],
  "startTime": "09:00:00",
  "endTime": "18:00:00",
  "timeZone": {"@odata.type": "#microsoft.graph.customTimeZone", ...}
}
```

---

## Availability Status Codes

| Code | Status |
|------|--------|
| `0` | Free |
| `1` | Tentative |
| `2` | Busy |
| `3` | Out of office |
| `4` | Working elsewhere |
| `?` | Unknown |

---

## Timezone Handling

### Detection Priority

1. **Microsoft Graph Mailbox Settings** (Primary) - Auto-detected from user's mailbox
2. **Environment Variable** (Secondary) - `USER_TIMEZONE` in `.env`
3. **UTC** (Default)

### Supported Timezones

All IANA timezone names are supported:
- `Asia/Shanghai` (CST, UTC+8)
- `America/New_York` (EST, UTC-5)
- `Europe/London` (GMT, UTC+0)
- `Asia/Tokyo` (JST, UTC+9)

### Windows to IANA Conversion

Windows timezone names (e.g., `"India Standard Time"`) are automatically converted to IANA format (e.g., `"Asia/Kolkata"`).

Supported Windows timezone mappings:
- `India Standard Time` → `Asia/Kolkata`
- `China Standard Time` → `Asia/Shanghai`
- `FLE Standard Time` → `Europe/Kiev` (Finland, Russia, Estonia, Latvia, Lithuania)
- `W. Europe Standard Time` → `Europe/Paris`
- `Romance Standard Time` → `Europe/Paris`
- `E. Europe Standard Time` → `Europe/Chisinau`
- `Central Europe Standard Time` → `Europe/Budapest`
- `Central European Standard Time` → `Europe/Warsaw`
- And more...

---

## Permissions

### Required Permissions

For accurate working hours retrieval from other users, the following Microsoft Graph permission is recommended:

| Permission | Type | Description |
|------------|------|-------------|
| `MailboxSettings.Read` | Delegated | Read user mailbox settings including working hours |

**Without this permission:**
- Working hours may be retrieved from the `getSchedule` API, which can sometimes return inaccurate `startTime` values
- The organizer's working hours will still be available via `/me/mailboxSettings`

**How to add permission:**
1. Go to Azure Portal → App Registrations → Your App
2. API Permissions → Add a permission → Microsoft Graph
3. Select **Delegated permissions** → MailboxSettings → `MailboxSettings.Read`
4. Grant admin consent (if required)

---

## Examples

### Single Attendee

```python
result = await check_attendee_availability(
    attendees=["colleague1@example.com"],
    date="2026-01-05"
)
```

### Multiple Attendees with Optional

```python
result = await check_attendee_availability(
    attendees=["colleague1@example.com"],
    optional_attendees=["colleague2@example.com"],
    date="2026-01-05",
    availability_view_interval=15,
    top_slots=3
)
```

### Custom Timezone Display

```python
result = await check_attendee_availability(
    attendees=["colleague1@example.com"],
    date="2026-01-05",
    time_zone="America/New_York",
    top_slots=5
)
```

---

## Troubleshooting

| Issue | Solutions |
|-------|-----------|
| All time slots showing as Free | Verify attendee has calendar events; check permissions; verify date format |
| Incorrect timezone conversion | Verify system timezone; check `USER_TIMEZONE` env var; test with known timezone |
| Working hours not displaying | Check mailbox settings permissions; verify attendee's timezone configuration |
| Organizer missing from attendees | May be due to API rate limiting (429 error); wait a moment and retry; user email is cached after first successful retrieval |
| "No time zone found" error | Check if timezone is in supported mappings; add missing timezone to `_convert_microsoft_timezone_to_iana()` |

---

## Related Documentation

- [CALENDAR.md](CALENDAR.md) - General calendar functionality
- [TIMEZONE_HANDLING.md](TIMEZONE_HANDLING.md) - Timezone handling details
- [TEST_README.md](TEST_README.md) - Testing guide
