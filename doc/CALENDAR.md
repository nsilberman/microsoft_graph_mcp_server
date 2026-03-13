# Calendar Functions Documentation

This document describes the calendar-related functions available in the Microsoft Graph MCP Server.

## Table of Contents

1. [Search Events](#search-events)
2. [Browse Events](#browse-events)
3. [Get Event Detail](#get-event-detail)
4. [Manage Event](#manage-event)
5. [Check Attendee Availability](#check-attendee-availability)

---

## Search Events

### Description
Search or list calendar events by keywords with configurable date range filtering. If no query is provided, lists events within the specified time range. Supports time range presets (today, tomorrow, this_week, next_week, this_month, next_month) or custom date ranges.

The `search_type` parameter allows you to specify what to search for:
- "subject": Search by event subject (default)
- "organizer": Search by organizer name (uses fuzzy matching)

### Parameters
- `query` (optional, string): Search query
  - For subject: subject text
  - For organizer: organizer name
  - If not provided, lists events within the time range
- `search_type` (optional, string): Type of search to perform
  - Values: "subject", "organizer"
  - Default: "subject"
  - If not provided, defaults to subject search
- `start_date` (optional, string): Start date in your local timezone
  - Format: "2024-01-01" or "2024-01-01T14:30"
  - Optional if time_range is provided
- `end_date` (optional, string): End date in your local timezone
  - Format: "2024-01-01" or "2024-01-01T23:59"
  - Optional if time_range is provided
- `time_range` (optional, string): Time range type in your local timezone
  - Values: "today", "tomorrow", "this_week", "next_week", "this_month", "next_month"
  - If provided, overrides start_date and end_date
  - Returns a user-friendly display string in the response (e.g., "Today", "This Week")

### Returns
```json
{
  "query": "meeting",
  "date_range": "This Week",
  "count": 15,
  "timezone": "Asia/Shanghai",
  "today": "Fri 01/01/2026 12:00 PM",
  "start_date": "Mon 12/29/2025 12:00 AM",
  "end_date": "Mon 01/05/2026 12:00 AM",
  "hint": "Found 15 events. Use browse_events to view the results."
}
```

### Return Fields
- `query`: Search query used (for searches)
- `date_range`: User-friendly display string (e.g., "Today", "This Week", "This Month")
- `count`: Number of events found
- `timezone`: User's timezone for reference
- `today`: Today's date in user's timezone
- `start_date`: Start date in user's timezone (if time_range or start_date provided)
- `end_date`: End date in user's timezone (if time_range or end_date provided)
- `hint`: Instructions for viewing results

### Example Usage

#### List Events by Time Range:
```python
# List today's events
result = await search_events(time_range="today")

# List this week's events
result = await search_events(time_range="this_week")

# List next month's events
result = await search_events(time_range="next_month")
```

#### List Events by Custom Date Range:
```python
# List events between specific dates
result = await search_events(
    start_date="2024-01-01",
    end_date="2024-01-31"
)

# List events with specific start/end times
result = await search_events(
    start_date="2024-01-01T09:00",
    end_date="2024-01-01T17:00"
)
```

#### Search Events by Keywords:
```python
# Search events by subject (default)
result = await search_events(query="meeting", time_range="this_week")

# Search events by subject explicitly
result = await search_events(query="meeting", search_type="subject", time_range="this_week")

# Search events by organizer (fuzzy matching)
result = await search_events(query="John", search_type="organizer", start_date="2024-01-01", end_date="2024-01-31")
```

### Notes
- Automatically clears the cache before performing search or loading events
- Loads search results or time range events into cache for browsing
- Use `browse_events` to view the results
- Time ranges are calculated in the user's local timezone and converted to UTC for API calls
- Recurring events are automatically expanded into individual occurrences within the specified date range
- Maximum MAX_EVENT_SEARCH_LIMIT events per search
- Subject search uses server-side filtering (contains) for exact matches
- Organizer search uses hybrid filtering: server-side for exact email matches, client-side for fuzzy name matches

---

## Browse Events

### Description
Browse events in the cache with pagination. Returns summary information with number column indicating position in cache. Automatically manages browsing state with disk cache for persistence.

### Parameters
- `page_number` (required, integer): Page number to view (starts at 1)
- `mode` (optional, string): Browsing mode
  - Values: "user" for human browsing (smaller page size, default 5)
  - Values: "llm" for LLM browsing (larger page size, default 20)
  - Default: "user"

### Configuration
The number of events per page is controlled by environment variables:

```env
# .env file
PAGE_SIZE=5           # For user mode
LLM_PAGE_SIZE=20      # For LLM mode
```

Default page sizes:
- User mode: 5 events per page
- LLM mode: 20 events per page

### Returns
```json
{
  "events": [...],
  "count": 5,
  "total_count": 15,
  "current_page": 1,
  "total_pages": 3,
  "page_size": 5,
  "mode": "user",
  "timezone": "Asia/Shanghai"
}
```

### Event Structure
Each event in the cache contains:
```json
{
  "number": 1,
  "id": "AAMkADc4MDRiMTA2...",
  "subject": "Team Meeting",
  "start": "Mon 01/06/2026 02:00 PM",
  "end": "Mon 01/06/2026 03:00 PM",
  "location": "Conference Room A",
  "organizer": {
    "name": "John Smith",
    "email": "john@example.com"
  },
  "attendees": 5,
  "isAllDay": false,
  "showAs": "busy",
  "importance": "normal",
  "type": "singleInstance",
  "recurrence": false,
  "responseStatus": {
    "response": "accepted",
    "time": "2026-01-01T10:00:00Z"
  },
  "sensitivity": "normal"
}
```

### Return Fields
- `events`: Array of event summaries
- `count`: Number of events on current page
- `total_count`: Total events in cache
- `current_page`: Current page number
- `total_pages`: Total pages available
- `page_size`: Number of events per page
- `mode`: Browsing mode (user or llm)
- `timezone`: User's timezone for reference

### Example Usage
```python
# View first page (user mode, default 5 events per page)
result = await browse_events(page_number=1)

# View second page (user mode)
result = await browse_events(page_number=2)

# View first page in LLM mode (20 events per page)
result = await browse_events(page_number=1, mode="llm")
```

### Notes
- Requires events to be loaded into cache first (via search_events)
- Automatically manages browsing state with disk cache for persistence
- Events are sorted by start time (latest first)
- Use the `number` field to reference specific events
- Cache persists for up to 8 hours, maximum age 24 hours

---

## Get Event Detail

### Description
Get detailed information for a specific calendar event by its cache number. Use the cache number from browse_events (e.g., 1, 2, 3) to retrieve complete event details.

### Parameters
- `cache_number` (required, string): Cache number from browse_events or search_events results (e.g., '1', '2', '3')

### Returns
```json
{
  "subject": "Team Meeting",
  "start": {
    "time": "Mon 01/06/2026 02:00 PM",
    "timeZone": "Asia/Shanghai"
  },
  "end": {
    "time": "Mon 01/06/2026 03:00 PM",
    "timeZone": "Asia/Shanghai"
  },
  "location": {
    "displayName": "Conference Room A",
    "locationUri": "",
    "locationType": "default"
  },
  "organizer": {
    "emailAddress": {
      "name": "John Smith",
      "address": "john@example.com"
    }
  },
  "attendees": [
    {
      "emailAddress": {
        "name": "Jane Doe",
        "address": "jane@example.com"
      },
      "status": {
        "response": "accepted",
        "time": "2026-01-01T10:00:00Z"
      },
      "type": "required"
    }
  ],
  "isAllDay": false,
  "showAs": "busy",
  "importance": "normal",
  "sensitivity": "normal",
  "body": {
    "contentType": "HTML",
    "content": "<html>...</html>"
  },
  "responseStatus": {
    "response": "organizer",
    "time": "2026-01-01T09:00:00Z"
  },
  "type": "singleInstance",
  "createdDateTime": "2026-01-01T08:00:00Z",
  "lastModifiedDateTime": "2026-01-01T08:30:00Z"
}
```

### Return Fields
- `subject`: Event subject
- `start`: Event start time with user-friendly format
  - `time`: Formatted time string in user's local timezone (e.g., "Mon 01/06/2026 02:00 PM")
  - `timeZone`: User's timezone (e.g., "Asia/Shanghai")
- `end`: Event end time with user-friendly format
  - `time`: Formatted time string in user's local timezone (e.g., "Mon 01/06/2026 03:00 PM")
  - `timeZone`: User's timezone (e.g., "Asia/Shanghai")
- `location`: Event location (displayName, locationUri, locationType)
- `organizer`: Event organizer (name, email)
- `attendees`: Array of attendees with response status
- `isAllDay`: Whether the event is all-day
- `showAs`: Show as status (busy, free, tentative, oof, workingElsewhere, unknown)
- `importance`: Event importance (low, normal, high)
- `sensitivity`: Event sensitivity (normal, personal, private, confidential)
- `body`: Event body content (contentType, content)
- `responseStatus`: User's response status (response, time)
- `type`: Event type (singleInstance, occurrence, exception, seriesMaster)
- `createdDateTime`: When the event was created
- `lastModifiedDateTime`: When the event was last modified

### Example Usage
```python
# Get event details by cache number
result = await get_event_detail(cache_number="1")

# Get details for another event
result = await get_event_detail(cache_number="5")
```

### Notes
- Requires valid event number from cache
- Returns complete event details including all fields from Microsoft Graph API
- Use `browse_events` to get the event number
- Event types:
  - `singleInstance`: Single occurrence event
  - `occurrence`: Occurrence of a recurring event
  - `exception`: Modified occurrence of a recurring event
  - `seriesMaster`: Master event for a recurring series

---

## Respond to Event

### Description
Respond to calendar events organized by others with multiple actions: accept, decline, tentatively accept, propose new time, delete cancelled events, and email attendees. This tool is for responding to events that you are invited to, not events you organized yourself.

### Smart Handling of Events Without Response Requests

When the organizer creates an event with `responseRequested: false` (meaning they don't require attendees to formally accept/decline), the system automatically handles these cases differently:

- **Accept**: Updates the event to `showAs: "busy"` and enables reminders
- **Tentatively Accept**: Updates the event to `showAs: "tentative"` and enables reminders
- **Decline**: Deletes the event from your calendar (since you explicitly declined)
- **Propose New Time**: Deletes the event and instructs you to contact the organizer directly (since proposals require responses)

This behavior matches how Outlook handles such events - the time is marked on your calendar as busy/tentative without sending a formal response to the organizer.

### Parameters

#### Common Parameters (All Actions)
- `action` (required, string): Action to perform
  - Values: "accept", "decline", "tentatively_accept", "propose_new_time", "delete"
- `cache_number` (required, string): Cache number from browse_events or search_events results (e.g., '1', '2', '3')

#### Accept Action Parameters
- `comment` (optional, string): Message to include in response
- `send_response` (optional, boolean): Send response to organizer (default: true)
- `series` (optional, boolean): Accept the entire recurring series (default: false)
  - If true, accepts all occurrences in the recurring series
  - Requires the series master event (event with type "seriesMaster")
  - If false, accepts only the single event occurrence

#### Decline Action Parameters
- `comment` (optional, string): Message to include in response
- `send_response` (optional, boolean): Send response to organizer (default: true)
- `series` (optional, boolean): Decline the entire recurring series (default: false)
  - If true, declines all occurrences in the recurring series
  - Requires the series master event (event with type "seriesMaster")
  - If false, declines only the single event occurrence

#### Tentatively Accept Action Parameters
- `comment` (optional, string): Message to include in response
- `send_response` (optional, boolean): Send response to organizer (default: true)
- `series` (optional, boolean): Tentatively accept the entire recurring series (default: false)
  - If true, tentatively accepts all occurrences in the recurring series
  - Requires the series master event (event with type "seriesMaster")
  - If false, tentatively accepts only the single event occurrence

#### Propose New Time Action Parameters
- `propose_new_time` (required, object): Proposed new time
  - `dateTime`: Proposed date and time in ISO format
  - `timeZone`: Time zone (e.g., 'UTC', 'America/New_York')
- `comment` (optional, string): Message to include with proposal
- `send_response` (optional, boolean): Send response to organizer (default: true)

#### Delete Action Parameters
- `comment` (optional, string): Message to include (optional)

### Returns

#### Accept/Decline/Tentatively Accept Actions Returns
```json
{
  "type": "text",
  "text": "Event accepted successfully."
}
```

#### Propose New Time Action Returns
```json
{
  "type": "text",
  "text": "New time proposed to organizer."
}
```

#### Delete Action Returns
```json
{
  "type": "text",
  "text": "Cancelled event deleted from calendar."
}
```

### Example Usage

#### Accept Event:
```python
result = await manage_event_as_attendee(
    action="accept",
    cache_number="1",
    comment="Looking forward to it"
)
```

#### Accept Entire Recurring Series:
```python
result = await manage_event_as_attendee(
    action="accept",
    cache_number="1",
    series=True,
    comment="I'll attend all meetings"
)
```

#### Decline Event:
```python
result = await manage_event_as_attendee(
    action="decline",
    cache_number="2",
    comment="I have a conflict at that time"
)
```

#### Decline Entire Recurring Series:
```python
result = await manage_event_as_attendee(
    action="decline",
    cache_number="3",
    series=True,
    comment="I won't be able to attend any of these"
)
```

#### Tentatively Accept Event:
```python
result = await manage_event_as_attendee(
    action="tentatively_accept",
    cache_number="4",
    comment="I'll try to make it"
)
```

#### Propose New Time:
```python
result = await manage_event_as_attendee(
    action="propose_new_time",
    cache_number="5",
    propose_new_time={
        "dateTime": "2026-01-06T15:00:00",
        "timeZone": "UTC"
    },
    comment="Can we meet an hour later?"
)
```

#### Delete Cancelled Event:
```python
result = await manage_event_as_attendee(
    action="delete_cancelled",
    cache_number="6",
    comment="Removing cancelled event"
)
```

#### Email Attendees (as Attendee):
```python
result = await manage_event_as_attendee(
    action="email_attendees",
    cache_number="7",
    email_subject="Question about the meeting",
    body="Hi everyone, I have a question about the agenda..."
)
```

> **Note**: As an attendee, `email_attendees` sends to the organizer + all other attendees (excluding yourself). Required attendees are in TO, optional attendees are in CC.

### Notes

#### Accept/Decline/Tentatively Accept Actions
- Sends response to organizer
- Updates user's response status for the event
- Comment is included in the response
- Set send_response to false to respond without notifying organizer
- **Series Parameter Support**:
  - When `series=true`, applies the action to all occurrences in the recurring series
  - Requires the series master event (event with type "seriesMaster")
  - When `series=false` (default), applies the action only to the single event occurrence
  - Use `search_events` or `browse_events` to identify series master events (type: "seriesMaster")
  - Individual occurrences have type "occurrence" or "exception"

#### Propose New Time Action
- Automatically declines the event
- Sends proposed new time to organizer
- Organizer can accept or decline the proposal
- Comment is included with the proposal
- This is the preferred way to propose alternative times

#### Delete Action
- Used to delete cancelled events organized by others from your calendar
- Only works for events where you are not the organizer
- Does not send notifications to the organizer
- Useful for cleaning up cancelled events from your calendar

#### General Notes
- Cache number is from browse_events or search_events
- Event must be in cache (use search_events first)
- For user-organized events, use the manage_event_as_organizer tool instead

---

## Manage Event As Organizer

### Description
Manage calendar events where you are the organizer (events you created). Actions: create, update, cancel, forward, email_attendees. This tool is for events that you organized yourself, not events you were invited to.

### Parameters

#### Common Parameters (All Actions)
- `action` (required, string): Action to perform
  - Values: "create", "update", "cancel", "forward", "email_attendees"

#### Create Action Parameters
- `subject` (required, string): Event subject
- `start` (required, string): Event start date and time in your local timezone
  - Format: "2026-01-06T14:30" or "2026-01-06 14:30"
  - The system will automatically convert to UTC using the timezone parameter or your timezone settings from your Microsoft 365 profile or .env configuration
- `end` (required, string): Event end date and time in your local timezone
  - Format: "2026-01-06T15:30" or "2026-01-06 15:30"
  - The system will automatically convert to UTC using the timezone parameter or your timezone settings from your Microsoft 365 profile or .env configuration
- `timezone` (optional, string): Timezone for the event in IANA format
  - Examples: 'Asia/Singapore', 'America/New_York', 'Europe/London', 'UTC'
  - If not provided, will use your timezone settings from your Microsoft 365 profile or .env configuration
- `attendees` (optional, array): List of required attendee email addresses
- `optional_attendees` (optional, array): List of optional attendee email addresses
- `body` (optional, string): Event body content
- `body_content_type` (optional, string): Body content type
  - Values: "Text" or "HTML"
  - Default: "HTML"
- `location` (optional, string): Event location
- `isOnlineMeeting` (optional, boolean): Whether to create the event as an online meeting (default: false)
- `onlineMeetingProvider` (optional, string): Online meeting provider
  - Values: "teamsForBusiness", "skypeForBusiness", "skypeForConsumer", "unknown"
  - Default: "teamsForBusiness" when isOnlineMeeting is true
- `recurrence` (optional, object): Recurrence pattern for the event
  - `pattern`: Recurrence pattern (how often the event repeats)
    - `type`: "daily", "weekly", "absoluteMonthly", "relativeMonthly", "absoluteYearly", "relativeYearly"
    - `interval`: Number of units between occurrences (e.g., 2 for every 2 weeks)
    - `dayOfMonth`: Day of month (1-31) for absoluteMonthly/absoluteYearly
    - `daysOfWeek`: Array of days for weekly/relativeMonthly/relativeYearly
      - Values: "sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"
    - `index`: "first", "second", "third", "fourth", "last" for relativeMonthly/relativeYearly
    - `month`: Month number (1-12) for absoluteYearly/relativeYearly
  - `range`: Recurrence range (how long the recurrence lasts)
    - `type`: "endDate", "noEnd", "numbered"
    - `startDate`: Start date in ISO format (e.g., "2024-01-01")
    - `endDate`: End date in ISO format for type "endDate"
    - `numberOfOccurrences`: Number of occurrences for type "numbered"

#### Update Action Parameters
- `cache_number` (required, string): Cache number from browse_events or search_events results (e.g., '1', '2', '3')
- `subject` (optional, string): Updated event subject
- `start` (optional, string): Updated event start date and time in your local timezone
  - The system will automatically convert to UTC using a timezone parameter or your timezone settings from your Microsoft 365 profile or .env configuration
- `end` (optional, string): Updated event end date and time in your local timezone
  - The system will automatically convert to UTC using a timezone parameter or your timezone settings from your Microsoft 365 profile or .env configuration
- `timezone` (optional, string): Timezone for event in IANA format
  - Examples: 'Asia/Singapore', 'America/New_York', 'Europe/London', 'UTC'
  - If not provided, will use your timezone settings from your Microsoft 365 profile or .env configuration
- `attendees` (optional, array): Updated list of required attendee email addresses
- `optional_attendees` (optional, array): Updated list of optional attendee email addresses
- `body` (optional, string): Updated event body content
- `body_content_type` (optional, string): Updated body content type
- `location` (optional, string): Updated event location
- `isOnlineMeeting` (optional, boolean): Whether to update as an online meeting
- `onlineMeetingProvider` (optional, string): Online meeting provider

#### Cancel Action Parameters
- `cache_number` (required, string): Cache number from browse_events or search_events results (e.g., '1', '2', '3')
- `comment` (optional, string): Message to include in cancellation notification

#### Forward Action Parameters
- `cache_number` (required, string): Cache number from browse_events or search_events results (e.g., '1', '2', '3')
- `attendees` (required, array): List of attendee email addresses to add as optional attendees
- `comment` (optional, string): Message to include in forward

#### Email Attendees Action Parameters
- `cache_number` (required, string): Cache number from browse_events or search_events results (e.g., '1', '2', '3')
- `email_subject` (optional, string): Email subject (default: "Re: Event")
- `body` (optional, string): Email body content (default: event body content)
- `to` (optional, array): List of TO recipient email addresses (default: required attendees)
- `cc` (optional, array): List of CC recipient email addresses (default: optional attendees)

### Returns

#### Create Action Returns
```json
{
  "id": "event-id",
  "subject": "Team Meeting",
  "start": {
    "dateTime": "2026-01-06T14:00:00",
    "timeZone": "UTC"
  },
  "end": {
    "dateTime": "2026-01-06T15:00:00",
    "timeZone": "UTC"
  },
  ...
}
```

#### Other Actions Returns
```json
{
  "type": "text",
  "text": "Event updated successfully."
}
```

### Example Usage

#### Create Event:
```python
result = await manage_event_as_organizer(
    action="create",
    subject="Team Meeting",
    start="2026-01-06T14:30",
    end="2026-01-06T15:30"
)
```

#### Create Event with Specific Timezone:
```python
result = await manage_event_as_organizer(
    action="create",
    subject="Global Team Meeting",
    start="2026-01-06T14:30",
    end="2026-01-06T15:30",
    timezone="America/New_York"
)
```

#### Create Recurring Event:
```python
result = await manage_event_as_organizer(
    action="create",
    subject="Weekly Team Meeting",
    start="2026-01-06T14:30",
    end="2026-01-06T15:30",
    recurrence={
        "pattern": {
            "type": "weekly",
            "interval": 1,
            "daysOfWeek": ["monday"]
        },
        "range": {
            "type": "noEnd",
            "startDate": "2026-01-06"
        }
    }
)
```

#### Create Recurring Event with Specific Timezone:
```python
result = await manage_event_as_organizer(
    action="create",
    subject="Weekly Global Meeting",
    start="2026-01-06T14:30",
    end="2026-01-06T15:30",
    timezone="Asia/Singapore",
    recurrence={
        "pattern": {
            "type": "weekly",
            "interval": 1,
            "daysOfWeek": ["monday"]
        },
        "range": {
            "type": "noEnd",
            "startDate": "2026-01-06"
        }
    }
)
```

#### Create Online Meeting:
```python
result = await manage_event_as_organizer(
    action="create",
    subject="Online Team Meeting",
    start="2026-01-06T14:30",
    end="2026-01-06T15:30",
    isOnlineMeeting=True,
    onlineMeetingProvider="teamsForBusiness"
)
```

#### Update Event:
```python
result = await manage_event_as_organizer(
    action="update",
    cache_number="1",
    subject="Updated Team Meeting",
    start="2026-01-06T15:30",
    end="2026-01-06T16:30"
)
```

#### Update Event with Specific Timezone:
```python
result = await manage_event_as_organizer(
    action="update",
    cache_number="1",
    subject="Updated Team Meeting",
    start="2026-01-06T15:30",
    end="2026-01-06T16:30",
    timezone="Europe/London"
)
```

#### Cancel Event:
```python
result = await manage_event_as_organizer(
    action="cancel",
    cache_number="2",
    comment="Meeting is no longer needed"
)
```

#### Forward Event:
```python
result = await manage_event_as_organizer(
    action="forward",
    cache_number="4",
    attendees=["new-attendee@example.com", "another@example.com"],
    comment="Please join this meeting"
)
```

#### Email Event Attendees:
```python
result = await manage_event_as_organizer(
    action="email_attendees",
    cache_number="5",
    email_subject="Re: Team Meeting",
    body="I'll be attending the meeting",
    to=["organizer@example.com"],
    cc=["optional-attendee@example.com"]
)
```

### Notes

#### Create Action
- Times should be in your local timezone (system automatically converts to UTC)
- Timezone is automatically detected from your Microsoft 365 profile or .env configuration, unless you provide the optional `timezone` parameter to override it
- Attendee types: required and optional
- Body content types: "Text" or "HTML"
- The event is created in your primary calendar
- Supports recurring events with flexible recurrence patterns
- Supports online meetings with Teams, Skype, or other providers

#### Update Action
- Only provided fields are updated
- Cache number is from browse_events or search_events
- Cannot change event type (single, recurring, etc.)
- Event must be in cache (use search_events first)

#### Cancel Action
- Sends cancellation notifications to all attendees
- Event is moved to deleted items
- Cannot be undone (must recreate if needed)

#### Forward Action
- Adds new attendees as optional attendees
- Original attendees remain unchanged
- Sends invitation to new attendees
- Comment is included in the invitation

#### Email Attendees Action
- Sends email to event attendees
- Uses event body as default email content if not provided
- TO recipients default to required attendees
- CC recipients default to optional attendees
- Event must be in cache (use search_events first)

#### General Notes
- Cache number is from browse_events or search_events
- For events organized by others, use the manage_event_as_attendee tool instead
- All date/time parameters are in your local timezone, unless you provide a optional `timezone` parameter to specify a different timezone

---

## Check Attendee Availability

### Description
Check availability of attendees for a given time range. Returns availability view string and schedule items for each attendee. Useful for finding optimal meeting times when creating or updating events.

### Parameters
- `attendees` (required, array): List of attendee email addresses to check availability for
  - Each attendee: email address string (e.g., "user@example.com")
- `start` (required, string): Start date and time in ISO format
  - Format: "2024-01-01T14:30:00"
  - Timezone is automatically detected from user's mailbox settings
- `end` (required, string): End date and time in ISO format
  - Format: "2024-01-01T15:30:00"
  - Timezone is automatically detected from user's mailbox settings
- `availability_view_interval` (optional, integer): Time interval in minutes for availability view
  - Default: 30 minutes
  - Controls the granularity of the availability view string
  - Valid values: 5, 6, 10, 15, 30, 60

### Returns
```json
{
  "type": "text",
  "text": "\nAttendee: user1@example.com\nAvailability View: 000011112222\nSchedule Items:\n  - busy: 2024-01-01T14:00:00 to 2024-01-01T14:30:00\n  - free: 2024-01-01T14:30:00 to 2024-01-01T15:00:00\n\nAttendee: user2@example.com\nAvailability View: 000000000000\nSchedule Items:\nNo scheduled items in this time range."
}
```

### Return Fields
- `type`: Content type (always "text")
- `text`: Formatted availability information including:
  - Attendee email addresses
  - Availability view string for each attendee
  - Schedule items showing busy/free periods

### Example Usage
```python
# Check availability for multiple attendees
result = await check_attendee_availability(
    attendees=["user1@example.com", "user2@example.com"],
    start="2024-01-01T14:00:00",
    end="2024-01-01T15:00:00"
)

# Check availability with custom interval
result = await check_attendee_availability(
    attendees=["user1@example.com"],
    start="2024-01-01T09:00:00",
    end="2024-01-01T17:00:00",
    availability_view_interval=15
)
```

### Notes
- Checks availability of multiple attendees for a given time range
- Returns availability view string and schedule items for each attendee
- Availability view string uses single-character codes for each time interval:
  - `0`: Free
  - `1`: Tentative
  - `2`: Busy
  - `3`: Out of office (OOF)
  - `4`: Working elsewhere
  - `?`: Unknown
- Schedule items provide detailed information about each attendee's scheduled events
- Timezone is automatically detected from user's mailbox settings
- The availability_view_interval parameter controls the granularity of the availability view
- Useful for finding optimal meeting times when creating or updating events

---

## Event Caching

### Overview
The calendar functionality uses a disk-based caching system to persist event data between sessions, improving performance and reducing API calls.

### Cache File Location
- Path: `~/.microsoft_graph_mcp_events.json` (in user's home directory)
- Format: JSON
- Version: 1.0

### Cache Expiration
- **Cache Expiry**: 8 hours
- **Maximum Age**: 24 hours
- Cache is automatically refreshed when expired or when search parameters change

### Cache Modes
The cache supports two browsing modes:
- **browse**: Default mode for browsing events by time range
- **search**: Mode for searching events by keywords

### Cache State Management
- Cache automatically invalidates when search parameters change (query, date range)
- Cache persists to disk for cross-session usage
- Cache is automatically refreshed when expired

### Cache Structure
```json
{
  "version": "1.0",
  "last_updated": "2026-01-01T12:00:00Z",
  "mode": "search",
  "browse_state": {
    "start_date": null,
    "end_date": null,
    "top": 20,
    "total_count": 0,
    "metadata": []
  },
  "search_state": {
    "query": "meeting",
    "start_date": "2026-01-01T00:00:00Z",
    "end_date": "2026-01-31T23:59:59Z",
    "top": 1000,
    "total_count": 15,
    "metadata": [...]
  },
  "metadata": {
    "user_id": null,
    "expires_at": "2026-01-01T20:00:00Z"
  }
}
```

---

## Time Range Presets

The system supports the following time range presets, all calculated in the user's local timezone:

### today
- Start: Today at 12:00 AM
- End: Tomorrow at 12:00 AM

### tomorrow
- Start: Tomorrow at 12:00 AM
- End: Day after tomorrow at 12:00 AM

### this_week
- Start: Monday of current week at 12:00 AM
- End: Monday of next week at 12:00 AM

### next_week
- Start: Monday of next week at 12:00 AM
- End: Monday of the week after at 12:00 AM

### this_month
- Start: First day of current month at 12:00 AM
- End: First day of next month at 12:00 AM

### next_month
- Start: First day of next month at 12:00 AM
- End: First day of the month after at 12:00 AM

---

## Recurring Events

### Overview
Recurring events are automatically expanded into individual occurrences within the specified date range using the Microsoft Graph API's calendarView endpoint.

### Event Types
- **singleInstance**: Single occurrence event (not recurring)
- **occurrence**: Occurrence of a recurring event
- **exception**: Modified occurrence of a recurring event
- **seriesMaster**: Master event for a recurring series

### Recurring Event Handling
- The system uses `/me/calendar/calendarView` endpoint which automatically expands recurring events
- Each occurrence appears as a separate event in the results
- Modified occurrences (exceptions) are correctly displayed
- The series master event is not included in the results (only occurrences)

### Recurrence Detection
The system automatically detects and marks events as recurring based on the following criteria:
- Events with a `recurrence` object are marked as `recurrence: true`
- Events of type "occurrence" or "seriesMaster" are marked as `recurrence: true`
- Events of type "exception" (modified occurrences) are also marked as `recurrence: true`
- For "occurrence" and "exception" types, the system fetches recurrence information from the series master event

### Example
For a weekly recurring meeting:
- Original: "Weekly Team Meeting" (seriesMaster)
- Expanded: "Weekly Team Meeting" on Jan 6, Jan 13, Jan 20, etc. (occurrence)
- Modified: "Weekly Team Meeting - Rescheduled" on Jan 20 (exception)
- All of these will have `recurrence: true` and include `recurrenceInfo`

---

## All-Day Events

### Overview
All-day events are treated specially by the system to ensure correct display and filtering behavior across different timezones.

### Time Display
- All-day events display "12:00 AM" for both start and end times
- No timezone conversion is applied to all-day event times
- This ensures consistent display regardless of user timezone

### Filtering Behavior
All-day events are filtered based on local midnight boundaries to ensure they appear in the correct date ranges:

- **Client-side filtering**: Events are filtered after being retrieved from the API
- **Local timezone normalization**: Event times are converted to the user's local timezone and normalized to midnight
- **Boundary exclusion**: Events that end exactly at the start of a date range are excluded from that range

### Example
Consider an all-day event "Great Cold" that spans Jan 20-21:

**UTC Storage**: 2026-01-20 00:00:00Z to 2026-01-21 00:00:00Z

**Asia/Shanghai Display** (UTC+8): 2026-01-20 08:00:00 to 2026-01-21 08:00:00

**Normalised for Filtering**: 2026-01-20 00:00:00 to 2026-01-21 00:00:00 (normalized to local midnight)

**Result**:
- Appears in "today" (Jan 20) because it overlaps with the range
- Does NOT appear in "tomorrow" (Jan 21) because it ends exactly at midnight

### Technical Implementation
1. Microsoft Graph stores all-day events as UTC midnight
2. When filtering, events are converted to user's local timezone
3. Event times are normalized to local midnight (hour=0, minute=0, second=0)
4. Events are included only if they have time strictly within the date range
5. Events ending exactly at the range start are excluded

## Timezone Handling

### Timezone Detection
The system automatically detects the user's timezone from Microsoft Graph mailbox settings. For create and update actions in manage_event_as_organizer, you can optionally provide a `timezone` parameter to override the automatic timezone detection and specify a custom timezone for the event.

### Timezone Conversion
- All date ranges are calculated in the user's local timezone
- Date ranges are converted to UTC for Microsoft Graph API calls
- Event times are displayed in the user's local timezone
- Original UTC times are preserved for accurate sorting

### Supported Timezones
The system supports all IANA timezone names:
- `Asia/Shanghai`
- `America/New_York`
- `Europe/London`
- `Asia/Tokyo`
- And many more...

### Timezone Format
- Windows timezone names are automatically converted to IANA format
- Example: "China Standard Time" → "Asia/Shanghai"

---

## Performance Considerations

### Maximum Events
- Maximum MAX_EVENT_SEARCH_LIMIT events per search
- Pagination allows efficient browsing of large result sets

### Caching Benefits
- Reduces API calls for repeated browsing
- Persists data across sessions
- Automatic cache refresh when expired

### Pagination
- User mode: 5 events per page (better for human browsing)
- LLM mode: 20 events per page (better for automated processing)
- Configurable via environment variables

---

## Troubleshooting

### Events Not Showing

**Problem**: Expected events are not appearing in search results.

**Solutions**:
1. Check if the date range is correct
2. Verify the timezone is set correctly
3. Try a broader date range
4. Check if events are in a different calendar

### Recurring Events Not Expanded

**Problem**: Recurring events appear only once or not at all.

**Solutions**:
1. Verify the date range includes the occurrences
2. Check if the recurrence pattern is valid
3. Ensure the calendarView endpoint is being used (automatic)

### Cache Issues

**Problem**: Events are not updating or showing stale data.

**Solutions**:
1. Wait for cache to expire (8 hours)
2. Delete cache file: `~/.microsoft_graph_mcp_events.json`
3. Perform a new search with different parameters to force refresh

### Timezone Issues

**Problem**: Event times appear incorrect.

**Solutions**:
1. Verify your timezone in Microsoft Graph mailbox settings
2. Check if `USER_TIMEZONE` environment variable is set
3. Ensure your system timezone is correct

---

## Related Documentation

- [README.md](../README.md) - Main project documentation
- [TIMEZONE_HANDLING.md](TIMEZONE_HANDLING.md) - Timezone conversion documentation
- [TEST_README.md](TEST_README.md) - Testing guide
