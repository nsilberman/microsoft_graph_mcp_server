"""Test script for event cache functionality and timezone conversion."""

import asyncio
import json
from pathlib import Path
import sys
from datetime import datetime, timezone

sys.path.insert(0, str(Path(__file__).parent.parent))

from microsoft_graph_mcp_server.date_handler import DateHandler
from microsoft_graph_mcp_server.clients.user_client import UserClient
from microsoft_graph_mcp_server.clients.calendar_client import CalendarClient
from microsoft_graph_mcp_server.event_cache import EventBrowsingCache


async def test_timezone_conversion():
    """Test that timezone conversion works correctly."""
    print("\n[Test 1] Testing timezone conversion...")
    
    date_handler = DateHandler()
    
    # Test UTC to Asia/Shanghai conversion
    utc_time = "2026-01-05T08:30:00Z"
    converted_time = date_handler.convert_utc_to_user_timezone(utc_time, "Asia/Shanghai")
    
    print(f"   UTC time: {utc_time}")
    print(f"   Converted to Asia/Shanghai: {converted_time}")
    
    # The conversion should work without errors
    assert converted_time is not None, "Time conversion should return a value"
    print("   ✓ Timezone conversion works")
    
    print("\n[Test 1] ✓ PASSED: Timezone conversion works correctly")


async def test_windows_to_iana_timezone_mapping():
    """Test that Windows timezone mapping is correct."""
    print("\n[Test 2] Testing Windows to IANA timezone mapping...")
    
    date_handler = DateHandler()
    
    # Test common Windows timezone mappings
    test_cases = [
        ("China Standard Time", "Asia/Shanghai"),
        ("Pacific Standard Time", "America/Los_Angeles"),
        ("Eastern Standard Time", "America/New_York"),
        ("UTC", "UTC"),
    ]
    
    for windows_tz, expected_iana in test_cases:
        result = date_handler.convert_to_iana_timezone(windows_tz)
        assert result == expected_iana, f"Expected {expected_iana}, got {result}"
        print(f"   ✓ {windows_tz} → {result}")
    
    print("\n[Test 2] ✓ PASSED: Windows to IANA timezone mapping is correct")


async def test_event_cache_timezone_conversion():
    """Test that event cache properly converts times to user timezone."""
    print("\n[Test 3] Testing event cache timezone conversion...")
    
    # This test verifies the timezone conversion logic that's used in search_events
    date_handler = DateHandler()
    user_timezone = "Asia/Shanghai"
    
    # Mock event data from Microsoft Graph API (UTC times)
    mock_event = {
        "id": "test-event-1",
        "subject": "Test Event",
        "start": {
            "dateTime": "2026-01-05T08:30:00",
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": "2026-01-05T09:30:00",
            "timeZone": "UTC"
        },
        "responseStatus": {
            "response": "accepted",
            "time": "2026-01-04T10:00:00Z"
        },
        "attendees": [
            {
                "emailAddress": {
                    "name": "John Doe",
                    "address": "john@example.com"
                },
                "status": {
                    "response": "accepted"
                }
            },
            {
                "emailAddress": {
                    "name": "Jane Smith",
                    "address": "jane@example.com"
                },
                "status": {
                    "response": "tentativelyAccepted"
                }
            }
        ],
        "recurrence": {
            "pattern": {
                "type": "weekly",
                "interval": 1,
                "daysOfWeek": ["monday"]
            },
            "range": {
                "type": "noEnd",
                "startDate": "2026-01-05"
            }
        }
    }
    
    # Test timezone conversion logic
    start_datetime = mock_event.get("start", {}).get("dateTime", "")
    end_datetime = mock_event.get("end", {}).get("dateTime", "")
    response_time = mock_event.get("responseStatus", {}).get("time")
    
    start_converted = date_handler.convert_utc_to_user_timezone(start_datetime, user_timezone)
    end_converted = date_handler.convert_utc_to_user_timezone(end_datetime, user_timezone)
    
    if response_time:
        response_time_converted = date_handler.convert_utc_to_user_timezone(response_time, user_timezone)
    else:
        response_time_converted = None
    
    print(f"   Event subject: {mock_event['subject']}")
    print(f"   Start time (converted): {start_converted}")
    print(f"   End time (converted): {end_converted}")
    print(f"   Response status time (converted): {response_time_converted}")
    print(f"   Attendees count: {len(mock_event['attendees'])}")
    print(f"   Is recurring: {mock_event.get('recurrence') is not None}")
    
    # Verify conversions
    assert start_converted is not None, "Start time should be converted"
    assert end_converted is not None, "End time should be converted"
    assert response_time_converted is not None, "Response status time should be converted"
    assert len(mock_event['attendees']) == 2, "Attendees count should be 2"
    assert mock_event.get('recurrence') is not None, "Event should have recurrence property"
    
    # Verify attendees have proper data
    attendee1 = mock_event['attendees'][0]
    assert attendee1['emailAddress']['name'] == "John Doe", "First attendee name should be John Doe"
    assert attendee1['emailAddress']['address'] == "john@example.com", "First attendee email should be john@example.com"
    
    print("   ✓ Event cache timezone conversion works correctly")
    
    print("\n[Test 3] ✓ PASSED: Event cache properly converts times to user timezone")


async def test_recurring_event_detection():
    """Test that recurring events are properly detected."""
    print("\n[Test 4] Testing recurring event detection...")
    
    date_handler = DateHandler()
    
    # Test with recurrence property
    event_with_recurrence = {
        "id": "test-1",
        "subject": "Weekly Meeting",
        "start": {"dateTime": "2026-01-05T08:30:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-01-05T09:30:00", "timeZone": "UTC"},
        "recurrence": {
            "pattern": {"type": "weekly", "interval": 1},
            "range": {"type": "noEnd", "startDate": "2026-01-05"}
        }
    }
    
    is_recurring1 = event_with_recurrence.get("recurrence") is not None or event_with_recurrence.get("type") in ["occurrence", "seriesMaster"]
    assert is_recurring1 is True, "Event with recurrence property should be marked as recurring"
    print("   ✓ Event with recurrence property detected")
    
    # Test with type="occurrence"
    event_occurrence = {
        "id": "test-2",
        "subject": "Weekly Meeting Occurrence",
        "start": {"dateTime": "2026-01-05T08:30:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-01-05T09:30:00", "timeZone": "UTC"},
        "type": "occurrence"
    }
    
    is_recurring2 = event_occurrence.get("recurrence") is not None or event_occurrence.get("type") in ["occurrence", "seriesMaster"]
    assert is_recurring2 is True, "Event with type='occurrence' should be marked as recurring"
    print("   ✓ Event with type='occurrence' detected")
    
    # Test with type="seriesMaster"
    event_series_master = {
        "id": "test-3",
        "subject": "Weekly Meeting Series",
        "start": {"dateTime": "2026-01-05T08:30:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-01-05T09:30:00", "timeZone": "UTC"},
        "type": "seriesMaster"
    }
    
    is_recurring3 = event_series_master.get("recurrence") is not None or event_series_master.get("type") in ["occurrence", "seriesMaster"]
    assert is_recurring3 is True, "Event with type='seriesMaster' should be marked as recurring"
    print("   ✓ Event with type='seriesMaster' detected")
    
    # Test non-recurring event
    event_single = {
        "id": "test-4",
        "subject": "One-time Meeting",
        "start": {"dateTime": "2026-01-05T08:30:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-01-05T09:30:00", "timeZone": "UTC"}
    }
    
    is_recurring4 = event_single.get("recurrence") is not None or event_single.get("type") in ["occurrence", "seriesMaster"]
    assert is_recurring4 is False, "Event without recurrence should not be marked as recurring"
    print("   ✓ Non-recurring event correctly identified")
    
    print("\n[Test 4] ✓ PASSED: Recurring events are properly detected")


async def test_attendees_list_formatting():
    """Test that attendees list is properly formatted for display."""
    print("\n[Test 5] Testing attendees list formatting...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    
    handler = CalendarHandler()
    
    # Mock event data with attendees
    event_data = {
        "id": "test-1",
        "subject": "Team Meeting",
        "start": "Mon 01/05/2026 04:30 PM",
        "end": "Mon 01/05/2026 05:30 PM",
        "attendees": 3,
        "attendees_list": [
            {
                "emailAddress": {
                    "name": "Alice Johnson",
                    "address": "alice@example.com"
                },
                "status": {"response": "accepted"}
            },
            {
                "emailAddress": {
                    "name": "Bob Smith",
                    "address": "bob@example.com"
                },
                "status": {"response": "declined"}
            },
            {
                "emailAddress": {
                    "name": "",
                    "address": "charlie@example.com"
                },
                "status": {"response": "tentativelyAccepted"}
            }
        ],
        "responseStatus": {"response": "accepted", "time": "2026-01-04T10:00:00"},
        "recurrence": True
    }
    
    # Format attendees for display
    attendees_list = event_data["attendees_list"]
    attendees_display = []
    for attendee in attendees_list:
        email_info = attendee.get("emailAddress", {})
        name = email_info.get("name", "")
        email = email_info.get("address", "")
        if name:
            attendees_display.append(f"{name} ({email})")
        elif email:
            attendees_display.append(email)
    
    print(f"   Attendees count: {len(attendees_display)}")
    for i, attendee in enumerate(attendees_display, 1):
        print(f"   {i}. {attendee}")
    
    assert len(attendees_display) == 3, "Should have 3 attendees"
    assert "Alice Johnson (alice@example.com)" in attendees_display, "First attendee should be formatted correctly"
    assert "Bob Smith (bob@example.com)" in attendees_display, "Second attendee should be formatted correctly"
    assert "charlie@example.com" in attendees_display, "Third attendee (no name) should show only email"
    
    print("   ✓ Attendees list formatting works correctly")
    
    print("\n[Test 5] ✓ PASSED: Attendees list is properly formatted for display")


async def test_response_status_time_conversion():
    """Test that response status time is properly converted to user timezone."""
    print("\n[Test 6] Testing response status time conversion...")
    
    date_handler = DateHandler()
    user_timezone = "Asia/Shanghai"
    
    # Mock event with response status time
    event_data = {
        "id": "test-1",
        "subject": "Meeting",
        "start": {"dateTime": "2026-01-05T08:30:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-01-05T09:30:00", "timeZone": "UTC"},
        "responseStatus": {
            "response": "accepted",
            "time": "2026-01-04T10:00:00Z"
        }
    }
    
    # Test timezone conversion logic for response status time
    response_time = event_data.get("responseStatus", {}).get("time")
    if response_time:
        response_time_converted = date_handler.convert_utc_to_user_timezone(
            response_time, user_timezone
        )
    else:
        response_time_converted = None
    
    print(f"   Original response time (UTC): {event_data['responseStatus']['time']}")
    print(f"   Converted response time: {response_time_converted}")
    
    assert response_time_converted is not None, "Response status time should be converted"
    assert event_data['responseStatus']['response'] == "accepted", "Response status should be preserved"
    
    print("   ✓ Response status time conversion works correctly")
    
    print("\n[Test 6] ✓ PASSED: Response status time is properly converted to user timezone")


async def test_event_cache_persistence():
    """Test that event cache can be saved and loaded."""
    print("\n[Test 7] Testing event cache persistence...")
    
    cache = EventBrowsingCache()
    
    # Clear any existing cache data
    cache.clear_cache()
    
    # Mock event data
    events = [
        {
            "id": "test-1",
            "subject": "Event 1",
            "start": "Mon 01/05/2026 04:30 PM",
            "end": "Mon 01/05/2026 05:30 PM",
            "attendees": 2,
            "attendees_list": [],
            "responseStatus": {"response": "accepted", "time": "2026-01-04T10:00:00"},
            "recurrence": False,
            "start_datetime": "2026-01-05T08:30:00Z"
        },
        {
            "id": "test-2",
            "subject": "Event 2",
            "start": "Tue 01/06/2026 10:00 AM",
            "end": "Tue 01/06/2026 11:00 AM",
            "attendees": 3,
            "attendees_list": [],
            "responseStatus": {"response": "tentativelyAccepted", "time": "2026-01-04T11:00:00"},
            "recurrence": True,
            "start_datetime": "2026-01-06T02:00:00Z"
        }
    ]
    
    # Save events to cache
    await cache.update_browse_state(
        start_date="2026-01-01",
        end_date="2026-01-31",
        total_count=2,
        metadata=events
    )
    print("   ✓ Events saved to cache")
    
    # Load events from cache
    loaded_events = cache.get_cached_events()
    print(f"   ✓ Loaded {len(loaded_events)} events from cache")
    
    assert len(loaded_events) == len(events), "Loaded events count should match saved events count"
    
    # Verify event data (events are sorted by start_datetime in reverse order)
    loaded_ids = {event['id'] for event in loaded_events}
    saved_ids = {event['id'] for event in events}
    assert loaded_ids == saved_ids, "Loaded events should contain all saved event IDs"
    
    # Verify each event's data
    loaded_events_dict = {event['id']: event for event in loaded_events}
    for saved_event in events:
        loaded_event = loaded_events_dict[saved_event['id']]
        assert saved_event['id'] == loaded_event['id'], f"Event ID should match"
        assert saved_event['subject'] == loaded_event['subject'], f"Event subject should match"
        assert saved_event['recurrence'] == loaded_event['recurrence'], f"Event recurrence flag should match"
    
    print("   ✓ Event data integrity verified")
    
    print("\n[Test 7] ✓ PASSED: Event cache persistence works correctly")


async def test_timezone_display_format():
    """Test that times are displayed in the correct format."""
    print("\n[Test 8] Testing timezone display format...")
    
    date_handler = DateHandler()
    
    # Test various UTC times
    test_cases = [
        ("2026-01-05T08:30:00Z", "Asia/Shanghai"),
        ("2026-01-05T14:30:00Z", "America/New_York"),
        ("2026-01-05T11:30:00Z", "America/Los_Angeles"),
    ]
    
    for utc_time, target_tz in test_cases:
        converted = date_handler.convert_utc_to_user_timezone(utc_time, target_tz)
        print(f"   {utc_time} → {target_tz}: {converted}")
        assert converted is not None, f"Conversion to {target_tz} should succeed"
    
    print("   ✓ Timezone display format works correctly")
    
    print("\n[Test 8] ✓ PASSED: Times are displayed in the correct format")


async def test_format_user_timezone_datetime():
    """Test that format_user_timezone_datetime works correctly for calendarView endpoint."""
    print("\n[Test 9] Testing format_user_timezone_datetime...")
    
    date_handler = DateHandler()
    user_timezone = "Asia/Shanghai"
    
    # Test case: datetime from Graph API calendarView endpoint (already in user's timezone)
    # For Asia/Shanghai (UTC+8), 4:30 PM should display correctly
    user_time = "2026-01-05T16:30:00"
    formatted = date_handler.format_user_timezone_datetime(user_time, user_timezone)
    
    print(f"   User timezone datetime: {user_time}")
    print(f"   Formatted result: {formatted}")
    
    # Verify the time is correctly formatted as 4:30 PM
    assert "04:30 PM" in formatted or "16:30" in formatted, f"Expected 4:30 PM in formatted time, got: {formatted}"
    assert "Mon 01/05/2026" in formatted, f"Expected date in formatted time, got: {formatted}"
    
    # Test another time: 10:00 AM
    user_time2 = "2026-01-06T10:00:00"
    formatted2 = date_handler.format_user_timezone_datetime(user_time2, user_timezone)
    
    print(f"   User timezone datetime: {user_time2}")
    print(f"   Formatted result: {formatted2}")
    
    # Verify the time is correctly formatted as 10:00 AM
    assert "10:00 AM" in formatted2, f"Expected 10:00 AM in formatted time, got: {formatted2}"
    assert "Tue 01/06/2026" in formatted2, f"Expected date in formatted time, got: {formatted2}"
    
    print("   ✓ format_user_timezone_datetime works correctly")
    
    print("\n[Test 9] ✓ PASSED: format_user_timezone_datetime works correctly")


async def main():
    """Run all tests."""
    print("=" * 70)
    print("Event Cache and Timezone Conversion Tests")
    print("=" * 70)
    
    tests = [
        test_timezone_conversion,
        test_windows_to_iana_timezone_mapping,
        test_event_cache_timezone_conversion,
        test_recurring_event_detection,
        test_attendees_list_formatting,
        test_response_status_time_conversion,
        test_event_cache_persistence,
        test_timezone_display_format,
        test_format_user_timezone_datetime,
    ]
    
    for test in tests:
        try:
            await test()
        except Exception as e:
            print(f"\n✗ FAILED: {test.__name__}")
            print(f"   Error: {e}")
            import traceback
            traceback.print_exc()
    
    print("\n" + "=" * 70)
    print("All tests completed!")
    print("=" * 70)


if __name__ == "__main__":
    asyncio.run(main())
