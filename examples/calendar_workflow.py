"""
Example: Complete calendar workflow for Microsoft Graph MCP Server.

This example demonstrates calendar operations including
availability checking, event creation, searching, browsing, and responding.
"""

import asyncio
import json


async def calendar_availability_and_create_workflow():
    """Demonstrate availability check and event creation workflow."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - CALENDAR AVAILABILITY & CREATE")
    print("=" * 70)
    print()

    # ============================================
    # STEP 1: Check attendee availability
    # ============================================
    print("STEP 1: Checking attendee availability...")
    print("-" * 70)

    availability_result = {
        "tool": "check_attendee_availability",
        "arguments": {
            "attendees": ["alice@example.com", "bob@example.com"],
            "date": "2024-01-15",
            "time_zone": "America/New_York",
        },
    }

    print(f"Calling: {json.dumps(availability_result, indent=2)}")
    print()

    # Expected response:
    expected_availability_response = {
        "success": True,
        "message": "Availability check completed",
        "availability_view": "0000011110000011110000",  # 0=Free, 2=Busy
        "schedule_items": [
            {
                "start": "2024-01-15T09:00:00",
                "end": "2024-01-15T10:00:00",
                "status": "busy",
            }
        ],
        "top_slots": [
            {
                "start": "2024-01-15T10:00:00",
                "end": "2024-01-15T11:00:00",
                "availability": "free",
            },
            {
                "start": "2024-01-15T14:00:00",
                "end": "2024-01-15T15:00:00",
                "availability": "free",
            },
        ],
    }

    print("Expected Response:")
    print(json.dumps(expected_availability_response, indent=2))
    print()
    print("💡 TIP: Use 'top_slots' to find best meeting times")

    # ============================================
    # STEP 2: Create meeting at free slot
    # ============================================
    print("STEP 2: Creating meeting at free slot...")
    print("-" * 70)

    create_result = {
        "tool": "manage_my_event",
        "arguments": {
            "action": "create",
            "subject": "Project Review",
            "start": "2024-01-15T10:00",
            "end": "2024-01-15T11:00",
            "timezone": "America/New_York",
            "location": "Conference Room A",
            "attendees": ["alice@example.com", "bob@example.com"],
            "body": "Please review project deliverables.",
            "isOnlineMeeting": True,
            "onlineMeetingProvider": "teamsForBusiness",
        },
    }

    print(f"Calling: {json.dumps(create_result, indent=2)}")
    print()

    # Expected response:
    expected_create_response = {
        "success": True,
        "message": "Event created successfully",
        "event": {
            "id": "...",
            "subject": "Project Review",
            "start": "2024-01-15T15:00:00Z",  # Converted to UTC
            "end": "2024-01-15T16:00:00Z",
            "location": {"displayName": "Conference Room A"},
            "attendees": [
                {
                    "email": {"address": "alice@example.com"},
                    "status": {"response": "accepted"},
                }
            ],
            "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/..."},
        },
    }

    print("Expected Response:")
    print(json.dumps(expected_create_response, indent=2))
    print()


async def calendar_search_and_browse_workflow():
    """Demonstrate calendar search and browse workflow."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - CALENDAR SEARCH & BROWSE")
    print("=" * 70)
    print()

    # ============================================
    # STEP 1: Search events
    # ============================================
    print("STEP 1: Searching events (loads into cache)...")
    print("-" * 70)

    search_result = {"tool": "search_events", "arguments": {"time_range": "this_week"}}

    print(f"Calling: {json.dumps(search_result, indent=2)}")
    print()

    # Expected response:
    expected_search_response = {
        "success": True,
        "events": [
            {
                "number": 1,
                "subject": "Project Review",
                "start": "2024-01-15T10:00",
                "end": "2024-01-15T11:00",
                "location": "Conference Room A",
                "organizer": "John Doe",
            }
            # ... more events
        ],
        "count": 50,
        "date_range": "This Week",
        "timezone": "America/New_York",
    }

    print("Expected Response:")
    print(json.dumps(expected_search_response, indent=2))
    print()

    # ============================================
    # STEP 2: Browse cache
    # ============================================
    print("STEP 2: Browsing event cache...")
    print("-" * 70)

    browse_result = {
        "tool": "browse_events",
        "arguments": {"page_number": 1, "mode": "llm"},
    }

    print(f"Calling: {json.dumps(browse_result, indent=2)}")
    print()

    expected_browse_response = {
        "success": True,
        "current_page": 1,
        "total_pages": 3,
        "count": 20,
        "total_count": 50,
        "events": [
            {
                "number": 1,  # ← Use this, not array index!
                "subject": "Project Review",
                "start": "2024-01-15T10:00",
                "end": "2024-01-15T11:00",
                "location": "Conference Room A",
            }
        ],
        "timezone": "America/New_York",
    }

    print("Expected Response:")
    print(json.dumps(expected_browse_response, indent=2))
    print()


async def calendar_actions_workflow():
    """Demonstrate calendar action operations."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - CALENDAR ACTIONS WORKFLOW")
    print("=" * 70)
    print()

    # ============================================
    # ACTION 1: Respond to event
    # ============================================
    print("ACTION 1: Accepting event invitation...")
    print("-" * 70)

    accept_result = {
        "tool": "respond_to_event",
        "arguments": {
            "action": "accept",
            "cache_number": 3,
            "comment": "I'll attend!",
            "send_response": True,
        },
    }

    print(f"Calling: {json.dumps(accept_result, indent=2)}")
    print()

    expected_accept_response = {
        "success": True,
        "message": "Event accepted successfully",
        "response": "accepted",
    }

    print("Expected Response:")
    print(json.dumps(expected_accept_response, indent=2))
    print()

    # ============================================
    # ACTION 2: Propose new time
    # ============================================
    print("ACTION 2: Proposing new time for event...")
    print("-" * 70)

    propose_result = {
        "tool": "respond_to_event",
        "arguments": {
            "action": "propose_new_time",
            "cache_number": 5,
            "comment": "Can we move to 2pm?",
            "propose_new_time": {
                "dateTime": "2024-01-16T14:00",  # New time
                "timeZone": "America/New_York",
            },
        },
    }

    print(f"Calling: {json.dumps(propose_result, indent=2)}")
    print()

    expected_propose_response = {
        "success": True,
        "message": "New time proposed successfully",
    }

    print("Expected Response:")
    print(json.dumps(expected_propose_response, indent=2))
    print()

    # ============================================
    # ACTION 3: Cancel own event
    # ============================================
    print("ACTION 3: Cancelling own event...")
    print("-" * 70)

    cancel_result = {
        "tool": "manage_my_event",
        "arguments": {
            "action": "cancel",
            "cache_number": 1,
            "comment": "Meeting cancelled due to conflict",
        },
    }

    print(f"Calling: {json.dumps(cancel_result, indent=2)}")
    print()

    expected_cancel_response = {
        "success": True,
        "message": "Event cancelled successfully",
        "cancellation_notification_sent": True,
    }

    print("Expected Response:")
    print(json.dumps(expected_cancel_response, indent=2))
    print()


async def recurring_event_example():
    """Demonstrate creating recurring event."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - RECURRING EVENT")
    print("=" * 70)
    print()

    print("CREATING WEEKLY RECURRING EVENT...")
    print("-" * 70)

    recurring_result = {
        "tool": "manage_my_event",
        "arguments": {
            "action": "create",
            "subject": "Weekly Team Standup",
            "start": "2024-01-15T10:00",
            "end": "2024-01-15T10:30",
            "timezone": "America/New_York",
            "attendees": ["team@company.com"],
            "body": "30-minute standup meeting",
            "isOnlineMeeting": True,
            "onlineMeetingProvider": "teamsForBusiness",
            "recurrence": {
                "pattern": {
                    "type": "weekly",
                    "interval": 1,
                    "daysOfWeek": ["monday", "wednesday", "friday"],
                },
                "range": {
                    "type": "endDate",
                    "startDate": "2024-01-15",
                    "endDate": "2024-06-30",
                },
            },
        },
    }

    print(f"Calling: {json.dumps(recurring_result, indent=2)}")
    print()

    print("💡 This creates a recurring event:")
    print("   - Every Monday, Wednesday, Friday")
    print("   - From Jan 15 to June 30, 2024")
    print("   - 10:00-10:30 AM Eastern Time")
    print()


async def common_mistakes_example():
    """Show common calendar mistakes and how to avoid them."""

    print()
    print("=" * 70)
    print("COMMON MISTAKES - CALENDAR WORKFLOWS")
    print("=" * 70)
    print()

    # ============================================
    # MISTAKE 1: Not checking availability
    # ============================================
    print("❌ MISTAKE 1: Not checking availability before creating meeting")
    print("-" * 70)

    print("Incorrect:")
    print("  Create meeting at random time")
    print("  manage_my_event(action='create', start='2024-01-15T10:00', ...)")
    print("  → May conflict with existing meetings")
    print()

    print("Correct:")
    print("  1. check_attendee_availability(date='2024-01-15', ...)")
    print("  2. Create meeting at free slot from top_slots")
    print("  3. manage_my_event(action='create', start=free_time, ...)")
    print()

    # ============================================
    # MISTAKE 2: Wrong time format
    # ============================================
    print("❌ MISTAKE 2: Using wrong time format")
    print("-" * 70)

    print("Incorrect:")
    print("  manage_my_event(action='create', start='Jan 15, 2024', ...)")
    print("  → Invalid time format")
    print()

    print("Correct:")
    print("  manage_my_event(action='create', start='2024-01-15T10:00', ...)")
    print("  OR start='2024-01-15 10:00'")
    print("  → ISO format or space-separated date/time")
    print()

    print("=" * 70)


if __name__ == "__main__":
    asyncio.run(calendar_availability_and_create_workflow())
    print()
    asyncio.run(calendar_search_and_browse_workflow())
    print()
    asyncio.run(calendar_actions_workflow())
    print()
    asyncio.run(recurring_event_example())
    print()
    asyncio.run(common_mistakes_example())
