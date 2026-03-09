"""Test script for manage_event_as_organizer tool functionality."""

import asyncio
import json
from pathlib import Path
import sys
from unittest.mock import AsyncMock, patch

sys.path.insert(0, str(Path(__file__).parent.parent))

from microsoft_graph_mcp_server.tools.registry import ToolRegistry
from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler


async def test_manage_my_event_schema():
    """Test that manage_event_as_organizer tool schema is correctly defined."""
    print("\n[Test 1] Checking manage_event_as_organizer tool schema...")

    tool = ToolRegistry.manage_event_as_organizer()

    assert tool.name == "manage_event_as_organizer", f"Tool name should be 'manage_event_as_organizer', got '{tool.name}'"
    print(f"   ✓ Tool name: {tool.name}")

    schema = tool.inputSchema
    assert schema["type"] == "object", "Schema should be an object"
    print("   ✓ Schema type: object")

    action_enum = schema["properties"]["action"]["enum"]
    expected_actions = ["create", "update", "cancel", "forward", "email_attendees"]
    
    for action in expected_actions:
        assert action in action_enum, f"Action '{action}' not found in enum"
    print(f"   ✓ All expected actions present: {', '.join(expected_actions)}")
    
    assert "recurrence" in schema["properties"], "recurrence property should exist"
    print("   ✓ recurrence property exists")
    
    recurrence_props = schema["properties"]["recurrence"]["properties"]
    assert "pattern" in recurrence_props, "recurrence should have pattern property"
    assert "range" in recurrence_props, "recurrence should have range property"
    print("   ✓ recurrence has required properties: pattern, range")
    
    required_fields = schema.get("required", [])
    assert "action" in required_fields, "action should be required"
    print("   ✓ Required field: action")
    
    print("\n[Test 1] ✓ PASSED: manage_event_as_organizer tool schema is correct")


async def test_handler_methods():
    """Test that all handler methods exist."""
    print("\n[Test 2] Checking handler methods...")
    
    handler = CalendarHandler()
    
    methods = [
        "handle_manage_my_event",
        "_handle_create_event_action",
        "_handle_update_event_action",
        "_handle_cancel_event_action",
        "_handle_delete_event_action",
        "_handle_forward_event_action",
        "_handle_reply_event_action",
        "_resolve_event_id",
    ]
    
    for method_name in methods:
        assert hasattr(handler, method_name), f"Handler method '{method_name}' not found"
        print(f"   ✓ Handler method exists: {method_name}")
    
    print("\n[Test 2] ✓ PASSED: All handler methods exist")


async def test_calendar_client_methods():
    """Test that calendar client methods exist."""
    print("\n[Test 3] Checking calendar client methods...")
    
    from microsoft_graph_mcp_server.clients.calendar_client import CalendarClient
    
    client = CalendarClient()
    
    methods = [
        "create_event",
        "update_event",
        "cancel_event",
        "delete_event",
        "forward_event",
    ]
    
    for method_name in methods:
        assert hasattr(client, method_name), f"Client method '{method_name}' not found"
        print(f"   ✓ Client method exists: {method_name}")
    
    print("\n[Test 3] ✓ PASSED: All calendar client methods exist")


async def test_create_event():
    """Test creating a new event."""
    print("\n[Test 4] Testing create event...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "create",
        "subject": "Team Meeting",
        "start": "2024-01-15T14:00",
        "end": "2024-01-15T15:00",
        "location": "Conference Room A",
        "body": "Weekly team sync",
        "attendees": ["user1@example.com", "user2@example.com"]
    }
    
    try:
        with patch('microsoft_graph_mcp_server.graph_client.graph_client.get_user_timezone', new_callable=AsyncMock) as mock_tz:
            mock_tz.return_value = "America/New_York"
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.create_event', new_callable=AsyncMock) as mock_create:
                mock_create.return_value = {"id": "new-event-id", "subject": "Team Meeting"}
                
                result = await handler.handle_manage_my_event(arguments)
                
                assert len(result) > 0, "Result should not be empty"
                assert "created successfully" in result[0].text.lower(), "Result should indicate success"
                print("   ✓ Create event works")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_update_event():
    """Test updating an existing event."""
    print("\n[Test 5] Testing update event...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "update",
        "event_id": "1",
        "subject": "Updated Meeting",
        "start": "2024-01-15T15:00",
        "end": "2024-01-15T16:00"
    }
    
    try:
        with patch.object(handler, '_resolve_event_id', new_callable=AsyncMock) as mock_resolve:
            mock_resolve.return_value = {
                "event_id": "test-event-id",
                "series_master_id": None,
                "is_recurring": False
            }
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.get_user_timezone', new_callable=AsyncMock) as mock_tz:
                mock_tz.return_value = "America/New_York"
                
                with patch('microsoft_graph_mcp_server.graph_client.graph_client.update_event', new_callable=AsyncMock) as mock_update:
                    mock_update.return_value = {"id": "test-event-id", "subject": "Updated Meeting"}
                    
                    result = await handler.handle_manage_my_event(arguments)
                    
                    assert len(result) > 0, "Result should not be empty"
                    assert "updated successfully" in result[0].text.lower(), "Result should indicate success"
                    print("   ✓ Update event works with cache number")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_cancel_event():
    """Test cancelling an event."""
    print("\n[Test 6] Testing cancel event...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "cancel",
        "event_id": "1",
        "comment": "Meeting cancelled due to scheduling conflict"
    }
    
    try:
        with patch.object(handler, '_resolve_event_id', new_callable=AsyncMock) as mock_resolve:
            mock_resolve.return_value = {
                "event_id": "test-event-id",
                "series_master_id": None,
                "is_recurring": False
            }
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.cancel_event', new_callable=AsyncMock) as mock_cancel:
                mock_cancel.return_value = {"status": "cancelled"}
                
                result = await handler.handle_manage_my_event(arguments)
                
                assert len(result) > 0, "Result should not be empty"
                assert "cancelled successfully" in result[0].text.lower(), "Result should indicate success"
                print("   ✓ Cancel event works with cache number")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_delete_event():
    """Test deleting an event."""
    print("\n[Test 7] Testing delete event...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "delete",
        "event_id": "1"
    }
    
    try:
        with patch.object(handler, '_resolve_event_id', new_callable=AsyncMock) as mock_resolve:
            mock_resolve.return_value = {
                "event_id": "test-event-id",
                "series_master_id": None,
                "is_recurring": False
            }
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.delete_event', new_callable=AsyncMock) as mock_delete:
                mock_delete.return_value = {"status": "deleted"}
                
                result = await handler.handle_manage_my_event(arguments)
                
                assert len(result) > 0, "Result should not be empty"
                assert "deleted successfully" in result[0].text.lower(), "Result should indicate success"
                print("   ✓ Delete event works with cache number")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_forward_event():
    """Test forwarding an event."""
    print("\n[Test 8] Testing forward event...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "forward",
        "event_id": "1",
        "attendees": ["newuser@example.com"],
        "comment": "Please attend this meeting"
    }
    
    try:
        with patch.object(handler, '_resolve_event_id', new_callable=AsyncMock) as mock_resolve:
            mock_resolve.return_value = {
                "event_id": "test-event-id",
                "series_master_id": None,
                "is_recurring": False
            }
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.forward_event', new_callable=AsyncMock) as mock_forward:
                mock_forward.return_value = {"status": "forwarded"}
                
                result = await handler.handle_manage_my_event(arguments)
                
                assert len(result) > 0, "Result should not be empty"
                assert "forwarded successfully" in result[0].text.lower(), "Result should indicate success"
                print("   ✓ Forward event works with cache number")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_reply_to_event():
    """Test replying to an event."""
    print("\n[Test 9] Testing reply to event...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "reply",
        "event_id": "1",
        "subject": "Re: Team Meeting",
        "body": "Thanks for the meeting invite!"
    }
    
    try:
        with patch.object(handler, '_resolve_event_id', new_callable=AsyncMock) as mock_resolve:
            mock_resolve.return_value = {
                "event_id": "test-event-id",
                "series_master_id": None,
                "is_recurring": False
            }
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.get_event', new_callable=AsyncMock) as mock_get_event:
                mock_get_event.return_value = {
                    "id": "test-event-id",
                    "subject": "Team Meeting",
                    "body": {"content": "Meeting details", "contentType": "Text"},
                    "attendees": [
                        {"emailAddress": {"address": "attendee1@example.com"}, "type": "required"},
                        {"emailAddress": {"address": "attendee2@example.com"}, "type": "optional"}
                    ]
                }
                
                with patch('microsoft_graph_mcp_server.graph_client.graph_client.send_email', new_callable=AsyncMock) as mock_send:
                    mock_send.return_value = {"status": "sent"}
                    
                    result = await handler.handle_manage_my_event(arguments)
                    
                    assert len(result) > 0, "Result should not be empty"
                    assert "sent successfully" in result[0].text.lower(), "Result should indicate success"
                    print("   ✓ Reply to event works with cache number")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_create_event_with_recurrence():
    """Test creating a recurring event."""
    print("\n[Test 10] Testing create event with recurrence...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "create",
        "subject": "Weekly Standup",
        "start": "2024-01-15T09:00",
        "end": "2024-01-15T09:30",
        "recurrence": {
            "pattern": {
                "type": "weekly",
                "interval": 1,
                "daysOfWeek": ["monday"]
            },
            "range": {
                "type": "noEnd",
                "startDate": "2024-01-15"
            }
        }
    }
    
    try:
        with patch('microsoft_graph_mcp_server.graph_client.graph_client.get_user_timezone', new_callable=AsyncMock) as mock_tz:
            mock_tz.return_value = "America/New_York"
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.create_event', new_callable=AsyncMock) as mock_create:
                mock_create.return_value = {"id": "new-event-id", "subject": "Weekly Standup"}
                
                result = await handler.handle_manage_my_event(arguments)
                
                assert len(result) > 0, "Result should not be empty"
                assert "created successfully" in result[0].text.lower(), "Result should indicate success"
                print("   ✓ Create event with recurrence works")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_invalid_action():
    """Test handling of invalid action."""
    print("\n[Test 11] Testing invalid action...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "invalid_action",
        "event_id": "1"
    }
    
    try:
        result = await handler.handle_manage_my_event(arguments)
        
        assert len(result) > 0, "Result should not be empty"
        assert "invalid action" in result[0].text.lower(), "Result should indicate invalid action"
        print("   ✓ Invalid action handled correctly")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def main():
    """Run all tests."""
    print("=" * 70)
    print("Testing manage_my_event tool functionality")
    print("=" * 70)
    
    tests = [
        test_manage_my_event_schema,
        test_handler_methods,
        test_calendar_client_methods,
        test_create_event,
        test_update_event,
        test_cancel_event,
        test_delete_event,
        test_forward_event,
        test_reply_to_event,
        test_create_event_with_recurrence,
        test_invalid_action,
    ]
    
    failed_tests = []
    
    for test in tests:
        try:
            await test()
        except Exception as e:
            print(f"\n✗ Test failed: {test.__name__}")
            print(f"  Error: {e}")
            failed_tests.append(test.__name__)
    
    print("\n" + "=" * 70)
    if failed_tests:
        print(f"✗ {len(failed_tests)} test(s) failed:")
        for test_name in failed_tests:
            print(f"  - {test_name}")
    else:
        print("✓ All tests passed!")
    print("=" * 70)


if __name__ == "__main__":
    asyncio.run(main())
