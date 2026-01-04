"""Test script for respond_to_event tool functionality."""

import asyncio
import json
from pathlib import Path
import sys
from unittest.mock import AsyncMock, patch

sys.path.insert(0, str(Path(__file__).parent.parent))

from microsoft_graph_mcp_server.tools.registry import ToolRegistry
from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler


async def test_respond_to_event_schema():
    """Test that respond_to_event tool schema is correctly defined."""
    print("\n[Test 1] Checking respond_to_event tool schema...")
    
    tool = ToolRegistry.respond_to_event()
    
    assert tool.name == "respond_to_event", f"Tool name should be 'respond_to_event', got '{tool.name}'"
    print(f"   ✓ Tool name: {tool.name}")
    
    schema = tool.inputSchema
    assert schema["type"] == "object", "Schema should be an object"
    print("   ✓ Schema type: object")
    
    action_enum = schema["properties"]["action"]["enum"]
    expected_actions = ["accept", "decline", "tentatively_accept", "propose_new_time", "delete"]
    
    for action in expected_actions:
        assert action in action_enum, f"Action '{action}' not found in enum"
    print(f"   ✓ All expected actions present: {', '.join(expected_actions)}")
    
    assert "propose_new_time" in schema["properties"], "propose_new_time property should exist"
    print("   ✓ propose_new_time property exists")
    
    propose_new_time_props = schema["properties"]["propose_new_time"]["properties"]
    assert "dateTime" in propose_new_time_props, "propose_new_time should have dateTime property"
    assert "timeZone" in propose_new_time_props, "propose_new_time should have timeZone property"
    print("   ✓ propose_new_time has required properties: dateTime, timeZone")
    
    required_fields = schema.get("required", [])
    assert "action" in required_fields, "action should be required"
    assert "event_id" in required_fields, "event_id should be required"
    print("   ✓ Required fields: action, event_id")
    
    print("\n[Test 1] ✓ PASSED: respond_to_event tool schema is correct")


async def test_handler_methods():
    """Test that all handler methods exist."""
    print("\n[Test 2] Checking handler methods...")
    
    handler = CalendarHandler()
    
    methods = [
        "handle_respond_to_event",
        "_handle_accept_event_action",
        "_handle_decline_event_action",
        "_handle_tentatively_accept_event_action",
        "_handle_propose_new_time_action",
        "_handle_delete_cancelled_event_action",
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
        "accept_event",
        "decline_event",
        "tentatively_accept_event",
        "propose_new_time",
    ]
    
    for method_name in methods:
        assert hasattr(client, method_name), f"Client method '{method_name}' not found"
        print(f"   ✓ Client method exists: {method_name}")
    
    print("\n[Test 3] ✓ PASSED: All calendar client methods exist")


async def test_accept_single_event():
    """Test accepting a single event."""
    print("\n[Test 4] Testing accept single event...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "accept",
        "event_id": "1",
        "comment": "I'll be there.",
        "send_response": True
    }
    
    try:
        with patch.object(handler, '_resolve_event_id', new_callable=AsyncMock) as mock_resolve:
            mock_resolve.return_value = {
                "event_id": "test-event-id",
                "series_master_id": None,
                "is_recurring": False
            }
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.accept_event', new_callable=AsyncMock) as mock_accept:
                mock_accept.return_value = {"status": "accepted"}
                
                result = await handler.handle_respond_to_event(arguments)
                
                assert len(result) > 0, "Result should not be empty"
                assert "accepted successfully" in result[0].text.lower(), "Result should indicate success"
                print("   ✓ Accept single event works")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_decline_single_event():
    """Test declining a single event."""
    print("\n[Test 5] Testing decline single event...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "decline",
        "event_id": "1",
        "comment": "I cannot attend.",
        "send_response": True
    }
    
    try:
        with patch.object(handler, '_resolve_event_id', new_callable=AsyncMock) as mock_resolve:
            mock_resolve.return_value = {
                "event_id": "test-event-id",
                "series_master_id": None,
                "is_recurring": False
            }
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.decline_event', new_callable=AsyncMock) as mock_decline:
                mock_decline.return_value = {"status": "declined"}
                
                result = await handler.handle_respond_to_event(arguments)
                
                assert len(result) > 0, "Result should not be empty"
                assert "declined successfully" in result[0].text.lower(), "Result should indicate success"
                print("   ✓ Decline single event works")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_tentatively_accept_single_event():
    """Test tentatively accepting a single event."""
    print("\n[Test 6] Testing tentatively accept single event...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "tentatively_accept",
        "event_id": "1",
        "comment": "I'll try to make it.",
        "send_response": True
    }
    
    try:
        with patch.object(handler, '_resolve_event_id', new_callable=AsyncMock) as mock_resolve:
            mock_resolve.return_value = {
                "event_id": "test-event-id",
                "series_master_id": None,
                "is_recurring": False
            }
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.tentatively_accept_event', new_callable=AsyncMock) as mock_tentative:
                mock_tentative.return_value = {"status": "tentativelyAccepted"}
                
                result = await handler.handle_respond_to_event(arguments)
                
                assert len(result) > 0, "Result should not be empty"
                assert "tentatively accepted" in result[0].text.lower(), "Result should indicate success"
                print("   ✓ Tentatively accept single event works")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_accept_series():
    """Test accepting an entire recurring series."""
    print("\n[Test 7] Testing accept series...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "accept",
        "event_id": "1",
        "series": True,
        "comment": "I'll attend all meetings.",
        "send_response": True
    }
    
    try:
        with patch.object(handler, '_resolve_event_id', new_callable=AsyncMock) as mock_resolve:
            mock_resolve.return_value = {
                "event_id": "test-event-id",
                "series_master_id": "test-series-id",
                "is_recurring": True
            }
            
            with patch('microsoft_graph_mcp_server.graph_client.graph_client.accept_event', new_callable=AsyncMock) as mock_accept:
                mock_accept.return_value = {"status": "accepted"}
                
                result = await handler.handle_respond_to_event(arguments)
                
                assert len(result) > 0, "Result should not be empty"
                assert "entire recurring series accepted" in result[0].text.lower(), "Result should indicate series acceptance"
                mock_accept.assert_called_once_with("test-series-id", "I'll attend all meetings.", True, True)
                print("   ✓ Accept series works with series_master_id")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_propose_new_time():
    """Test proposing a new time for an event."""
    print("\n[Test 8] Testing propose new time...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    from unittest.mock import patch, AsyncMock
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "propose_new_time",
        "event_id": "1",
        "propose_new_time": {
            "dateTime": "2024-12-31T14:30",
            "timeZone": "America/New_York"
        },
        "comment": "Can we move to this time?",
        "send_response": True
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
                
                with patch('microsoft_graph_mcp_server.graph_client.graph_client.propose_new_time', new_callable=AsyncMock) as mock_propose:
                    mock_propose.return_value = {"status": "proposed"}
                    
                    result = await handler.handle_respond_to_event(arguments)
                    
                    assert len(result) > 0, "Result should not be empty"
                    assert "proposed new time" in result[0].text.lower(), "Result should indicate success"
                    print("   ✓ Propose new time works")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_delete_cancelled_event():
    """Test deleting a cancelled event."""
    print("\n[Test 9] Testing delete cancelled event...")
    
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
                
                result = await handler.handle_respond_to_event(arguments)
                
                assert len(result) > 0, "Result should not be empty"
                assert "deleted successfully" in result[0].text.lower(), "Result should indicate success"
                print("   ✓ Delete cancelled event works")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def test_invalid_action():
    """Test handling of invalid action."""
    print("\n[Test 10] Testing invalid action...")
    
    from microsoft_graph_mcp_server.handlers.calendar_handlers import CalendarHandler
    
    handler = CalendarHandler()
    
    arguments = {
        "action": "invalid_action",
        "event_id": "1"
    }
    
    try:
        result = await handler.handle_respond_to_event(arguments)
        
        assert len(result) > 0, "Result should not be empty"
        assert "invalid action" in result[0].text.lower(), "Result should indicate invalid action"
        print("   ✓ Invalid action handled correctly")
    except Exception as e:
        print(f"   ✗ Error: {e}")
        raise


async def main():
    """Run all tests."""
    print("=" * 70)
    print("Testing respond_to_event tool functionality")
    print("=" * 70)
    
    tests = [
        test_respond_to_event_schema,
        test_handler_methods,
        test_calendar_client_methods,
        test_accept_single_event,
        test_decline_single_event,
        test_tentatively_accept_single_event,
        test_accept_series,
        test_propose_new_time,
        test_delete_cancelled_event,
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
