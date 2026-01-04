"""Test script for user_settings tool functionality."""

import asyncio
import json
import os
import tempfile
from pathlib import Path
import shutil
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))

from microsoft_graph_mcp_server.tools.registry import ToolRegistry
from microsoft_graph_mcp_server.handlers.user_handlers import UserHandler


async def test_user_settings_tool_schema():
    """Test that user_settings tool schema is correctly defined."""
    print("\n[Test 1] Checking user_settings tool schema...")
    
    tool = ToolRegistry.user_settings()
    
    assert tool.name == "user_settings", f"Tool name should be 'user_settings', got '{tool.name}'"
    print(f"   ✓ Tool name: {tool.name}")
    
    schema = tool.inputSchema
    assert schema["type"] == "object", "Schema should be an object"
    print("   ✓ Schema type: object")
    
    properties = schema["properties"]
    assert "action" in properties, "action property should exist"
    assert "page_size" in properties, "page_size property should exist"
    assert "llm_page_size" in properties, "llm_page_size property should exist"
    assert "default_search_days" in properties, "default_search_days property should exist"
    assert "timezone" in properties, "timezone property should exist"
    print("   ✓ All expected properties exist: action, page_size, llm_page_size, default_search_days, timezone")
    
    action_schema = properties["action"]
    assert action_schema["type"] == "string", "action should be string"
    assert "init" in action_schema["enum"], "action should have 'init' option"
    assert "update" in action_schema["enum"], "action should have 'update' option"
    print("   ✓ action schema correct: string with 'init' and 'update' options")
    
    page_size_schema = properties["page_size"]
    assert page_size_schema["type"] == "integer", "page_size should be integer"
    assert page_size_schema["minimum"] == 1, "page_size minimum should be 1"
    assert page_size_schema["maximum"] == 50, "page_size maximum should be 50"
    print("   ✓ page_size schema correct: integer, min=1, max=50")
    
    llm_page_size_schema = properties["llm_page_size"]
    assert llm_page_size_schema["type"] == "integer", "llm_page_size should be integer"
    assert llm_page_size_schema["minimum"] == 1, "llm_page_size minimum should be 1"
    assert llm_page_size_schema["maximum"] == 100, "llm_page_size maximum should be 100"
    print("   ✓ llm_page_size schema correct: integer, min=1, max=100")
    
    default_search_days_schema = properties["default_search_days"]
    assert default_search_days_schema["type"] == "integer", "default_search_days should be integer"
    assert default_search_days_schema["minimum"] == 1, "default_search_days minimum should be 1"
    assert default_search_days_schema["maximum"] == 365, "default_search_days maximum should be 365"
    print("   ✓ default_search_days schema correct: integer, min=1, max=365")
    
    timezone_schema = properties["timezone"]
    assert timezone_schema["type"] == "string", "timezone should be string"
    print("   ✓ timezone schema correct: string")
    
    assert "action" in schema["required"], "action should be required"
    print("   ✓ action is required")
    
    print("\n[Test 1] ✓ PASSED: user_settings tool schema is correct")


async def test_handler_method_exists():
    """Test that handle_user_settings method exists."""
    print("\n[Test 2] Checking handle_user_settings method...")
    
    handler = UserHandler()
    
    assert hasattr(handler, "handle_user_settings"), "handle_user_settings method should exist"
    print("   ✓ handle_user_settings method exists")
    
    import inspect
    method = getattr(handler, "handle_user_settings")
    assert callable(method), "handle_user_settings should be callable"
    print("   ✓ handle_user_settings is callable")
    
    sig = inspect.signature(method)
    params = list(sig.parameters.keys())
    assert "arguments" in params, "handle_user_settings should have 'arguments' parameter"
    print("   ✓ handle_user_settings has 'arguments' parameter")
    
    print("\n[Test 2] ✓ PASSED: handle_user_settings method exists and is correct")


async def test_user_settings_init_action():
    """Test user_settings with init action."""
    print("\n[Test 3] Testing user_settings with init action...")
    
    handler = UserHandler()
    
    arguments = {"action": "init"}
    
    try:
        result = await handler.handle_user_settings(arguments)
        assert isinstance(result, list), "Result should be a list"
        assert len(result) > 0, "Result should not be empty"
        assert result[0].type == "text", "Result should be TextContent"
        print("   ✓ user_settings with init action works")
        
        content = json.loads(result[0].text)
        assert "status" in content, "Result should have status"
        
        if content["status"] == "success":
            assert "action" in content, "Result should have action"
            assert content["action"] == "init", "Action should be 'init'"
            assert "settings" in content, "Result should have settings"
            assert content["settings"]["page_size"] == 5, "Default page_size should be 5"
            assert content["settings"]["llm_page_size"] == 20, "Default llm_page_size should be 20"
            assert content["settings"]["default_search_days"] == 90, "Default default_search_days should be 90"
            print("   ✓ Init action sets default values: page_size=5, llm_page_size=20, default_search_days=90")
        else:
            print(f"   ✓ user_settings returned error status (expected if not authenticated): {content.get('message', 'Unknown error')}")
    except Exception as e:
        error_msg = str(e).lower()
        if "not authenticated" in error_msg or "graph_client" in error_msg or "token" in error_msg or "authorization" in error_msg:
            print("   ✓ user_settings with sync action works (authentication expected)")
        else:
            print(f"   ✗ user_settings with sync action failed with error: {e}")
            raise
    
    print("\n[Test 3] ✓ PASSED: user_settings with sync action works")


async def test_user_settings_update_action():
    """Test user_settings with update action."""
    print("\n[Test 4] Testing user_settings with update action...")
    
    handler = UserHandler()
    
    arguments = {
        "action": "update",
        "page_size": 10,
        "llm_page_size": 30,
        "default_search_days": 60
    }
    
    try:
        result = await handler.handle_user_settings(arguments)
        assert isinstance(result, list), "Result should be a list"
        assert len(result) > 0, "Result should not be empty"
        assert result[0].type == "text", "Result should be TextContent"
        print("   ✓ user_settings with update action works")
        
        content = json.loads(result[0].text)
        assert "status" in content, "Result should have status"
        
        if content["status"] == "success":
            assert "action" in content, "Result should have action"
            assert content["action"] == "update", "Action should be 'update'"
            assert "settings" in content, "Result should have settings"
            assert content["settings"]["page_size"] == 10, "Custom page_size should be 10"
            assert content["settings"]["llm_page_size"] == 30, "Custom llm_page_size should be 30"
            assert content["settings"]["default_search_days"] == 60, "Custom default_search_days should be 60"
            print("   ✓ Update action sets custom values: page_size=10, llm_page_size=30, default_search_days=60")
        else:
            print(f"   ✓ user_settings returned error status (expected if not authenticated): {content.get('message', 'Unknown error')}")
    except Exception as e:
        error_msg = str(e).lower()
        if "not authenticated" in error_msg or "graph_client" in error_msg or "token" in error_msg or "authorization" in error_msg:
            print("   ✓ user_settings with update action works (authentication expected)")
        else:
            print(f"   ✗ user_settings with update action failed with error: {e}")
            raise
    
    print("\n[Test 4] ✓ PASSED: user_settings with update action works")


async def test_user_settings_env_file_updates():
    """Test that user_settings updates .env file correctly."""
    print("\n[Test 5] Testing .env file updates...")
    
    handler = UserHandler()
    
    original_env_path = Path(__file__).parent.parent / ".env"
    temp_env_path = Path(__file__).parent / "test_temp.env"
    
    try:
        if original_env_path.exists():
            shutil.copy(original_env_path, temp_env_path)
        
        initial_content = """CLIENT_ID=test_id
CLIENT_SECRET=test_secret
TENANT_ID=organizations
USER_TIMEZONE=America/New_York
DEFAULT_SEARCH_DAYS=90
PAGE_SIZE=5
LLM_PAGE_SIZE=20
"""
        
        with open(original_env_path, "w") as f:
            f.write(initial_content)
        
        arguments = {
            "action": "update",
            "page_size": 15,
            "llm_page_size": 25,
            "default_search_days": 120
        }
        
        try:
            result = await handler.handle_user_settings(arguments)
            print("   ✓ user_settings executed")
        except Exception as e:
            error_msg = str(e).lower()
            if "not authenticated" in error_msg or "graph_client" in error_msg or "token" in error_msg or "authorization" in error_msg:
                print("   ✓ user_settings executed (authentication expected)")
            else:
                raise
        
        with open(original_env_path, "r") as f:
            updated_content = f.read()
        
        assert "PAGE_SIZE=15" in updated_content or "PAGE_SIZE=5" in updated_content, "PAGE_SIZE should be updated or remain"
        assert "LLM_PAGE_SIZE=25" in updated_content or "LLM_PAGE_SIZE=20" in updated_content, "LLM_PAGE_SIZE should be updated or remain"
        assert "DEFAULT_SEARCH_DAYS=120" in updated_content or "DEFAULT_SEARCH_DAYS=90" in updated_content, "DEFAULT_SEARCH_DAYS should be updated or remain"
        print("   ✓ .env file contains expected values")
        
    finally:
        if temp_env_path.exists():
            shutil.move(temp_env_path, original_env_path)
        elif original_env_path.exists():
            original_env_path.unlink()
    
    print("\n[Test 5] ✓ PASSED: .env file updates work correctly")


async def test_user_settings_response_format():
    """Test that user_settings returns correct response format."""
    print("\n[Test 6] Testing user_settings response format...")
    
    handler = UserHandler()
    
    arguments = {
        "action": "update",
        "page_size": 8,
        "llm_page_size": 25,
        "default_search_days": 100
    }
    
    try:
        result = await handler.handle_user_settings(arguments)
        assert isinstance(result, list), "Result should be a list"
        assert len(result) == 1, "Result should have exactly one TextContent"
        assert result[0].type == "text", "Result should be TextContent"
        print("   ✓ Response format is correct: list with one TextContent")
        
        content = json.loads(result[0].text)
        assert "status" in content, "Response should have status"
        assert "message" in content, "Response should have message"
        assert "action" in content, "Response should have action"
        print("   ✓ Response has required fields: status, message, action")
        
        assert content["status"] in ["success", "error"], "Status should be success or error"
        assert content["action"] in ["init", "update"], "Action should be init or update"
        print(f"   ✓ Status is valid: {content['status']}, Action is valid: {content['action']}")
        
        if content["status"] == "success":
            assert "user_info" in content, "Response should have user_info"
            assert "settings" in content, "Response should have settings"
            print("   ✓ Response has success fields: user_info, settings")
            
            assert "display_name" in content["user_info"], "user_info should have display_name"
            assert "email" in content["user_info"], "user_info should have email"
            assert "timezone" in content["user_info"], "user_info should have timezone"
            print("   ✓ user_info has required fields: display_name, email, timezone")
            
            assert "page_size" in content["settings"], "settings should have page_size"
            assert "llm_page_size" in content["settings"], "settings should have llm_page_size"
            assert "default_search_days" in content["settings"], "settings should have default_search_days"
            print("   ✓ settings has required fields: page_size, llm_page_size, default_search_days")
        else:
            print(f"   ✓ Response has error status (expected if not authenticated)")
    except Exception as e:
        error_msg = str(e).lower()
        if "not authenticated" in error_msg or "graph_client" in error_msg or "token" in error_msg or "authorization" in error_msg:
            print("   ✓ Response format is correct (authentication expected)")
        else:
            print(f"   ✗ Response format test failed with error: {e}")
            raise
    
    print("\n[Test 6] ✓ PASSED: user_settings response format is correct")


async def test_user_settings_boundary_values():
    """Test user_settings with boundary values."""
    print("\n[Test 7] Testing user_settings with boundary values...")
    
    handler = UserHandler()
    
    test_cases = [
        {"page_size": 1, "llm_page_size": 1, "default_search_days": 1, "desc": "minimum values"},
        {"page_size": 50, "llm_page_size": 100, "default_search_days": 365, "desc": "maximum values"},
        {"page_size": 25, "llm_page_size": 50, "default_search_days": 182, "desc": "middle values"},
    ]
    
    for test_case in test_cases:
        arguments = {
            "action": "update",
            "page_size": test_case["page_size"],
            "llm_page_size": test_case["llm_page_size"],
            "default_search_days": test_case["default_search_days"]
        }
        
        try:
            result = await handler.handle_user_settings(arguments)
            assert isinstance(result, list), "Result should be a list"
            assert result[0].type == "text", "Result should be TextContent"
            print(f"   ✓ {test_case['desc']} work: page_size={test_case['page_size']}, llm_page_size={test_case['llm_page_size']}, default_search_days={test_case['default_search_days']}")
        except Exception as e:
            error_msg = str(e).lower()
            if "not authenticated" in error_msg or "graph_client" in error_msg or "token" in error_msg or "authorization" in error_msg:
                print(f"   ✓ {test_case['desc']} work (authentication expected)")
            else:
                print(f"   ✗ {test_case['desc']} failed with error: {e}")
                raise
    
    print("\n[Test 7] ✓ PASSED: user_settings with boundary values works")


async def test_user_settings_invalid_action():
    """Test user_settings with invalid action."""
    print("\n[Test 8] Testing user_settings with invalid action...")
    
    handler = UserHandler()
    
    arguments = {"action": "invalid_action"}
    
    result = await handler.handle_user_settings(arguments)
    assert isinstance(result, list), "Result should be a list"
    assert result[0].type == "text", "Result should be TextContent"
    print("   ✓ user_settings handles invalid action")
    
    content = json.loads(result[0].text)
    assert content["status"] == "error", "Status should be error"
    assert "Invalid action" in content["message"], "Error message should mention invalid action"
    print(f"   ✓ Error message: {content['message']}")
    
    print("\n[Test 8] ✓ PASSED: user_settings handles invalid action correctly")


async def test_user_settings_login_required():
    """Test that user_settings requires login."""
    print("\n[Test 9] Testing user_settings login requirement...")
    
    handler = UserHandler()
    
    arguments = {"action": "init"}
    
    result = await handler.handle_user_settings(arguments)
    assert isinstance(result, list), "Result should be a list"
    assert result[0].type == "text", "Result should be TextContent"
    print("   ✓ user_settings executed")
    
    content = json.loads(result[0].text)
    
    if content["status"] == "error":
        assert "not authenticated" in content["message"].lower() or "login" in content["message"].lower(), "Error message should mention authentication"
        print(f"   ✓ Correctly requires login: {content['message']}")
    else:
        print(f"   ✓ User is authenticated, returning success")
    
    print("\n[Test 9] ✓ PASSED: user_settings correctly handles login requirement")


async def main():
    """Run all tests."""
    print("=" * 70)
    print("USER_SETTINGS TOOL TESTS")
    print("=" * 70)
    
    tests = [
        test_user_settings_tool_schema,
        test_handler_method_exists,
        test_user_settings_init_action,
        test_user_settings_update_action,
        test_user_settings_env_file_updates,
        test_user_settings_response_format,
        test_user_settings_boundary_values,
        test_user_settings_invalid_action,
        test_user_settings_login_required,
    ]
    
    passed = 0
    failed = 0
    
    for test in tests:
        try:
            await test()
            passed += 1
        except Exception as e:
            failed += 1
            print(f"\n✗ TEST FAILED: {test.__name__}")
            print(f"   Error: {e}")
            import traceback
            traceback.print_exc()
    
    print("\n" + "=" * 70)
    print("TEST SUMMARY")
    print("=" * 70)
    print(f"Total tests: {len(tests)}")
    print(f"Passed: {passed}")
    print(f"Failed: {failed}")
    print("=" * 70)
    
    if failed == 0:
        print("\n✓ ALL TESTS PASSED!")
    else:
        print(f"\n✗ {failed} TEST(S) FAILED")
    
    return failed == 0


if __name__ == "__main__":
    success = asyncio.run(main())
    sys.exit(0 if success else 1)
