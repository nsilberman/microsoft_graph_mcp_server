"""
Example: Complete authentication workflow for Microsoft Graph MCP Server.

This example demonstrates the complete authentication flow including
login, completion, status checking, token extension, and logout.
"""

import asyncio
import json


async def auth_workflow_example():
    """Demonstrate complete authentication workflow."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - AUTHENTICATION WORKFLOW EXAMPLE")
    print("=" * 70)
    print()

    # ============================================
    # STEP 1: Start login process
    # ============================================
    print("STEP 1: Starting login process...")
    print("-" * 70)

    login_result = {"tool": "auth", "arguments": {"action": "login"}}

    print(f"Calling: {json.dumps(login_result, indent=2)}")
    print()

    # In real MCP client:
    # result = await mcp_client.call_tool("auth", {"action": "login"})

    # Expected response:
    expected_login_response = {
        "success": True,
        "status": "pending",
        "message": "Visit https://microsoft.com/devicelogin and enter code ABC123",
        "verification_url": "https://microsoft.com/devicelogin",
        "user_code": "ABC123",
        "expires_in": 900,
        "interval": 5,
    }

    print("Expected Response:")
    print(json.dumps(expected_login_response, indent=2))
    print()

    # User completes authentication in browser
    print("ACTION REQUIRED: User must now:")
    print(f"  1. Visit: {expected_login_response['verification_url']}")
    print(f"  2. Enter code: {expected_login_response['user_code']}")
    print()

    input("Press Enter after completing authentication in browser...")

    # ============================================
    # STEP 2: Complete login (CRITICAL!)
    # ============================================
    print("STEP 2: Completing login process...")
    print("-" * 70)
    print("⚠️  IMPORTANT: You MUST call complete_login!")
    print("    Without this, authentication will fail and you can't use other tools.")
    print()

    complete_login_result = {
        "tool": "auth",
        "arguments": {
            "action": "complete_login",
        },
    }

    print(f"Calling: {json.dumps(complete_login_result, indent=2)}")
    print()

    # Expected response:
    expected_complete_response = {
        "success": True,
        "status": "authenticated",
        "message": "Authentication completed successfully",
        "authenticated": True,
        "token_expiry": "1704777600",
        "user_info": {
            "id": "...",
            "displayName": "John Doe",
            "mail": "john.doe@example.com",
            "userPrincipalName": "john.doe@example.com",
        },
    }

    print("Expected Response:")
    print(json.dumps(expected_complete_response, indent=2))
    print()

    # ============================================
    # STEP 3: Check authentication status (optional)
    # ============================================
    print("STEP 3: Checking authentication status (optional)...")
    print("-" * 70)

    check_status_result = {"tool": "auth", "arguments": {"action": "check_status"}}

    print(f"Calling: {json.dumps(check_status_result, indent=2)}")
    print()

    # Expected response:
    expected_status_response = {
        "success": True,
        "authenticated": True,
        "token_expiry": "1704777600",
        "user_info": {"displayName": "John Doe", "mail": "john.doe@example.com"},
    }

    print("Expected Response:")
    print(json.dumps(expected_status_response, indent=2))
    print()

    # ============================================
    # STEP 4: Use authenticated tools (example)
    # ============================================
    print("STEP 4: Now you can use authenticated tools...")
    print("-" * 70)

    search_emails_result = {
        "tool": "search_emails",
        "arguments": {"days": 7, "folder": "Inbox"},
    }

    print(f"Example - Calling: {json.dumps(search_emails_result, indent=2)}")
    print()
    print("✅ Authentication successful! You can now use any tool.")

    # ============================================
    # STEP 5: Extend token (when expires)
    # ============================================
    print()
    print("STEP 5: Extend token (optional, when token expires)...")
    print("-" * 70)
    print("Access tokens expire after 1 hour.")
    print("You can extend the token without user login:")
    print()

    extend_token_result = {"tool": "auth", "arguments": {"action": "extend_token"}}

    print(f"Calling: {json.dumps(extend_token_result, indent=2)}")
    print()

    # Expected response:
    expected_extend_response = {
        "success": True,
        "message": "Token extended successfully",
        "token_expiry": "1704781200",  # New expiry (1 hour later)
    }

    print("Expected Response:")
    print(json.dumps(expected_extend_response, indent=2))
    print()

    # ============================================
    # STEP 6: Logout (when done)
    # ============================================
    print("STEP 6: Logout (when done)...")
    print("-" * 70)

    logout_result = {"tool": "auth", "arguments": {"action": "logout"}}

    print(f"Calling: {json.dumps(logout_result, indent=2)}")
    print()

    # Expected response:
    expected_logout_response = {
        "success": True,
        "message": "Logged out successfully",
        "authenticated": False,
    }

    print("Expected Response:")
    print(json.dumps(expected_logout_response, indent=2))
    print()

    print("=" * 70)
    print("WORKFLOW COMPLETE")
    print("=" * 70)


async def common_mistakes_example():
    """Show common mistakes and how to avoid them."""

    print()
    print("=" * 70)
    print("COMMON MISTAKES - HOW TO AVOID")
    print("=" * 70)
    print()

    # ============================================
    # MISTAKE 1: Forgetting complete_login
    # ============================================
    print("❌ MISTAKE 1: Forgetting to call complete_login")
    print("-" * 70)

    print("Incorrect workflow:")
    print(
        json.dumps(
            {
                "step 1": "auth(action='login')",
                "step 2": "User completes browser auth",
                "step 3": "search_emails(days=7)  # ← FAILS! Not authenticated",
            },
            indent=2,
        )
    )
    print()

    print("Correct workflow:")
    print(
        json.dumps(
            {
                "step 1": "auth(action='login')",
                "step 2": "User completes browser auth",
                "step 3": "auth(action='complete_login')  # ← CRITICAL STEP!",
                "step 4": "search_emails(days=7)  # Now works!",
            },
            indent=2,
        )
    )
    print()

    # ============================================
    # MISTAKE 2: Not waiting for user completion
    # ============================================
    print("❌ MISTAKE 2: Not waiting for user to complete browser auth")
    print("-" * 70)

    print("Incorrect:")
    print("  1. auth(action='login')")
    print(
        "  2. auth(action='complete_login')  # ← Too fast! User hasn't completed auth"
    )
    print()

    print("Correct:")
    print("  1. auth(action='login')")
    print("  2. Wait for user to visit URL and enter code")
    print("  3. auth(action='complete_login')  # ← Wait for user!")
    print()

    print("=" * 70)


if __name__ == "__main__":
    asyncio.run(auth_workflow_example())
    print()
    asyncio.run(common_mistakes_example())
