"""
Example: Complete authentication workflow for Microsoft Graph MCP Server.

This example demonstrates the simplified authentication flow with 4 actions:
- start: Begin login flow
- complete: Finish login after browser auth
- refresh: Manually refresh token
- logout: Clear tokens
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

    start_result = {"tool": "auth", "arguments": {"action": "start"}}

    print(f"Calling: {json.dumps(start_result, indent=2)}")
    print()

    # Expected response:
    expected_start_response = {
        "status": "pending",
        "message": "Visit https://microsoft.com/devicelogin and enter code ABC123",
        "verification_uri": "https://microsoft.com/devicelogin",
        "user_code": "ABC123",
        "expires_in": 900,
    }

    print("Expected Response:")
    print(json.dumps(expected_start_response, indent=2))
    print()

    # User completes authentication in browser
    print("ACTION REQUIRED: User must now:")
    print(f"  1. Visit: {expected_start_response['verification_uri']}")
    print(f"  2. Enter code: {expected_start_response['user_code']}")
    print()

    input("Press Enter after completing authentication in browser...")

    # ============================================
    # STEP 2: Complete login
    # ============================================
    print("STEP 2: Completing login process...")
    print("-" * 70)
    print("⚠️  IMPORTANT: You MUST call complete after browser auth!")
    print()

    complete_result = {
        "tool": "auth",
        "arguments": {
            "action": "complete",
        },
    }

    print(f"Calling: {json.dumps(complete_result, indent=2)}")
    print()

    # Expected response:
    expected_complete_response = {
        "status": "success",
        "authenticated": True,
        "message": "Successfully authenticated with Microsoft Graph. Token expires in 59m.",
        "time_remaining": {"seconds": 3540, "display": "59m"},
    }

    print("Expected Response:")
    print(json.dumps(expected_complete_response, indent=2))
    print()

    # ============================================
    # STEP 3: Use authenticated tools
    # ============================================
    print("STEP 3: Now you can use authenticated tools...")
    print("-" * 70)

    search_emails_result = {
        "tool": "search_emails",
        "arguments": {"days": 7, "folder": "Inbox"},
    }

    print(f"Example - Calling: {json.dumps(search_emails_result, indent=2)}")
    print()
    print("✅ Authentication successful! You can now use any tool.")
    print()
    print("NOTE: Access tokens auto-refresh when expired. No manual refresh needed!")

    # ============================================
    # STEP 4: Manual refresh (optional)
    # ============================================
    print()
    print("STEP 4: Manual refresh (optional)...")
    print("-" * 70)
    print("Access tokens auto-refresh, but you can manually refresh:")
    print()

    refresh_result = {"tool": "auth", "arguments": {"action": "refresh"}}

    print(f"Calling: {json.dumps(refresh_result, indent=2)}")
    print()

    expected_refresh_response = {
        "status": "refreshed",
        "authenticated": True,
        "message": "Token refreshed successfully. Expires in 1h 0m.",
        "time_remaining": {"seconds": 3600, "display": "1h 0m"},
    }

    print("Expected Response:")
    print(json.dumps(expected_refresh_response, indent=2))
    print()

    # ============================================
    # STEP 5: Logout (when done)
    # ============================================
    print("STEP 5: Logout (when done)...")
    print("-" * 70)

    logout_result = {"tool": "auth", "arguments": {"action": "logout"}}

    print(f"Calling: {json.dumps(logout_result, indent=2)}")
    print()

    expected_logout_response = {
        "status": "logged_out",
        "authenticated": False,
        "message": "Successfully logged out.",
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
    # MISTAKE 1: Forgetting complete
    # ============================================
    print("❌ MISTAKE 1: Forgetting to call complete")
    print("-" * 70)

    print("Incorrect workflow:")
    print(
        json.dumps(
            {
                "step 1": "auth(action='start')",
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
                "step 1": "auth(action='start')",
                "step 2": "User completes browser auth",
                "step 3": "auth(action='complete')  # ← CRITICAL STEP!",
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
    print("  1. auth(action='start')")
    print("  2. auth(action='complete')  # ← Too fast! User hasn't completed auth")
    print()

    print("Correct:")
    print("  1. auth(action='start')")
    print("  2. Wait for user to visit URL and enter code")
    print("  3. auth(action='complete')  # ← Wait for user!")
    print()

    print("=" * 70)


if __name__ == "__main__":
    asyncio.run(auth_workflow_example())
    print()
    asyncio.run(common_mistakes_example())
