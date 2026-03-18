"""
Example: Complete email workflow for Microsoft Graph MCP Server.

This example demonstrates complete email operations including
search, browse, read content, and actions (reply, forward, move, delete).
"""

import asyncio
import json


async def email_search_and_browse_workflow():
    """Demonstrate email search and browse workflow."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - EMAIL SEARCH & BROWSE WORKFLOW")
    print("=" * 70)
    print()

    # ============================================
    # STEP 1: Search emails
    # ============================================
    print("STEP 1: Searching emails (loads into cache)...")
    print("-" * 70)

    search_result = {
        "tool": "search_emails",
        "arguments": {
            "query": "meeting",
            "search_type": "subject",
            "folder": "Inbox",
            "days": 7,
        },
    }

    print(f"Calling: {json.dumps(search_result, indent=2)}")
    print()

    # Expected response:
    expected_search_response = {
        "success": True,
        "emails": [
            {
                "number": 1,
                "subject": "Project Review Meeting",
                "from": "manager@company.com",
                "to": "team@company.com",
                "received_date": "2024-01-08 14:30",
                "folder": "Inbox",
                "is_read": False,
                "has_attachments": True,
            },
            # ... more emails
        ],
        "count": 50,
        "date_range": "Jan 1 - Jan 7, 2024",
        "filter_date_range": "last 7 days",
        "timezone": "America/New_York",
    }

    print("Expected Response:")
    print(json.dumps(expected_search_response, indent=2))
    print()

    # ============================================
    # STEP 2: Browse cache
    # ============================================
    print("STEP 2: Browsing email cache (pagination)...")
    print("-" * 70)

    browse_result = {
        "tool": "browse_email_cache",
        "arguments": {
            "page_number": 1,
            "mode": "llm",  # LLM mode: 20 emails per page
        },
    }

    print(f"Calling: {json.dumps(browse_result, indent=2)}")
    print()

    # Expected response:
    expected_browse_response = {
        "success": True,
        "current_page": 1,
        "total_pages": 3,
        "count": 20,
        "total_count": 50,
        "emails": [
            {
                "number": 1,  # ← Use this number, not array index!
                "subject": "Project Review Meeting",
                "from": "manager@company.com",
                "to": "team@company.com",
                "received_date": "2024-01-08 14:30",
                "folder": "Inbox",
                "is_read": False,
                "has_attachments": True,
            },
            # ... emails 2-20 on this page
        ],
        "page_size": 20,
        "mode": "llm",
        "timezone": "America/New_York",
    }

    print("Expected Response:")
    print(json.dumps(expected_browse_response, indent=2))
    print()
    print("💡 IMPORTANT: Use the 'number' field, not array index!")
    print("   If browse returns emails[0].number=1, use cache_number=1")

    # ============================================
    # STEP 3: Get full email content
    # ============================================
    print("STEP 3: Getting full email content...")
    print("-" * 70)

    content_result = {
        "tool": "get_email_content",
        "arguments": {
            "cache_number": 1,  # ← Use number from browse results!
            "return_html": True,  # Get full HTML content with attachments
        },
    }

    print(f"Calling: {json.dumps(content_result, indent=2)}")
    print()

    # Expected response:
    expected_content_response = {
        "success": True,
        "subject": "Project Review Meeting",
        "from": {"email": "manager@company.com", "name": "John Manager"},
        "to": [{"email": "team@company.com", "name": "Team Members"}],
        "cc": [],
        "bcc": [],
        "body": "<html>...</html>",
        "attachments": [
            {
                "name": "meeting_agenda.pdf",
                "size": 12345,
                "content_type": "application/pdf",
            }
        ],
        "sent_date": "2024-01-08 14:30",
        "received_date": "2024-01-08 14:30",
    }

    print("Expected Response:")
    print(json.dumps(expected_content_response, indent=2))
    print()


async def email_actions_workflow():
    """Demonstrate email action operations."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - EMAIL ACTIONS WORKFLOW")
    print("=" * 70)
    print()

    # Assume we've already browsed emails and have cache_number=1

    # ============================================
    # ACTION 1: Reply to email
    # ============================================
    print("ACTION 1: Replying to email...")
    print("-" * 70)

    reply_result = {
        "tool": "send_email",
        "arguments": {
            "action": "reply",
            "cache_number": 1,
            "htmlbody": "<p>Thank you for the meeting agenda!</p><p>I'll be prepared.</p>",
        },
    }

    print(f"Calling: {json.dumps(reply_result, indent=2)}")
    print()

    expected_reply_response = {
        "success": True,
        "message": "Reply sent successfully",
        "sent_count": 1,
        "failed_count": 0,
        "recipients": ["manager@company.com"],
    }

    print("Expected Response:")
    print(json.dumps(expected_reply_response, indent=2))
    print()

    # ============================================
    # ACTION 2: Forward email
    # ============================================
    print("ACTION 2: Forwarding email...")
    print("-" * 70)

    forward_result = {
        "tool": "send_email",
        "arguments": {
            "action": "forward",
            "cache_number": 1,
            "to": ["colleague@company.com"],
            "htmlbody": "<p>FYI - See agenda below.</p>",
        },
    }

    print(f"Calling: {json.dumps(forward_result, indent=2)}")
    print()

    expected_forward_response = {
        "success": True,
        "message": "Forward sent successfully",
        "sent_count": 1,
        "failed_count": 0,
        "recipients": ["colleague@company.com"],
    }

    print("Expected Response:")
    print(json.dumps(expected_forward_response, indent=2))
    print()

    # ============================================
    # ACTION 3: Move email
    # ============================================
    print("ACTION 3: Moving email to folder...")
    print("-" * 70)

    move_result = {
        "tool": "manage_emails",
        "arguments": {
            "action": "move_single",
            "cache_number": 1,
            "destination_folder": "Archive/2024",
        },
    }

    print(f"Calling: {json.dumps(move_result, indent=2)}")
    print()

    expected_move_response = {
        "success": True,
        "message": "Email moved successfully",
        "moved_count": 1,
        "failed_count": 0,
        "errors": [],
    }

    print("Expected Response:")
    print(json.dumps(expected_move_response, indent=2))
    print()

    # ============================================
    # ACTION 4: Flag email
    # ============================================
    print("ACTION 4: Flagging email for follow-up...")
    print("-" * 70)

    flag_result = {
        "tool": "manage_emails",
        "arguments": {
            "action": "flag_single",
            "cache_number": 1,
            "flag_status": "flagged",
        },
    }

    print(f"Calling: {json.dumps(flag_result, indent=2)}")
    print()

    expected_flag_response = {"success": True, "message": "Email flagged successfully"}

    print("Expected Response:")
    print(json.dumps(expected_flag_response, indent=2))
    print()

    # ============================================
    # ACTION 5: Delete email
    # ============================================
    print("ACTION 5: Deleting email...")
    print("-" * 70)

    delete_result = {
        "tool": "manage_emails",
        "arguments": {"action": "delete_single", "cache_number": 1},
    }

    print(f"Calling: {json.dumps(delete_result, indent=2)}")
    print()

    expected_delete_response = {
        "success": True,
        "message": "Email deleted successfully",
        "deleted_count": 1,
        "failed_count": 0,
        "errors": [],
    }

    print("Expected Response:")
    print(json.dumps(expected_delete_response, indent=2))
    print()


async def send_new_email_with_bcc_csv():
    """Demonstrate sending new email with BCC CSV file."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - SEND NEW EMAIL WITH BCC CSV")
    print("=" * 70)
    print()

    # ============================================
    # Send new email with BCC CSV
    # ============================================
    print("SENDING NEW EMAIL WITH BCC CSV FILE...")
    print("-" * 70)

    send_result = {
        "tool": "send_email",
        "arguments": {
            "action": "send_new",
            "to": ["manager@company.com"],
            "subject": "Monthly Newsletter - January 2024",
            "htmlbody": """
                <p>Dear Team,</p>
                <p>Please find below the monthly newsletter for January 2024.</p>
                <p>Best regards,<br>Marketing Team</p>
            """,
            "bcc_csv_file": "/path/to/recipients.csv",  # Large recipient list
        },
    }

    print(f"Calling: {json.dumps(send_result, indent=2)}")
    print()

    # Expected response (auto-batches to 500 recipients per email):
    expected_send_response = {
        "success": True,
        "message": "Email sent successfully to 1250 recipients (3 batches)",
        "sent_count": 3,
        "failed_count": 0,
        "recipients": [
            "manager@company.com",
            # BCC recipients from CSV (auto-batched)
        ],
    }

    print("Expected Response:")
    print(json.dumps(expected_send_response, indent=2))
    print()
    print("💡 TIP: BCC CSV files automatically batch to 500 recipients per email")
    print("   This is the preferred method for large recipient lists")


async def folder_management_workflow():
    """Demonstrate folder management operations."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - FOLDER MANAGEMENT WORKFLOW")
    print("=" * 70)
    print()

    # ============================================
    # STEP 1: List all folders
    # ============================================
    print("STEP 1: Listing all mail folders...")
    print("-" * 70)

    list_result = {"tool": "manage_mail_folder", "arguments": {"action": "list"}}

    print(f"Calling: {json.dumps(list_result, indent=2)}")
    print()

    expected_list_response = {
        "success": True,
        "folders": [
            {
                "displayName": "Inbox",
                "path": "Inbox",
                "totalItemCount": 150,
                "unreadItemCount": 45,
                "childFolderCount": 3,
            },
            {
                "displayName": "Archive",
                "path": "Archive",
                "totalItemCount": 2000,
                "unreadItemCount": 0,
                "childFolderCount": 5,
            },
        ],
    }

    print("Expected Response:")
    print(json.dumps(expected_list_response, indent=2))
    print()

    # ============================================
    # STEP 2: Create new folder
    # ============================================
    print("STEP 2: Creating new folder...")
    print("-" * 70)

    create_result = {
        "tool": "manage_mail_folder",
        "arguments": {
            "action": "create",
            "folder_name": "Projects/2024",
            "parent_folder": "Archive",
        },
    }

    print(f"Calling: {json.dumps(create_result, indent=2)}")
    print()

    expected_create_response = {
        "success": True,
        "message": "Folder created successfully",
        "path": "Archive/Projects/2024",
        "displayName": "2024",
        "totalItemCount": 0,
        "unreadItemCount": 0,
        "childFolderCount": 0,
    }

    print("Expected Response:")
    print(json.dumps(expected_create_response, indent=2))
    print()

    print("=" * 70)
    print("WORKFLOW COMPLETE")
    print("=" * 70)


async def common_mistakes_example():
    """Show common email mistakes and how to avoid them."""

    print()
    print("=" * 70)
    print("COMMON MISTAKES - EMAIL WORKFLOWS")
    print("=" * 70)
    print()

    # ============================================
    # MISTAKE 1: Using array index instead of cache number
    # ============================================
    print("❌ MISTAKE 1: Using array index instead of cache number")
    print("-" * 70)

    print("Incorrect:")
    print("  browse_email_cache(page_number=1, mode='llm')")
    print("  # Returns emails[0].number=21, emails[1].number=22, ...")
    print("  get_email_content(cache_number=0)  # Wrong! Array index")
    print()

    print("Correct:")
    print("  browse_email_cache(page_number=1, mode='llm')")
    print("  # Returns emails[0].number=21, emails[1].number=22, ...")
    print("  get_email_content(cache_number=21)  # Correct! Use number field")
    print()

    # ============================================
    # MISTAKE 2: Searching in wrong tool
    # ============================================
    print("❌ MISTAKE 2: Searching emails in wrong tool")
    print("-" * 70)

    print("Incorrect:")
    print("  search_contacts(query='meeting@example.com')  # Wrong! Searches people")
    print()

    print("Correct:")
    print(
        "  search_emails(query='meeting@example.com', search_type='sender')  # Correct!"
    )
    print()

    # ============================================
    # MISTAKE 3: Not browsing first
    # ============================================
    print("❌ MISTAKE 3: Not browsing cache before using cache_number")
    print("-" * 70)

    print("Incorrect:")
    print("  get_email_content(cache_number=5)  # Cache is empty!")
    print()

    print("Correct:")
    print("  search_emails(days=7)  # Load emails into cache first")
    print("  browse_email_cache(page_number=1, mode='llm')  # Browse cache")
    print("  get_email_content(cache_number=5)  # Now works!")
    print()

    print("=" * 70)


if __name__ == "__main__":
    asyncio.run(email_search_and_browse_workflow())
    print()
    asyncio.run(email_actions_workflow())
    print()
    asyncio.run(send_new_email_with_bcc_csv())
    print()
    asyncio.run(folder_management_workflow())
    print()
    asyncio.run(common_mistakes_example())
