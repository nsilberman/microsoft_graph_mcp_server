"""
Example: Complete template workflow for Microsoft Graph MCP Server.

This example demonstrates template management including
creation, listing, viewing (simple text and full HTML), editing, and sending.
"""

import asyncio
import json


async def template_lifecycle_workflow():
    """Demonstrate complete template lifecycle workflow."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - TEMPLATE LIFECYCLE WORKFLOW")
    print("=" * 70)
    print()

    # ============================================
    # STEP 1: Search for email to use as template
    # ============================================
    print("STEP 1: Searching for email to create template from...")
    print("-" * 70)

    search_result = {
        "tool": "search_emails",
        "arguments": {"query": "newsletter", "folder": "Inbox", "days": 30},
    }

    print(f"Calling: {json.dumps(search_result, indent=2)}")
    print()

    expected_search_response = {
        "success": True,
        "emails": [
            {
                "number": 5,
                "subject": "Monthly Newsletter - December 2023",
                "from": "marketing@company.com",
            }
        ],
        "count": 1,
    }

    print("Expected Response:")
    print(json.dumps(expected_search_response, indent=2))
    print()

    # ============================================
    # STEP 2: Browse email cache
    # ============================================
    print("STEP 2: Browsing email cache...")
    print("-" * 70)

    browse_result = {
        "tool": "browse_email_cache",
        "arguments": {"page_number": 1, "mode": "llm"},
    }

    print(f"Calling: {json.dumps(browse_result, indent=2)}")
    print()

    expected_browse_response = {
        "success": True,
        "emails": [
            {
                "number": 5,  # ← Use this number!
                "subject": "Monthly Newsletter - December 2023",
                "from": "marketing@company.com",
            }
        ],
        "total_count": 1,
    }

    print("Expected Response:")
    print(json.dumps(expected_browse_response, indent=2))
    print()

    # ============================================
    # STEP 3: Create template from email
    # ============================================
    print("STEP 3: Creating template from email...")
    print("-" * 70)

    create_template_result = {
        "tool": "manage_templates",
        "arguments": {
            "action": "create_from_email",
            "cache_number": 5,  # ← Use number from browse results!
        },
    }

    print(f"Calling: {json.dumps(create_template_result, indent=2)}")
    print()

    expected_create_response = {
        "success": True,
        "message": "Template created successfully",
        "template_number": 1,
        "subject": "Template: Monthly Newsletter - December 2023",
    }

    print("Expected Response:")
    print(json.dumps(expected_create_response, indent=2))
    print()

    # ============================================
    # STEP 4: List templates
    # ============================================
    print("STEP 4: Listing templates...")
    print("-" * 70)

    list_templates_result = {
        "tool": "manage_templates",
        "arguments": {"action": "list", "page_number": 1},
    }

    print(f"Calling: {json.dumps(list_templates_result, indent=2)}")
    print()

    expected_list_response = {
        "success": True,
        "templates": [
            {
                "number": 1,
                "subject": "Template: Monthly Newsletter - December 2023",
                "to": ["recipients@company.com"],
            }
        ],
        "total_count": 1,
    }

    print("Expected Response:")
    print(json.dumps(expected_list_response, indent=2))
    print()


async def template_editing_workflow():
    """Demonstrate template viewing and editing workflow."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - TEMPLATE EDITING WORKFLOW")
    print("=" * 70)
    print()

    # Assume template_number=1 exists from previous workflow

    # ============================================
    # STEP 1: View template as simple text (user view)
    # ============================================
    print("STEP 1: Viewing template as simple text (for user review)...")
    print("-" * 70)

    get_simple_result = {
        "tool": "manage_templates",
        "arguments": {
            "action": "get",
            "template_number": 1,
            "return_html": False,  # ← Simple text only
        },
    }

    print(f"Calling: {json.dumps(get_simple_result, indent=2)}")
    print()

    expected_simple_response = {
        "success": True,
        "subject": "Template: Monthly Newsletter - December 2023",
        "body": """
Dear Team,

Here is our monthly newsletter for December 2023.

Highlights:
- New product launches
- Company achievements
- Upcoming events

Best regards,
Marketing Team
""",
        "to": ["recipients@company.com"],
    }

    print("Expected Response (Simple Text - User View):")
    print(json.dumps(expected_simple_response, indent=2))
    print()
    print("💡 User reviews this and provides update instructions to LLM...")

    input("Press Enter after user provides update instructions...")

    # ============================================
    # STEP 2: LLM views template as full HTML
    # ============================================
    print("STEP 2: LLM retrieving full HTML (for editing)...")
    print("-" * 70)

    get_html_result = {
        "tool": "manage_templates",
        "arguments": {
            "action": "get",
            "template_number": 1,
            "return_html": True,  # ← Full HTML!
        },
    }

    print(f"Calling: {json.dumps(get_html_result, indent=2)}")
    print()

    expected_html_response = {
        "success": True,
        "subject": "Template: Monthly Newsletter - December 2023",
        "body": """
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; }
        .highlight { color: #0066cc; font-weight: bold; }
    </style>
</head>
<body>
    <p>Dear Team,</p>
    <p>Here is our monthly newsletter for <span class="highlight">December 2023</span>.</p>
    <h3>Highlights:</h3>
    <ul>
        <li>New product launches</li>
        <li>Company achievements</li>
        <li>Upcoming events</li>
    </ul>
    <p>Best regards,<br>Marketing Team</p>
</body>
</html>
        """,
        "to": ["recipients@company.com"],
    }

    print("Expected Response (Full HTML - LLM View):")
    print(json.dumps(expected_html_response, indent=2))
    print()
    print("💡 LLM now has full HTML to apply updates...")

    # ============================================
    # STEP 3: LLM updates template
    # ============================================
    print("STEP 3: LLM updating template with changes...")
    print("-" * 70)

    update_result = {
        "tool": "manage_templates",
        "arguments": {
            "action": "update",
            "template_number": 1,
            "subject": "Template: Monthly Newsletter - January 2024",  # ← Updated subject
            "htmlbody": """
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; }
        .highlight { color: #009900; font-weight: bold; }  # ← Updated color
        .new-section { background-color: #f0f0f0; padding: 10px; }  # ← New section
    </style>
</head>
<body>
    <p>Dear Team,</p>
    <p>Here is our monthly newsletter for <span class="highlight">January 2024</span>.</p>
    <div class="new-section">
        <h3>🎉 Happy New Year!</h3>
    </div>
    <h3>Highlights:</h3>
    <ul>
        <li>New year kickoff events</li>
        <li>Q1 planning sessions</li>
        <li>Upcoming holidays</li>
    </ul>
    <p>Best regards,<br>Marketing Team</p>
</body>
</html>
            """,  # ← Updated content
        },
    }

    print(f"Calling: {json.dumps(update_result, indent=2)}")
    print()

    expected_update_response = {
        "success": True,
        "message": "Template updated successfully",
        "template_number": 1,
    }

    print("Expected Response:")
    print(json.dumps(expected_update_response, indent=2))
    print()

    # ============================================
    # STEP 4: User verifies changes
    # ============================================
    print("STEP 4: User verifying changes (view as simple text)...")
    print("-" * 70)

    verify_result = {
        "tool": "manage_templates",
        "arguments": {"action": "get", "template_number": 1, "return_html": False},
    }

    print(f"Calling: {json.dumps(verify_result, indent=2)}")
    print()

    print("User reviews updated content and confirms changes look correct...")
    print()


async def template_send_workflow():
    """Demonstrate sending template to recipients."""

    print("=" * 70)
    print("MICROSOFT GRAPH MCP SERVER - TEMPLATE SENDING WORKFLOW")
    print("=" * 70)
    print()

    # Assume template_number=1 exists

    # ============================================
    # STEP 1: Send template
    # ============================================
    print("STEP 1: Sending template to recipients...")
    print("-" * 70)

    send_result = {
        "tool": "manage_templates",
        "arguments": {
            "action": "send",
            "template_number": 1,
            "to": [
                "employee1@company.com",
                "employee2@company.com",
                "employee3@company.com",
            ],
        },
    }

    print(f"Calling: {json.dumps(send_result, indent=2)}")
    print()

    expected_send_response = {
        "success": True,
        "message": "Template sent successfully",
        "sent_count": 1,
        "recipients": [
            "employee1@company.com",
            "employee2@company.com",
            "employee3@company.com",
        ],
    }

    print("Expected Response:")
    print(json.dumps(expected_send_response, indent=2))
    print()

    print("💡 TIP: Sending a template creates a copy and sends it.")
    print("   The original template is preserved for future use.")
    print("   You can send the same template multiple times to different recipients.")


async def common_mistakes_example():
    """Show common template mistakes and how to avoid them."""

    print()
    print("=" * 70)
    print("COMMON MISTAKES - TEMPLATE WORKFLOWS")
    print("=" * 70)
    print()

    # ============================================
    # MISTAKE 1: Using wrong template number
    # ============================================
    print("❌ MISTAKE 1: Using invalid template_number")
    print("-" * 70)

    print("Incorrect:")
    print("  manage_templates(action='get', template_number=999)")
    print("  # Template #999 doesn't exist!")
    print()

    print("Correct:")
    print("  manage_templates(action='list', page_number=1)")
    print("  # List templates to see valid template_numbers")
    print("  manage_templates(action='get', template_number=1)")
    print("  # Use template_number from list results")
    print()

    # ============================================
    # MISTAKE 2: Confusing return_html with htmlbody
    # ============================================
    print("❌ MISTAKE 2: Confusing return_html with htmlbody")
    print("-" * 70)

    print("Incorrect:")
    print("  # User view: get with return_html=false (simple text)")
    print("  manage_templates(action='get', template_number=1, return_html=false)")
    print("  # Returns simple text body, not HTML")
    print()

    print("Correct:")
    print("  # User view: get with return_html=false (simple text)")
    print("  manage_templates(action='get', template_number=1, return_html=false)")
    print("  # Returns simple text body for user review")
    print()
    print("  # LLM view: get with return_html=true (full HTML)")
    print("  manage_templates(action='get', template_number=1, return_html=true)")
    print("  # Returns full HTML body for LLM editing")
    print()

    # ============================================
    # MISTAKE 3: Not verifying changes before sending
    # ============================================
    print("❌ MISTAKE 3: Not verifying changes before sending")
    print("-" * 70)

    print("Incorrect:")
    print("  1. LLM updates template")
    print("  2. Immediately send template")
    print("  # User never sees the changes!")
    print()

    print("Correct:")
    print("  1. LLM updates template")
    print("  2. User verifies: get with return_html=false")
    print("  3. User approves changes")
    print("  4. Send template")
    print()

    print("=" * 70)


if __name__ == "__main__":
    asyncio.run(template_lifecycle_workflow())
    print()
    asyncio.run(template_editing_workflow())
    print()
    asyncio.run(template_send_workflow())
    print()
    asyncio.run(common_mistakes_example())
