"""MCP Server implementation for Microsoft Graph API."""

import asyncio
import csv
import json
from pathlib import Path
from typing import Any, Dict, List, Optional

from mcp.server import Server
from mcp.server.models import InitializationOptions
from mcp.server.lowlevel.server import NotificationOptions
import mcp.server.stdio
import mcp.types as types

from .graph_client import graph_client
from .config import settings
from .auth import auth_manager
from .email_cache import email_cache
from .date_handler import date_handler


def read_bcc_from_csv(csv_file_path: str) -> List[str]:
    """Read BCC email addresses from a CSV file.
    
    The CSV file should have a single column with header "Email" or "email".
    
    Args:
        csv_file_path: Path to the CSV file
        
    Returns:
        List of email addresses
        
    Raises:
        FileNotFoundError: If the CSV file doesn't exist
        ValueError: If the CSV file doesn't have the required header
    """
    csv_path = Path(csv_file_path)
    
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file not found: {csv_file_path}")
    
    bcc_emails = []
    
    with open(csv_path, 'r', newline='', encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile)
        
        if not reader.fieldnames:
            raise ValueError("CSV file is empty or has no headers")
        
        email_column = None
        for header in reader.fieldnames:
            if header.strip().lower() == "email":
                email_column = header
                break
        
        if email_column is None:
            raise ValueError("CSV file must have a column named 'Email' or 'email'")
        
        for row in reader:
            email = row[email_column].strip()
            if email:
                bcc_emails.append(email)
    
    return bcc_emails


class MicrosoftGraphMCPServer:
    """MCP Server for Microsoft Graph API integration."""
    
    def __init__(self):
        self.server = Server("microsoft-graph-mcp-server")
        self._register_handlers()
    
    def _register_handlers(self):
        """Register MCP tool handlers."""
        
        @self.server.list_tools()
        async def handle_list_tools() -> list[types.Tool]:
            """List all available tools."""
            return [
                types.Tool(
                    name="login",
                    description="Authenticate with Microsoft Graph using device code flow. First call returns verification link and code. Open the link, enter the code, and sign in. Then call login again to verify authentication (this second call will timeout quickly if you haven't completed authentication yet). If already authenticated, returns remaining time until token expires. Run this first before using other tools.",
                    inputSchema={
                        "type": "object",
                        "properties": {}
                    }
                ),
                types.Tool(
                    name="check_login_status",
                    description="Check current authentication status with Microsoft Graph. Shows if authenticated, not authenticated, or expired, along with token expiry time. Does not initiate authentication.",
                    inputSchema={
                        "type": "object",
                        "properties": {}
                    }
                ),
                types.Tool(
                    name="logout",
                    description="Logout from Microsoft Graph and clear authentication state. Useful for security, testing, or when you want to switch accounts.",
                    inputSchema={
                        "type": "object",
                        "properties": {}
                    }
                ),
                types.Tool(
                    name="get_user_info",
                    description="Get current user information from Microsoft Graph",
                    inputSchema={
                        "type": "object",
                        "properties": {}
                    }
                ),
                types.Tool(
                    name="search_contacts",
                    description="Search contacts and people relevant to you. Returns people you interact with most, including organization users and your personal contacts. Use this to find specific people by name or email.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "query": {
                                "type": "string",
                                "description": "Search query (name or email)"
                            }
                        },
                        "required": ["query"]
                    }
                ),
                types.Tool(
                    name="list_mail_folders",
                    description="List all mail folders with their paths (e.g., 'Inbox', 'Inbox/Projects', 'Archive/2024').",
                    inputSchema={
                        "type": "object",
                        "properties": {}
                    }
                ),
                types.Tool(
                    name="list_recent_emails",
                    description="List recent emails from Inbox with optional days parameter. Loads emails from the last N days (default: 1 day, maximum: 7 days) into cache for browsing.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "days": {
                                "type": "integer",
                                "description": "Number of days to look back (default: 1, maximum: 7)",
                                "default": 1,
                                "minimum": 1,
                                "maximum": 7
                            }
                        }
                    }
                ),
                types.Tool(
                    name="load_emails_by_folder",
                    description="Load emails from a folder into cache. Can filter by days or limit by top number (mutually exclusive). Use 'days' to get emails from last N days, or 'top' to get the most recent N emails. Loads emails into cache for browsing.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "folder": {
                                "type": "string",
                                "description": "Mail folder path (e.g., 'Inbox', 'Inbox/Projects', 'Archive/2024'). Default: Inbox",
                                "default": "Inbox"
                            },
                            "days": {
                                "type": "integer",
                                "description": "Number of days to look back (e.g., 7 for last 7 days). Cannot be used with 'top'.",
                                "minimum": 1
                            },
                            "top": {
                                "type": "integer",
                                "description": "Maximum number of emails to load (e.g., 50). Cannot be used with 'days'.",
                                "minimum": 1
                            }
                        }
                    }
                ),
                types.Tool(
                    name="clear_email_cache",
                    description="Clear the email browsing cache. This removes all cached emails from memory and disk. Use this to free up memory or start fresh.",
                    inputSchema={
                        "type": "object",
                        "properties": {}
                    }
                ),
                types.Tool(
                    name="browse_email_cache",
                    description="Browse emails in the cache with pagination. Returns summary information with number column indicating position in cache. Use page_number to navigate. Automatically manages browsing state with disk cache for persistence.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "page_number": {
                                "type": "integer",
                                "description": "Page number to view (starts at 1)",
                                "minimum": 1
                            }
                        },
                        "required": ["page_number"]
                    }
                ),
                types.Tool(
                    name="search_emails",
                    description="Unified search tool for emails. Search by sender, recipient, subject, or body text. Returns email numbers found in cache. Use browse_email_cache to view the results. All searches return only the count and hint to use browse tool.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "search_type": {
                                "type": "string",
                                "enum": ["sender", "recipient", "subject", "body"],
                                "description": "Type of search to perform"
                            },
                            "query": {
                                "type": "string",
                                "description": "Search query (sender name/email, recipient name/email, subject text, or body text)"
                            },
                            "folder": {
                                "type": "string",
                                "description": "Optional folder path to search (e.g., 'Inbox', 'Inbox/Projects', 'Archive/2024'). Default: Inbox",
                                "default": "Inbox"
                            },
                            "days": {
                                "type": "integer",
                                "description": "Number of days to search back (default: 90). Set to null to search all emails.",
                                "default": 90
                            }
                        },
                        "required": ["search_type", "query"]
                    }
                ),
                types.Tool(
                    name="get_email_content",
                    description="Get full email content by cache number. Use the email number from browse_email_cache (e.g., 1, 2, 3) to retrieve complete email with body, attachments, and all details.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "emailNumber": {
                                "type": "integer",
                                "description": "Email number from browse_email_cache (e.g., 1, 2, 3)"
                            },
                            "text_only": {
                                "type": "boolean",
                                "description": "If true, return only text content without embedded images and attachments. If false, return full content including embedded images and attachments.",
                                "default": True
                            }
                        },
                        "required": ["emailNumber"]
                    }
                ),
                types.Tool(
                    name="send_message",
                    description="Send an email message",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "to": {
                                "type": "string",
                                "description": "Recipient email address"
                            },
                            "subject": {
                                "type": "string",
                                "description": "Email subject"
                            },
                            "body": {
                                "type": "string",
                                "description": "Email body content"
                            },
                            "cc": {
                                "type": "string",
                                "description": "CC recipient email address (optional)"
                            }
                        },
                        "required": ["to", "subject", "body"]
                    }
                ),
                types.Tool(
                    name="compose_email",
                    description="Compose and send a new email. Supports multiple recipients, CC, and BCC.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "to": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of recipient email addresses"
                            },
                            "subject": {
                                "type": "string",
                                "description": "Email subject"
                            },
                            "body": {
                                "type": "string",
                                "description": "Email body content"
                            },
                            "cc": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of CC recipient email addresses (optional)"
                            },
                            "bcc": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of BCC recipient email addresses (optional)"
                            },
                            "body_content_type": {
                                "type": "string",
                                "enum": ["Text", "HTML"],
                                "description": "Content type for the email body (default: HTML)",
                                "default": "HTML"
                            }
                        },
                        "required": ["to", "subject", "body"]
                    }
                ),
                types.Tool(
                    name="reply_email",
                    description="Reply to an existing email. The reply will be linked to the original email thread. If only emailNumber is provided, it will show the original email content for preview. If reply parameters are provided, it will send the reply with the email thread included. IMPORTANT: The body must be HTML format.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "emailNumber": {
                                "type": "integer",
                                "description": "Email number from browse_email_cache (e.g., 1, 2, 3)"
                            },
                            "to": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of recipient email addresses (optional - if not provided, will show email preview)"
                            },
                            "subject": {
                                "type": "string",
                                "description": "Email subject (optional - if not provided, will show email preview)"
                            },
                            "body": {
                                "type": "string",
                                "description": "Email body content (optional - if not provided, will show email preview). MUST be HTML format."
                            },
                            "cc": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of CC recipient email addresses (optional)"
                            },
                            "bcc": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of BCC recipient email addresses (optional)"
                            }
                        },
                        "required": ["emailNumber"]
                    }
                ),
                types.Tool(
                    name="batch_forward_email",
                    description="Forward an email to batch recipients. The original email will be included in the forwarded message with 'FW:' prefix on the subject. You can add a message before the forwarded content. BCC recipients can be provided via a CSV file with a single 'Email' or 'email' column. If BCC recipients exceed the limit (default 500), they will be sent in batches.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "emailNumber": {
                                "type": "integer",
                                "description": "Email number from browse_email_cache (e.g., 1, 2, 3)"
                            },
                            "to": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of recipient email addresses"
                            },
                            "subject": {
                                "type": "string",
                                "description": "Email subject (optional, defaults to 'FW: ' + original subject)"
                            },
                            "body": {
                                "type": "string",
                                "description": "Email body content (optional, message to add before forwarded content)"
                            },
                            "cc": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of CC recipient email addresses (optional)"
                            },
                            "bcc": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of BCC recipient email addresses (optional)"
                            },
                            "bcc_csv_file": {
                                "type": "string",
                                "description": "Path to CSV file containing BCC recipients. CSV must have a single column with header 'Email' or 'email' (optional)"
                            },
                            "body_content_type": {
                                "type": "string",
                                "enum": ["Text", "HTML"],
                                "description": "Content type for the email body (default: Text)",
                                "default": "Text"
                            }
                        },
                        "required": ["emailNumber", "to"]
                    }
                ),
                types.Tool(
                    name="browse_events",
                    description="Browse calendar events with pagination. Returns summary information (id, subject, start, end, location, organizer, attendees, isAllDay, showAs, importance). Use page_number to navigate.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "page_number": {
                                "type": "integer",
                                "description": "Page number to view (starts at 1)",
                                "minimum": 1
                            },
                            "start_date": {
                                "type": "string",
                                "description": "Start date in ISO format (optional)"
                            },
                            "end_date": {
                                "type": "string",
                                "description": "End date in ISO format (optional)"
                            }
                        },
                        "required": ["page_number"]
                    }
                ),
                types.Tool(
                    name="get_event",
                    description="Get full calendar event by ID. Use the event ID from browse_events or search_events to retrieve complete event with body, attachments, and all details.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "event_id": {
                                "type": "string",
                                "description": "Event ID from browse_events or search_events"
                            }
                        },
                        "required": ["event_id"]
                    }
                ),
                types.Tool(
                    name="search_events",
                    description="Search calendar events by keywords. Returns summary information. Note: Pagination with skip is not supported with search.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "query": {
                                "type": "string",
                                "description": "Search keywords"
                            },
                            "start_date": {
                                "type": "string",
                                "description": "Start date in ISO format (optional)"
                            },
                            "end_date": {
                                "type": "string",
                                "description": "End date in ISO format (optional)"
                            }
                        },
                        "required": ["query"]
                    }
                ),
                types.Tool(
                    name="create_event",
                    description="Create a calendar event",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "subject": {
                                "type": "string",
                                "description": "Event subject"
                            },
                            "start": {
                                "type": "string",
                                "description": "Start time in ISO format"
                            },
                            "end": {
                                "type": "string",
                                "description": "End time in ISO format"
                            },
                            "attendees": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of attendee email addresses"
                            }
                        },
                        "required": ["subject", "start", "end"]
                    }
                ),
                types.Tool(
                    name="list_files",
                    description="List files and folders from OneDrive",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "folder_path": {
                                "type": "string",
                                "description": "Folder path (default: root)",
                                "default": ""
                            }
                        }
                    }
                ),
                types.Tool(
                    name="get_teams",
                    description="Get list of Microsoft Teams",
                    inputSchema={
                        "type": "object",
                        "properties": {}
                    }
                ),
                types.Tool(
                    name="get_team_channels",
                    description="Get channels for a specific Team",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "team_id": {
                                "type": "string",
                                "description": "Team ID"
                            }
                        },
                        "required": ["team_id"]
                    }
                )
            ]
        
        @self.server.call_tool()
        async def handle_call_tool(
            name: str,
            arguments: dict | None
        ) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
            """Handle tool execution requests."""
            
            if arguments is None:
                arguments = {}
            
            try:
                if name == "login":
                    result = await auth_manager.login()
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(result, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "check_login_status":
                    result = await auth_manager.check_login_status()
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(result, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "logout":
                    result = await auth_manager.logout()
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(result, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "get_user_info":
                    result = await graph_client.get_me()
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(result, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "search_contacts":
                    query = arguments["query"]
                    page_size = settings.page_size
                    contacts = await graph_client.search_contacts(query, page_size)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(contacts, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "list_mail_folders":
                    folders = await graph_client.list_mail_folders()
                    return [types.TextContent(
                        type="text",
                        text=json.dumps({
                            "message": f"Found {len(folders)} mail folders",
                            "folders": folders,
                            "count": len(folders)
                        }, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "list_recent_emails":
                    days = arguments.get("days", 1)
                    
                    if days > 7:
                        return [types.TextContent(
                            type="text",
                            text="Error: Days parameter must be 7 or less."
                        )]
                    
                    email_cache.clear_cache()
                    result = await graph_client.load_emails_by_folder("Inbox", days, None)
                    
                    user_timezone = await graph_client.get_user_timezone()
                    
                    await email_cache.set_mode("list")
                    await email_cache.update_list_state(
                        folder="Inbox",
                        days=days,
                        top=None,
                        total_count=result["count"],
                        metadata=result["metadata"]
                    )
                    
                    email_date_range = date_handler.format_email_date_range(result["metadata"], user_timezone)
                    filter_date_range = date_handler.format_filter_date_range(days, user_timezone)
                    
                    response_data = {
                        "message": f"Loaded {result['count']} recent emails from Inbox (last {days} day(s))",
                        "folder": "Inbox",
                        "days": days,
                        "count": result["count"],
                        "timezone": user_timezone,
                        "date_range": email_date_range,
                        "filter_date_range": filter_date_range,
                        "hint": "Use browse_email_cache to view the loaded emails"
                    }
                    
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(response_data, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "load_emails_by_folder":
                    folder = arguments.get("folder", "Inbox")
                    days = arguments.get("days")
                    top = arguments.get("top")
                    
                    if days is not None and top is not None:
                        return [types.TextContent(
                            type="text",
                            text="Error: Cannot specify both 'days' and 'top' parameters simultaneously. Please use only one."
                        )]
                    
                    email_cache.clear_cache()
                    result = await graph_client.load_emails_by_folder(folder, days, top)
                    
                    user_timezone = await graph_client.get_user_timezone()
                    
                    await email_cache.set_mode("list")
                    await email_cache.update_list_state(
                        folder=folder,
                        days=days,
                        top=top,
                        total_count=result["count"],
                        metadata=result["metadata"]
                    )
                    
                    filter_date_range = date_handler.format_filter_date_range(days, user_timezone)
                    
                    response_data = {
                        "message": f"Loaded {result['count']} emails from {folder}",
                        "folder": folder,
                        "count": result["count"],
                        "timezone": user_timezone,
                        "date_range": filter_date_range,
                        "hint": "Use browse_email_cache to view the loaded emails"
                    }
                    
                    if days is not None:
                        response_data["days"] = days
                    if top is not None:
                        response_data["top"] = top
                    
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(response_data, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "clear_email_cache":
                    email_cache.clear_cache()
                    return [types.TextContent(
                        type="text",
                        text=json.dumps({
                            "message": "Email cache cleared successfully",
                            "status": "success"
                        }, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "browse_email_cache":
                    page_number = arguments["page_number"]
                    
                    cached_emails = email_cache.get_cached_emails()
                    total_count = len(cached_emails)
                    
                    if total_count == 0:
                        return [types.TextContent(
                            type="text",
                            text=json.dumps({
                                "message": "No emails in cache. Use load_emails_by_folder to load emails first.",
                                "emails": [],
                                "count": 0
                            }, indent=2, ensure_ascii=False)
                        )]
                    
                    page_size = settings.page_size
                    start_idx = (page_number - 1) * page_size
                    end_idx = start_idx + page_size
                    page_emails = cached_emails[start_idx:end_idx]
                    
                    user_timezone = await graph_client.get_user_timezone()
                    
                    filtered_emails = []
                    for idx, email in enumerate(page_emails):
                        filtered_email = {k: v for k, v in email.items() if k not in ["id", "metadata"]}
                        filtered_email["number"] = start_idx + idx + 1
                        filtered_emails.append(filtered_email)
                    
                    return [types.TextContent(
                        type="text",
                        text=json.dumps({
                            "emails": filtered_emails,
                            "count": len(filtered_emails),
                            "total_count": total_count,
                            "current_page": page_number,
                            "total_pages": (total_count + page_size - 1) // page_size,
                            "timezone": user_timezone
                        }, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "search_emails":
                    search_type = arguments["search_type"]
                    query = arguments["query"]
                    folder = arguments.get("folder", "Inbox")
                    page_size = settings.page_size
                    days = arguments.get("days", 90)
                    
                    email_cache.clear_cache()
                    
                    user_timezone = await graph_client.get_user_timezone()
                    
                    if search_type == "sender":
                        result = await graph_client.search_emails_by_sender(query, folder, page_size, days)
                    elif search_type == "recipient":
                        result = await graph_client.search_emails_by_recipient(query, folder, page_size, days)
                    elif search_type == "subject":
                        result = await graph_client.search_emails_by_subject(query, folder, page_size, days)
                    elif search_type == "body":
                        result = await graph_client.search_emails_by_body(query, folder, page_size, days)
                    else:
                        return [types.TextContent(
                            type="text",
                            text=f"Error: Invalid search_type '{search_type}'. Must be one of: sender, recipient, subject, body"
                        )]
                    
                    await email_cache.set_mode("search")
                    await email_cache.update_search_state(
                    query=query,
                    folder=folder,
                    top=page_size,
                    days=days,
                    search_type=search_type,
                    total_count=result["count"],
                    metadata=result["metadata"]
                )
                    
                    response_data = {
                        "search_type": search_type,
                        "query": query,
                        "folder": folder,
                        "count": result["count"],
                        "timezone": user_timezone,
                        "date_range": result.get("date_range"),
                        "filter_date_range": result.get("filter_date_range"),
                        "hint": f"Found {result['count']} emails. Use browse_email_cache to view the results."
                    }
                    
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(response_data, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "get_email_content":
                    emailNumber = arguments["emailNumber"]
                    text_only = arguments.get("text_only", True)
                    
                    cached_emails = email_cache.get_cached_emails()
                    
                    if emailNumber < 1 or emailNumber > len(cached_emails):
                        return [types.TextContent(
                            type="text",
                            text=json.dumps({
                                "error": f"Email number {emailNumber} is out of range. Please choose a number between 1 and {len(cached_emails)}."
                            }, indent=2, ensure_ascii=False)
                        )]
                    
                    email = cached_emails[emailNumber - 1]
                    email_id = email.get("id")
                    
                    if not email_id:
                        return [types.TextContent(
                            type="text",
                            text=json.dumps({
                                "error": f"Email number {emailNumber} does not have a valid Graph ID in cache."
                            }, indent=2, ensure_ascii=False)
                        )]
                    
                    email_content = await graph_client.get_email(email_id, emailNumber, text_only=text_only)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(email_content["content"], indent=2, ensure_ascii=False)
                    )]
                
                elif name == "send_message":
                    message_data = {
                        "subject": arguments["subject"],
                        "body": {
                            "contentType": "Text",
                            "content": arguments["body"]
                        },
                        "toRecipients": [
                            {
                                "emailAddress": {
                                    "address": arguments["to"]
                                }
                            }
                        ]
                    }
                    
                    if "cc" in arguments:
                        message_data["ccRecipients"] = [
                            {
                                "emailAddress": {
                                    "address": arguments["cc"]
                                }
                            }
                        ]
                    
                    result = await graph_client.send_message(message_data)
                    return [types.TextContent(
                        type="text",
                        text=f"Message sent successfully: {json.dumps(result, indent=2, ensure_ascii=False)}"
                    )]
                
                elif name == "compose_email":
                    to_recipients = arguments["to"]
                    subject = arguments["subject"]
                    body = arguments["body"]
                    cc_recipients = arguments.get("cc")
                    bcc_recipients = arguments.get("bcc")
                    body_content_type = arguments.get("body_content_type", "Text")
                    
                    result = await graph_client.send_email(
                        to_recipients=to_recipients,
                        subject=subject,
                        body=body,
                        cc_recipients=cc_recipients,
                        bcc_recipients=bcc_recipients,
                        body_content_type=body_content_type
                    )
                    return [types.TextContent(
                        type="text",
                        text=f"Email composed and sent successfully: {json.dumps(result, indent=2, ensure_ascii=False)}"
                    )]
                
                elif name == "reply_email":
                    emailNumber = arguments["emailNumber"]
                    to_recipients = arguments.get("to")
                    subject = arguments.get("subject")
                    body = arguments.get("body")
                    cc_recipients = arguments.get("cc")
                    bcc_recipients = arguments.get("bcc")

                    # DEBUG: Log incoming body
                    import sys
                    print(f"[DEBUG] reply_email: body type={type(body).__name__}, first 100 chars={repr(body[:100]) if body else None}", file=sys.stderr)
                    print(f"[DEBUG] reply_email: body has {repr(body.count(chr(10)) if body else 0)} newlines", file=sys.stderr)
                    
                    cached_emails = email_cache.get_cached_emails()
                    if emailNumber < 1 or emailNumber > len(cached_emails):
                        return [types.TextContent(
                            type="text",
                            text=f"Error: Email number {emailNumber} is out of range. Please use a number between 1 and {len(cached_emails)}."
                        )]
                    
                    email = cached_emails[emailNumber - 1]
                    email_id = email["id"]
                    
                    if not to_recipients or not subject or not body:
                        full_email = await graph_client.get_email(email_id, emailNumber, text_only=True)
                        return [types.TextContent(
                            type="text",
                            text=json.dumps({
                                "message": "Original email content for preview. To reply, provide to, subject, and body parameters.",
                                "email": full_email
                            }, indent=2, ensure_ascii=False)
                        )]
                    
                    result = await graph_client.send_email(
                        to_recipients=to_recipients,
                        subject=subject,
                        body=body,
                        cc_recipients=cc_recipients,
                        bcc_recipients=bcc_recipients,
                        reply_to_message_id=email_id,
                        body_content_type="HTML"
                    )
                    return [types.TextContent(
                        type="text",
                        text=f"Reply email sent successfully: {json.dumps(result, indent=2, ensure_ascii=False)}"
                    )]
                
                elif name == "batch_forward_email":
                    email_number = arguments["emailNumber"]
                    to_recipients = arguments["to"]
                    subject = arguments.get("subject")
                    body = arguments.get("body", "")
                    cc_recipients = arguments.get("cc")
                    bcc_recipients = arguments.get("bcc")
                    bcc_csv_file = arguments.get("bcc_csv_file")
                    body_content_type = arguments.get("body_content_type", "HTML")
                    
                    cached_emails = email_cache.get_cached_emails()
                    total_count = len(cached_emails)
                    
                    if total_count == 0:
                        return [types.TextContent(
                            type="text",
                            text="Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
                        )]
                    
                    if email_number < 1 or email_number > total_count:
                        return [types.TextContent(
                            type="text",
                            text=f"Error: Invalid email number: {email_number}. Please use valid number from browse_email_cache (1-{total_count})."
                        )]
                    
                    email = cached_emails[email_number - 1]
                    email_id = email.get("id")
                    
                    if not email_id:
                        return [types.TextContent(
                            type="text",
                            text="Error: No valid email ID found. Please check the cache and try again."
                        )]
                    
                    if bcc_csv_file:
                        try:
                            csv_bcc = read_bcc_from_csv(bcc_csv_file)
                            if bcc_recipients:
                                bcc_recipients = bcc_recipients + csv_bcc
                            else:
                                bcc_recipients = csv_bcc
                        except Exception as e:
                            return [types.TextContent(
                                type="text",
                                text=f"Error reading BCC CSV file: {str(e)}"
                            )]
                    
                    all_bcc_recipients = bcc_recipients or []
                    max_bcc = settings.max_bcc_recipients
                    total_bcc = len(all_bcc_recipients)
                    
                    if total_bcc > max_bcc:
                        num_batches = (total_bcc + max_bcc - 1) // max_bcc
                        results = []
                        
                        for i in range(num_batches):
                            start_idx = i * max_bcc
                            end_idx = start_idx + max_bcc
                            batch_bcc = all_bcc_recipients[start_idx:end_idx]
                            
                            result = await graph_client.batch_forward_emails(
                                to_recipients=to_recipients,
                                subject=subject,
                                body=body,
                                email_ids=[email_id],
                                cc_recipients=cc_recipients,
                                bcc_recipients=batch_bcc,
                                body_content_type=body_content_type
                            )
                            results.append({
                                "batch": i + 1,
                                "bcc_count": len(batch_bcc),
                                "result": result
                            })
                        
                        response_message = f"Email forwarded successfully in {num_batches} batches (total {total_bcc} BCC recipients): {json.dumps(results, indent=2, ensure_ascii=False)}"
                    else:
                        result = await graph_client.batch_forward_emails(
                            to_recipients=to_recipients,
                            subject=subject,
                            body=body,
                            email_ids=[email_id],
                            cc_recipients=cc_recipients,
                            bcc_recipients=bcc_recipients,
                            body_content_type=body_content_type
                        )
                        
                        response_message = f"Email forwarded successfully: {json.dumps(result, indent=2, ensure_ascii=False)}"
                        if bcc_recipients:
                            response_message = f"Email forwarded successfully to {len(bcc_recipients)} BCC recipients: {json.dumps(result, indent=2, ensure_ascii=False)}"
                    
                    return [types.TextContent(
                        type="text",
                        text=response_message
                    )]
                
                elif name == "browse_events":
                    page_number = arguments["page_number"]
                    start_date = arguments.get("start_date")
                    end_date = arguments.get("end_date")
                    page_size = settings.page_size
                    skip = (page_number - 1) * page_size
                    events = await graph_client.browse_events(start_date, end_date, page_size, skip)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(events, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "get_event":
                    event_id = arguments["event_id"]
                    event = await graph_client.get_event(event_id)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(event, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "search_events":
                    query = arguments["query"]
                    start_date = arguments.get("start_date")
                    end_date = arguments.get("end_date")
                    page_size = settings.page_size
                    events = await graph_client.search_events(query, start_date, end_date, page_size)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(events, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "create_event":
                    event_data = {
                        "subject": arguments["subject"],
                        "start": {
                            "dateTime": arguments["start"],
                            "timeZone": "UTC"
                        },
                        "end": {
                            "dateTime": arguments["end"],
                            "timeZone": "UTC"
                        }
                    }
                    
                    if "attendees" in arguments:
                        event_data["attendees"] = [
                            {
                                "emailAddress": {
                                    "address": email
                                },
                                "type": "required"
                            }
                            for email in arguments["attendees"]
                        ]
                    
                    result = await graph_client.create_event(event_data)
                    return [types.TextContent(
                        type="text",
                        text=f"Event created successfully: {json.dumps(result, indent=2, ensure_ascii=False)}"
                    )]
                
                elif name == "list_files":
                    folder_path = arguments.get("folder_path", "")
                    items = await graph_client.get_drive_items(folder_path)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(items, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "get_teams":
                    teams = await graph_client.get_teams()
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(teams, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "get_team_channels":
                    team_id = arguments["team_id"]
                    channels = await graph_client.get_team_channels(team_id)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(channels, indent=2, ensure_ascii=False)
                    )]
                
                else:
                    raise ValueError(f"Unknown tool: {name}")
            
            except Exception as e:
                return [types.TextContent(
                    type="text",
                    text=f"Error executing tool {name}: {str(e)}"
                )]
    
    async def run(self):
        """Run the MCP server."""
        async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream,
                write_stream,
                InitializationOptions(
                    server_name=settings.server_name,
                    server_version=settings.server_version,
                    capabilities=self.server.get_capabilities(
                        notification_options=NotificationOptions(),
                        experimental_capabilities={}
                    )
                )
            )