"""MCP Server implementation for Microsoft Graph API."""

import asyncio
import json
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
                            },
                            "top": {
                                "type": "integer",
                                "description": "Number of results to return (default: 10)",
                                "default": 10
                            }
                        },
                        "required": ["query"]
                    }
                ),
                types.Tool(
                    name="browse_emails",
                    description="Browse emails in a folder with pagination. Returns summary information (id, subject, from, date, read status, attachments, importance). Use page_number to navigate. Automatically manages browsing state with disk cache for persistence.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "page_number": {
                                "type": "integer",
                                "description": "Page number to view (starts at 1)",
                                "minimum": 1
                            },
                            "folder": {
                                "type": "string",
                                "description": "Mail folder name (default: Inbox, cached from previous call)",
                                "default": "Inbox"
                            },
                            "top": {
                                "type": "integer",
                                "description": "Number of emails per page (default: 20, cached from previous call)",
                                "default": 20
                            },
                            "filter": {
                                "type": "string",
                                "description": "Optional OData filter query (e.g., 'isRead eq false' for unread only)"
                            }
                        },
                        "required": ["page_number"]
                    }
                ),
                types.Tool(
                    name="get_email",
                    description="Get full email content by ID. Use the email ID from browse_emails or search_emails to retrieve complete email with body, attachments, and all details.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "email_id": {
                                "type": "string",
                                "description": "Email ID from browse_emails or search_emails"
                            }
                        },
                        "required": ["email_id"]
                    }
                ),
                types.Tool(
                    name="search_emails",
                    description="Search emails by keywords across all folders or specific folder. Returns summary information. Automatically manages search state with disk cache for persistence.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "query": {
                                "type": "string",
                                "description": "Search keywords"
                            },
                            "page_number": {
                                "type": "integer",
                                "description": "Page number to view (starts at 1, cached from previous call)",
                                "default": 1,
                                "minimum": 1
                            },
                            "folder": {
                                "type": "string",
                                "description": "Optional folder to search (default: all folders)"
                            },
                            "top": {
                                "type": "integer",
                                "description": "Number of emails per page (default: 20, cached from previous call)",
                                "default": 20
                            }
                        },
                        "required": ["query"]
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
                            },
                            "top": {
                                "type": "integer",
                                "description": "Number of events per page (default: 20)",
                                "default": 20
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
                            },
                            "top": {
                                "type": "integer",
                                "description": "Number of events to return (default: 20)",
                                "default": 20
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
                    top = arguments.get("top", 10)
                    contacts = await graph_client.search_contacts(query, top)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(contacts, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "browse_emails":
                    page_number = arguments["page_number"]
                    folder = arguments.get("folder", "Inbox")
                    top = arguments.get("top", 20)
                    filter_query = arguments.get("filter")
                    emails = await graph_client.browse_emails(page_number, folder, top, filter_query)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(emails, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "get_email":
                    email_id = arguments["email_id"]
                    email = await graph_client.get_email(email_id)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(email, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "search_emails":
                    query = arguments["query"]
                    page_number = arguments.get("page_number", 1)
                    folder = arguments.get("folder")
                    top = arguments.get("top", 20)
                    emails = await graph_client.search_emails(query, page_number, folder, top)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(emails, indent=2, ensure_ascii=False)
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
                
                elif name == "browse_events":
                    page_number = arguments["page_number"]
                    start_date = arguments.get("start_date")
                    end_date = arguments.get("end_date")
                    top = arguments.get("top", 20)
                    skip = (page_number - 1) * top
                    events = await graph_client.browse_events(start_date, end_date, top, skip)
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
                    top = arguments.get("top", 20)
                    events = await graph_client.search_events(query, start_date, end_date, top)
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