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
                    name="list_users",
                    description="List users in the organization",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "filter": {
                                "type": "string",
                                "description": "Optional OData filter query"
                            }
                        }
                    }
                ),
                types.Tool(
                    name="get_messages",
                    description="Get email messages from specified folder",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "folder": {
                                "type": "string",
                                "description": "Mail folder name (default: Inbox)",
                                "default": "Inbox"
                            },
                            "top": {
                                "type": "integer",
                                "description": "Number of messages to retrieve (default: 10)",
                                "default": 10
                            },
                            "filter": {
                                "type": "string",
                                "description": "Optional OData filter query"
                            }
                        }
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
                    name="get_events",
                    description="Get calendar events",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "start_date": {
                                "type": "string",
                                "description": "Start date in ISO format (optional)"
                            },
                            "end_date": {
                                "type": "string",
                                "description": "End date in ISO format (optional)"
                            }
                        }
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
                
                elif name == "list_users":
                    filter_query = arguments.get("filter")
                    users = await graph_client.get_users(filter_query)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(users, indent=2, ensure_ascii=False)
                    )]
                
                elif name == "get_messages":
                    folder = arguments.get("folder", "Inbox")
                    top = arguments.get("top", 10)
                    filter_query = arguments.get("filter")
                    messages = await graph_client.get_messages(folder, top, filter_query)
                    return [types.TextContent(
                        type="text",
                        text=json.dumps(messages, indent=2, ensure_ascii=False)
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
                
                elif name == "get_events":
                    start_date = arguments.get("start_date")
                    end_date = arguments.get("end_date")
                    events = await graph_client.get_events(start_date, end_date)
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