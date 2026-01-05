"""MCP Server implementation for Microsoft Graph API."""

import asyncio
import json
import logging
from typing import Any, Dict, List, Optional

from mcp.server import Server
from mcp.server.models import InitializationOptions
from mcp.server.lowlevel.server import NotificationOptions
import mcp.server.stdio
import mcp.types as types

from .handlers import (
    AuthHandler,
    UserHandler,
    EmailHandler,
    CalendarHandler,
    FileHandler,
    TeamsHandler,
)
from .tools import ToolRegistry
from .config import settings
from .utils import read_bcc_from_csv

logger = logging.getLogger(__name__)


class MicrosoftGraphMCPServer:
    """MCP Server for Microsoft Graph API integration."""

    def __init__(self):
        logger.info("=" * 70)
        logger.info("MicrosoftGraphMCPServer.__init__() called")
        logger.info(f"Server name: {settings.server_name}")
        logger.info(f"Server version: {settings.server_version}")
        logger.info("=" * 70)

        self.server = Server("microsoft-graph-mcp-server")
        self.auth_handler = AuthHandler()
        self.user_handler = UserHandler()
        self.email_handler = EmailHandler()
        self.calendar_handler = CalendarHandler()
        self.file_handler = FileHandler()
        self.teams_handler = TeamsHandler()
        self._register_handlers()

        logger.info("Handlers registered successfully")

    def _register_handlers(self):
        """Register MCP tool handlers."""

        @self.server.list_tools()
        async def handle_list_tools() -> list[types.Tool]:
            """List all available tools."""
            return ToolRegistry.get_all_tools()

        @self.server.call_tool()
        async def handle_call_tool(
            name: str, arguments: dict | None
        ) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
            """Handle tool execution requests."""

            if arguments is None:
                arguments = {}

            logger.info(f"Server: Tool called - {name} with args: {arguments}")

            try:
                if name == "auth":
                    logger.info(
                        f"Server: Routing to auth_handler with action: {arguments.get('action')}"
                    )
                    return await self.auth_handler.handle_auth(arguments)

                elif name == "user_settings":
                    return await self.user_handler.handle_user_settings(arguments)

                elif name == "search_contacts":
                    return await self.user_handler.handle_search_contacts(arguments)

                elif name == "manage_mail_folder":
                    return await self.email_handler.handle_manage_mail_folder(arguments)

                elif name == "manage_emails":
                    return await self.email_handler.handle_manage_emails(arguments)

                elif name == "browse_email_cache":
                    return await self.email_handler.handle_browse_email_cache(arguments)

                elif name == "search_emails":
                    return await self.email_handler.handle_search_emails(arguments)

                elif name == "get_email_content":
                    return await self.email_handler.handle_get_email_content(arguments)

                elif name == "send_email":
                    return await self.email_handler.handle_send_email(arguments)

                elif name == "browse_events":
                    return await self.calendar_handler.handle_browse_events(arguments)

                elif name == "get_event_detail":
                    return await self.calendar_handler.handle_get_event(arguments)

                elif name == "search_events":
                    return await self.calendar_handler.handle_search_events(arguments)

                elif name == "check_attendee_availability":
                    return (
                        await self.calendar_handler.handle_check_attendee_availability(
                            arguments
                        )
                    )

                elif name == "respond_to_event":
                    return await self.calendar_handler.handle_respond_to_event(
                        arguments
                    )

                elif name == "manage_my_event":
                    return await self.calendar_handler.handle_manage_my_event(arguments)

                elif name == "list_files":
                    return await self.file_handler.handle_list_files(arguments)

                elif name == "get_teams":
                    return await self.teams_handler.handle_get_teams(arguments)

                elif name == "get_team_channels":
                    return await self.teams_handler.handle_get_team_channels(arguments)

                else:
                    raise ValueError(f"Unknown tool: {name}")

            except Exception as e:
                return [
                    types.TextContent(
                        type="text", text=f"Error executing tool {name}: {str(e)}"
                    )
                ]

    async def run(self):
        """Run the MCP server."""
        logger.info("=" * 70)
        logger.info("MicrosoftGraphMCPServer.run() called - Starting MCP server")
        logger.info("=" * 70)

        try:
            async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
                logger.info("stdio_server acquired, starting server.run()...")
                await self.server.run(
                    read_stream,
                    write_stream,
                    InitializationOptions(
                        server_name=settings.server_name,
                        server_version=settings.server_version,
                        capabilities=self.server.get_capabilities(
                            notification_options=NotificationOptions(),
                            experimental_capabilities={},
                        ),
                    ),
                )
        except Exception as e:
            logger.error(f"Error in server.run(): {e}", exc_info=True)
            raise
