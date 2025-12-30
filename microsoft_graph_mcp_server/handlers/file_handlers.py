"""File handlers for MCP tools."""

import json
from mcp import types

from .base import BaseHandler
from ..graph_client import graph_client


class FileHandler(BaseHandler):
    """Handler for file-related tools."""

    async def handle_list_files(self, arguments: dict) -> list[types.TextContent]:
        """Handle list_files tool."""
        folder_path = arguments.get("folder_path", "")
        items = await graph_client.get_drive_items(folder_path)
        return self._format_response(items)
