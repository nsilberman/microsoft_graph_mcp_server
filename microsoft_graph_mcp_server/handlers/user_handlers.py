"""User handlers for MCP tools."""

from .base import BaseHandler
import mcp.types as types
from ..graph_client import graph_client
from ..config import settings


class UserHandler(BaseHandler):
    """Handler for user-related tools."""
    
    async def handle_get_user_info(self, arguments: dict) -> list[types.TextContent]:
        """Handle get_user_info tool."""
        result = await graph_client.get_me()
        return self._format_response(result)
    
    async def handle_search_contacts(self, arguments: dict) -> list[types.TextContent]:
        """Handle search_contacts tool."""
        query = arguments["query"]
        page_size = settings.page_size
        contacts = await graph_client.search_contacts(query, page_size)
        return self._format_response(contacts)
