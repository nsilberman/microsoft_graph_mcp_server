"""Authentication handlers for MCP tools."""

from .base import BaseHandler
import mcp.types as types
from ..auth import auth_manager


class AuthHandler(BaseHandler):
    """Handler for authentication-related tools."""
    
    async def handle_login(self, arguments: dict) -> list[types.TextContent]:
        """Handle login tool."""
        result = await auth_manager.login()
        return self._format_response(result)
    
    async def handle_check_login_status(self, arguments: dict) -> list[types.TextContent]:
        """Handle check_login_status tool."""
        result = await auth_manager.check_login_status()
        return self._format_response(result)
    
    async def handle_logout(self, arguments: dict) -> list[types.TextContent]:
        """Handle logout tool."""
        result = await auth_manager.logout()
        return self._format_response(result)
