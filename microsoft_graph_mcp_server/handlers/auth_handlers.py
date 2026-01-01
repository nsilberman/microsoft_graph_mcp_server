"""Authentication handlers for MCP tools."""

import mcp.types as types
from .base import BaseHandler
from ..auth_modules.auth_manager import auth_manager


class AuthHandler(BaseHandler):
    """Handler for authentication-related tools."""
    
    async def handle_auth(self, arguments: dict) -> list[types.TextContent]:
        """Handle auth tool with login, check_status, and logout actions."""
        action = arguments.get("action")
        
        if action == "login":
            return await self._handle_login()
        elif action == "check_status":
            return await self._handle_check_status()
        elif action == "logout":
            return await self._handle_logout()
        else:
            return self._format_error(f"Invalid action: {action}. Must be 'login', 'check_status', or 'logout'.")
    
    async def _handle_login(self) -> list[types.TextContent]:
        """Handle login action."""
        try:
            result = await auth_manager.login()
            return self._format_response(result)
        except Exception as e:
            return self._format_error(f"Login failed: {str(e)}")
    
    async def _handle_check_status(self) -> list[types.TextContent]:
        """Handle check_status action."""
        try:
            result = await auth_manager.check_login_status()
            return self._format_response(result)
        except Exception as e:
            return self._format_error(f"Failed to check login status: {str(e)}")
    
    async def _handle_logout(self) -> list[types.TextContent]:
        """Handle logout action."""
        try:
            result = await auth_manager.logout()
            return self._format_response(result)
        except Exception as e:
            return self._format_error(f"Logout failed: {str(e)}")