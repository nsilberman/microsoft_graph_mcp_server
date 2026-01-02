"""Authentication handlers for MCP tools."""

import logging
from typing import Optional

import mcp.types as types
from .base import BaseHandler
from ..auth_modules.auth_manager import auth_manager

logger = logging.getLogger(__name__)


class AuthHandler(BaseHandler):
    """Handler for authentication-related tools."""

    async def handle_auth(self, arguments: dict) -> list[types.TextContent]:
        """Handle auth tool with login, complete_login, check_status, and logout actions."""
        action = arguments.get("action")

        if action == "login":
            return await self._handle_login()
        elif action == "complete_login":
            return await self._handle_complete_login(arguments)
        elif action == "check_status":
            return await self._handle_check_status()
        elif action == "logout":
            return await self._handle_logout()
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'login', 'complete_login', 'check_status', or 'logout'."
            )

    async def _handle_login(self) -> list[types.TextContent]:
        """Handle login action."""
        try:
            logger.info("AuthHandler: Handling login request")
            result = await auth_manager.login()
            logger.info(f"AuthHandler: Login result: {result.get('status')}")
            return self._format_response(result)
        except Exception as e:
            logger.error(f"AuthHandler: Login failed with exception: {str(e)}")
            return self._format_error(f"Login failed: {str(e)}")

    async def _handle_complete_login(self, arguments: dict) -> list[types.TextContent]:
        """Handle complete_login action."""
        try:
            device_code = arguments.get("device_code")
            logger.info(f"AuthHandler: Handling complete_login with device_code: {device_code[:20] if device_code else 'None'}...")
            result = await auth_manager.complete_login(device_code)
            logger.info(f"AuthHandler: Complete login result: {result.get('status')}")
            return self._format_response(result)
        except Exception as e:
            logger.error(f"AuthHandler: Complete login failed with exception: {str(e)}")
            return self._format_error(f"Failed to complete login: {str(e)}")

    async def _handle_check_status(self) -> list[types.TextContent]:
        """Handle check_status action (read-only)."""
        try:
            logger.info("AuthHandler: Handling check_status (read-only)")
            result = await auth_manager.check_status()
            logger.info(f"AuthHandler: Check status result: {result.get('status')}")
            return self._format_response(result)
        except Exception as e:
            logger.error(f"AuthHandler: Check status failed with exception: {str(e)}")
            return self._format_error(f"Failed to check authentication status: {str(e)}")

    async def _handle_logout(self) -> list[types.TextContent]:
        """Handle logout action."""
        try:
            result = await auth_manager.logout()
            return self._format_response(result)
        except Exception as e:
            return self._format_error(f"Logout failed: {str(e)}")
