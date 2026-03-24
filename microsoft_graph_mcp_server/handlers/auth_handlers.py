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
        """Handle auth tool with start, complete, check_status, and logout actions.
        
        Simplified authentication workflow:
        - start: Initiate device code flow, returns verification URL and code
        - complete: Complete the login process after user authenticates in browser
        - check_status: Check auth status and auto-refresh if needed
        - logout: Clear all authentication tokens
        """
        action = arguments.get("action")

        if action == "start":
            return await self._handle_start()
        elif action == "complete":
            return await self._handle_complete(arguments)
        elif action == "check_status":
            return await self._handle_check_status()
        elif action == "logout":
            return await self._handle_logout()
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'start', 'complete', 'check_status', or 'logout'."
            )

    async def _handle_start(self) -> list[types.TextContent]:
        """Handle start action - initiate device code flow."""
        try:
            logger.info("AuthHandler: Handling start request")
            result = await auth_manager.start_auth()
            logger.info(f"AuthHandler: Start result: {result.get('status')}")
            return self._format_response(result)
        except Exception as e:
            logger.error(f"AuthHandler: Start failed with exception: {str(e)}")
            return self._format_error(f"Start failed: {str(e)}")

    async def _handle_complete(self, arguments: dict) -> list[types.TextContent]:
        """Handle complete action - finalize the login process."""
        try:
            device_code = arguments.get("device_code")
            logger.info(
                f"AuthHandler: Handling complete with device_code: {device_code[:20] if device_code else 'None'}..."
            )
            result = await auth_manager.complete_auth(device_code)
            logger.info(f"AuthHandler: Complete result: {result.get('status')}")
            return self._format_response(result)
        except Exception as e:
            logger.error(f"AuthHandler: Complete failed with exception: {str(e)}")
            return self._format_error(f"Failed to complete authentication: {str(e)}")

    async def _handle_check_status(self) -> list[types.TextContent]:
        """Handle check_status action - check auth status and auto-refresh if needed."""
        try:
            logger.info("AuthHandler: Handling check_status request")
            result = await auth_manager.check_status()
            logger.info(f"AuthHandler: Check_status result: {result.get('status')}")
            return self._format_response(result)
        except Exception as e:
            logger.error(f"AuthHandler: Check_status failed with exception: {str(e)}")
            return self._format_error(f"Failed to check status: {str(e)}")

    async def _handle_logout(self) -> list[types.TextContent]:
        """Handle logout action - clear all tokens."""
        try:
            logger.info("AuthHandler: Handling logout request")
            result = await auth_manager.logout()
            logger.info("AuthHandler: Logout successful")
            return self._format_response(result)
        except Exception as e:
            logger.error(f"AuthHandler: Logout failed with exception: {str(e)}")
            return self._format_error(f"Logout failed: {str(e)}")
