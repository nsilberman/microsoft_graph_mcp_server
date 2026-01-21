"""Authentication manager for Microsoft Graph API."""

import asyncio
import logging
import time
from datetime import datetime
from typing import Any, Dict, Optional
from urllib.parse import urlencode
from zoneinfo import ZoneInfo

import httpx
from msal import PublicClientApplication

from .token_manager import TokenManager
from .device_flow import DeviceFlowManager
from ..config import settings

logger = logging.getLogger(__name__)


class GraphAuthManager:
    """Manages authentication for Microsoft Graph API using device code flow."""

    def __init__(self):
        self.client_app = PublicClientApplication(
            client_id=settings.client_id,
            authority=f"{settings.auth_url}/{settings.tenant_id}",
        )
        self.token_manager = TokenManager()
        self.device_flow_manager = DeviceFlowManager(
            self.client_app, self.token_manager
        )

    async def get_access_token(self) -> str:
        """Get a valid access token, refreshing if necessary."""
        if not self.token_manager.authenticated or not self.token_manager.access_token:
            raise Exception(
                "Not authenticated. Please call the login tool first to authenticate with Microsoft Graph."
            )

        if self.token_manager.is_token_valid():
            return self.token_manager.access_token

        if self.token_manager.refresh_token:
            await self._refresh_token()
        else:
            raise Exception(
                "Token expired and no refresh token available. Please call the login tool again."
            )

        return self.token_manager.access_token

    async def extend_token(self) -> Dict[str, Any]:
        """Refresh the access token using the refresh token.

        This method explicitly refreshes the access token without requiring user login.
        It uses the stored refresh token to obtain a new access token from Microsoft.
        The new token will have a fresh lifetime (typically 1 hour).

        Returns:
            Dict with refresh status and token information including:
            - status: 'refreshed' if successful
            - authenticated: true
            - message: success message
            - token_expires_at: ISO timestamp of new token expiry in user's local timezone
            - time_remaining: dict with remaining time in seconds/minutes/hours
            - refresh_available: boolean indicating if refresh token is still available

        Raises:
            Exception: If not authenticated or no refresh token available
        """
        logger.info("AuthManager: extend_token called")

        if not self.token_manager.authenticated or not self.token_manager.access_token:
            raise Exception(
                "Not authenticated. Please call the login tool first to authenticate with Microsoft Graph."
            )

        if not self.token_manager.refresh_token:
            raise Exception(
                "No refresh token available. Please call the login tool to authenticate."
            )

        logger.info("AuthManager: Refreshing token")
        await self._refresh_token()

        expiry_info = self.token_manager.get_token_expiry_info()

        from ..graph_client import graph_client

        user_timezone = await graph_client.get_user_timezone()

        utc_expiry = datetime.utcfromtimestamp(self.token_manager.token_expiry)
        local_expiry = utc_expiry.replace(tzinfo=ZoneInfo("UTC")).astimezone(
            ZoneInfo(user_timezone)
        )
        token_expires_at_local = local_expiry.strftime("%Y-%m-%dT%H:%M:%S%z")

        return {
            "status": "refreshed",
            "authenticated": True,
            "message": "Successfully refreshed access token.",
            "token_expires_at": token_expires_at_local,
            "time_remaining": {
                "seconds": expiry_info["remaining_seconds"],
                "minutes": expiry_info["remaining_minutes"],
                "hours": expiry_info["remaining_hours"],
            },
            "refresh_available": self.token_manager.refresh_token is not None,
            "timezone": user_timezone,
        }

    async def logout(self) -> Dict[str, Any]:
        """Logout and clear authentication state."""
        self.token_manager.clear_tokens()
        self.device_flow_manager.clear_device_flow()

        return {
            "status": "logged_out",
            "message": "Successfully logged out from Microsoft Graph. Authentication state has been cleared.",
        }

    async def check_status(self) -> Dict[str, Any]:
        """Check current authentication status (read-only).

        Returns information about authentication state and token expiry without
        triggering any actions. Useful for debugging and monitoring.

        Returns:
            Dict with status information including:
            - authenticated: boolean indicating if authenticated
            - token_expires_at: ISO timestamp of token expiry in user's local timezone (if authenticated)
            - time_remaining: dict with remaining time in seconds/minutes/hours
            - refresh_available: boolean indicating if refresh token is available
            - timezone: user's timezone (if authenticated)
        """
        logger.info("AuthManager: check_status called (read-only)")

        if not self.token_manager.authenticated or not self.token_manager.access_token:
            return {
                "status": "not_authenticated",
                "authenticated": False,
                "message": "Not authenticated with Microsoft Graph. Please call the login tool first.",
            }

        expiry_info = self.token_manager.get_token_expiry_info()

        if expiry_info["remaining_seconds"] <= 0:
            return {
                "status": "token_expired",
                "authenticated": False,
                "message": "Authentication token has expired. Please call the login tool again.",
            }

        from ..graph_client import graph_client

        user_timezone = await graph_client.get_user_timezone()

        utc_expiry = datetime.utcfromtimestamp(self.token_manager.token_expiry)
        local_expiry = utc_expiry.replace(tzinfo=ZoneInfo("UTC")).astimezone(
            ZoneInfo(user_timezone)
        )
        token_expires_at_local = local_expiry.strftime("%Y-%m-%dT%H:%M:%S%z")

        return {
            "status": "authenticated",
            "authenticated": True,
            "message": "Successfully authenticated with Microsoft Graph.",
            "token_expires_at": token_expires_at_local,
            "time_remaining": {
                "seconds": expiry_info["remaining_seconds"],
                "minutes": expiry_info["remaining_minutes"],
                "hours": expiry_info["remaining_hours"],
            },
            "refresh_available": self.token_manager.refresh_token is not None,
            "timezone": user_timezone,
        }

    async def complete_login(self, device_code: Optional[str] = None) -> Dict[str, Any]:
        """Complete the login process by checking authentication status.

        This method waits for the user to complete browser authentication and
        finalizes the login process by acquiring the access token.

        Args:
            device_code: Optional device_code to load device flow from disk
        """
        logger.info(
            f"AuthManager: complete_login called with device_code: {device_code[:20] if device_code else 'None'}..."
        )
        return await self.device_flow_manager.check_login_status(device_code)

    async def login(self) -> Dict[str, Any]:
        """Explicit login method for authentication - creates device flow and returns URL/code."""
        logger.info("AuthManager: login called")
        # Clear all previous authentication state (user wants fresh login)
        self.token_manager.clear_tokens()
        self.device_flow_manager.clear_device_flow()

        # Create new device flow
        logger.info("AuthManager: Initiating device flow")
        result = await self.device_flow_manager.initiate_device_flow_only()
        logger.info(
            f"AuthManager: Device flow initiated with status: {result.get('status')}"
        )
        return result

    async def _acquire_token(self) -> None:
        """Acquire a new access token using device code flow."""
        try:
            flow = self.client_app.initiate_device_flow(
                scopes=["https://graph.microsoft.com/.default"]
            )

            if "error" in flow:
                raise Exception(f"Failed to initiate device flow: {flow['error']}")

            logger.info("=" * 70)
            logger.info("MICROSOFT GRAPH AUTHENTICATION")
            logger.info("=" * 70)
            logger.info(f"To sign in, use a web browser to open the page: {flow['verification_uri']}")
            logger.info(f"And enter the code: {flow['user_code']}")
            logger.info("=" * 70)
            logger.info("Waiting for authentication...")
            logger.info("=" * 70)

            result = self.client_app.acquire_token_by_device_flow(flow)

            if "access_token" in result:
                self.token_manager.update_token(
                    access_token=result["access_token"],
                    expires_in=result.get("expires_in", 3600),
                    refresh_token=result.get("refresh_token"),
                )
                logger.info("✓ Authentication successful!")
            else:
                raise Exception(
                    f"Failed to acquire token: {result.get('error_description', 'Unknown error')}"
                )
        except Exception as e:
            raise Exception(f"Authentication failed: {str(e)}")

    async def _refresh_token(self) -> None:
        """Refresh the access token using the refresh token.

        Raises:
            Exception: If refresh token is not available or refresh fails
        """
        if not self.token_manager.refresh_token:
            raise Exception(
                "No refresh token available. Please call the login tool to authenticate."
            )

        try:
            result = self.client_app.acquire_token_by_refresh_token(
                self.token_manager.refresh_token,
                scopes=["https://graph.microsoft.com/.default"],
            )

            if "access_token" in result:
                self.token_manager.update_token(
                    access_token=result["access_token"],
                    expires_in=result.get("expires_in", 3600),
                    refresh_token=result.get(
                        "refresh_token", self.token_manager.refresh_token
                    ),
                )
            else:
                error_msg = result.get("error_description", "Unknown error")
                raise Exception(f"Failed to refresh token: {error_msg}")
        except Exception as e:
            raise Exception(f"Failed to refresh token: {str(e)}")

    def get_auth_url(self, state: Optional[str] = None) -> str:
        """Get the authorization URL for interactive authentication."""
        auth_url = f"{settings.auth_url}/{settings.tenant_id}/oauth2/v2.0/authorize"

        params = {
            "client_id": settings.client_id,
            "response_type": "code",
            "scope": "https://graph.microsoft.com/.default",
            "state": state or "default_state",
        }

        return f"{auth_url}?{urlencode(params)}"

    async def exchange_code_for_token(
        self, auth_code: str, redirect_uri: str
    ) -> Dict[str, Any]:
        """Exchange authorization code for access token."""
        token_url = f"{settings.auth_url}/{settings.tenant_id}/oauth2/v2.0/token"

        data = {
            "client_id": settings.client_id,
            "code": auth_code,
            "grant_type": "authorization_code",
            "redirect_uri": redirect_uri,
        }

        async with httpx.AsyncClient() as client:
            response = await client.post(token_url, data=data)

            if response.status_code == 200:
                result = response.json()
                self.token_manager.update_token(
                    access_token=result["access_token"],
                    expires_in=result.get("expires_in", 3600),
                    refresh_token=result.get("refresh_token"),
                )
                return result
            else:
                raise Exception(f"Token exchange failed: {response.text}")


# Global auth manager instance
auth_manager = GraphAuthManager()
