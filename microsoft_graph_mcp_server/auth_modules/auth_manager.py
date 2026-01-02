"""Authentication manager for Microsoft Graph API."""

import asyncio
import logging
import time
from typing import Any, Dict, Optional
from urllib.parse import urlencode

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
            - token_expires_at: ISO timestamp of token expiry (if authenticated)
            - time_remaining: dict with remaining time in seconds/minutes/hours
            - refresh_available: boolean indicating if refresh token is available
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

        return {
            "status": "authenticated",
            "authenticated": True,
            "message": "Successfully authenticated with Microsoft Graph.",
            "token_expires_at": time.strftime(
                "%Y-%m-%dT%H:%M:%SZ", time.gmtime(self.token_manager.token_expiry)
            ),
            "time_remaining": {
                "seconds": expiry_info["remaining_seconds"],
                "minutes": expiry_info["remaining_minutes"],
                "hours": expiry_info["remaining_hours"],
            },
            "refresh_available": self.token_manager.refresh_token is not None,
        }

    async def complete_login(self, device_code: Optional[str] = None) -> Dict[str, Any]:
        """Complete the login process by checking authentication status.

        This method waits for the user to complete browser authentication and
        finalizes the login process by acquiring the access token.

        Args:
            device_code: Optional device_code to load device flow from disk
        """
        logger.info(f"AuthManager: complete_login called with device_code: {device_code[:20] if device_code else 'None'}...")
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
        logger.info(f"AuthManager: Device flow initiated with status: {result.get('status')}")
        return result

    async def _acquire_token(self) -> None:
        """Acquire a new access token using device code flow."""
        try:
            flow = self.client_app.initiate_device_flow(
                scopes=["https://graph.microsoft.com/.default"]
            )

            if "error" in flow:
                raise Exception(f"Failed to initiate device flow: {flow['error']}")

            print("\n" + "=" * 70)
            print("MICROSOFT GRAPH AUTHENTICATION")
            print("=" * 70)
            print(f"\nTo sign in, use a web browser to open the page:")
            print(f"\n{flow['verification_uri']}")
            print(f"\nAnd enter the code:")
            print(f"\n{flow['user_code']}")
            print("\n" + "=" * 70)
            print("Waiting for authentication...")
            print("=" * 70 + "\n")

            result = self.client_app.acquire_token_by_device_flow(flow)

            if "access_token" in result:
                self.token_manager.update_token(
                    access_token=result["access_token"],
                    expires_in=result.get("expires_in", 3600),
                    refresh_token=result.get("refresh_token"),
                )
                print("\n✓ Authentication successful!\n")
            else:
                raise Exception(
                    f"Failed to acquire token: {result.get('error_description', 'Unknown error')}"
                )
        except Exception as e:
            raise Exception(f"Authentication failed: {str(e)}")

    async def _refresh_token(self) -> None:
        """Refresh the access token using the refresh token."""
        if not self.token_manager.refresh_token:
            await self._acquire_token()
            return

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
                await self._acquire_token()
        except Exception:
            await self._acquire_token()

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
