"""Authentication manager for Microsoft Graph API.

Simplified authentication workflow with 4 actions:
- start: Initiate device code flow, returns verification URL and code
- complete: Complete the login process after user authenticates in browser  
- check_status: Check authentication status and refresh if needed
- logout: Clear all authentication tokens

Auto-refresh: Access tokens are automatically refreshed when expired if a refresh_token exists.
"""

import logging
from typing import Any, Dict, Optional

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

    # =========================================================================
    # Public API - 4 Main Actions
    # =========================================================================

    async def start_auth(self) -> Dict[str, Any]:
        """Start authentication - initiate device code flow.
        
        Returns verification URL and user code for browser authentication.
        After user completes authentication in browser, call complete_auth().
        
        Returns:
            Dict with:
            - status: 'pending'
            - verification_uri: URL to open in browser
            - user_code: Code to enter on the verification page
            - expires_in: Time until the code expires (seconds)
            - message: Instructions for the user
        """
        logger.info("AuthManager: start_auth called")
        
        # Check if already authenticated with valid token
        self.token_manager.load_tokens_from_disk()
        if self.token_manager.is_token_valid():
            logger.info("Already authenticated with valid token")
            return {
                "status": "authenticated",
                "authenticated": True,
                "message": "Already authenticated with Microsoft Graph.",
            }
        
        # Check if we have a refresh_token that can be used
        if self.token_manager.refresh_token:
            logger.info("Found refresh_token, attempting auto-refresh instead of new login")
            try:
                await self._refresh_token_internal()
                return {
                    "status": "authenticated",
                    "authenticated": True,
                    "message": "Token refreshed automatically.",
                }
            except Exception as e:
                logger.warning(f"Auto-refresh failed: {e}, proceeding with new login")
        
        # Clear previous state and start new device flow
        self.token_manager.clear_tokens()
        self.device_flow_manager.clear_device_flow()
        
        logger.info("Initiating new device flow")
        result = await self.device_flow_manager.initiate_device_flow_only()
        logger.info(f"Device flow initiated with status: {result.get('status')}")
        return result

    async def complete_auth(self, device_code: Optional[str] = None) -> Dict[str, Any]:
        """Complete authentication - finalize login after browser authentication.
        
        Call this after user completes authentication in browser.
        
        Args:
            device_code: Optional device_code (will auto-load if not provided)
            
        Returns:
            Dict with authentication status and token info
        """
        logger.info(
            f"AuthManager: complete_auth called with device_code: {device_code[:20] if device_code else 'None'}..."
        )
        return await self.device_flow_manager.check_login_status(device_code)

    async def check_status(self) -> Dict[str, Any]:
        """Check authentication status and refresh token if needed.
        
        Logic:
        1. If token is still valid → return "authenticated" (no refresh needed)
        2. If token expired → try refresh using refresh_token
           - Success → return "authenticated" 
           - Fail → return status indicating user needs to login
        
        Returns:
            Dict with:
            - status: 'authenticated' | 'not_authenticated'
            - authenticated: boolean
            - message: Status message
        """
        logger.info("AuthManager: check_status called")
        
        # Load latest tokens
        self.token_manager.load_tokens_from_disk()
        
        # No refresh_token available → need to login
        if not self.token_manager.refresh_token:
            return {
                "status": "not_authenticated",
                "authenticated": False,
                "message": "Not authenticated. Please call auth with action='start' to login.",
            }
        
        # Token is still valid → already authenticated
        if self.token_manager.is_token_valid():
            logger.info("Token still valid")
            
            return {
                "status": "authenticated",
                "authenticated": True,
                "message": "Already authenticated.",
            }
        
        # Token expired → try to refresh
        logger.info("Token expired, attempting refresh...")
        try:
            await self._refresh_token_internal()
            logger.info("Token refresh successful")
            
            return {
                "status": "authenticated",
                "authenticated": True,
                "message": "Token refreshed.",
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Token refresh failed: {error_msg}")
            
            if "invalid_grant" in error_msg.lower() or "expired" in error_msg.lower():
                self.token_manager.clear_tokens()
            
            return {
                "status": "not_authenticated",
                "authenticated": False,
                "message": "Session expired. Please call auth with action='start' to login again.",
            }

    async def logout(self) -> Dict[str, Any]:
        """Logout - clear all authentication tokens.
        
        Returns:
            Dict with logout confirmation
        """
        logger.info("AuthManager: logout called")
        self.token_manager.clear_tokens()
        self.device_flow_manager.clear_device_flow()
        
        return {
            "status": "logged_out",
            "authenticated": False,
            "message": "Successfully logged out. Authentication tokens have been cleared.",
        }

    # =========================================================================
    # Internal Methods
    # =========================================================================

    async def get_access_token(self) -> str:
        """Get a valid access token, automatically refreshing if necessary.

        This is used internally by other modules. It will automatically use
        refresh_token to get a new access_token if the current token is expired.
        """
        # Load tokens from disk
        self.token_manager.load_tokens_from_disk()

        # If we have a valid token, return it directly
        if self.token_manager.is_token_valid():
            return self.token_manager.access_token

        # Token expired but we have refresh_token, try to refresh automatically
        if self.token_manager.refresh_token:
            logger.info("Access token expired, attempting auto-refresh using refresh_token...")
            try:
                await self._refresh_token_internal()
                logger.info("Token auto-refresh successful")
                return self.token_manager.access_token
            except Exception as e:
                logger.error(f"Auto-refresh failed: {e}")
                if "invalid_grant" in str(e).lower() or "expired" in str(e).lower():
                    self.token_manager.clear_tokens()
                raise Exception(
                    f"Token refresh failed: {str(e)}. Please login again with auth action='start'."
                )

        # No valid token and no refresh_token available
        raise Exception(
            "Not authenticated. Please call auth with action='start' to login."
        )

    async def _refresh_token_internal(self) -> None:
        """Internal method to refresh the access token using the refresh token."""
        if not self.token_manager.refresh_token:
            raise Exception("No refresh token available.")

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
            logger.info("Token refreshed successfully")
        else:
            error = result.get("error", "unknown")
            error_description = result.get("error_description", "Unknown error")

            if error == "invalid_grant":
                logger.warning("Refresh token invalid or expired, clearing tokens")
                self.token_manager.clear_tokens()
                raise Exception("Refresh token expired or invalid. Please login again.")

            raise Exception(f"Failed to refresh token: {error_description}")


# Global auth manager instance
auth_manager = GraphAuthManager()
