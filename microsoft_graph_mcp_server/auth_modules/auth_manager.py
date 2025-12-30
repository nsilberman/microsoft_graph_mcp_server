"""Authentication manager for Microsoft Graph API."""

import asyncio
import time
from typing import Any, Dict, Optional
from urllib.parse import urlencode

import httpx
from msal import PublicClientApplication

from .token_manager import TokenManager
from .device_flow import DeviceFlowManager
from ..config import settings


class GraphAuthManager:
    """Manages authentication for Microsoft Graph API using device code flow."""
    
    def __init__(self):
        self.client_app = PublicClientApplication(
            client_id=settings.client_id,
            authority=f"{settings.auth_url}/{settings.tenant_id}"
        )
        self.token_manager = TokenManager()
        self.device_flow_manager = DeviceFlowManager(self.client_app, self.token_manager)
    
    async def get_access_token(self) -> str:
        """Get a valid access token, refreshing if necessary."""
        if not self.token_manager.authenticated or not self.token_manager.access_token:
            raise Exception("Not authenticated. Please call the login tool first to authenticate with Microsoft Graph.")
        
        if self.token_manager.is_token_valid():
            return self.token_manager.access_token
        
        if self.token_manager.refresh_token:
            await self._refresh_token()
        else:
            raise Exception("Token expired and no refresh token available. Please call the login tool again.")
        
        return self.token_manager.access_token
    
    async def logout(self) -> Dict[str, Any]:
        """Logout and clear authentication state."""
        self.token_manager.clear_tokens()
        self.device_flow_manager.clear_device_flow()
        
        return {
            "status": "logged_out",
            "message": "Successfully logged out from Microsoft Graph. Authentication state has been cleared."
        }
    
    async def check_login_status(self) -> Dict[str, Any]:
        """Check the current login status."""
        return await self.device_flow_manager.check_login_status()
    
    async def login(self) -> Dict[str, Any]:
        """Explicit login method for authentication."""
        if self.token_manager.authenticated and self.token_manager.access_token and self.token_manager.is_token_valid():
            expiry_info = self.token_manager.get_token_expiry_info()
            expiry_datetime = __import__("datetime").datetime.fromtimestamp(self.token_manager.token_expiry)
            expiry_str = expiry_datetime.strftime("%Y-%m-%d %H:%M:%S")
            
            remaining_hours = expiry_info["remaining_hours"]
            remaining_minutes = expiry_info["remaining_minutes"]
            
            if remaining_hours > 0:
                time_remaining = f"{remaining_hours} hour{'s' if remaining_hours > 1 else ''} and {remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
            else:
                time_remaining = f"{remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
            
            return {
                "status": "already_authenticated",
                "message": f"You are already authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                "token_expiry": self.token_manager.token_expiry,
                "expiry_datetime": expiry_str,
                **expiry_info
            }
        
        if self.device_flow_manager.device_flow is None:
            return await self.device_flow_manager.initiate_device_code()
        else:
            return await self.device_flow_manager.check_authentication_status()
    
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
                    refresh_token=result.get("refresh_token")
                )
                print("\n✓ Authentication successful!\n")
            else:
                raise Exception(f"Failed to acquire token: {result.get('error_description', 'Unknown error')}")
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
                scopes=["https://graph.microsoft.com/.default"]
            )
            
            if "access_token" in result:
                self.token_manager.update_token(
                    access_token=result["access_token"],
                    expires_in=result.get("expires_in", 3600),
                    refresh_token=result.get("refresh_token", self.token_manager.refresh_token)
                )
            else:
                await self._acquire_token()
        except Exception:
            await self._acquire_token()
    
    def get_auth_url(self, state: Optional[str] = None) -> str:
        """Get the authorization URL for interactive authentication."""
        auth_url = (
            f"{settings.auth_url}/{settings.tenant_id}/oauth2/v2.0/authorize"
        )
        
        params = {
            "client_id": settings.client_id,
            "response_type": "code",
            "scope": "https://graph.microsoft.com/.default",
            "state": state or "default_state"
        }
        
        return f"{auth_url}?{urlencode(params)}"
    
    async def exchange_code_for_token(self, auth_code: str, redirect_uri: str) -> Dict[str, Any]:
        """Exchange authorization code for access token."""
        token_url = f"{settings.auth_url}/{settings.tenant_id}/oauth2/v2.0/token"
        
        data = {
            "client_id": settings.client_id,
            "code": auth_code,
            "grant_type": "authorization_code",
            "redirect_uri": redirect_uri
        }
        
        async with httpx.AsyncClient() as client:
            response = await client.post(token_url, data=data)
            
            if response.status_code == 200:
                result = response.json()
                self.token_manager.update_token(
                    access_token=result["access_token"],
                    expires_in=result.get("expires_in", 3600),
                    refresh_token=result.get("refresh_token")
                )
                return result
            else:
                raise Exception(f"Token exchange failed: {response.text}")


# Global auth manager instance
auth_manager = GraphAuthManager()
