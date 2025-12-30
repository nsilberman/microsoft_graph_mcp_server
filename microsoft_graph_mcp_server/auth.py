"""Authentication module for Microsoft Graph API."""

import asyncio
import datetime
import json
import os
import time
from pathlib import Path
from typing import Dict, Optional, Any
from urllib.parse import urlencode

import httpx
from msal import PublicClientApplication

from .config import settings


TOKEN_FILE = Path.home() / ".microsoft_graph_mcp_tokens.json"


class GraphAuthManager:
    """Manages authentication for Microsoft Graph API using device code flow."""
    
    def __init__(self):
        self.client_app = PublicClientApplication(
            client_id=settings.client_id,
            authority=f"{settings.auth_url}/{settings.tenant_id}"
        )
        self.access_token: Optional[str] = None
        self.token_expiry: float = 0
        self.refresh_token: Optional[str] = None
        self.authenticated: bool = False
        self.device_flow: Optional[Dict[str, Any]] = None
        
        self._load_tokens_from_disk()
    
    def _save_tokens_to_disk(self) -> None:
        """Save authentication tokens to disk."""
        try:
            token_data = {
                "access_token": self.access_token,
                "refresh_token": self.refresh_token,
                "token_expiry": self.token_expiry,
                "authenticated": self.authenticated
            }
            with open(TOKEN_FILE, "w") as f:
                json.dump(token_data, f, indent=2)
        except Exception as e:
            print(f"Warning: Failed to save tokens to disk: {e}")
    
    def _load_tokens_from_disk(self) -> None:
        """Load authentication tokens from disk."""
        if not TOKEN_FILE.exists():
            return
        
        try:
            with open(TOKEN_FILE, "r") as f:
                token_data = json.load(f)
            
            self.access_token = token_data.get("access_token")
            self.refresh_token = token_data.get("refresh_token")
            self.token_expiry = token_data.get("token_expiry", 0)
            self.authenticated = token_data.get("authenticated", False)
            
            if self.authenticated and self.access_token:
                current_time = time.time()
                if current_time >= self.token_expiry - 60:
                    self.authenticated = False
                    self.access_token = None
                    self._delete_tokens_from_disk()
        except Exception as e:
            print(f"Warning: Failed to load tokens from disk: {e}")
            self._delete_tokens_from_disk()
    
    def _delete_tokens_from_disk(self) -> None:
        """Delete authentication tokens from disk."""
        try:
            if TOKEN_FILE.exists():
                os.remove(TOKEN_FILE)
        except Exception as e:
            print(f"Warning: Failed to delete tokens from disk: {e}")
    
    async def get_access_token(self) -> str:
        """Get a valid access token, refreshing if necessary."""
        if not self.authenticated or not self.access_token:
            raise Exception("Not authenticated. Please call the login tool first to authenticate with Microsoft Graph.")
        
        if time.time() < self.token_expiry - 60:
            return self.access_token
        
        if self.refresh_token:
            await self._refresh_token()
        else:
            raise Exception("Token expired and no refresh token available. Please call the login tool again.")
        
        return self.access_token
    
    async def logout(self) -> Dict[str, Any]:
        """Logout and clear authentication state."""
        self.access_token = None
        self.token_expiry = 0
        self.refresh_token = None
        self.authenticated = False
        self.device_flow = None
        
        self._delete_tokens_from_disk()
        
        return {
            "status": "logged_out",
            "message": "Successfully logged out from Microsoft Graph. Authentication state has been cleared."
        }
    
    async def check_login_status(self) -> Dict[str, Any]:
        if self.device_flow is not None:
            try:
                loop = asyncio.get_event_loop()
                
                async def acquire_with_timeout():
                    try:
                        result = await asyncio.wait_for(
                            loop.run_in_executor(
                                None,
                                self.client_app.acquire_token_by_device_flow,
                                self.device_flow
                            ),
                            timeout=3.0
                        )
                        return result
                    except asyncio.TimeoutError:
                        return {"error": "timeout", "error_description": "Authentication still pending"}
                
                result = await acquire_with_timeout()
                
                if "access_token" in result:
                    self.access_token = result["access_token"]
                    self.token_expiry = time.time() + result.get("expires_in", 3600)
                    self.refresh_token = result.get("refresh_token")
                    self.authenticated = True
                    self.device_flow = None
                    
                    self._save_tokens_to_disk()
                    
                    remaining_seconds = int(self.token_expiry - time.time())
                    remaining_minutes = remaining_seconds // 60
                    remaining_hours = remaining_minutes // 60
                    remaining_minutes = remaining_minutes % 60
                    
                    expiry_datetime = datetime.datetime.fromtimestamp(self.token_expiry)
                    expiry_str = expiry_datetime.strftime("%Y-%m-%d %H:%M:%S")
                    
                    if remaining_hours > 0:
                        time_remaining = f"{remaining_hours} hour{'s' if remaining_hours > 1 else ''} and {remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
                    else:
                        time_remaining = f"{remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
                    
                    return {
                        "status": "authenticated",
                        "message": f"Successfully authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                        "token_expiry": self.token_expiry,
                        "expiry_datetime": expiry_str,
                        "remaining_seconds": remaining_seconds,
                        "remaining_minutes": remaining_minutes,
                        "remaining_hours": remaining_hours
                    }
                else:
                    error = result.get("error", "")
                    if error in ["authorization_pending", "timeout"]:
                        return {
                            "status": "pending",
                            "message": "Authentication is pending. Please complete authentication in the browser using the verification link and code, then call login again to verify.",
                            "verification_uri": self.device_flow.get("verification_uri", ""),
                            "user_code": self.device_flow.get("user_code", "")
                        }
                    else:
                        self.device_flow = None
                        return {
                            "status": "failed",
                            "message": f"Authentication failed: {result.get('error_description', 'Unknown error')}"
                        }
            except Exception as e:
                self.device_flow = None
                return {
                    "status": "error",
                    "message": f"Authentication check failed: {str(e)}"
                }
        
        if not self.authenticated or not self.access_token:
            return {
                "status": "not_authenticated",
                "message": "Not authenticated with Microsoft Graph. Please call the login tool first."
            }
        
        if time.time() >= self.token_expiry - 60:
            return {
                "status": "expired",
                "message": "Authentication token has expired. Please call the login tool to re-authenticate."
            }
        
        remaining_seconds = int(self.token_expiry - time.time())
        remaining_minutes = remaining_seconds // 60
        remaining_hours = remaining_minutes // 60
        remaining_minutes = remaining_minutes % 60
        
        expiry_datetime = datetime.datetime.fromtimestamp(self.token_expiry)
        expiry_str = expiry_datetime.strftime("%Y-%m-%d %H:%M:%S")
        
        if remaining_hours > 0:
            time_remaining = f"{remaining_hours} hour{'s' if remaining_hours > 1 else ''} and {remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
        else:
            time_remaining = f"{remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
        
        return {
            "status": "authenticated",
            "message": f"Authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
            "token_expiry": self.token_expiry,
            "expiry_datetime": expiry_str,
            "remaining_seconds": remaining_seconds,
            "remaining_minutes": remaining_minutes,
            "remaining_hours": remaining_hours
        }
    
    async def login(self) -> Dict[str, Any]:
        """Explicit login method for authentication."""
        if self.authenticated and self.access_token and time.time() < self.token_expiry - 60:
            remaining_seconds = int(self.token_expiry - time.time())
            remaining_minutes = remaining_seconds // 60
            remaining_hours = remaining_minutes // 60
            remaining_minutes = remaining_minutes % 60
            
            expiry_datetime = datetime.datetime.fromtimestamp(self.token_expiry)
            expiry_str = expiry_datetime.strftime("%Y-%m-%d %H:%M:%S")
            
            if remaining_hours > 0:
                time_remaining = f"{remaining_hours} hour{'s' if remaining_hours > 1 else ''} and {remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
            else:
                time_remaining = f"{remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
            
            return {
                "status": "already_authenticated",
                "message": f"You are already authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                "token_expiry": self.token_expiry,
                "expiry_datetime": expiry_str,
                "remaining_seconds": remaining_seconds,
                "remaining_minutes": remaining_minutes,
                "remaining_hours": remaining_hours
            }
        
        if self.device_flow is None:
            return await self._initiate_device_code()
        else:
            return await self._check_authentication_status()
    
    async def _initiate_device_code(self) -> Dict[str, Any]:
        """Initiate device code flow and return authentication info."""
        try:
            flow = self.client_app.initiate_device_flow(
                scopes=["https://graph.microsoft.com/.default"]
            )
            
            if "error" in flow:
                raise Exception(f"Failed to initiate device flow: {flow['error']}")
            
            self.device_flow = flow
            
            return {
                "status": "pending",
                "message": "Please complete authentication using the link and code below. After completing authentication, call login again to verify.",
                "verification_uri": flow.get("verification_uri", ""),
                "user_code": flow.get("user_code", ""),
                "expires_in": flow.get("expires_in", 900),
                "interval": flow.get("interval", 5)
            }
        except Exception as e:
            raise Exception(f"Authentication failed: {str(e)}")
    
    async def _check_authentication_status(self) -> Dict[str, Any]:
        """Check if authentication is complete by attempting to acquire token with short timeout."""
        if self.device_flow is None:
            return {
                "status": "error",
                "message": "No pending device code flow. Please call login again."
            }
        
        try:
            loop = asyncio.get_event_loop()
            
            async def acquire_with_timeout():
                try:
                    result = await asyncio.wait_for(
                        loop.run_in_executor(
                            None,
                            self.client_app.acquire_token_by_device_flow,
                            self.device_flow
                        ),
                        timeout=3.0
                    )
                    return result
                except asyncio.TimeoutError:
                    return {"error": "timeout", "error_description": "Authentication still pending"}
            
            result = await acquire_with_timeout()
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                self.token_expiry = time.time() + result.get("expires_in", 3600)
                self.refresh_token = result.get("refresh_token")
                self.authenticated = True
                self.device_flow = None
                
                self._save_tokens_to_disk()
                
                remaining_seconds = int(self.token_expiry - time.time())
                remaining_minutes = remaining_seconds // 60
                remaining_hours = remaining_minutes // 60
                remaining_minutes = remaining_minutes % 60
                
                expiry_datetime = datetime.datetime.fromtimestamp(self.token_expiry)
                expiry_str = expiry_datetime.strftime("%Y-%m-%d %H:%M:%S")
                
                if remaining_hours > 0:
                    time_remaining = f"{remaining_hours} hour{'s' if remaining_hours > 1 else ''} and {remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
                else:
                    time_remaining = f"{remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
                
                return {
                    "status": "success",
                    "message": f"Successfully authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                    "token_expiry": self.token_expiry,
                    "expiry_datetime": expiry_str,
                    "remaining_seconds": remaining_seconds,
                    "remaining_minutes": remaining_minutes,
                    "remaining_hours": remaining_hours
                }
            else:
                error = result.get("error", "")
                if error in ["authorization_pending", "timeout"]:
                    return {
                        "status": "pending",
                        "message": "Authentication still pending. Please complete authentication in the browser and call login again.",
                        "verification_uri": self.device_flow.get("verification_uri", ""),
                        "user_code": self.device_flow.get("user_code", "")
                    }
                else:
                    self.device_flow = None
                    return {
                        "status": "failed",
                        "message": f"Authentication failed: {result.get('error_description', 'Unknown error')}"
                    }
        except Exception as e:
            self.device_flow = None
            raise Exception(f"Authentication failed: {str(e)}")
    
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
                self.access_token = result["access_token"]
                self.token_expiry = time.time() + result.get("expires_in", 3600)
                self.refresh_token = result.get("refresh_token")
                self.authenticated = True
                self._save_tokens_to_disk()
                print("\n✓ Authentication successful!\n")
            else:
                raise Exception(f"Failed to acquire token: {result.get('error_description', 'Unknown error')}")
        except Exception as e:
            raise Exception(f"Authentication failed: {str(e)}")
    
    async def _refresh_token(self) -> None:
        """Refresh the access token using the refresh token."""
        if not self.refresh_token:
            await self._acquire_token()
            return
        
        try:
            result = self.client_app.acquire_token_by_refresh_token(
                self.refresh_token,
                scopes=["https://graph.microsoft.com/.default"]
            )
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                self.token_expiry = time.time() + result.get("expires_in", 3600)
                self.refresh_token = result.get("refresh_token", self.refresh_token)
                self._save_tokens_to_disk()
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
                self.access_token = result["access_token"]
                self.token_expiry = time.time() + result.get("expires_in", 3600)
                self.refresh_token = result.get("refresh_token")
                return result
            else:
                raise Exception(f"Token exchange failed: {response.text}")


# Global auth manager instance
auth_manager = GraphAuthManager()