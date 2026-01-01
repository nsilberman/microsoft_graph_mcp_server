"""Device flow authentication for Microsoft Graph API."""

import asyncio
import datetime
import time
from typing import Any, Dict, Optional

from .token_manager import TokenManager


class DeviceFlowManager:
    """Manages device code flow authentication for Microsoft Graph API."""

    def __init__(self, client_app, token_manager: TokenManager):
        self.client_app = client_app
        self.token_manager = token_manager
        self.device_flow: Optional[Dict[str, Any]] = None

    async def initiate_device_code(self) -> Dict[str, Any]:
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
                "interval": flow.get("interval", 5),
            }
        except Exception as e:
            raise Exception(f"Authentication failed: {str(e)}")

    async def check_authentication_status(self) -> Dict[str, Any]:
        """Check if authentication is complete by attempting to acquire token with short timeout."""
        if self.device_flow is None:
            return {
                "status": "error",
                "message": "No pending device code flow. Please call login again.",
            }

        try:
            loop = asyncio.get_event_loop()

            async def acquire_with_timeout():
                try:
                    result = await asyncio.wait_for(
                        loop.run_in_executor(
                            None,
                            self.client_app.acquire_token_by_device_flow,
                            self.device_flow,
                        ),
                        timeout=3.0,
                    )
                    return result
                except asyncio.TimeoutError:
                    return {
                        "error": "timeout",
                        "error_description": "Authentication still pending",
                    }

            result = await acquire_with_timeout()

            if "access_token" in result:
                self.token_manager.update_token(
                    access_token=result["access_token"],
                    expires_in=result.get("expires_in", 3600),
                    refresh_token=result.get("refresh_token"),
                )
                self.device_flow = None

                expiry_info = self.token_manager.get_token_expiry_info()
                expiry_datetime = datetime.datetime.fromtimestamp(
                    self.token_manager.token_expiry
                )
                expiry_str = expiry_datetime.strftime("%Y-%m-%d %H:%M:%S")

                remaining_hours = expiry_info["remaining_hours"]
                remaining_minutes = expiry_info["remaining_minutes"]

                if remaining_hours > 0:
                    time_remaining = f"{remaining_hours} hour{'s' if remaining_hours > 1 else ''} and {remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
                else:
                    time_remaining = f"{remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"

                return {
                    "status": "success",
                    "message": f"Successfully authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                    "token_expiry": self.token_manager.token_expiry,
                    "expiry_datetime": expiry_str,
                    **expiry_info,
                }
            else:
                error = result.get("error", "")
                if error in ["authorization_pending", "timeout"]:
                    return {
                        "status": "pending",
                        "message": "Authentication still pending. Please complete authentication in the browser and call login again.",
                        "verification_uri": self.device_flow.get(
                            "verification_uri", ""
                        ),
                        "user_code": self.device_flow.get("user_code", ""),
                    }
                else:
                    self.device_flow = None
                    return {
                        "status": "failed",
                        "message": f"Authentication failed: {result.get('error_description', 'Unknown error')}",
                    }
        except Exception as e:
            self.device_flow = None
            raise Exception(f"Authentication failed: {str(e)}")

    async def check_login_status(self) -> Dict[str, Any]:
        """Check the current login status."""
        if self.device_flow is not None:
            try:
                loop = asyncio.get_event_loop()

                async def acquire_with_timeout():
                    try:
                        result = await asyncio.wait_for(
                            loop.run_in_executor(
                                None,
                                self.client_app.acquire_token_by_device_flow,
                                self.device_flow,
                            ),
                            timeout=3.0,
                        )
                        return result
                    except asyncio.TimeoutError:
                        return {
                            "error": "timeout",
                            "error_description": "Authentication still pending",
                        }

                result = await acquire_with_timeout()

                if "access_token" in result:
                    self.token_manager.update_token(
                        access_token=result["access_token"],
                        expires_in=result.get("expires_in", 3600),
                        refresh_token=result.get("refresh_token"),
                    )
                    self.device_flow = None

                    expiry_info = self.token_manager.get_token_expiry_info()
                    expiry_datetime = datetime.datetime.fromtimestamp(
                        self.token_manager.token_expiry
                    )
                    expiry_str = expiry_datetime.strftime("%Y-%m-%d %H:%M:%S")

                    remaining_hours = expiry_info["remaining_hours"]
                    remaining_minutes = expiry_info["remaining_minutes"]

                    if remaining_hours > 0:
                        time_remaining = f"{remaining_hours} hour{'s' if remaining_hours > 1 else ''} and {remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
                    else:
                        time_remaining = f"{remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"

                    return {
                        "status": "authenticated",
                        "message": f"Successfully authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                        "token_expiry": self.token_manager.token_expiry,
                        "expiry_datetime": expiry_str,
                        **expiry_info,
                    }
                else:
                    error = result.get("error", "")
                    if error in ["authorization_pending", "timeout"]:
                        return {
                            "status": "pending",
                            "message": "Authentication is pending. Please complete authentication in the browser using the verification link and code, then call login again to verify.",
                            "verification_uri": self.device_flow.get(
                                "verification_uri", ""
                            ),
                            "user_code": self.device_flow.get("user_code", ""),
                        }
                    else:
                        self.device_flow = None
                        return {
                            "status": "failed",
                            "message": f"Authentication failed: {result.get('error_description', 'Unknown error')}",
                        }
            except Exception as e:
                self.device_flow = None
                return {
                    "status": "error",
                    "message": f"Authentication check failed: {str(e)}",
                }

        if not self.token_manager.authenticated or not self.token_manager.access_token:
            return {
                "status": "not_authenticated",
                "message": "Not authenticated with Microsoft Graph. Please call the login tool first.",
            }

        if time.time() >= self.token_manager.token_expiry - 60:
            return {
                "status": "expired",
                "message": "Authentication token has expired. Please call the login tool to re-authenticate.",
            }

        expiry_info = self.token_manager.get_token_expiry_info()
        expiry_datetime = datetime.datetime.fromtimestamp(
            self.token_manager.token_expiry
        )
        expiry_str = expiry_datetime.strftime("%Y-%m-%d %H:%M:%S")

        remaining_hours = expiry_info["remaining_hours"]
        remaining_minutes = expiry_info["remaining_minutes"]

        if remaining_hours > 0:
            time_remaining = f"{remaining_hours} hour{'s' if remaining_hours > 1 else ''} and {remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
        else:
            time_remaining = (
                f"{remaining_minutes} minute{'s' if remaining_minutes != 1 else ''}"
            )

        return {
            "status": "authenticated",
            "message": f"Authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
            "token_expiry": self.token_manager.token_expiry,
            "expiry_datetime": expiry_str,
            **expiry_info,
        }

    def clear_device_flow(self) -> None:
        """Clear the device flow state."""
        self.device_flow = None
