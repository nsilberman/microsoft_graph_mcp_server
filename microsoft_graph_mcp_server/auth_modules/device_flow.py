"""Device flow authentication for Microsoft Graph API."""

import asyncio
import datetime
import logging
import time
from typing import Any, Dict, Optional

from .token_manager import TokenManager

logger = logging.getLogger(__name__)

# Timeout for waiting for authentication completion (in seconds)
# This gives users enough time to complete authentication in the browser
AUTH_VERIFICATION_TIMEOUT = 60.0

# Timeout for waiting for Microsoft to issue the access token after user completes authentication (in seconds)
# This accounts for Microsoft's processing time, network latency, and system performance variations
TOKEN_ACQUISITION_TIMEOUT = 15.0


class DeviceFlowManager:
    """Manages device code flow authentication for Microsoft Graph API."""

    def __init__(self, client_app, token_manager: TokenManager):
        self.client_app = client_app
        self.token_manager = token_manager
        self.device_flow: Optional[Dict[str, Any]] = None

    async def initiate_device_code(self) -> Dict[str, Any]:
        """Initiate device code flow and return authentication info."""
        max_retries = 3
        retry_delay = 2

        for attempt in range(max_retries):
            try:
                # Clear any existing device flow before initiating a new one
                self.device_flow = None
                logger.info(
                    f"Initiating device code flow (attempt {attempt + 1}/{max_retries})"
                )

                loop = asyncio.get_event_loop()
                flow = await asyncio.wait_for(
                    loop.run_in_executor(
                        None,
                        self.client_app.initiate_device_flow,
                        ["https://graph.microsoft.com/.default"],
                    ),
                    timeout=30.0,
                )

                if "error" in flow:
                    raise Exception(f"Failed to initiate device flow: {flow['error']}")

                self.device_flow = flow
                device_code = flow.get("device_code", "")
                user_code = flow.get("user_code", "")
                logger.info(
                    f"Device code flow initiated successfully. user_code: {user_code}, device_code: {device_code[:50]}"
                )

                # Add expires_at timestamp to the flow
                expires_in = flow.get("expires_in", 900)
                flow["expires_at"] = time.time() + expires_in

                # Save device flow to disk for later retrieval using device_code
                if device_code:
                    self.token_manager.save_device_flow(device_code, flow)
                    self.token_manager.save_latest_device_code(device_code)

                print("\n" + "=" * 70)
                print("MICROSOFT GRAPH AUTHENTICATION")
                print("=" * 70)
                print(f"\nTo sign in, use a web browser to open the page:")
                print(f"\n{flow['verification_uri']}")
                print(f"\nAnd enter the code:")
                print(f"\n{flow['user_code']}")
                print("\n" + "=" * 70)
                print("NOTE: Previous tokens have been cleared.")
                print(
                    "IMPORTANT: After completing authentication, you MUST call complete_login"
                )
                print("to complete the authentication process.")
                print("=" * 70 + "\n")

                return {
                    "status": "pending",
                    "message": "Please complete authentication using the link and code below. NOTE: Previous tokens have been cleared. IMPORTANT: After completing authentication, you MUST call complete_login to complete the authentication process.",
                    "verification_uri": flow.get("verification_uri", ""),
                    "user_code": user_code,
                    "expires_in": flow.get("expires_in", 900),
                    "interval": flow.get("interval", 5),
                }
            except asyncio.TimeoutError:
                if attempt < max_retries - 1:
                    print(
                        f"Connection timeout. Retrying... (Attempt {attempt + 1}/{max_retries})"
                    )
                    await asyncio.sleep(retry_delay)
                else:
                    raise Exception(
                        "Authentication failed: Connection timeout. Please check your internet connection and try again."
                    )
            except Exception as e:
                error_msg = str(e)
                error_lower = error_msg.lower()

                network_errors = [
                    "connection aborted",
                    "remote disconnected",
                    "connection refused",
                    "connection reset",
                    "network unreachable",
                    "timed out",
                    "timeout",
                    "no route to host",
                    "name or service not known",
                    "temporary failure",
                ]

                is_network_error = any(err in error_lower for err in network_errors)

                if is_network_error and attempt < max_retries - 1:
                    print(f"Network error: {error_msg}")
                    print(f"Retrying... (Attempt {attempt + 1}/{max_retries})")
                    await asyncio.sleep(retry_delay)
                elif is_network_error:
                    return {
                        "status": "failed",
                        "message": f"Authentication failed: Network connection error. Please check your internet connection and try again. Error: {error_msg}",
                    }
                else:
                    return {
                        "status": "failed",
                        "message": f"Authentication failed: {error_msg}",
                    }

        # This should never be reached, but added to satisfy type checker
        raise Exception("Unexpected: No return in initiate_device_code")

    async def initiate_device_flow_only(self) -> Dict[str, Any]:
        """Initiate device code flow without waiting for completion.

        Returns:
            Dict with authentication status and verification details
        """
        return await self.initiate_device_code()

    async def initiate_and_wait_for_completion(
        self, max_wait_time: int = 5, poll_interval: int = 1, progress_interval: int = 5
    ) -> Dict[str, Any]:
        """Initiate device code flow and automatically wait for authentication completion with progress updates.

        Args:
            max_wait_time: Maximum time to wait for authentication in seconds (default: 5)
            poll_interval: How often to check for authentication completion in seconds (default: 1)
            progress_interval: How often to show progress updates in seconds (default: 5)

        Returns:
            Dict with authentication status and details
        """
        max_retries = 3
        retry_delay = 2

        for attempt in range(max_retries):
            try:
                # Clear any existing device flow before initiating a new one
                self.device_flow = None

                loop = asyncio.get_event_loop()
                flow = await asyncio.wait_for(
                    loop.run_in_executor(
                        None,
                        self.client_app.initiate_device_flow,
                        ["https://graph.microsoft.com/.default"],
                    ),
                    timeout=30.0,
                )

                if "error" in flow:
                    raise Exception(f"Failed to initiate device flow: {flow['error']}")

                self.device_flow = flow

                print("\n" + "=" * 70)
                print("MICROSOFT GRAPH AUTHENTICATION")
                print("=" * 70)
                print(f"\nTo sign in, use a web browser to open the page:")
                print(f"\n{flow['verification_uri']}")
                print(f"\nAnd enter the code:")
                print(f"\n{flow['user_code']}")
                print("\n" + "=" * 70)
                print(
                    f"Waiting for authentication (timeout: {max_wait_time} seconds)..."
                )
                print("=" * 70 + "\n")

                elapsed = 0
                last_progress = 0

                while elapsed < max_wait_time:
                    try:
                        result = await asyncio.wait_for(
                            loop.run_in_executor(
                                None,
                                self.client_app.acquire_token_by_device_flow,
                                flow,
                            ),
                            timeout=poll_interval,
                        )

                        if "access_token" in result:
                            self.token_manager.update_token(
                                access_token=result["access_token"],
                                expires_in=int(result.get("expires_in", 3600)),
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

                            print("\n✓ Authentication successful!\n")

                            return {
                                "status": "success",
                                "message": f"Successfully authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                                "token_expiry": self.token_manager.token_expiry,
                                "expiry_datetime": expiry_str,
                                **expiry_info,
                            }
                        else:
                            error = result.get("error", "")
                            error_description = result.get(
                                "error_description", ""
                            ).lower()
                            if error not in ["authorization_pending", "timeout"]:
                                self.device_flow = None
                                if "already redeemed" in error_description:
                                    self.token_manager.load_tokens_from_disk()
                                    if (
                                        self.token_manager.authenticated
                                        and self.token_manager.access_token
                                    ):
                                        expiry_info = (
                                            self.token_manager.get_token_expiry_info()
                                        )
                                        expiry_datetime = (
                                            datetime.datetime.fromtimestamp(
                                                self.token_manager.token_expiry
                                            )
                                        )
                                        expiry_str = expiry_datetime.strftime(
                                            "%Y-%m-%d %H:%M:%S"
                                        )

                                        remaining_hours = expiry_info["remaining_hours"]
                                        remaining_minutes = expiry_info[
                                            "remaining_minutes"
                                        ]

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
                                        return {
                                            "status": "failed",
                                            "message": "Authentication code was already used but no valid token found. Please call login again to start a new authentication.",
                                        }
                                return {
                                    "status": "failed",
                                    "message": f"Authentication failed: {result.get('error_description', 'Unknown error')}",
                                }
                    except asyncio.TimeoutError:
                        pass

                    elapsed += poll_interval

                    if elapsed - last_progress >= progress_interval:
                        remaining = max_wait_time - elapsed
                        print(
                            f"Still waiting for authentication... {remaining} seconds remaining"
                        )
                        last_progress = elapsed

                remaining = max_wait_time - elapsed
                print(f"\nAuthentication timed out after {max_wait_time} seconds.")

                return {
                    "status": "pending",
                    "message": f"Authentication timed out. Please complete authentication in the browser using the verification link and code, then call login again to verify.",
                    "verification_uri": flow.get("verification_uri", ""),
                    "user_code": flow.get("user_code", ""),
                    "expires_in": flow.get("expires_in", 900),
                }
            except asyncio.TimeoutError:
                if attempt < max_retries - 1:
                    print(
                        f"Connection timeout. Retrying... (Attempt {attempt + 1}/{max_retries})"
                    )
                    await asyncio.sleep(retry_delay)
                else:
                    raise Exception(
                        "Authentication failed: Connection timeout. Please check your internet connection and try again."
                    )
            except Exception as e:
                error_msg = str(e)
                error_lower = error_msg.lower()

                network_errors = [
                    "connection aborted",
                    "remote disconnected",
                    "connection refused",
                    "connection reset",
                    "network unreachable",
                    "timed out",
                    "timeout",
                    "no route to host",
                    "name or service not known",
                    "temporary failure",
                ]

                is_network_error = any(err in error_lower for err in network_errors)

                if is_network_error and attempt < max_retries - 1:
                    print(f"Network error: {error_msg}")
                    print(f"Retrying... (Attempt {attempt + 1}/{max_retries})")
                    await asyncio.sleep(retry_delay)
                elif is_network_error:
                    raise Exception(
                        f"Authentication failed: Network connection error. Please check your internet connection and try again. Error: {error_msg}"
                    )
                else:
                    self.device_flow = None
                    if "already redeemed" in error_lower:
                        self.token_manager.load_tokens_from_disk()
                        if (
                            self.token_manager.authenticated
                            and self.token_manager.access_token
                        ):
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
                            return {
                                "status": "failed",
                                "message": "Authentication code was already used but no valid token found. Please call login again to start a new authentication.",
                            }
                    else:
                        return {
                            "status": "failed",
                            "message": f"Authentication failed: {error_msg}",
                        }

        # This should never be reached, but added to satisfy type checker
        raise Exception("Unexpected: No return in initiate_and_wait_for_completion")

    async def check_authentication_status(self) -> Dict[str, Any]:
        """Check if authentication is complete by attempting to acquire token with short timeout."""
        if self.device_flow is None:
            logger.warning("check_authentication_status called but device_flow is None")
            return {
                "status": "error",
                "message": "No pending device code flow. Please call login again.",
            }

        try:
            loop = asyncio.get_event_loop()
            device_code_from_flow = self.device_flow.get("device_code", "")
            logger.info(
                f"check_authentication_status called with device_code: {device_code_from_flow[:50]}"
            )

            async def acquire_with_timeout():
                try:
                    result = await asyncio.wait_for(
                        loop.run_in_executor(
                            None,
                            self.client_app.acquire_token_by_device_flow,
                            self.device_flow,
                        ),
                        timeout=TOKEN_ACQUISITION_TIMEOUT,
                    )
                    return result
                except asyncio.TimeoutError:
                    return {
                        "error": "timeout",
                        "error_description": "Authentication still pending",
                    }

            result = await acquire_with_timeout()
            logger.info(
                f"acquire_token_by_device_flow result error: {result.get('error', 'none')}"
            )

            if "access_token" in result:
                self.token_manager.update_token(
                    access_token=result["access_token"],
                    expires_in=int(result.get("expires_in", 3600)),
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
                elif "already redeemed" in result.get("error_description", "").lower():
                    self.device_flow = None
                    self.token_manager.load_tokens_from_disk()
                    if (
                        self.token_manager.authenticated
                        and self.token_manager.access_token
                    ):
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
                        return {
                            "status": "failed",
                            "message": "Authentication code was already used but no valid token found. Please call login again to start a new authentication.",
                        }
                else:
                    self.device_flow = None
                    return {
                        "status": "failed",
                        "message": f"Authentication failed: {result.get('error_description', 'Unknown error')}",
                    }
        except Exception as e:
            error_msg = str(e)
            if "already redeemed" in error_msg.lower():
                # Device code was already used, try to load tokens from disk
                self.token_manager.load_tokens_from_disk()
                if self.token_manager.authenticated and self.token_manager.access_token:
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
                    return {
                        "status": "failed",
                        "message": "Authentication code was already used but no valid token found. Please call login again to start a new authentication.",
                    }
            else:
                # For other errors, clear the device flow
                self.device_flow = None
                return {
                    "status": "failed",
                    "message": f"Authentication failed: {error_msg}",
                }

    async def check_login_status(
        self, device_code: Optional[str] = None
    ) -> Dict[str, Any]:
        """Check the current login status.

        Args:
            device_code: Optional device_code to load device flow from disk and check online status
                         If not provided, will automatically load the latest device_code from disk
        """
        logger.info("=" * 70)
        logger.info("check_login_status called!")
        logger.info(f"  device_code: {device_code[:50] if device_code else 'None'}")

        # If device_code is not provided, try to load the latest one from disk
        if not device_code:
            device_code = self.token_manager.get_latest_device_code()
            logger.info(
                f"  Auto-loaded device_code from disk: {device_code[:50] if device_code else 'None'}"
            )

        logger.info("=" * 70)

        # Clean up expired device flows first
        self.token_manager.cleanup_expired_device_flows()

        # If device_code is provided, first check if we're already authenticated
        # This prevents "already redeemed" errors when check_status is called multiple times
        if device_code:
            self.token_manager.load_tokens_from_disk()
            logger.info(
                f"Token loaded from disk - authenticated: {self.token_manager.authenticated}, has_token: {bool(self.token_manager.access_token)}"
            )

            if self.token_manager.authenticated and self.token_manager.access_token:
                # Already have a valid token, return authentication status
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
                    "message": f"Authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                    "token_expiry": self.token_manager.token_expiry,
                    "expiry_datetime": expiry_str,
                    **expiry_info,
                }

            # Not authenticated, proceed to load device flow from disk
            # Always load the device flow from disk when device_code is provided
            # This ensures we're using the saved flow state, not any in-memory state
            loaded_flow = self.token_manager.load_device_flow(device_code)
            if loaded_flow:
                # Check if the loaded flow is expired
                expires_at = loaded_flow.get("expires_at")
                if expires_at and time.time() >= expires_at:
                    # Flow is expired, delete it and return error
                    self.token_manager.delete_device_flow(device_code)
                    self.device_flow = None
                    return {
                        "status": "failed",
                        "message": "The device code has expired. Please call login again to start a new authentication.",
                    }
                self.device_flow = loaded_flow
            else:
                # No flow found in disk, it may have been deleted
                # Try loading tokens to see if authentication was already completed
                self.token_manager.load_tokens_from_disk()
                if self.token_manager.authenticated and self.token_manager.access_token:
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
                        "message": f"Authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                        "token_expiry": self.token_manager.token_expiry,
                        "expiry_datetime": expiry_str,
                        **expiry_info,
                    }
                else:
                    self.device_flow = None
                    return {
                        "status": "failed",
                        "message": "Device flow not found. Please call login again to start a new authentication.",
                    }

            # Check online status if there's a pending device flow
            if self.device_flow is not None:
                # Save the device_code before calling check_authentication_status
                # because it will set self.device_flow = None on success
                flow_device_code = self.device_flow.get("device_code", "")
                logger.info(
                    f"Processing device flow with device_code: {flow_device_code[:50]}"
                )

                # CRITICAL FIX: Check if flow still exists on disk before calling MSAL
                # If it doesn't exist, it means it was already used and deleted
                flow_still_on_disk = (
                    self.token_manager.load_device_flow(flow_device_code) is not None
                )
                logger.info(f"Flow still on disk: {flow_still_on_disk}")

                if not flow_still_on_disk:
                    # Flow was already used, try to load tokens
                    logger.info("Flow not on disk, attempting to load tokens directly")
                    self.token_manager.load_tokens_from_disk()
                    if (
                        self.token_manager.authenticated
                        and self.token_manager.access_token
                    ):
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
                            "message": f"Authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                            "token_expiry": self.token_manager.token_expiry,
                            "expiry_datetime": expiry_str,
                            **expiry_info,
                        }
                    else:
                        self.device_flow = None
                        return {
                            "status": "failed",
                            "message": "Device flow was already used but no valid token found. Please call login again to start a new authentication.",
                        }

                try:
                    result = await self.check_authentication_status()

                    if result.get("status") == "success":
                        # Tokens already saved to disk by check_authentication_status
                        # Delete the device flow from disk since authentication is complete
                        if flow_device_code:
                            self.token_manager.delete_device_flow(flow_device_code)
                        return result
                    elif result.get("status") == "pending":
                        return result
                    else:
                        # Authentication failed
                        if flow_device_code:
                            self.token_manager.delete_device_flow(flow_device_code)
                        # Don't clear device_flow here, let check_authentication_status handle it
                        return result
                except Exception as e:
                    # If check_authentication_status raised an exception, handle it here
                    error_msg = str(e)
                    logger.error(
                        f"check_authentication_status exception: {error_msg[:200]}"
                    )
                    if "already redeemed" in error_msg.lower():
                        # Device code was already used, try to load tokens from disk
                        self.token_manager.load_tokens_from_disk()
                        if (
                            self.token_manager.authenticated
                            and self.token_manager.access_token
                        ):
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

                            # Delete the device flow from disk since authentication is complete
                            if flow_device_code:
                                self.token_manager.delete_device_flow(flow_device_code)

                            return {
                                "status": "success",
                                "message": f"Successfully authenticated with Microsoft Graph. Token expires in {time_remaining} at {expiry_str}",
                                "token_expiry": self.token_manager.token_expiry,
                                "expiry_datetime": expiry_str,
                                **expiry_info,
                            }
                        else:
                            return {
                                "status": "failed",
                                "message": "Authentication code was already used but no valid token found. Please call login again to start a new authentication.",
                            }
                    else:
                        # For other errors, return the error message
                        return {
                            "status": "failed",
                            "message": f"Authentication failed: {error_msg}",
                        }
        else:
            # No device_code provided, clear any in-memory device flow to prevent using stale flows
            if self.device_flow is not None:
                self.clear_device_flow()

        # No device_code provided or flow failed, check disk for existing tokens
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
