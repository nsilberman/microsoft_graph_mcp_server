"""Authentication module for Microsoft Graph API.

This module provides backward compatibility by re-exporting the auth_manager
from the new modular auth_modules package.
"""

from .auth_modules import (
    auth_manager,
    GraphAuthManager,
    TokenManager,
    DeviceFlowManager,
)

__all__ = ["auth_manager", "GraphAuthManager", "TokenManager", "DeviceFlowManager"]
