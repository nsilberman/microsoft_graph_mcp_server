"""Token management for Microsoft Graph API authentication."""

import json
import os
import time
from pathlib import Path
from typing import Optional

from ..config import settings


TOKEN_FILE = Path.home() / ".microsoft_graph_mcp_tokens.json"


class TokenManager:
    """Manages authentication tokens for Microsoft Graph API."""

    def __init__(self):
        self.access_token: Optional[str] = None
        self.token_expiry: float = 0
        self.refresh_token: Optional[str] = None
        self.authenticated: bool = False

        self.load_tokens_from_disk()

    def save_tokens_to_disk(self) -> None:
        """Save authentication tokens to disk."""
        try:
            token_data = {
                "access_token": self.access_token,
                "refresh_token": self.refresh_token,
                "token_expiry": self.token_expiry,
                "authenticated": self.authenticated,
            }
            with open(TOKEN_FILE, "w") as f:
                json.dump(token_data, f, indent=2)
        except Exception as e:
            print(f"Warning: Failed to save tokens to disk: {e}")

    def load_tokens_from_disk(self) -> None:
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
                    self.delete_tokens_from_disk()
        except Exception as e:
            print(f"Warning: Failed to load tokens from disk: {e}")
            self.delete_tokens_from_disk()

    def delete_tokens_from_disk(self) -> None:
        """Delete authentication tokens from disk."""
        try:
            if TOKEN_FILE.exists():
                os.remove(TOKEN_FILE)
        except Exception as e:
            print(f"Warning: Failed to delete tokens from disk: {e}")

    def update_token(
        self,
        access_token: str,
        expires_in: int = 3600,
        refresh_token: Optional[str] = None,
    ) -> None:
        """Update the access token and related information."""
        self.access_token = access_token
        self.token_expiry = time.time() + expires_in
        self.refresh_token = refresh_token or self.refresh_token
        self.authenticated = True
        self.save_tokens_to_disk()

    def clear_tokens(self) -> None:
        """Clear all authentication tokens."""
        self.access_token = None
        self.token_expiry = 0
        self.refresh_token = None
        self.authenticated = False
        self.delete_tokens_from_disk()

    def is_token_valid(self) -> bool:
        """Check if the current token is valid and not expired."""
        if not self.authenticated or not self.access_token:
            return False
        return time.time() < self.token_expiry - 60

    def get_token_expiry_info(self) -> dict:
        """Get token expiry information."""
        remaining_seconds = int(self.token_expiry - time.time())
        remaining_minutes = remaining_seconds // 60
        remaining_hours = remaining_minutes // 60
        remaining_minutes = remaining_minutes % 60

        return {
            "remaining_seconds": remaining_seconds,
            "remaining_minutes": remaining_minutes,
            "remaining_hours": remaining_hours,
        }
