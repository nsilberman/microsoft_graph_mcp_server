"""Event browsing cache management module."""

import asyncio
import json
import os
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional


class EventBrowsingCache:
    """Manages event browsing state with disk persistence."""

    CACHE_VERSION = "1.0"
    CACHE_EXPIRY_HOURS = 8
    CACHE_MAX_AGE_HOURS = 24

    def __init__(self):
        self.cache_file = self._get_cache_file_path()
        self.cache = self._load_cache()

    def _get_cache_file_path(self) -> Path:
        """Get the cache file path in user's home directory."""
        home_dir = Path.home()
        return home_dir / ".microsoft_graph_mcp_events.json"

    def _load_cache(self) -> Dict[str, Any]:
        """Load cache from disk or create new cache."""
        if self.cache_file.exists():
            try:
                with open(self.cache_file, "r", encoding="utf-8") as f:
                    cache = json.load(f)

                if self._is_cache_valid(cache):
                    return cache
                else:
                    self._delete_cache_file()
                    return self._create_new_cache()
            except (json.JSONDecodeError, IOError):
                self._delete_cache_file()
                return self._create_new_cache()
        else:
            return self._create_new_cache()

    def _create_new_cache(self) -> Dict[str, Any]:
        """Create a new cache structure with default values."""
        now = datetime.utcnow()
        return {
            "version": self.CACHE_VERSION,
            "last_updated": now.isoformat() + "Z",
            "mode": "browse",
            "browse_state": {
                "start_date": None,
                "end_date": None,
                "top": 20,
                "total_count": 0,
                "metadata": [],
            },
            "search_state": {
                "query": None,
                "start_date": None,
                "end_date": None,
                "top": 20,
                "total_count": 0,
                "metadata": [],
            },
            "metadata": {
                "user_id": None,
                "expires_at": (
                    now + timedelta(hours=self.CACHE_EXPIRY_HOURS)
                ).isoformat()
                + "Z",
            },
        }

    def _is_cache_valid(self, cache: Dict[str, Any]) -> bool:
        """Check if cache is valid and not too old."""
        try:
            last_updated = datetime.fromisoformat(
                cache["last_updated"].replace("Z", "")
            )
            age = datetime.utcnow() - last_updated

            if age > timedelta(hours=self.CACHE_MAX_AGE_HOURS):
                return False

            return True
        except (KeyError, ValueError):
            return False

    def _is_cache_expired(self) -> bool:
        """Check if cache has expired (> 8 hours)."""
        try:
            expires_at = datetime.fromisoformat(
                self.cache["metadata"]["expires_at"].replace("Z", "")
            )
            return datetime.utcnow() > expires_at
        except (KeyError, ValueError):
            return True

    async def _save_cache(self):
        """Save cache to disk asynchronously using threading."""
        self.cache["last_updated"] = datetime.utcnow().isoformat() + "Z"
        try:
            await asyncio.to_thread(self._save_cache_sync)
        except IOError:
            pass

    def _save_cache_sync(self):
        """Save cache to disk synchronously (called from thread)."""
        with open(self.cache_file, "w", encoding="utf-8") as f:
            json.dump(self.cache, f, indent=2, ensure_ascii=False)

    def _delete_cache_file(self):
        """Delete cache file from disk."""
        try:
            if self.cache_file.exists():
                self.cache_file.unlink()
        except IOError:
            pass

    async def invalidate_browse_state(self):
        """Invalidate browse state (e.g., when date range changes)."""
        self.cache["browse_state"]["total_count"] = 0
        self.cache["browse_state"]["metadata"] = []
        await self._save_cache()

    async def invalidate_search_state(self):
        """Invalidate search state (e.g., when query changes)."""
        self.cache["search_state"]["total_count"] = 0
        self.cache["search_state"]["metadata"] = []
        await self._save_cache()

    async def set_mode(self, mode: str):
        """Set browsing mode ('browse' or 'search')."""
        if self.cache["mode"] != mode:
            self.cache["mode"] = mode
            await self._save_cache()

    def get_mode(self) -> str:
        """Get current browsing mode."""
        return self.cache["mode"]

    async def update_browse_state(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: Optional[int] = None,
        total_count: Optional[int] = None,
        metadata: Optional[List[Dict[str, Any]]] = None,
    ):
        """Update browse state parameters."""
        state = self.cache["browse_state"]

        if start_date != state["start_date"]:
            state["start_date"] = start_date
            state["total_count"] = 0
            state["metadata"] = []

        if end_date != state["end_date"]:
            state["end_date"] = end_date
            state["total_count"] = 0
            state["metadata"] = []

        if top is not None and top != state["top"]:
            state["top"] = top

        if total_count is not None:
            state["total_count"] = total_count

        if metadata is not None:
            state["metadata"] = metadata

        await self._save_cache()

    async def update_search_state(
        self,
        query: str,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: Optional[int] = None,
        total_count: Optional[int] = None,
        metadata: Optional[List[Dict[str, Any]]] = None,
    ):
        """Update search state parameters."""
        state = self.cache["search_state"]

        if query != state["query"]:
            state["query"] = query
            state["total_count"] = 0
            state["metadata"] = []

        if start_date != state["start_date"]:
            state["start_date"] = start_date
            state["total_count"] = 0
            state["metadata"] = []

        if end_date != state["end_date"]:
            state["end_date"] = end_date
            state["total_count"] = 0
            state["metadata"] = []

        if top is not None and top != state["top"]:
            state["top"] = top

        if total_count is not None:
            state["total_count"] = total_count

        if metadata is not None:
            state["metadata"] = metadata

        await self._save_cache()

    def get_browse_state(self) -> Dict[str, Any]:
        """Get current browse state."""
        return self.cache["browse_state"].copy()

    def get_search_state(self) -> Dict[str, Any]:
        """Get current search state."""
        return self.cache["search_state"].copy()

    def get_cached_events(self) -> List[Dict[str, Any]]:
        """Get cached events from current mode, sorted by start time (earliest first)."""
        if self.cache["mode"] == "browse":
            metadata = self.cache["browse_state"]["metadata"].copy()
        else:
            metadata = self.cache["search_state"]["metadata"].copy()

        sorted_events = sorted(metadata, key=lambda x: x.get("start_datetime", ""))

        for idx, event in enumerate(sorted_events):
            event["number"] = idx + 1

        return sorted_events

    def should_refresh_total_count(self) -> bool:
        """Check if total_count needs to be refreshed."""
        state = (
            self.cache["browse_state"]
            if self.cache["mode"] == "browse"
            else self.cache["search_state"]
        )
        return state["total_count"] == 0 or self._is_cache_expired()

    def clear_cache(self):
        """Clear cache and delete file."""
        self.cache = self._create_new_cache()
        self._delete_cache_file()

    def get_cache_info(self) -> Dict[str, Any]:
        """Get cache information for debugging."""
        return {
            "cache_file": str(self.cache_file),
            "exists": self.cache_file.exists(),
            "mode": self.cache["mode"],
            "last_updated": self.cache["last_updated"],
            "expires_at": self.cache["metadata"]["expires_at"],
            "is_expired": self._is_cache_expired(),
            "browse_state": self.cache["browse_state"],
            "search_state": self.cache["search_state"],
        }


# Global cache instance
event_cache = EventBrowsingCache()
