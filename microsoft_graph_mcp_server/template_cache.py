"""Template management cache module."""

import asyncio
import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional


class TemplateCache:
    """Manages template browsing state with disk persistence."""

    CACHE_VERSION = "1.0"
    CACHE_EXPIRY_HOURS = 1
    CACHE_MAX_AGE_HOURS = 24

    def __init__(self):
        self.cache_file = self._get_cache_file_path()
        self.cache = self._load_cache()

    def _get_cache_file_path(self) -> Path:
        """Get the cache file path in user's home directory."""
        home_dir = Path.home()
        return home_dir / ".microsoft_graph_mcp_templates.json"

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
            "templates": [],
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
        """Check if cache has expired (> 1 hour)."""
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

    async def add_template(self, template: Dict[str, Any]):
        """Add a template to the cache.

        Args:
            template: Template dictionary with id, subject, to, cc, etc.
        """
        self.cache["templates"].append(template)
        await self._save_cache()

    async def remove_template(self, template_id: str):
        """Remove a template from the cache by its ID.

        Args:
            template_id: ID of the template to remove from cache
        """
        self.cache["templates"] = [
            template
            for template in self.cache["templates"]
            if template.get("id") != template_id
        ]
        await self._save_cache()

    async def update_template(self, template_id: str, updates: Dict[str, Any]):
        """Update a template in the cache.

        Args:
            template_id: ID of the template to update
            updates: Dictionary of fields to update
        """
        for template in self.cache["templates"]:
            if template.get("id") == template_id:
                template.update(updates)
                break
        await self._save_cache()

    async def set_templates(self, templates: List[Dict[str, Any]]):
        """Set all templates in the cache.

        Args:
            templates: List of template dictionaries
        """
        self.cache["templates"] = templates
        await self._save_cache()

    def get_cached_templates(self) -> List[Dict[str, Any]]:
        """Get all cached templates."""
        return self.cache["templates"].copy()

    def get_template_by_id(self, template_id: str) -> Optional[Dict[str, Any]]:
        """Get a template by its ID.

        Args:
            template_id: ID of the template to retrieve

        Returns:
            Template dictionary or None if not found
        """
        for template in self.cache["templates"]:
            if template.get("id") == template_id:
                return template.copy()
        return None

    def get_template_by_number(self, number: int) -> Optional[Dict[str, Any]]:
        """Get a template by its cache number.

        Args:
            number: Template cache number (1-indexed)

        Returns:
            Template dictionary or None if not found
        """
        templates = self.cache["templates"]
        if 1 <= number <= len(templates):
            return templates[number - 1].copy()
        return None

    def should_refresh_cache(self) -> bool:
        """Check if cache needs to be refreshed."""
        return len(self.cache["templates"]) == 0 or self._is_cache_expired()

    async def clear_cache(self):
        """Clear cache and delete file."""
        self.cache = self._create_new_cache()
        await asyncio.to_thread(self._delete_cache_file)

    def get_cache_info(self) -> Dict[str, Any]:
        """Get cache information for debugging."""
        return {
            "cache_file": str(self.cache_file),
            "exists": self.cache_file.exists(),
            "last_updated": self.cache["last_updated"],
            "expires_at": self.cache["metadata"]["expires_at"],
            "is_expired": self._is_cache_expired(),
            "template_count": len(self.cache["templates"]),
        }


# Global cache instance
template_cache = TemplateCache()
