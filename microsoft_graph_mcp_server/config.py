"""Configuration module for Microsoft Graph MCP Server."""

import os
from typing import Optional
try:
    from pydantic_settings import BaseSettings
except ImportError:
    from pydantic import BaseSettings


class Settings(BaseSettings):
    """Application settings."""
    
    # Microsoft Graph API configuration
    client_id: str = os.getenv("CLIENT_ID", "d3590ed6-52b3-4102-aeff-aad2292ab01c")
    client_secret: str = os.getenv("CLIENT_SECRET", "")
    tenant_id: str = os.getenv("TENANT_ID", "organizations")
    
    # User timezone configuration (fallback if API fails)
    user_timezone: str = os.getenv("USER_TIMEZONE", "UTC")
    
    # MCP Server configuration
    server_name: str = "microsoft-graph-mcp-server"
    server_version: str = "0.1.0"
    
    # API endpoints
    graph_api_base_url: str = "https://graph.microsoft.com/v1.0"
    auth_url: str = "https://login.microsoftonline.com"
    
    # Cache settings
    cache_ttl: int = 300  # 5 minutes
    
    # Search settings
    default_search_days: int = int(os.getenv("DEFAULT_SEARCH_DAYS", "90"))  # Default search range in days
    
    # Pagination settings
    page_size: int = int(os.getenv("PAGE_SIZE", "5"))  # Number of items per page for browsing
    
    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"


# Global settings instance
settings = Settings()