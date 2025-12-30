"""Base handler class for MCP tool handlers."""

import json
from typing import Any, Dict, Optional
import mcp.types as types


class BaseHandler:
    """Base class for all MCP tool handlers."""
    
    def __init__(self):
        pass
    
    def _format_response(self, data: Any) -> list[types.TextContent]:
        """Format response data as MCP text content.
        
        Args:
            data: Data to format
            
        Returns:
            List of TextContent objects
        """
        return [types.TextContent(
            type="text",
            text=json.dumps(data, indent=2, ensure_ascii=False)
        )]
    
    def _format_error(self, error_message: str) -> list[types.TextContent]:
        """Format error message as MCP text content.
        
        Args:
            error_message: Error message to format
            
        Returns:
            List of TextContent objects
        """
        return [types.TextContent(
            type="text",
            text=error_message
        )]
    
    def _format_success(self, message: str, **kwargs) -> list[types.TextContent]:
        """Format success message with optional additional data.
        
        Args:
            message: Success message
            **kwargs: Additional data to include
            
        Returns:
            List of TextContent objects
        """
        response = {"message": message, "status": "success"}
        response.update(kwargs)
        return self._format_response(response)
