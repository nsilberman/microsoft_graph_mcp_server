"""Entry point when package is run as module (e.g., python -m microsoft_graph_mcp_server)."""

# CRITICAL: Disable PYTHONSTARTUP to prevent it from corrupting stdio MCP protocol
# This must be done BEFORE any other imports
import os
import sys
if 'PYTHONSTARTUP' in os.environ:
    del os.environ['PYTHONSTARTUP']

# Import and run the main function
from .main import main

if __name__ == "__main__":
    main()