"""Entry point when package is run as module (e.g., python -m microsoft_graph_mcp_server)."""

# IMMEDIATE STARTUP LOGGING - This runs BEFORE any imports
import sys
import os

# Print startup message immediately to ensure visibility
print("=" * 70, file=sys.stdout, flush=True)
print("[MCP] Microsoft Graph MCP Server starting...", file=sys.stdout, flush=True)
print("=" * 70, file=sys.stdout, flush=True)

# Now import and run the main function
from .main import main

if __name__ == "__main__":
    main()
