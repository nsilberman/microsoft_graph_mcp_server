"""Main entry point for Microsoft Graph MCP Server."""

# IMMEDIATE STARTUP LOGGING - Write BEFORE any imports
import sys
import os

try:
    from pathlib import Path

    # Place log file in project root (parent of the microsoft_graph_mcp_server directory)
    log_file = Path(__file__).resolve().parent.parent / "mcp_server.log"
except:
    log_file = Path.cwd() / "mcp_server.log"

# Ensure log directory exists
log_file.parent.mkdir(parents=True, exist_ok=True)

# Write startup message to log file IMMEDIATELY
with open(log_file, "a") as f:
    f.write(f"\n{'='*70}\n")
    f.write(f"MCP SERVER STARTED\n")
    f.write(f"Timestamp: {os.path.getmtime(__file__)}\n")
    f.write(f"Log file: {log_file}\n")
    f.write(f"{'='*70}\n\n")
    f.flush()

# CRITICAL: Configure logging BEFORE importing any other modules!
import logging

# Create file handler ONLY - no console handler for stdio MCP server
file_handler = logging.FileHandler(log_file, mode="a")

# Set format
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
file_handler.setFormatter(formatter)

# Configure root logger FIRST - only file handler
root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)
root_logger.addHandler(file_handler)

# NOW we can import other modules
import asyncio
from typing import Optional

import click

from .server import MicrosoftGraphMCPServer
from .config import settings

# Configure specific loggers AFTER they are imported
device_flow_logger = logging.getLogger(
    "microsoft_graph_mcp_server.auth_modules.device_flow"
)
device_flow_logger.setLevel(logging.DEBUG)
device_flow_logger.addHandler(file_handler)
device_flow_logger.propagate = False  # Don't duplicate logs

auth_manager_logger = logging.getLogger(
    "microsoft_graph_mcp_server.auth_modules.auth_manager"
)
auth_manager_logger.setLevel(logging.DEBUG)
auth_manager_logger.addHandler(file_handler)
auth_manager_logger.propagate = False

# Log that logging is configured
logging.info("=" * 70)
logging.info("Logging initialized successfully")
logging.info(f"Log file: {log_file}")
logging.info(f"Server: {settings.server_name} v{settings.server_version}")
logging.info("=" * 70)


@click.command()
@click.option(
    "--client-id",
    help="Microsoft Graph client ID (default: Microsoft's public client ID)",
    default=lambda: settings.client_id or "",
)
@click.option(
    "--tenant-id",
    help="Microsoft Graph tenant ID (default: organizations)",
    default=lambda: settings.tenant_id or "",
)
def main(client_id: str, tenant_id: str):
    """Microsoft Graph MCP Server - Provides access to Microsoft 365 ecosystem."""

    # Log that main() is called
    logging.info("main() function called")
    logging.info(f"Client ID: {client_id[:20]}... if provided")
    logging.info(f"Tenant ID: {tenant_id}")

    # Update settings with provided values
    if client_id:
        settings.client_id = client_id
        logging.info(f"Client ID updated: {client_id[:20]}...")
    if tenant_id:
        settings.tenant_id = tenant_id
        logging.info(f"Tenant ID updated: {tenant_id}")

    logging.info("Initializing MicrosoftGraphMCPServer...")

    try:
        server = MicrosoftGraphMCPServer()
        logging.info("Starting server.run()...")
        asyncio.run(server.run())
    except KeyboardInterrupt:
        logging.info("Server stopped by user (KeyboardInterrupt)")
    except Exception as e:
        logging.error(f"Server error: {str(e)}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
