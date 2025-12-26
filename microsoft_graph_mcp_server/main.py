"""Main entry point for Microsoft Graph MCP Server."""

import asyncio
import sys
from typing import Optional

import click

from .server import MicrosoftGraphMCPServer
from .config import settings


@click.command()
@click.option(
    "--client-id",
    help="Microsoft Graph client ID (default: Microsoft's public client ID)",
    default=lambda: settings.client_id or ""
)
@click.option(
    "--tenant-id",
    help="Microsoft Graph tenant ID (default: organizations)",
    default=lambda: settings.tenant_id or ""
)
def main(
    client_id: str,
    tenant_id: str
):
    """Microsoft Graph MCP Server - Provides access to Microsoft 365 ecosystem."""
    
    # Update settings with provided values
    if client_id:
        settings.client_id = client_id
    if tenant_id:
        settings.tenant_id = tenant_id
    
    click.echo("Starting Microsoft Graph MCP Server...")
    click.echo(f"Server Name: {settings.server_name}")
    click.echo(f"Server Version: {settings.server_version}")
    click.echo(f"Graph API Base URL: {settings.graph_api_base_url}")
    click.echo(f"Authentication: Device Code Flow (Interactive)")
    click.echo("-" * 50)
    
    try:
        server = MicrosoftGraphMCPServer()
        asyncio.run(server.run())
    except KeyboardInterrupt:
        click.echo("\nServer stopped by user.")
    except Exception as e:
        click.echo(f"Error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()