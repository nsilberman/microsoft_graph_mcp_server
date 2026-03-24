#!/bin/bash
echo "========================================"
echo "  Microsoft Graph MCP Server - Reinstall"
echo "========================================"
echo

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "[1/2] Uninstalling old version..."
uv tool uninstall microsoft-graph-mcp-server 2>/dev/null

echo "[2/2] Installing new version..."
uv tool install --force "$SCRIPT_DIR"

echo
echo "========================================"
echo "  Done! MCP Server has been reinstalled."
echo "========================================"
echo
echo "Please restart Claude Desktop to apply changes."
