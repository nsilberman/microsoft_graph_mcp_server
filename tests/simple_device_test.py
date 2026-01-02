"""Simple test - don't clear log file."""

import asyncio
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

print("[Test] Importing auth_manager...")
from microsoft_graph_mcp_server.auth_modules.auth_manager import GraphAuthManager

print("[Test] Creating auth_manager...")
auth_manager = GraphAuthManager()
print("   ✓ Created")

print("[Test] Calling check_login_status...")
result = asyncio.run(auth_manager.check_login_status())
print(f"   Status: {result['status']}")

print("\n[Test] Done!")
print("Now check the log file for device_flow logs.")
