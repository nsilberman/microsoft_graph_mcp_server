"""Minimal test to check if logging works."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

print("[Test] Clearing log...")
log_file = Path(__file__).parent / 'microsoft_graph_mcp_server' / 'mcp_server_auth.log'
with open(log_file, 'w') as f:
    f.write("")

print("\n[Test 1] Importing main (should trigger logging)...")
import microsoft_graph_mcp_server.main

print("\n[Test 2] Importing device_flow...")
from microsoft_graph_mcp_server.auth_modules.device_flow import DeviceFlowManager

print("\n[Test 3] Importing auth_manager...")
from microsoft_graph_mcp_server.auth_modules.auth_manager import GraphAuthManager

print("\n[Test 4] Creating auth_manager...")
auth_manager = GraphAuthManager()
print("   ✓ Auth manager created")

print("\n[Test 5] Logging test messages...")
import logging
root_logger = logging.getLogger()
root_logger.info("This is a test log message from root logger")
print("   ✓ Logged to root logger")

print("\n[Test 6] Checking log file...")
if log_file.exists():
    with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        print(f"   ✓ Log file has {len(content)} characters")
        print(f"\n[Complete] Log file content:")
        print("=" * 70)
        print(content)
        print("=" * 70)
else:
    print("   ✗ Log file not found")
