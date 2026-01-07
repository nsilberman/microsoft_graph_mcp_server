"""Test login status to see if logging works."""

import asyncio
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

# Clear log
log_file = Path(__file__).parent / 'microsoft_graph_mcp_server' / 'mcp_server_auth.log'
print(f"Clearing log: {log_file}")
with open(log_file, 'w') as f:
    f.write("")

print("\n[1] Importing auth_manager...")
from microsoft_graph_mcp_server.auth_modules.auth_manager import GraphAuthManager

print("[2] Creating auth_manager instance...")
auth_manager = GraphAuthManager()

print("[3] Calling check_login_status (no device_code)...")
result1 = asyncio.run(auth_manager.check_login_status())
print(f"   Status: {result1['status']}")

print("\n[4] Calling check_login_status (with device_code)...")
print("   Note: This will initiate login first...")
result2 = asyncio.run(auth_manager.check_login_status("test_device_code"))
print(f"   Status: {result2['status']}")

print("\n[5] Checking log file...")
if log_file.exists():
    with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        lines = content.count('\n')
        print(f"   Log file has {lines} lines")
        if lines > 20:
            print(f"   Last 20 lines:")
            print("=" * 70)
            all_lines = content.split('\n')
            for line in all_lines[-20:]:
                print(line)
        else:
            print(f"   All lines:")
            print("=" * 70)
            print(content)
else:
    print("   (Log file not found)")

print("\n✓ Test complete!")
