"""Quick test to check device_flow logging."""

import asyncio
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

# Clear log
log_file = Path(__file__).parent / 'microsoft_graph_mcp_server' / 'mcp_server_auth.log'
print(f"Clearing log: {log_file}")
with open(log_file, 'w') as f:
    f.write("")

print("\n[Test] Creating auth_manager...")
from microsoft_graph_mcp_server.auth_modules.auth_manager import GraphAuthManager

auth_manager = GraphAuthManager()
print("   ✓ Auth manager created")

print("\n[Test] Calling check_login_status (no device_code)...")
result = asyncio.run(auth_manager.check_login_status())
print(f"   Status: {result['status']}")

print("\n[Test] Checking log file...")
if log_file.exists():
    with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        lines = content.split('\n')
        print(f"   Log file has {len(lines)} lines")
        if len(lines) > 30:
            print(f"   Last 30 lines:")
            print("=" * 70)
            for line in lines[-30:]:
                print(line)
        else:
            print(f"   All {len(lines)} lines:")
            print("=" * 70)
            print(content)
else:
    print("   (Log file not found)")
