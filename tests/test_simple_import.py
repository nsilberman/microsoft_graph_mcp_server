"""Very simple test - just import device_flow module."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

print("[Test] Importing device_flow module...")
from microsoft_graph_mcp_server.auth_modules.device_flow import DeviceFlowManager

print("[Test] Import complete!")
print("\nNow check the log file:")
log_file = Path(__file__).parent / 'microsoft_graph_mcp_server' / 'mcp_server_auth.log'
print(f"  {log_file}")

if log_file.exists():
    with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        lines = content.count('\n')
        print(f"\nLog file has {lines} lines")
        print(f"Last 20 lines:")
        print("=" * 70)
        all_lines = content.split('\n')
        for line in all_lines[-20:]:
            print(line)
else:
    print("(Log file not found)")
