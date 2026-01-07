"""Test script to simulate uvx startup and verify logging."""

import sys
import os
from pathlib import Path

print("=" * 70)
print("TESTING UVX STARTUP LOGGING")
print("=" * 70)

# Clear log file first
log_file = Path(__file__).parent / 'microsoft_graph_mcp_server' / 'mcp_server_auth.log'
if log_file.exists():
    os.remove(log_file)
    print(f"\n✓ Cleared log file: {log_file}")
else:
    print(f"\n✗ Log file not found: {log_file}")

print("\nSimulating: uvx microsoft-graph-mcp-server")
print("=" * 70)

# Import main module - this should trigger logging configuration
print("\nImporting main module...")
try:
    from microsoft_graph_mcp_server.main import main
    print("✓ Main module imported successfully")
except Exception as e:
    print(f"✗ Import failed: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Check if log file was created
print("\n" + "=" * 70)
print("CHECKING LOG FILE")
print("=" * 70)

if log_file.exists():
    print(f"\n✓ Log file exists: {log_file}")
    with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        if content:
            print(f"\nLog file size: {len(content)} bytes")
            print("\n--- Log file content ---")
            print(content)
            print("--- End of log file ---")
        else:
            print("\n✗ Log file is empty")
else:
    print(f"\n✗ Log file was NOT created")
    print(f"Expected location: {log_file}")

print("\n" + "=" * 70)
print("TEST COMPLETE")
print("=" * 70)
