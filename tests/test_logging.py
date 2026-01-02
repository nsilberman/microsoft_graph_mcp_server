"""Simple test to verify logging works."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

# Clear log file
log_file = Path(__file__).parent / 'microsoft_graph_mcp_server' / 'mcp_server_auth.log'
print(f"Clearing log file: {log_file}")
with open(log_file, 'w') as f:
    f.write("")

print("\n[1] Importing main module...")
from microsoft_graph_mcp_server import main

print("[2] Importing device_flow module...")
from microsoft_graph_mcp_server.auth_modules.device_flow import DeviceFlowManager

print("[3] Importing auth_manager module...")
from microsoft_graph_mcp_server.auth_modules.auth_manager import GraphAuthManager

print("\n[4] Checking log file after imports...")

if log_file.exists():
    with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        if content:
            print("Log file has content!")
            print(f"Last 10 lines:\n{content}")
        else:
            print("⚠️ Log file is empty after imports!")
else:
    print("⚠️ Log file not found!")

print("\n[5] Checking if device_flow logger has handlers...")
import logging
device_flow_logger = logging.getLogger('microsoft_graph_mcp_server.auth_modules.device_flow')
print(f"   Logger name: {device_flow_logger.name}")
print(f"   Logger level: {device_flow_logger.level}")
print(f"   Logger handlers: {device_flow_logger.handlers}")
print(f"   Logger propagate: {device_flow_logger.propagate}")

print("\n✓ Test complete!")
print(f"\nNow check the log file: {log_file}")
