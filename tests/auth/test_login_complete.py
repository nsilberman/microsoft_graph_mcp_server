"""Complete test for login and status check - displays all logs."""

import sys
import os
import logging
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

# CRITICAL: Configure logging BEFORE importing any other modules!
log_file = Path(__file__).parent / 'microsoft_graph_mcp_server' / 'mcp_server_auth.log'
log_file.parent.mkdir(parents=True, exist_ok=True)

# Create handlers
file_handler = logging.FileHandler(log_file, mode='a')
console_handler = logging.StreamHandler(sys.stdout)

# Set format
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Configure root logger
root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)
root_logger.addHandler(file_handler)
root_logger.addHandler(console_handler)

# Configure specific loggers
device_flow_logger = logging.getLogger('microsoft_graph_mcp_server.auth_modules.device_flow')
device_flow_logger.setLevel(logging.DEBUG)
device_flow_logger.addHandler(file_handler)
device_flow_logger.addHandler(console_handler)
device_flow_logger.propagate = False

auth_manager_logger = logging.getLogger('microsoft_graph_mcp_server.auth_modules.auth_manager')
auth_manager_logger.setLevel(logging.DEBUG)
auth_manager_logger.addHandler(file_handler)
auth_manager_logger.addHandler(console_handler)
auth_manager_logger.propagate = False

# Enable MSAL debug logging
msal_logger = logging.getLogger('msal')
msal_logger.setLevel(logging.DEBUG)
msal_logger.addHandler(file_handler)
msal_logger.addHandler(console_handler)
msal_logger.propagate = False

logging.info("=" * 70)
logging.info("Test script logging initialized")
logging.info(f"Log file: {log_file}")
logging.info("=" * 70)

# NOW we can import other modules
import asyncio
import json
import time

print("=" * 70)
print("MCP SERVER LOGIN & STATUS CHECK TEST")
print("=" * 70)

# Display file locations
log_file = Path(__file__).parent / 'microsoft_graph_mcp_server' / 'mcp_server_auth.log'
token_file = Path.home() / '.microsoft_graph_mcp_tokens.json'
device_flow_file = Path.home() / '.microsoft_graph_mcp_device_flows.json'

print(f"\nFile locations:")
print(f"  Log file: {log_file}")
print(f"  Token file: {token_file}")
print(f"  Device flow file: {device_flow_file}")

# Run tests
print("\n" + "-" * 70)
print("Running tests...")
print("-" * 70 + "\n")

# Import triggers logging
from microsoft_graph_mcp_server.auth_modules.auth_manager import GraphAuthManager

auth_manager = GraphAuthManager()

print("1. Checking initial login status...")
s1 = asyncio.run(auth_manager.check_login_status())
print(f"   Status: {s1['status']}")
print(f"   Message: {s1['message'][:80]}...")

print("\n2. Initiating login...")
l1 = asyncio.run(auth_manager.login())
print(f"   Status: {l1['status']}")
device_code = l1.get('device_code', '')
user_code = l1.get('user_code', '')
verification_uri = l1.get('verification_uri', '')
print(f"   User Code: {user_code}")
print(f"   Verification URI: {verification_uri}")
print(f"   Device Code (first 50): {device_code[:50]}...")

print("\n3. Checking login status with device_code...")
s2 = asyncio.run(auth_manager.check_login_status(device_code))
print(f"   Status: {s2['status']}")
print(f"   Message: {s2['message'][:100]}...")

print("\n4. Checking login status without device_code...")
s3 = asyncio.run(auth_manager.check_login_status())
print(f"   Status: {s3['status']}")
print(f"   Message: {s3['message'][:100]}...")

# Display logs
print("\n" + "=" * 70)
print("LOG FILE CONTENTS")
print("=" * 70)

if log_file.exists():
    with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        if content:
            print(content)
        else:
            print("(Log file is empty)")
else:
    print("(Log file not found)")

# Display token files
print("\n" + "=" * 70)
print("TOKEN FILES")
print("=" * 70)

if token_file.exists():
    with open(token_file, 'r') as f:
        tokens = json.load(f)
        print(f"\n✓ Token file: {token_file}")
        print(f"  Authenticated: {tokens.get('authenticated')}")
        print(f"  Has access_token: {bool(tokens.get('access_token'))}")
        if tokens.get('token_expiry'):
            remaining = tokens['token_expiry'] - time.time()
            print(f"  Expires in: {remaining:.0f}s ({remaining/60:.0f}m)")
else:
    print(f"\n✗ Token file not found: {token_file}")

if device_flow_file.exists():
    with open(device_flow_file, 'r') as f:
        flows = json.load(f)
        print(f"\n✓ Device flow file: {device_flow_file}")
        print(f"  Number of flows: {len(flows)}")
        if flows:
            first_key = list(flows.keys())[0]
            flow = flows[first_key]
            print(f"  First flow:")
            print(f"    User code: {flow.get('user_code')}")
            print(f"    Expires in: {flow.get('expires_in')}s")
else:
    print(f"\n✗ Device flow file not found: {device_flow_file}")

print("\n" + "=" * 70)
print("✓ TEST COMPLETE")
print("=" * 70)

print("\nNext steps to complete authentication:")
print("  1. Open the verification URL in a browser")
print("  2. Enter the user code")
print("  3. Run check_status again with the device_code")
print("=" * 70)
