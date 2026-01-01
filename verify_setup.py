"""Quick verification script to confirm MCP server is properly configured."""

import sys


def main():
    """Verify the MCP server can be imported and instantiated."""
    print("=" * 60)
    print("Microsoft Graph MCP Server - Quick Verification")
    print("=" * 60)
    
    # Test 1: Import the package
    print("\n1. Testing package import...")
    try:
        import microsoft_graph_mcp_server
        print("   ✅ Package imported successfully")
    except ImportError as e:
        print(f"   ❌ Failed to import package: {e}")
        return False
    
    # Test 2: Import the server
    print("\n2. Testing server import...")
    try:
        from microsoft_graph_mcp_server.server import MicrosoftGraphMCPServer
        print("   ✅ Server class imported successfully")
    except ImportError as e:
        print(f"   ❌ Failed to import server: {e}")
        return False
    
    # Test 3: Import settings
    print("\n3. Testing settings import...")
    try:
        from microsoft_graph_mcp_server.config import settings
        print(f"   ✅ Settings loaded: {settings.server_name}")
    except ImportError as e:
        print(f"   ❌ Failed to import settings: {e}")
        return False
    
    # Test 4: Create server instance
    print("\n4. Testing server instantiation...")
    try:
        from microsoft_graph_mcp_server.server import MicrosoftGraphMCPServer
        server = MicrosoftGraphMCPServer()
        print("   ✅ Server instance created successfully")
    except Exception as e:
        print(f"   ❌ Failed to create server: {e}")
        return False
    
    # Test 5: Check entry point
    print("\n5. Testing entry point...")
    try:
        from microsoft_graph_mcp_server.main import main
        print("   ✅ Main entry point imported successfully")
    except ImportError as e:
        print(f"   ❌ Failed to import main: {e}")
        return False
    
    print("\n" + "=" * 60)
    print("✅ All verification checks passed!")
    print("=" * 60)
    print("\nThe MCP server is properly configured and ready to use.")
    print("\nNext steps:")
    print("1. Configure Claude Desktop with the UVX command")
    print("2. Restart Claude Desktop")
    print("3. Use the Microsoft Graph tools in Claude")
    print("\nConfiguration for Claude Desktop:")
    print('  "command": "uvx"')
    print('  "args": ["--from", ".", "microsoft-graph-mcp-server"]')
    print("\n" + "=" * 60)
    
    return True


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
