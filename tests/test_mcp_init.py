"""Test MCP server initialization without STDIO"""

import asyncio
from microsoft_graph_mcp_server.server import MicrosoftGraphMCPServer
from microsoft_graph_mcp_server.config import settings

async def test_server():
    """Test server initialization"""
    print("Testing MCP Server initialization...")
    print(f"Server Name: {settings.server_name}")
    print(f"Server Version: {settings.server_version}")
    print(f"Client ID: {settings.client_id}")
    print(f"Tenant ID: {settings.tenant_id}")
    
    try:
        server = MicrosoftGraphMCPServer()
        print("✓ Server initialized successfully")
        
        # Test tool listing
        tools = await server.server._handlers["tools/list"]()
        print(f"✓ Tools loaded: {len(tools)} tools available")
        
        for tool in tools:
            print(f"  - {tool.name}: {tool.description}")
        
        return True
    except Exception as e:
        print(f"✗ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = asyncio.run(test_server())
    exit(0 if success else 1)
