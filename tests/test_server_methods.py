import asyncio
import sys
from microsoft_graph_mcp_server.server import MicrosoftGraphMCPServer

async def test():
    try:
        server = MicrosoftGraphMCPServer()
        print('Server created successfully')
        print(f'Server object type: {type(server.server)}')
        
        public_methods = [m for m in dir(server.server) if not m.startswith('_')]
        print(f'Server public methods: {public_methods}')
        
        return True
    except Exception as e:
        print(f'Error: {e}')
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    result = asyncio.run(test())
    sys.exit(0 if result else 1)
