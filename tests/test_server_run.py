import asyncio
import sys
import traceback
from microsoft_graph_mcp_server.server import MicrosoftGraphMCPServer
import mcp.server.stdio

async def test_run():
    try:
        server = MicrosoftGraphMCPServer()
        print('Server created successfully')
        
        print('Attempting to run server with stdio...')
        try:
            async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
                print('STDIO server created')
                
                from mcp.server.models import InitializationOptions
                init_options = InitializationOptions(
                    server_name="microsoft-graph-mcp-server",
                    server_version="0.1.0",
                    capabilities=server.server.get_capabilities(
                        notification_options=None,
                        experimental_capabilities=None
                    )
                )
                print(f'Initialization options: {init_options}')
                
                print('Starting server.run()...')
                await server.server.run(
                    read_stream,
                    write_stream,
                    init_options
                )
        except Exception as e:
            print(f'Error during server run: {e}')
            traceback.print_exc()
            return False
            
    except Exception as e:
        print(f'Error during server creation: {e}')
        traceback.print_exc()
        return False
    
    return True

if __name__ == "__main__":
    try:
        result = asyncio.run(test_run())
        sys.exit(0 if result else 1)
    except Exception as e:
        print(f'Unhandled exception: {e}')
        traceback.print_exc()
        sys.exit(1)
