import asyncio
import sys
sys.path.insert(0, '.')

from microsoft_graph_mcp_server.clients.email_client import EmailClient

async def check_deleted_items():
    client = EmailClient()
    emails = await client.load_emails_by_folder('Deleted Items', top=50)
    print(f'Deleted Items count: {emails["count"]}')
    print('\nRecent emails in Deleted Items:')
    for e in emails['metadata'][:20]:
        print(f'{e["number"]}. {e["subject"]}')

if __name__ == "__main__":
    asyncio.run(check_deleted_items())
