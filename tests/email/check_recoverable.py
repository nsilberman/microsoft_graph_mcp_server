import asyncio
import sys

sys.path.insert(0, ".")

from microsoft_graph_mcp_server.clients.email_client import EmailClient


async def check_recoverable_items():
    client = EmailClient()

    try:
        emails = await client.load_emails_by_folder("Recoverable Items", top=50)
        print(f'Recoverable Items count: {emails["count"]}')
        print("\nRecent emails in Recoverable Items:")
        for e in emails["metadata"][:20]:
            print(f'{e["number"]}. {e["subject"]}')
    except Exception as e:
        print(f"Error accessing Recoverable Items: {e}")


if __name__ == "__main__":
    asyncio.run(check_recoverable_items())
