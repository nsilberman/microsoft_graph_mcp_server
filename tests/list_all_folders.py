import asyncio
import sys

sys.path.insert(0, ".")

from microsoft_graph_mcp_server.clients.email_client import EmailClient


async def list_all_folders():
    client = EmailClient()
    folders = await client.list_mail_folders()

    print("All folders:")
    for folder in folders:
        if "test" in folder["path"].lower() or "appricate" in folder["path"].lower():
            print(f"  *** {folder['path']}: {folder['emailCount']} emails")
        else:
            print(f"  {folder['path']}: {folder['emailCount']} emails")


if __name__ == "__main__":
    asyncio.run(list_all_folders())
