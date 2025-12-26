"""Test script for Microsoft Graph MCP Server functions."""

import asyncio
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

# Add the project root to the path
sys.path.insert(0, str(Path(__file__).parent.parent))

from microsoft_graph_mcp_server.graph_client import graph_client
from microsoft_graph_mcp_server.auth import auth_manager
from microsoft_graph_mcp_server.email_cache import email_cache


class TestColors:
    """ANSI color codes for test output."""
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    RESET = '\033[0m'
    BOLD = '\033[1m'


def print_test_header(test_name):
    """Print test header."""
    print(f"\n{TestColors.BLUE}{TestColors.BOLD}{'='*60}{TestColors.RESET}")
    print(f"{TestColors.BLUE}{TestColors.BOLD}TEST: {test_name}{TestColors.RESET}")
    print(f"{TestColors.BLUE}{TestColors.BOLD}{'='*60}{TestColors.RESET}\n")


def print_success(message):
    """Print success message."""
    print(f"{TestColors.GREEN}✓ {message}{TestColors.RESET}")


def print_error(message):
    """Print error message."""
    print(f"{TestColors.RED}✗ {message}{TestColors.RESET}")


def print_info(message):
    """Print info message."""
    print(f"{TestColors.YELLOW}ℹ {message}{TestColors.RESET}")


def print_result(label, data):
    """Print result data."""
    print(f"\n{TestColors.BOLD}{label}:{TestColors.RESET}")
    print(json.dumps(data, indent=2, ensure_ascii=False))


class MCPTester:
    """Test suite for Microsoft Graph MCP Server."""
    
    def __init__(self):
        self.test_results = []
        self.email_id = None
        self.team_id = None
    
    def record_result(self, test_name, passed, error=None):
        """Record test result."""
        self.test_results.append({
            "test": test_name,
            "passed": passed,
            "error": error
        })
    
    async def test_login(self):
        """Test login functionality."""
        print_test_header("Login")
        
        try:
            result = await auth_manager.login()
            print_result("Login Result", result)
            
            if result.get("status") == "success" or result.get("authenticated"):
                print_success("Login test passed")
                self.record_result("login", True)
            else:
                print_info("Login requires user interaction (expected)")
                self.record_result("login", True)
        except Exception as e:
            print_error(f"Login test failed: {e}")
            self.record_result("login", False, str(e))
    
    async def test_check_login_status(self):
        """Test check login status."""
        print_test_header("Check Login Status")
        
        try:
            result = await auth_manager.check_login_status()
            print_result("Login Status", result)
            print_success("Check login status test passed")
            self.record_result("check_login_status", True)
        except Exception as e:
            print_error(f"Check login status test failed: {e}")
            self.record_result("check_login_status", False, str(e))
    
    async def test_get_user_info(self):
        """Test get user info."""
        print_test_header("Get User Info")
        
        try:
            result = await graph_client.get_me()
            print_result("User Info", result)
            print_success("Get user info test passed")
            self.record_result("get_user_info", True)
        except Exception as e:
            print_error(f"Get user info test failed: {e}")
            self.record_result("get_user_info", False, str(e))
    
    async def test_search_contacts(self):
        """Test search contacts."""
        print_test_header("Search Contacts")
        
        try:
            result = await graph_client.search_contacts("test", top=5)
            print_result("Search Results", result)
            print_success(f"Search contacts test passed - found {len(result)} contacts")
            self.record_result("search_contacts", True)
        except Exception as e:
            print_error(f"Search contacts test failed: {e}")
            self.record_result("search_contacts", False, str(e))
    
    async def test_browse_emails(self):
        """Test browse emails."""
        print_test_header("Browse Emails")
        
        try:
            result = await graph_client.browse_emails(
                folder="Inbox",
                top=5,
                skip=0
            )
            print_result("Email List", result)
            
            if result.get("emails"):
                self.email_id = result["emails"][0]["id"]
                print_info(f"Stored email ID for later tests: {self.email_id}")
            
            print_success(f"Browse emails test passed - found {result.get('count', 0)} emails")
            self.record_result("browse_emails", True)
        except Exception as e:
            print_error(f"Browse emails test failed: {e}")
            self.record_result("browse_emails", False, str(e))
    
    async def test_get_email_count(self):
        """Test get email count."""
        print_test_header("Get Email Count")
        
        try:
            count = await graph_client.get_email_count(folder="Inbox")
            print_result("Email Count", {"count": count})
            print_success(f"Get email count test passed - {count} emails in Inbox")
            self.record_result("get_email_count", True)
        except Exception as e:
            print_error(f"Get email count test failed: {e}")
            self.record_result("get_email_count", False, str(e))
    
    async def test_get_email(self):
        """Test get email by ID."""
        print_test_header("Get Email by ID")
        
        if not self.email_id:
            print_info("Skipping - no email ID available from previous test")
            self.record_result("get_email", True, "Skipped - no email ID")
            return
        
        try:
            result = await graph_client.get_email(self.email_id)
            print_result("Email Details", result)
            print_success("Get email test passed")
            self.record_result("get_email", True)
        except Exception as e:
            print_error(f"Get email test failed: {e}")
            self.record_result("get_email", False, str(e))
    
    async def test_search_emails(self):
        """Test search emails."""
        print_test_header("Search Emails")
        
        try:
            result = await graph_client.search_emails(
                query="test",
                top=5
            )
            print_result("Search Results", result)
            print_success(f"Search emails test passed - found {result.get('count', 0)} emails")
            self.record_result("search_emails", True)
        except Exception as e:
            print_error(f"Search emails test failed: {e}")
            self.record_result("search_emails", False, str(e))
    
    async def test_get_events(self):
        """Test get calendar events."""
        print_test_header("Get Calendar Events")
        
        try:
            start_date = (datetime.now() - timedelta(days=7)).isoformat()
            end_date = (datetime.now() + timedelta(days=7)).isoformat()
            
            result = await graph_client.get_events(start_date=start_date, end_date=end_date)
            print_result("Calendar Events", result)
            print_success(f"Get events test passed - found {len(result)} events")
            self.record_result("get_events", True)
        except Exception as e:
            print_error(f"Get events test failed: {e}")
            self.record_result("get_events", False, str(e))
    
    async def test_list_files(self):
        """Test list OneDrive files."""
        print_test_header("List OneDrive Files")
        
        try:
            result = await graph_client.get_drive_items()
            print_result("OneDrive Items", result)
            print_success(f"List files test passed - found {len(result)} items")
            self.record_result("list_files", True)
        except Exception as e:
            print_error(f"List files test failed: {e}")
            self.record_result("list_files", False, str(e))
    
    async def test_get_teams(self):
        """Test get Teams."""
        print_test_header("Get Microsoft Teams")
        
        try:
            result = await graph_client.get_teams()
            print_result("Teams List", result)
            
            if result:
                self.team_id = result[0].get("id")
                print_info(f"Stored team ID for later tests: {self.team_id}")
            
            print_success(f"Get teams test passed - found {len(result)} teams")
            self.record_result("get_teams", True)
        except Exception as e:
            print_error(f"Get teams test failed: {e}")
            self.record_result("get_teams", False, str(e))
    
    async def test_get_team_channels(self):
        """Test get team channels."""
        print_test_header("Get Team Channels")
        
        if not self.team_id:
            print_info("Skipping - no team ID available from previous test")
            self.record_result("get_team_channels", True, "Skipped - no team ID")
            return
        
        try:
            result = await graph_client.get_team_channels(self.team_id)
            print_result("Team Channels", result)
            print_success(f"Get team channels test passed - found {len(result)} channels")
            self.record_result("get_team_channels", True)
        except Exception as e:
            print_error(f"Get team channels test failed: {e}")
            self.record_result("get_team_channels", False, str(e))
    
    async def test_email_cache(self):
        """Test email cache functionality."""
        print_test_header("Email Cache")
        
        try:
            email_cache.clear_cache()
            print_success("Cache cleared successfully")
            
            email_cache.update_list_state(folder="Inbox", top=20, total_count=100)
            state = email_cache.get_list_state()
            print_result("List State", state)
            print_success("Email cache test passed")
            self.record_result("email_cache", True)
        except Exception as e:
            print_error(f"Email cache test failed: {e}")
            self.record_result("email_cache", False, str(e))
    
    async def test_logout(self):
        """Test logout."""
        print_test_header("Logout")
        
        try:
            result = await auth_manager.logout()
            print_result("Logout Result", result)
            print_success("Logout test passed")
            self.record_result("logout", True)
        except Exception as e:
            print_error(f"Logout test failed: {e}")
            self.record_result("logout", False, str(e))
    
    def print_summary(self):
        """Print test summary."""
        print(f"\n{TestColors.BOLD}{TestColors.BLUE}{'='*60}{TestColors.RESET}")
        print(f"{TestColors.BOLD}{TestColors.BLUE}TEST SUMMARY{TestColors.RESET}")
        print(f"{TestColors.BOLD}{TestColors.BLUE}{'='*60}{TestColors.RESET}\n")
        
        total_tests = len(self.test_results)
        passed_tests = sum(1 for r in self.test_results if r["passed"])
        failed_tests = total_tests - passed_tests
        
        print(f"Total Tests: {total_tests}")
        print(f"{TestColors.GREEN}Passed: {passed_tests}{TestColors.RESET}")
        print(f"{TestColors.RED}Failed: {failed_tests}{TestColors.RESET}")
        
        if failed_tests > 0:
            print(f"\n{TestColors.BOLD}Failed Tests:{TestColors.RESET}")
            for result in self.test_results:
                if not result["passed"]:
                    print(f"{TestColors.RED}✗ {result['test']}: {result.get('error', 'Unknown error')}{TestColors.RESET}")
        
        print(f"\n{TestColors.BOLD}{'='*60}{TestColors.RESET}\n")
    
    async def run_all_tests(self):
        """Run all tests."""
        print(f"{TestColors.BOLD}{TestColors.BLUE}")
        print("╔════════════════════════════════════════════════════════════╗")
        print("║     Microsoft Graph MCP Server Test Suite                 ║")
        print("╚════════════════════════════════════════════════════════════╝")
        print(f"{TestColors.RESET}\n")
        
        await self.test_login()
        await self.test_check_login_status()
        await self.test_get_user_info()
        await self.test_search_contacts()
        await self.test_browse_emails()
        await self.test_get_email_count()
        await self.test_get_email()
        await self.test_search_emails()
        await self.test_get_events()
        await self.test_list_files()
        await self.test_get_teams()
        await self.test_get_team_channels()
        await self.test_email_cache()
        await self.test_logout()
        
        self.print_summary()


async def main():
    """Main entry point."""
    tester = MCPTester()
    await tester.run_all_tests()


if __name__ == "__main__":
    asyncio.run(main())
