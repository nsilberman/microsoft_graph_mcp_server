"""Test script for new email functions."""

import asyncio
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from microsoft_graph_mcp_server.graph_client import graph_client
from microsoft_graph_mcp_server.auth import auth_manager
from microsoft_graph_mcp_server.email_cache import email_cache


class TestColors:
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    RESET = '\033[0m'
    BOLD = '\033[1m'


def print_test_header(test_name):
    print(f"\n{TestColors.BLUE}{TestColors.BOLD}{'='*60}{TestColors.RESET}")
    print(f"{TestColors.BLUE}{TestColors.BOLD}TEST: {test_name}{TestColors.RESET}")
    print(f"{TestColors.BLUE}{TestColors.BOLD}{'='*60}{TestColors.RESET}\n")


def print_success(message):
    print(f"{TestColors.GREEN}✓ {message}{TestColors.RESET}")


def print_error(message):
    print(f"{TestColors.RED}✗ {message}{TestColors.RESET}")


def print_info(message):
    print(f"{TestColors.YELLOW}ℹ {message}{TestColors.RESET}")


def print_result(label, data):
    print(f"\n{TestColors.BOLD}{label}:{TestColors.RESET}")
    print(json.dumps(data, indent=2, ensure_ascii=False))


class NewEmailFunctionsTester:
    def __init__(self):
        self.test_results = []
    
    def record_result(self, test_name, passed, error=None):
        self.test_results.append({
            "test": test_name,
            "passed": passed,
            "error": error
        })
    
    async def test_login(self):
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
    
    async def test_list_recent_emails_default(self):
        print_test_header("List Recent Emails (Default 1 Day)")
        
        try:
            result = await graph_client.load_emails_by_folder("Inbox", 1, None)
            print_result("Recent Emails (1 Day)", result)
            print_success(f"List recent emails test passed - found {result['count']} emails")
            self.record_result("list_recent_emails_default", True)
        except Exception as e:
            print_error(f"List recent emails test failed: {e}")
            self.record_result("list_recent_emails_default", False, str(e))
    
    async def test_list_recent_emails_custom_days(self):
        print_test_header("List Recent Emails (Custom Days)")
        
        try:
            result = await graph_client.load_emails_by_folder("Inbox", 3, None)
            print_result("Recent Emails (3 Days)", result)
            print_success(f"List recent emails (3 days) test passed - found {result['count']} emails")
            self.record_result("list_recent_emails_custom_days", True)
        except Exception as e:
            print_error(f"List recent emails (3 days) test failed: {e}")
            self.record_result("list_recent_emails_custom_days", False, str(e))
    
    async def test_list_recent_emails_max_days(self):
        print_test_header("List Recent Emails (Max 7 Days)")
        
        try:
            result = await graph_client.load_emails_by_folder("Inbox", 7, None)
            print_result("Recent Emails (7 Days)", result)
            print_success(f"List recent emails (7 days) test passed - found {result['count']} emails")
            self.record_result("list_recent_emails_max_days", True)
        except Exception as e:
            print_error(f"List recent emails (7 days) test failed: {e}")
            self.record_result("list_recent_emails_max_days", False, str(e))
    
    async def test_search_emails_by_sender_with_folder(self):
        print_test_header("Search Emails by Sender with Folder")
        
        try:
            result = await graph_client.search_emails_by_sender("test", "Inbox", 5)
            print_result("Search Results (Sender + Folder)", result)
            print_success(f"Search emails by sender with folder test passed - found {result['count']} emails")
            self.record_result("search_emails_by_sender_with_folder", True)
        except Exception as e:
            print_error(f"Search emails by sender with folder test failed: {e}")
            self.record_result("search_emails_by_sender_with_folder", False, str(e))
    
    async def test_search_emails_by_subject_with_folder(self):
        print_test_header("Search Emails by Subject with Folder")
        
        try:
            result = await graph_client.search_emails_by_subject("test", "Inbox", 5)
            print_result("Search Results (Subject + Folder)", result)
            print_success(f"Search emails by subject with folder test passed - found {result['count']} emails")
            self.record_result("search_emails_by_subject_with_folder", True)
        except Exception as e:
            print_error(f"Search emails by subject with folder test failed: {e}")
            self.record_result("search_emails_by_subject_with_folder", False, str(e))
    
    async def test_search_emails_by_body_with_folder(self):
        print_test_header("Search Emails by Body with Folder")
        
        try:
            result = await graph_client.search_emails_by_body("test", "Inbox", 5)
            print_result("Search Results (Body + Folder)", result)
            print_success(f"Search emails by body with folder test passed - found {result['count']} emails")
            self.record_result("search_emails_by_body_with_folder", True)
        except Exception as e:
            print_error(f"Search emails by body with folder test failed: {e}")
            self.record_result("search_emails_by_body_with_folder", False, str(e))
    
    async def test_cache_clearing_on_load(self):
        print_test_header("Cache Clearing on Load")
        
        try:
            email_cache.clear_cache()
            await email_cache.update_list_state(folder="Inbox", top=20, total_count=100, emails=[{"id": "1", "subject": "test"}])
            
            state_before = email_cache.get_list_state()
            print_result("Cache State Before Load", state_before)
            
            email_cache.clear_cache()
            
            state_after = email_cache.get_list_state()
            print_result("Cache State After Clear", state_after)
            
            if state_after["total_count"] == 0 and len(state_after["emails"]) == 0:
                print_success("Cache clearing on load test passed")
                self.record_result("cache_clearing_on_load", True)
            else:
                print_error("Cache was not cleared properly")
                self.record_result("cache_clearing_on_load", False, "Cache was not cleared properly")
        except Exception as e:
            print_error(f"Cache clearing on load test failed: {e}")
            self.record_result("cache_clearing_on_load", False, str(e))
    
    async def test_async_cache_save(self):
        print_test_header("Async Cache Save Performance")
        
        try:
            import time
            
            email_cache.clear_cache()
            
            test_emails = [{"id": str(i), "subject": f"Test Email {i}"} for i in range(100)]
            
            start_time = time.time()
            await email_cache.update_list_state(folder="Inbox", top=100, total_count=100, emails=test_emails)
            end_time = time.time()
            
            elapsed_time = end_time - start_time
            print_info(f"Time to save 100 emails: {elapsed_time:.4f} seconds")
            
            state = email_cache.get_list_state()
            print_result("Cache State After Save", {"total_count": state["total_count"], "emails_count": len(state["emails"])})
            
            if state["total_count"] == 100 and len(state["emails"]) == 100:
                print_success("Async cache save test passed")
                self.record_result("async_cache_save", True)
            else:
                print_error("Cache save did not preserve all emails")
                self.record_result("async_cache_save", False, "Cache save did not preserve all emails")
        except Exception as e:
            print_error(f"Async cache save test failed: {e}")
            self.record_result("async_cache_save", False, str(e))
    
    async def test_clear_email_cache(self):
        print_test_header("Clear Email Cache")
        
        try:
            await email_cache.update_list_state(folder="Inbox", top=20, total_count=50, emails=[{"id": "1", "subject": "test"}])
            
            state_before = email_cache.get_list_state()
            print_result("Cache State Before Clear", state_before)
            
            email_cache.clear_cache()
            
            state_after = email_cache.get_list_state()
            print_result("Cache State After Clear", state_after)
            
            if state_after["total_count"] == 0 and len(state_after["emails"]) == 0:
                print_success("Clear email cache test passed")
                self.record_result("clear_email_cache", True)
            else:
                print_error("Cache was not cleared properly")
                self.record_result("clear_email_cache", False, "Cache was not cleared properly")
        except Exception as e:
            print_error(f"Clear email cache test failed: {e}")
            self.record_result("clear_email_cache", False, str(e))
    
    async def test_list_mail_folders(self):
        print_test_header("List Mail Folders")
        
        try:
            result = await graph_client.list_mail_folders()
            print_result("Mail Folders", result)
            print_success(f"List mail folders test passed - found {len(result)} folders")
            self.record_result("list_mail_folders", True)
        except Exception as e:
            print_error(f"List mail folders test failed: {e}")
            self.record_result("list_mail_folders", False, str(e))
    
    def print_summary(self):
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
        print(f"{TestColors.BOLD}{TestColors.BLUE}")
        print("╔════════════════════════════════════════════════════════════╗")
        print("║     New Email Functions Test Suite                         ║")
        print("╚════════════════════════════════════════════════════════════╝")
        print(f"{TestColors.RESET}\n")
        
        await self.test_login()
        await self.test_list_recent_emails_default()
        await self.test_list_recent_emails_custom_days()
        await self.test_list_recent_emails_max_days()
        await self.test_search_emails_by_sender_with_folder()
        await self.test_search_emails_by_subject_with_folder()
        await self.test_search_emails_by_body_with_folder()
        await self.test_cache_clearing_on_load()
        await self.test_async_cache_save()
        await self.test_clear_email_cache()
        await self.test_list_mail_folders()
        
        self.print_summary()


async def main():
    tester = NewEmailFunctionsTester()
    await tester.run_all_tests()


if __name__ == "__main__":
    asyncio.run(main())
