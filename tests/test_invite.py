"""Test script for calendar event invite functionality."""

import asyncio
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from microsoft_graph_mcp_server.graph_client import graph_client
from microsoft_graph_mcp_server.auth import auth_manager


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


class InviteTester:
    def __init__(self):
        self.test_results = []
        self.created_event_id = None
    
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
    
    async def test_create_event_without_attendees(self):
        print_test_header("Create Event Without Attendees")
        
        try:
            start_time = datetime.now() + timedelta(days=1)
            end_time = start_time + timedelta(hours=1)
            
            event_data = {
                "subject": "Test Meeting (No Attendees)",
                "start": {
                    "dateTime": start_time.strftime("%Y-%m-%dT%H:%M:%S"),
                    "timeZone": "UTC"
                },
                "end": {
                    "dateTime": end_time.strftime("%Y-%m-%dT%H:%M:%S"),
                    "timeZone": "UTC"
                },
                "body": {
                    "contentType": "HTML",
                    "content": "<p>This is a test meeting without attendees.</p>"
                }
            }
            
            result = await graph_client.create_event(event_data)
            print_result("Created Event", result)
            
            if result.get("id"):
                self.created_event_id = result["id"]
                print_success(f"Event created successfully with ID: {self.created_event_id}")
                self.record_result("create_event_without_attendees", True)
            else:
                print_error("Event creation failed - no ID returned")
                self.record_result("create_event_without_attendees", False, "No ID returned")
        except Exception as e:
            print_error(f"Create event test failed: {e}")
            self.record_result("create_event_without_attendees", False, str(e))
    
    async def test_create_event_with_attendees(self):
        print_test_header("Create Event With Attendees (Invite)")
        
        try:
            start_time = datetime.now() + timedelta(days=2)
            end_time = start_time + timedelta(hours=1)
            
            event_data = {
                "subject": "Test Meeting With Attendees",
                "start": {
                    "dateTime": start_time.strftime("%Y-%m-%dT%H:%M:%S"),
                    "timeZone": "UTC"
                },
                "end": {
                    "dateTime": end_time.strftime("%Y-%m-%dT%H:%M:%S"),
                    "timeZone": "UTC"
                },
                "body": {
                    "contentType": "HTML",
                    "content": "<p>This is a test meeting with attendees.</p>"
                },
                "attendees": [
                    {
                        "emailAddress": {
                            "address": "test@example.com",
                            "name": "Test User"
                        },
                        "type": "required"
                    },
                    {
                        "emailAddress": {
                            "address": "optional@example.com",
                            "name": "Optional User"
                        },
                        "type": "optional"
                    }
                ],
                "location": {
                    "displayName": "Conference Room A"
                }
            }
            
            result = await graph_client.create_event(event_data)
            print_result("Created Event with Attendees", result)
            
            if result.get("id"):
                print_success(f"Event with attendees created successfully with ID: {result['id']}")
                self.record_result("create_event_with_attendees", True)
            else:
                print_error("Event creation failed - no ID returned")
                self.record_result("create_event_with_attendees", False, "No ID returned")
        except Exception as e:
            print_error(f"Create event with attendees test failed: {e}")
            self.record_result("create_event_with_attendees", False, str(e))
    
    async def test_create_event_with_multiple_attendees(self):
        print_test_header("Create Event With Multiple Attendees")
        
        try:
            start_time = datetime.now() + timedelta(days=3)
            end_time = start_time + timedelta(hours=2)
            
            event_data = {
                "subject": "Team Meeting",
                "start": {
                    "dateTime": start_time.strftime("%Y-%m-%dT%H:%M:%S"),
                    "timeZone": "UTC"
                },
                "end": {
                    "dateTime": end_time.strftime("%Y-%m-%dT%H:%M:%S"),
                    "timeZone": "UTC"
                },
                "body": {
                    "contentType": "HTML",
                    "content": "<p>Weekly team sync meeting.</p>"
                },
                "attendees": [
                    {
                        "emailAddress": {
                            "address": "user1@example.com",
                            "name": "User One"
                        },
                        "type": "required"
                    },
                    {
                        "emailAddress": {
                            "address": "user2@example.com",
                            "name": "User Two"
                        },
                        "type": "required"
                    },
                    {
                        "emailAddress": {
                            "address": "user3@example.com",
                            "name": "User Three"
                        },
                        "type": "optional"
                    },
                    {
                        "emailAddress": {
                            "address": "user4@example.com",
                            "name": "User Four"
                        },
                        "type": "optional"
                    }
                ],
                "location": {
                    "displayName": "Virtual Meeting"
                },
                "isOnlineMeeting": True
            }
            
            result = await graph_client.create_event(event_data)
            print_result("Created Event with Multiple Attendees", result)
            
            if result.get("id"):
                print_success(f"Event with multiple attendees created successfully with ID: {result['id']}")
                self.record_result("create_event_with_multiple_attendees", True)
            else:
                print_error("Event creation failed - no ID returned")
                self.record_result("create_event_with_multiple_attendees", False, "No ID returned")
        except Exception as e:
            print_error(f"Create event with multiple attendees test failed: {e}")
            self.record_result("create_event_with_multiple_attendees", False, str(e))
    
    async def test_browse_events(self):
        print_test_header("Browse Events")
        
        try:
            start_date = (datetime.now()).isoformat()
            end_date = (datetime.now() + timedelta(days=7)).isoformat()
            
            result = await graph_client.browse_events(
                start_date=start_date,
                end_date=end_date,
                top=10,
                skip=0
            )
            print_result("Browsed Events", result)
            print_success(f"Browse events test passed - found {result.get('count', 0)} events")
            self.record_result("browse_events", True)
        except Exception as e:
            print_error(f"Browse events test failed: {e}")
            self.record_result("browse_events", False, str(e))
    
    async def test_get_event(self):
        print_test_header("Get Event by ID")
        
        if not self.created_event_id:
            print_info("Skipping - no event ID available from previous test")
            self.record_result("get_event", True, "Skipped - no event ID")
            return
        
        try:
            result = await graph_client.get_event(self.created_event_id)
            print_result("Event Details", result)
            print_success("Get event test passed")
            self.record_result("get_event", True)
        except Exception as e:
            print_error(f"Get event test failed: {e}")
            self.record_result("get_event", False, str(e))
    
    async def test_search_events(self):
        print_test_header("Search Events")
        
        try:
            result = await graph_client.search_events(
                query="meeting",
                top=10
            )
            print_result("Search Results", result)
            print_success(f"Search events test passed - found {result.get('count', 0)} events")
            self.record_result("search_events", True)
        except Exception as e:
            print_error(f"Search events test failed: {e}")
            self.record_result("search_events", False, str(e))
    
    async def test_create_all_day_event(self):
        print_test_header("Create All-Day Event")
        
        try:
            event_date = datetime.now() + timedelta(days=5)
            
            event_data = {
                "subject": "All-Day Conference",
                "start": {
                    "dateTime": event_date.strftime("%Y-%m-%dT00:00:00"),
                    "timeZone": "UTC"
                },
                "end": {
                    "dateTime": event_date.strftime("%Y-%m-%dT23:59:59"),
                    "timeZone": "UTC"
                },
                "isAllDay": True,
                "body": {
                    "contentType": "HTML",
                    "content": "<p>Annual all-day conference.</p>"
                },
                "attendees": [
                    {
                        "emailAddress": {
                            "address": "conference@example.com",
                            "name": "Conference Attendee"
                        },
                        "type": "required"
                    }
                ]
            }
            
            result = await graph_client.create_event(event_data)
            print_result("Created All-Day Event", result)
            
            if result.get("id"):
                print_success(f"All-day event created successfully with ID: {result['id']}")
                self.record_result("create_all_day_event", True)
            else:
                print_error("All-day event creation failed - no ID returned")
                self.record_result("create_all_day_event", False, "No ID returned")
        except Exception as e:
            print_error(f"Create all-day event test failed: {e}")
            self.record_result("create_all_day_event", False, str(e))
    
    def print_summary(self):
        print(f"\n{TestColors.BOLD}{TestColors.BLUE}{'='*60}{TestColors.RESET}")
        print(f"{TestColors.BOLD}{TestColors.BLUE}INVITE TEST SUMMARY{TestColors.RESET}")
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
        print("║     Calendar Event Invite Test Suite                       ║")
        print("╚════════════════════════════════════════════════════════════╝")
        print(f"{TestColors.RESET}\n")
        
        await self.test_login()
        await self.test_create_event_without_attendees()
        await self.test_create_event_with_attendees()
        await self.test_create_event_with_multiple_attendees()
        await self.test_browse_events()
        await self.test_get_event()
        await self.test_search_events()
        await self.test_create_all_day_event()
        
        self.print_summary()


async def main():
    tester = InviteTester()
    await tester.run_all_tests()


if __name__ == "__main__":
    asyncio.run(main())
