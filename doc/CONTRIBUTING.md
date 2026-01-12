# Contributing Guide

Thank you for your interest in contributing to the Microsoft Graph MCP Server! We welcome all forms of contributions.

## Table of Contents

- [Development Setup](#development-setup)
- [Code Style Standards](#code-style-standards)
- [Adding a New Tool](#adding-a-new-tool)
- [Running Tests](#running-tests)
- [Committing Changes](#committing-changes)
- [Reporting Issues](#reporting-issues)

---

## Development Setup

### 1. Clone the Repository

```bash
git clone https://github.com/your-org/microsoft-graph-mcp-server.git
cd microsoft-graph-mcp-server
```

### 2. Create Virtual Environment

```bash
# Create virtual environment
python -m venv .venv

# Activate virtual environment
# On Linux/Mac:
source .venv/bin/activate

# On Windows:
.venv\Scripts\activate
```

### 3. Install Dependencies

```bash
# Install from requirements.txt
pip install -r requirements.txt

# Or install in development mode
pip install -e .
```

### 4. Configure Environment

Create a `.env` file in the project root:

```bash
# Copy example environment file
cp .env.example .env

# Edit .env with your configuration
# Microsoft Graph API credentials (optional - uses public client ID by default)
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
TENANT_ID=your-tenant-id

# User configuration
USER_TIMEZONE=America/New_York

# Pagination settings
PAGE_SIZE=5
LLM_PAGE_SIZE=20

# Search settings
DEFAULT_SEARCH_DAYS=90

# Contact search settings
CONTACT_SEARCH_LIMIT=10
```

---

## Code Style Standards

### Formatting

- Use **Black** for code formatting
- Configure Black to use 100 character line length
- Run Black before committing

```bash
# Install Black
pip install black

# Format code
black microsoft_graph_mcp_server/

# Format specific file
black microsoft_graph_mcp_server/handlers/email_handlers.py
```

### Import Sorting

- Use **isort** for import sorting
- Configure isort to match Black style
- Run isort before committing

```bash
# Install isort
pip install isort

# Sort imports
isort microsoft_graph_mcp_server/

# Sort specific file
isort microsoft_graph_mcp_server/handlers/email_handlers.py
```

### Code Conventions

Follow PEP 8 style guide:

1. **Docstrings**: Use Google-style docstrings
   ```python
   def my_function(param1: str, param2: int) -> bool:
       """Brief description.

       Extended description spanning multiple lines.

       Args:
           param1: Description of param1
           param2: Description of param2

       Returns:
           Description of return value

       Raises:
           ValueError: If param1 is invalid

       Examples:
           >>> my_function("test", 5)
           True
       """
   ```

2. **Type Hints**: Use type hints for all functions
   ```python
   def my_function(name: str, age: int) -> dict:
       return {"name": name, "age": age}
   ```

3. **Naming Conventions**:
   - Functions and variables: `snake_case`
   - Classes: `PascalCase`
   - Constants: `UPPER_SNAKE_CASE`
   - Private methods: `_prefix_method_name`

4. **Error Handling**:
   - Use specific exceptions when possible
   - Always handle exceptions in handlers and return formatted errors
   - Never suppress exceptions with bare `except:`

5. **Logging**:
   - Use Python's logging module (not print statements)
   - Use appropriate log levels: DEBUG, INFO, WARNING, ERROR, CRITICAL
   - Include context in log messages

6. **File Organization**:
   - Keep related functionality together
   - Use clear, descriptive filenames
   - Follow existing directory structure

---

## Adding a New Tool

### 1. Add Tool Definition

Add your tool definition to `microsoft_graph_mcp_server/tools/registry.py`:

```python
@staticmethod
def your_new_tool() -> types.Tool:
    """Your new tool definition."""
    return types.Tool(
        name="your_new_tool",
        description="""
        Clear description of what the tool does.

        WORKFLOW: When to use this tool.

        EXAMPLES:
            1) Basic usage:
               your_new_tool(param1="value")
            
            2) Advanced usage:
               your_new_tool(param1="value", param2="value")
        
        Returns: {success: boolean, data: object, message: string}
        """,
        inputSchema={
            "type": "object",
            "properties": {
                "param1": {
                    "type": "string",
                    "description": "Description of param1"
                },
                "param2": {
                    "type": "integer",
                    "description": "Description of param2",
                    "minimum": 1,
                    "maximum": 100
                }
            },
            "required": ["param1"]
        }
    )
```

Add to `get_all_tools()` method:

```python
@staticmethod
def get_all_tools() -> List[types.Tool]:
    """Get all available tools."""
    return [
        # ... existing tools ...
        ToolRegistry.your_new_tool(),  # Add your tool here
    ]
```

### 2. Add Handler Method

Add handler method to appropriate handler in `microsoft_graph_mcp_server/handlers/`:

```python
async def handle_your_new_tool(
    self, arguments: dict
) -> list[types.TextContent]:
    """Handle your_new_tool.
    
    Args:
        arguments: Dictionary containing:
            - param1 (str, required): Description of param1
            - param2 (int, optional): Description of param2
    
    Returns:
        List of TextContent containing JSON with:
        - success (bool): Operation success
        - data (object): Result data
        - message (str): Status message
        
    Raises:
        ValueError: If param1 is invalid
        Exception: If operation fails
        
    Examples:
        >>> await handler.handle_your_new_tool({"param1": "test"})
        
        >>> await handler.handle_your_new_tool({
        ...     "param1": "test",
        ...     "param2": 50
        ... })
    """
    try:
        param1 = arguments.get("param1")
        param2 = arguments.get("param2")
        
        # Validate input
        if not param1:
            return self._format_error("param1 is required")
        
        # Your logic here
        result = await self._do_something(param1, param2)
        
        return self._format_response({
            "success": True,
            "data": result,
            "message": "Operation successful"
        })
        
    except Exception as e:
        logger.error(f"Error in your_new_tool: {e}", exc_info=True)
        return self._format_error(f"Error: {str(e)}")
```

### 3. Add to Dispatch Table

Add tool to dispatch table in `microsoft_graph_mcp_server/server.py`:

```python
def _build_dispatch_table(self):
    """Build tool dispatch table for O(1) lookup."""
    self.tool_dispatch = {
        # ... existing entries ...
        "your_new_tool": (self.your_handler, "handle_your_new_tool"),
    }
```

### 4. Write Tests

Add tests in `tests/` directory:

```python
import pytest
from microsoft_graph_mcp_server.handlers.your_handler import YourHandler

class TestYourNewTool:
    """Test your_new_tool functionality."""
    
    @pytest.mark.asyncio
    async def test_basic_usage(self):
        """Test basic usage of your_new_tool."""
        handler = YourHandler()
        result = await handler.handle_your_new_tool({
            "param1": "test"
        })
        assert result is not None
    
    @pytest.mark.asyncio
    async def test_missing_param(self):
        """Test error handling for missing parameter."""
        handler = YourHandler()
        result = await handler.handle_your_new_tool({})
        assert "error" in result[0].text.lower()
    
    @pytest.mark.asyncio
    async def test_invalid_param(self):
        """Test error handling for invalid parameter."""
        handler = YourHandler()
        result = await handler.handle_your_new_tool({"param1": ""})
        assert "error" in result[0].text.lower()
```

### 5. Update Documentation

Update relevant documentation:

- Add tool to `README.md` in Available Tools section
- Add to `doc/TOOL_DEPENDENCIES.md` if tool has dependencies
- Add to `doc/ERROR_CODES.md` if tool introduces new errors
- Add examples to `examples/` directory

---

## Running Tests

### Run All Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=microsoft_graph_mcp_server

# Run with verbose output
pytest -v
```

### Run Specific Tests

```bash
# Run specific test file
pytest tests/test_email_handlers.py

# Run specific test class
pytest tests/test_email_handlers.py::TestEmailHandler

# Run specific test
pytest tests/test_email_handlers.py::TestEmailHandler::test_search_emails
```

### Test Structure

```
tests/
├── unit/                          # Unit tests
│   ├── test_auth_handlers.py
│   ├── test_email_handlers.py
│   ├── test_calendar_handlers.py
│   └── ...
├── integration/                    # Integration tests (TODO)
│   ├── test_auth_workflow.py
│   ├── test_email_workflow.py
│   ├── test_calendar_workflow.py
│   └── test_template_workflow.py
└── conftest.py                     # Pytest configuration
```

---

## Committing Changes

### 1. Create Feature Branch

```bash
git checkout -b feature/your-feature-name
```

### 2. Format Code

```bash
# Format with Black
black microsoft_graph_mcp_server/

# Sort imports with isort
isort microsoft_graph_mcp_server/
```

### 3. Run Tests

```bash
# Run tests
pytest

# Ensure all tests pass
```

### 4. Commit Changes

```bash
# Stage changes
git add .

# Commit with descriptive message
git commit -m "feat: Add your_new_tool for X functionality

- Add tool definition in tools/registry.py
- Add handler method in handlers/your_handler.py
- Add to dispatch table in server.py
- Add unit tests in tests/unit/
- Update documentation

Fixes #123"
```

### Commit Message Format

Use conventional commit messages:

- `feat:` New feature
- `fix:` Bug fix
- `docs:` Documentation changes
- `style:` Code style changes (formatting, no logic change)
- `refactor:` Code refactoring
- `test:` Adding or updating tests
- `chore:` Maintenance tasks

Examples:

```
feat: Add template management for email workflows

fix: Correct timezone handling in search_emails

docs: Update CONTRIBUTING.md with new tool guide

style: Format code with Black

refactor: Simplify email cache logic

test: Add integration tests for auth workflow

chore: Update dependencies
```

---

## Reporting Issues

### Issue Template

When reporting issues, please include:

1. **Description**: Clear description of the issue
2. **Steps to Reproduce**: Detailed steps to reproduce
3. **Expected Behavior**: What you expected to happen
4. **Actual Behavior**: What actually happened
5. **Environment Information**:
   - OS: Windows/Linux/Mac
   - Python version: `python --version`
   - Package version: `pip show microsoft-graph-mcp-server`
   - MCP client: Claude/Cursor/ChatGPT/etc.

### Example Issue Report

```markdown
## Issue: Tool X fails when parameter Y is empty

### Description
When calling `your_new_tool` with `param1` as empty string, the tool crashes instead of returning an error message.

### Steps to Reproduce

1. Call tool: `your_new_tool(param1="")`
2. Tool raises exception

### Expected Behavior

Tool should return:
```json
{
  "success": false,
  "message": "param1 cannot be empty"
}
```

### Actual Behavior

Tool raises `ValueError: param1 cannot be empty`

### Environment

- OS: Windows 11
- Python version: 3.11
- Package version: 0.1.0
- MCP client: Claude
```

---

## Project Structure

```
microsoft_graph_mcp_server/
├── microsoft_graph_mcp_server/    # Main package
│   ├── handlers/                    # MCP tool handlers
│   │   ├── auth_handlers.py
│   │   ├── email_handlers.py
│   │   ├── calendar_handlers.py
│   │   ├── user_handlers.py
│   │   ├── file_handlers.py
│   │   └── teams_handlers.py
│   ├── clients/                     # Graph API clients
│   │   ├── base_client.py
│   │   ├── email_client.py
│   │   ├── calendar_client.py
│   │   ├── file_client.py
│   │   ├── teams_client.py
│   │   └── user_client.py
│   ├── auth_modules/                 # Authentication modules
│   │   ├── auth_manager.py
│   │   ├── device_flow.py
│   │   └── token_manager.py
│   ├── cache/                       # Caching modules
│   │   ├── email_cache.py
│   │   ├── event_cache.py
│   │   └── template_cache.py
│   ├── tools/                       # Tool definitions
│   │   └── registry.py
│   ├── utils/                       # Utility functions
│   │   ├── date_handler.py
│   │   └── csv_utils.py
│   ├── validation/                  # Input validation
│   │   └── common.py
│   ├── config.py                    # Configuration
│   ├── graph_client.py              # Graph API client
│   ├── server.py                    # MCP server
│   └── main.py                     # Entry point
├── doc/                            # Documentation
│   ├── CONTRIBUTING.md             # This file
│   ├── TOOL_DEPENDENCIES.md       # Tool call sequences
│   ├── ERROR_CODES.md              # Error reference
│   └── ...                         # Other documentation
├── examples/                        # Usage examples (NEW)
│   ├── __init__.py
│   ├── auth_workflow.py
│   ├── email_workflow.py
│   ├── calendar_workflow.py
│   └── template_workflow.py
├── tests/                           # Tests
│   ├── unit/
│   └── integration/                # TODO
├── requirements.txt                  # Dependencies
├── .env.example                     # Environment template
├── README.md                       # Project README
└── pyproject.toml                  # Project configuration
```

---

## Pull Request Process

### 1. Fork Repository

Fork the repository and create a feature branch.

### 2. Make Changes

Follow the code style standards and add tests.

### 3. Run Tests

Ensure all tests pass before submitting PR.

### 4. Update Documentation

Update README and relevant documentation.

### 5. Submit Pull Request

Create a pull request with:
- Clear description of changes
- Reference any related issues
- Include screenshots if applicable
- Ensure CI/CD checks pass

---

## Code Review Guidelines

When reviewing pull requests:

1. **Code Style**: Check for adherence to PEP 8
2. **Docstrings**: Ensure all public functions have docstrings
3. **Tests**: Verify tests cover new functionality
4. **Error Handling**: Check for proper error handling
5. **Documentation**: Ensure documentation is updated

---

## Getting Help

If you need help with contributing:

1. Check existing issues and pull requests
2. Read the documentation in `doc/` directory
3. Review the code examples in `examples/` directory
4. Ask questions in GitHub Discussions

---

**Thank you for contributing!**

**Document Version:** 1.0
**Last Updated:** 2025-01-09
