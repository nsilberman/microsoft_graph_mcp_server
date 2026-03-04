"""HTML utilities for Microsoft Graph MCP Server."""

import re
from typing import Optional


def normalize_email_html(html: Optional[str]) -> Optional[str]:
    """Normalize HTML to fix common spacing issues in email clients.
    
    Email clients often convert whitespace/newlines between block elements to <br> tags,
    causing excessive spacing. This function removes unnecessary whitespace between
    block-level elements to prevent this issue.
    
    Fixes:
    - Remove whitespace between block-level elements (prevents email client from adding <br>)
    - Remove <br> tags between block elements
    - Normalize multiple consecutive <br> tags
    
    Args:
        html: HTML content to normalize
        
    Returns:
        Normalized HTML content, or None if input is None
        
    Example:
        >>> normalize_email_html('<p>Hello</p>\\n\\n<p>World</p>')
        '<p>Hello</p><p>World</p>'
    """
    if not html:
        return html
    
    # Block-level elements that should have no whitespace between them
    block_elements = ['p', 'div', 'ul', 'ol', 'li', 'blockquote', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'tr', 'td', 'th', 'hr', 'pre', 'dl', 'dt', 'dd']
    
    # Build pattern for block elements (for opening tags)
    # Matches: <p>, <p class="foo">, etc.
    open_block_pattern = rf'<(?:{"|".join(block_elements)})(?:\s[^>]*)?>'
    # Matches closing tags: </p>, </div>, etc.
    close_block_pattern = rf'</(?:{"|".join(block_elements)})>'
    
    # 1. Remove whitespace (including newlines) between closing block tag and opening block tag
    # Pattern: </tag> [whitespace] <tag>
    # We need to capture both the closing tag and the opening tag
    html = re.sub(
        rf'({close_block_pattern})\s+({open_block_pattern})',
        r'\1\2',
        html,
        flags=re.IGNORECASE
    )
    
    # 2. Remove whitespace between closing block tag and <br>
    html = re.sub(
        rf'({close_block_pattern})\s+(<br\s*/?>)',
        r'\1\2',
        html,
        flags=re.IGNORECASE
    )
    
    # 3. Remove whitespace between <br> and opening block tag
    html = re.sub(
        rf'(<br\s*/?>)\s+({open_block_pattern})',
        r'\1\2',
        html,
        flags=re.IGNORECASE
    )
    
    # 4. Remove <br> between closing block tag and opening block tag
    # Pattern: </tag> <br> <tag>  -> </tag><tag>
    html = re.sub(
        rf'({close_block_pattern})\s*<br\s*/?>\s*({open_block_pattern})',
        r'\1\2',
        html,
        flags=re.IGNORECASE
    )
    
    # 5. Remove <br> at the beginning of block tags (inside)
    # Pattern: <p><br>  -> <p>
    html = re.sub(
        rf'({open_block_pattern})<br\s*/?>',
        r'\1',
        html,
        flags=re.IGNORECASE
    )
    
    # 6. Remove <br> at the end of block tags (before closing)
    # Pattern: <br></p> -> </p>
    html = re.sub(
        rf'<br\s*/?>({close_block_pattern})',
        r'\1',
        html,
        flags=re.IGNORECASE
    )
    
    # 7. Normalize multiple consecutive <br> tags to a single <br>
    html = re.sub(
        r'(<br\s*/?>\s*){2,}',
        '<br>',
        html,
        flags=re.IGNORECASE
    )
    
    return html
