import re

# Simulate the HTML cleaning process
original_body = """
<div><img alt="Banner" data-outlook-trace="F:1|T:1" src="cid:banner"> </div>
<div>
<p style="margin:0"></p>
<p style="margin:0">Dear Meng Ning IL Luo,</p>
<br>
<p style="margin:0">Your approval is needed for the following certification request.</p>
<br>
<p style="margin:0">Learner: Amit Srivastava</p>
<p style="margin:0">Certification: Certified Kubernetes Administrator (CKA)</p>
<p style="margin:0">Vendor: CNCF</p>
<br>
<p style="margin:0"><strong>Action Required:</strong>&nbsp;Please Approve or Deny this request within 3 working days in the Learning Request Tool via this link:
<a href="https://lrt.yourlearning.ibm.com/requests/view/CRT274064?mode=approval">
https://lrt.yourlearning.ibm.com/requests/view/CRT274064?mode=approval</a></p>
<br>
<p style="margin:0">Thank you,</p>
<br>
<p style="margin:0">Learning and Knowledge India</p>
<p></p>
</div>
"""

print("Original HTML:")
print(original_body)
print("\n" + "="*80 + "\n")

body_content = original_body

# Step 1: Extract body content
body_match = re.search(r'<body[^>]*>(.*?)</body>', original_body, re.DOTALL | re.IGNORECASE)
if body_match:
    body_content = body_match.group(1)
    print("After extracting body:")
    print(body_content)
    print("\n" + "="*80 + "\n")

# Step 2: Remove images with cid
body_content = re.sub(r'<img[^>]+src=["\']cid:[^"\']*["\'][^>]*>', '', body_content, flags=re.IGNORECASE)
print("After removing images:")
print(body_content)
print("\n" + "="*80 + "\n")

# Step 3: Remove o:p tags
body_content = re.sub(r'<o:p>\s*</o:p>', '', body_content, flags=re.IGNORECASE)
body_content = re.sub(r'<o:p>.*?</o:p>', '', body_content, flags=re.IGNORECASE | re.DOTALL)
print("After removing o:p tags:")
print(body_content)
print("\n" + "="*80 + "\n")

# Step 4: Remove span tags
body_content = re.sub(r'<span[^>]*>(.*?)</span>', r'\1', body_content, flags=re.IGNORECASE | re.DOTALL)
print("After removing span tags:")
print(body_content)
print("\n" + "="*80 + "\n")

# Step 5: Convert links
body_content = re.sub(r'<a[^>]+href=["\']([^"\']*)["\'][^>]*>(.*?)</a>', r'\2 (\1)', body_content, flags=re.IGNORECASE | re.DOTALL)
print("After converting links:")
print(body_content)
print("\n" + "="*80 + "\n")

# Step 6: Remove p tags
body_content = re.sub(r'<p[^>]*>', '', body_content, flags=re.IGNORECASE)
body_content = re.sub(r'</p>', '\n\n', body_content, flags=re.IGNORECASE)
print("After removing p tags:")
print(body_content)
print("\n" + "="*80 + "\n")

# Step 7: Remove div tags
body_content = re.sub(r'<div[^>]*>', '', body_content, flags=re.IGNORECASE)
body_content = re.sub(r'</div>', '\n\n', body_content, flags=re.IGNORECASE)
print("After removing div tags:")
print(body_content)
print("\n" + "="*80 + "\n")

# Step 8: Convert br to newlines
body_content = re.sub(r'<br\s*/?>', '\n', body_content, flags=re.IGNORECASE)
print("After converting br to newlines:")
print(body_content)
print("\n" + "="*80 + "\n")

# Step 9: Remove remaining HTML tags
body_content = re.sub(r'<[^>]+>', '', body_content, flags=re.IGNORECASE)
print("After removing remaining HTML tags:")
print(body_content)
print("\n" + "="*80 + "\n")

# Step 10: Convert HTML entities
body_content = re.sub(r'&nbsp;', ' ', body_content)
body_content = re.sub(r'&lt;', '<', body_content)
body_content = re.sub(r'&gt;', '>', body_content)
body_content = re.sub(r'&amp;', '&', body_content)
print("After converting HTML entities:")
print(body_content)
print("\n" + "="*80 + "\n")

# Step 11: Normalize newlines
body_content = re.sub(r'\n{3,}', '\n\n', body_content)
print("After normalizing newlines:")
print(body_content)
print("\n" + "="*80 + "\n")

# Step 12: Convert to HTML
paragraphs = body_content.replace('\r\n', '\n').split('\n\n')
html_body = '<br><br>'.join([p.strip().replace('\n', '<br>') for p in paragraphs if p.strip()])
print("Final HTML output:")
print(html_body)
