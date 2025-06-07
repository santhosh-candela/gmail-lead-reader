from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import base64
import re
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime

print("🚀 Starting Gmail to Excel Export Script...")

# Step 1: Auth
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
print("🔑 Authenticating...")
flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
creds = flow.run_local_server(port=0)
service = build('gmail', 'v1', credentials=creds)
print("✅ Authentication successful.")

# Step 2: Get label map
print("📥 Fetching Gmail labels...")
label_results = service.users().labels().list(userId='me').execute()
label_map = {label['name']: label['id'] for label in label_results['labels']}
print("🏷️ Your Gmail labels:")
for k, v in label_map.items():
    print(f"   • {k}: {v}")

# Step 3: Define custom labels (including 'Qualified' as shown in your sample)
custom_labels = ['Follow Up', 'Sales Made', 'Callback', 'Quoted', 'Qualified']
custom_label_ids = [label_map[label] for label in custom_labels if label in label_map]
print(f"\n📌 Using custom labels: {custom_labels}")
print(f"🔎 Found label IDs: {custom_label_ids}")

if not custom_label_ids:
    print("❌ No matching custom labels found.")
    exit()

# Step 4: Fetch messages
print("\n📬 Fetching messages...")
all_messages = []
for label_id in custom_label_ids:
    print(f"🔎 Searching messages with label ID: {label_id}")
    result = service.users().messages().list(userId='me', labelIds=[label_id], maxResults=500).execute()
    fetched = result.get('messages', [])
    print(f"   ➤ Fetched {len(fetched)} messages")
    all_messages += fetched

print(f"\n📦 Total messages before deduplication: {len(all_messages)}")

# Step 5: Deduplicate messages
unique_messages = {msg['id']: msg for msg in all_messages}
messages = list(unique_messages.values())
print(f"✅ Unique messages after deduplication: {len(messages)}")

# Step 6: Extract details
data = []

print("\n📝 Parsing message details...\n")

# Improved extraction function
def clean_extract(text, field_name):
    """Improved extraction that stops at common delimiters and handles various formats"""
    patterns = {
        'name': [
            r'(?:Customer\s+Name|Full\s+Name|Name)\s*[:=]\s*([^\n•<>]+?)(?=\s*(?:[•\n]|Email|Phone|$))',
            r'(?:Name)\s*[:=]\s*([^\n•<>]+?)(?=\s*(?:[•\n]|$))',
            r'(?:From|Sender)\s*[:=]\s*([^\n•<>]+?)(?=\s*(?:[•\n]|$))'
        ],
        'email': [
            r'(?:Email|E-Mail|Email\s+Address)\s*[:=]\s*([^\s•<>\n]+@[^\s•<>\n]+)',
            r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
        ],
        'phone': [
            r'(?:Phone|Mobile|Contact|Phone\s+Number|Contact\s+Number)\s*[:=]\s*(\+?1?\s*\(?[0-9]{3}\)?\s*[0-9]{3}[-.\s]*[0-9]{4})',
            r'(?:Tel|Telephone)\s*[:=]\s*(\+?1?\s*\(?[0-9]{3}\)?\s*[0-9]{3}[-.\s]*[0-9]{4})',
            r'(\+1\s*\([0-9]{3}\)\s*[0-9]{3}\s*[0-9]{4})'
        ]
    }
    
    field_patterns = patterns.get(field_name.lower(), [])
    for pattern in field_patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            result = match.group(1).strip()
            # Clean up common artifacts
            result = re.sub(r'[<>"\']', '', result)
            result = result.strip()
            if result and len(result) > 1:  # Ensure we have meaningful content
                return result
    return ""

for index, msg in enumerate(messages, 1):
    print(f"📧 Parsing message {index}/{len(messages)}: ID = {msg['id']}")
    msg_detail = service.users().messages().get(userId='me', id=msg['id'], format='full').execute()
    headers = msg_detail['payload']['headers']
    msg_labels = msg_detail.get('labelIds', [])

    # Extract message date and format it
    date = ''
    for header in headers:
        if header['name'].lower() == 'date':
            date_str = header['value']
            try:
                # Parse the date string into a datetime object
                date_obj = datetime.strptime(date_str.split('(')[0].strip(), '%a, %d %b %Y %H:%M:%S %z')
                # Format as YYYY-MM-DD HH:MM:SS
                date = date_obj.strftime('%Y-%m-%d %H:%M:%S')
            except:
                date = date_str
            break

    # Extract body (plain or HTML)
    def extract_body(payload):
        body = ''
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain':
                    data = part['body'].get('data')
                    if data:
                        body += base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                elif part['mimeType'] == 'text/html':
                    data = part['body'].get('data')
                    if data:
                        html = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                        body += BeautifulSoup(html, 'html.parser').get_text()
                elif 'parts' in part:
                    # Handle nested parts
                    body += extract_body(part)
        elif 'body' in payload and 'data' in payload['body']:
            body = base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8', errors='ignore')
        
        # Clean up the body text
        body = body.replace('\r', ' ').replace('\n', ' ')
        body = ' '.join(body.split())  # Remove extra whitespace
        return body

    try:
        body = extract_body(msg_detail['payload'])
        print("📄 Message body preview:")
        print(body[:300] + '...' if len(body) > 300 else body)
    except Exception as e:
        print(f"❌ Failed to extract body: {e}")
        body = ''

    # Extract email from headers
    sender_email = ''
    for header in headers:
        if header['name'].lower() == 'from':
            # Extract email from "Name <email@domain.com>" format
            match = re.search(r'<([^>]+)>', header['value'])
            if match:
                sender_email = match.group(1)
            else:
                # If no angle brackets, use the whole value if it looks like an email
                if '@' in header['value']:
                    sender_email = header['value'].strip()
            break

    # Extract name from headers
    sender_name = ''
    for header in headers:
        if header['name'].lower() == 'from':
            # Extract name from "Name <email@domain.com>" format
            match = re.match(r'^([^<]+)', header['value'])
            if match:
                sender_name = match.group(1).strip().strip('"')
            break

    # Use improved clean_extract function
    name = clean_extract(body, 'name') or sender_name
    email = clean_extract(body, 'email') or sender_email
    phone = clean_extract(body, 'phone')
    
    # If no phone found, try fallback pattern
    if not phone:
        phone_match = re.search(r'(\+1\s*\([0-9]{3}\)\s*[0-9]{3}\s*[0-9]{4})', body)
        if phone_match:
            phone = phone_match.group(1)

    # Label to status - use 'Qualified' as default like in your sample
    status = 'Qualified'
    for lbl, lid in label_map.items():
        if lid in msg_labels and lid in custom_label_ids and lbl != 'Qualified':
            status = lbl  # Only override if it's one of our other custom labels
            break

    # Print in the requested format
    print(f"Customer Name: {name}")
    print(f"Email: {email}")
    print(f"Phone: {phone}")
    print(f"Created At: {date}")
    print(f"Status: {status}")
    print("-" * 60)

    data.append([name, email, phone, date, status])

# Step 7: Export to Excel
print("📤 Exporting to gmail_leads.xlsx...")
df = pd.DataFrame(data, columns=[
    "Customer Name", "Email", "Phone", "Created At", "status"
])

# Save to Excel - using openpyxl as it's more reliable for formatting
try:
    df.to_excel("gmail_leads.xlsx", index=False, engine='openpyxl')
    print("✅ Export complete. File saved as gmail_leads.xlsx")
except ImportError:
    # Fallback if openpyxl not installed
    df.to_excel("gmail_leads.xlsx", index=False)
    print("✅ Export complete (using default engine). File saved as gmail_leads.xlsx")

print("🚀 Gmail to Excel Export Script completed.")