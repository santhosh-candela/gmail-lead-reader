# from googleapiclient.discovery import build
# from google_auth_oauthlib.flow import InstalledAppFlow
# import base64
# import re
# import pandas as pd
# from bs4 import BeautifulSoup
# from datetime import datetime
# import pytz

# print("🚀 Starting Gmail to Excel Export Script...")

# # Step 1: Auth
# SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
# print("🔑 Authenticating...")
# flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
# creds = flow.run_local_server(port=0)
# service = build('gmail', 'v1', credentials=creds)
# print("✅ Authentication successful.")

# # Step 2: Get label map
# print("📥 Fetching Gmail labels...")
# label_results = service.users().labels().list(userId='me').execute()
# label_map = {label['name']: label['id'] for label in label_results['labels']}
# print("🏷️ Your Gmail labels:")
# for k, v in label_map.items():
#     print(f"   • {k}: {v}")

# # Step 2.5: Office hours check function
# def is_within_office_hours(date_str):
#     """
#     Check if the message timestamp is today between 7AM-4PM PST, Monday-Friday
#     """
#     try:
#         # Set up PST timezone
#         pst = pytz.timezone('US/Pacific')
        
#         # Parse the email date
#         if '(' in date_str:
#             date_str = date_str.split('(')[0].strip()
        
#         # Parse date and convert to PST
#         date_obj = datetime.strptime(date_str, '%a, %d %b %Y %H:%M:%S %z')
#         pst_date = date_obj.astimezone(pst)
        
#         # Get current date in PST
#         now_pst = datetime.now(pst)
#         today_pst = now_pst.date()
        
#         # Check if message is from today
#         if pst_date.date() != today_pst:
#             return False, f"Not today (message date: {pst_date.date()}, today: {today_pst})"
        
#         # Check if it's Monday-Friday (0=Monday, 6=Sunday)
#         if pst_date.weekday() > 4:  # 5=Saturday, 6=Sunday
#             return False, f"Weekend (day: {pst_date.strftime('%A')})"
        
#         # Check if it's between 7AM-4PM PST
#         hour = pst_date.hour
#         if hour < 7 or hour >= 16:  # 16 = 4PM (exclusive)
#             return False, f"Outside office hours (time: {pst_date.strftime('%I:%M %p PST')})"
        
#         return True, f"Within office hours ({pst_date.strftime('%A, %I:%M %p PST')})"
        
#     except Exception as e:
#         return False, f"Error parsing date: {e}"

# print(f"\n⏰ Office Hours Filter: Monday-Friday, 7AM-4PM PST")
# print(f"📅 Current PST time: {datetime.now(pytz.timezone('US/Pacific')).strftime('%A, %B %d, %Y %I:%M %p PST')}")

# # Step 3: Define custom labels (including 'Qualified' as shown in your sample)
# custom_labels = ['Follow Up', 'Sales Made', 'Callback', 'Quoted', 'Qualified']
# custom_label_ids = [label_map[label] for label in custom_labels if label in label_map]
# print(f"\n📌 Using custom labels: {custom_labels}")
# print(f"🔎 Found label IDs: {custom_label_ids}")

# if not custom_label_ids:
#     print("❌ No matching custom labels found.")
#     exit()

# # Step 4: Fetch messages
# print("\n📬 Fetching messages...")
# all_messages = []
# for label_id in custom_label_ids:
#     print(f"🔎 Searching messages with label ID: {label_id}")
#     result = service.users().messages().list(userId='me', labelIds=[label_id], maxResults=500).execute()
#     fetched = result.get('messages', [])
#     print(f"   ➤ Fetched {len(fetched)} messages")
#     all_messages += fetched

# print(f"\n📦 Total messages before deduplication: {len(all_messages)}")

# # Step 5: Deduplicate messages
# unique_messages = {msg['id']: msg for msg in all_messages}
# messages = list(unique_messages.values())
# print(f"✅ Unique messages after deduplication: {len(messages)}")

# # Step 6: Filter by office hours and extract details
# data = []
# filtered_count = 0
# office_hours_count = 0

# print("\n📝 Parsing message details and applying office hours filter...\n")

# # Improved extraction function
# def clean_extract(text, field_name):
#     """Improved extraction that stops at common delimiters and handles various formats"""
#     patterns = {
#         'name': [
#             r'(?:Customer\s+Name|Full\s+Name|Name)\s*[:=]\s*([^\n•<>]+?)(?=\s*(?:[•\n]|Email|Phone|$))',
#             r'(?:Name)\s*[:=]\s*([^\n•<>]+?)(?=\s*(?:[•\n]|$))',
#             r'(?:From|Sender)\s*[:=]\s*([^\n•<>]+?)(?=\s*(?:[•\n]|$))'
#         ],
#         'email': [
#             r'(?:Email|E-Mail|Email\s+Address)\s*[:=]\s*([^\s•<>\n]+@[^\s•<>\n]+)',
#             r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
#         ],
#         'phone': [
#             r'(?:Phone|Mobile|Contact|Phone\s+Number|Contact\s+Number)\s*[:=]\s*(\+?1?\s*\(?[0-9]{3}\)?\s*[0-9]{3}[-.\s]*[0-9]{4})',
#             r'(?:Tel|Telephone)\s*[:=]\s*(\+?1?\s*\(?[0-9]{3}\)?\s*[0-9]{3}[-.\s]*[0-9]{4})',
#             r'(\+1\s*\([0-9]{3}\)\s*[0-9]{3}\s*[0-9]{4})'
#         ]
#     }
    
#     field_patterns = patterns.get(field_name.lower(), [])
#     for pattern in field_patterns:
#         match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
#         if match:
#             result = match.group(1).strip()
#             # Clean up common artifacts
#             result = re.sub(r'[<>"\']', '', result)
#             result = result.strip()
#             if result and len(result) > 1:  # Ensure we have meaningful content
#                 return result
#     return ""

# for index, msg in enumerate(messages, 1):
#     print(f"📧 Parsing message {index}/{len(messages)}: ID = {msg['id']}")
#     msg_detail = service.users().messages().get(userId='me', id=msg['id'], format='full').execute()
#     headers = msg_detail['payload']['headers']
#     msg_labels = msg_detail.get('labelIds', [])

#     # Extract message date and check office hours
#     date = ''
#     date_str = ''
#     within_hours = False
#     hours_reason = ''
    
#     for header in headers:
#         if header['name'].lower() == 'date':
#             date_str = header['value']
#             within_hours, hours_reason = is_within_office_hours(date_str)
            
#             try:
#                 # Parse the date string into a datetime object
#                 date_obj = datetime.strptime(date_str.split('(')[0].strip(), '%a, %d %b %Y %H:%M:%S %z')
#                 # Format as YYYY-MM-DD HH:MM:SS
#                 date = date_obj.strftime('%Y-%m-%d %H:%M:%S')
#             except:
#                 date = date_str
#             break
    
#     print(f"📅 Date check: {hours_reason}")
    
#     # Skip if not within office hours
#     if not within_hours:
#         print(f"⏭️  Skipping message (outside office hours)")
#         print("-" * 60)
#         filtered_count += 1
#         continue
    
#     office_hours_count += 1
#     print(f"✅ Message is within office hours - processing...")

#     # Extract body (plain or HTML)
#     def extract_body(payload):
#         body = ''
#         if 'parts' in payload:
#             for part in payload['parts']:
#                 if part['mimeType'] == 'text/plain':
#                     data = part['body'].get('data')
#                     if data:
#                         body += base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
#                 elif part['mimeType'] == 'text/html':
#                     data = part['body'].get('data')
#                     if data:
#                         html = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
#                         body += BeautifulSoup(html, 'html.parser').get_text()
#                 elif 'parts' in part:
#                     # Handle nested parts
#                     body += extract_body(part)
#         elif 'body' in payload and 'data' in payload['body']:
#             body = base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8', errors='ignore')
        
#         # Clean up the body text
#         body = body.replace('\r', ' ').replace('\n', ' ')
#         body = ' '.join(body.split())  # Remove extra whitespace
#         return body

#     try:
#         body = extract_body(msg_detail['payload'])
#         print("📄 Message body preview:")
#         print(body[:300] + '...' if len(body) > 300 else body)
#     except Exception as e:
#         print(f"❌ Failed to extract body: {e}")
#         body = ''

#     # Extract email from headers
#     sender_email = ''
#     for header in headers:
#         if header['name'].lower() == 'from':
#             # Extract email from "Name <email@domain.com>" format
#             match = re.search(r'<([^>]+)>', header['value'])
#             if match:
#                 sender_email = match.group(1)
#             else:
#                 # If no angle brackets, use the whole value if it looks like an email
#                 if '@' in header['value']:
#                     sender_email = header['value'].strip()
#             break

#     # Extract name from headers
#     sender_name = ''
#     for header in headers:
#         if header['name'].lower() == 'from':
#             # Extract name from "Name <email@domain.com>" format
#             match = re.match(r'^([^<]+)', header['value'])
#             if match:
#                 sender_name = match.group(1).strip().strip('"')
#             break

#     # Use improved clean_extract function
#     name = clean_extract(body, 'name') or sender_name
#     email = clean_extract(body, 'email') or sender_email
#     phone = clean_extract(body, 'phone')
    
#     # If no phone found, try fallback pattern
#     if not phone:
#         phone_match = re.search(r'(\+1\s*\([0-9]{3}\)\s*[0-9]{3}\s*[0-9]{4})', body)
#         if phone_match:
#             phone = phone_match.group(1)

#     # Label to status - use 'Qualified' as default like in your sample
#     status = 'Qualified'
#     for lbl, lid in label_map.items():
#         if lid in msg_labels and lid in custom_label_ids and lbl != 'Qualified':
#             status = lbl  # Only override if it's one of our other custom labels
#             break

#     # Print in the requested format
#     print(f"Customer Name: {name}")
#     print(f"Email: {email}")
#     print(f"Phone: {phone}")
#     print(f"Created At: {date}")
#     print(f"Status: {status}")
#     print("-" * 60)

#     data.append([name, email, phone, date, status])

# # Step 7: Export to Excel
# print(f"\n📊 Summary:")
# print(f"   • Total messages found: {len(messages)}")
# print(f"   • Messages outside office hours: {filtered_count}")
# print(f"   • Messages within office hours: {office_hours_count}")
# print(f"   • Messages to export: {len(data)}")

# print("\n📤 Exporting to gmail_leads.xlsx...")
# df = pd.DataFrame(data, columns=[
#     "Customer Name", "Email", "Phone", "Created At", "status"
# ])

# # Save to Excel - using openpyxl as it's more reliable for formatting
# try:
#     df.to_excel("gmail_leads.xlsx", index=False, engine='openpyxl')
#     print("✅ Export complete. File saved as gmail_leads.xlsx")
# except ImportError:
#     # Fallback if openpyxl not installed
#     df.to_excel("gmail_leads.xlsx", index=False)
#     print("✅ Export complete (using default engine). File saved as gmail_leads.xlsx")

# print("🚀 Gmail to Excel Export Script completed.")

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import base64
import re
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
import pytz

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

# Step 2.5: Office hours check function
def is_within_office_hours(date_str):
    """
    Check if the message timestamp is today between 7AM-4PM PST, Monday-Friday
    """
    try:
        # Set up PST timezone
        pst = pytz.timezone('US/Pacific')
        
        # Parse the email date
        if '(' in date_str:
            date_str = date_str.split('(')[0].strip()
        
        # Parse date and convert to PST
        date_obj = datetime.strptime(date_str, '%a, %d %b %Y %H:%M:%S %z')
        pst_date = date_obj.astimezone(pst)
        
        # Get current date in PST
        now_pst = datetime.now(pst)
        today_pst = now_pst.date()
        
        # Check if message is from today
        if pst_date.date() != today_pst:
            return False, f"Not today (message date: {pst_date.date()}, today: {today_pst})"
        
        # Check if it's Monday-Friday (0=Monday, 6=Sunday)
        if pst_date.weekday() > 4:  # 5=Saturday, 6=Sunday
            return False, f"Weekend (day: {pst_date.strftime('%A')})"
        
        # Check if it's between 7AM-4PM PST
        hour = pst_date.hour
        if hour < 7 or hour >= 16:  # 16 = 4PM (exclusive)
            return False, f"Outside office hours (time: {pst_date.strftime('%I:%M %p PST')})"
        
        return True, f"Within office hours ({pst_date.strftime('%A, %I:%M %p PST')})"
        
    except Exception as e:
        return False, f"Error parsing date: {e}"

print(f"\n⏰ Office Hours Filter: Monday-Friday, 7AM-4PM PST")
print(f"📅 Current PST time: {datetime.now(pytz.timezone('US/Pacific')).strftime('%A, %B %d, %Y %I:%M %p PST')}")

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

# Step 6: Filter by office hours and extract details
data = []
filtered_count = 0
office_hours_count = 0

print("\n📝 Parsing message details and applying office hours filter...\n")

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

    # Extract message date and check office hours
    date = ''
    date_str = ''
    within_hours = False
    hours_reason = ''
    
    for header in headers:
        if header['name'].lower() == 'date':
            date_str = header['value']
            within_hours, hours_reason = is_within_office_hours(date_str)
            
            try:
                # Parse the date string into a datetime object
                date_obj = datetime.strptime(date_str.split('(')[0].strip(), '%a, %d %b %Y %H:%M:%S %z')
                # Format as YYYY-MM-DD HH:MM:SS
                date = date_obj.strftime('%Y-%m-%d %H:%M:%S')
            except:
                date = date_str
            break
    
    print(f"📅 Date check: {hours_reason}")
    
    # Skip if not within office hours
    if not within_hours:
        print(f"⏭️  Skipping message (outside office hours)")
        print("-" * 60)
        filtered_count += 1
        continue
    
    office_hours_count += 1
    print(f"✅ Message is within office hours - processing...")

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

# Step 7: Export to Excel (only if there's data)
print(f"\n📊 Summary:")
print(f"   • Total messages found: {len(messages)}")
print(f"   • Messages outside office hours: {filtered_count}")
print(f"   • Messages within office hours: {office_hours_count}")
print(f"   • Messages to export: {len(data)}")

if len(data) == 0:
    print("\n❌ No data to export!")
    print("💡 Reasons this might happen:")
    print("   • No messages received today during office hours (7AM-4PM PST, Mon-Fri)")
    print("   • All messages were filtered out due to time/date constraints")
    print("   • No messages with the specified custom labels")
    print("\n🚀 Gmail to Excel Export Script completed (no file created).")
else:
    print(f"\n📤 Exporting {len(data)} records to gmail_leads.xlsx...")
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