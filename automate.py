


from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import base64
import re
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime, timedelta

print("\U0001F680 Starting Gmail to Excel Export Script...")

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
print("\U0001F511 Authenticating...")
flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
creds = flow.run_local_server(port=0)
service = build('gmail', 'v1', credentials=creds)
print("✅ Authentication successful.")

print("\U0001F4E5 Fetching Gmail labels...")
label_results = service.users().labels().list(userId='me').execute()
label_map = {label['name']: label['id'] for label in label_results['labels']}

custom_labels = ['SALE MADE','Quoted','Call Back','Follow up']
custom_label_ids = [label_map[label] for label in custom_labels if label in label_map]
print(f"\U0001F50E Using custom labels: {custom_labels}")

if not custom_label_ids:
    print("❌ No matching custom labels found.")
    exit()

yesterday = (datetime.now() - timedelta(1)).strftime('%Y/%m/%d')
print(f"\n\U0001F4C5 Fetching messages from today (after {yesterday})...")

all_messages = []
for label_id in custom_label_ids:
    result = service.users().messages().list(
        userId='me',
        labelIds=[label_id],
        q=f"after:{yesterday}",
        maxResults=500
    ).execute()
    messages = result.get('messages', [])
    all_messages.extend(messages)

unique_messages = {msg['id']: msg for msg in all_messages}
messages = list(unique_messages.values())
print(f"\n\U0001F4E6 Found {len(messages)} messages to process")

def clean_phone(phone):
    phone = re.sub(r'[^\d+]', '', phone)
    if phone.startswith('+1'):
        return phone
    elif phone.startswith('1') and len(phone) == 11:
        return f"+{phone}"
    elif len(phone) == 10:
        return f"+1{phone}"
    return phone

def clean_extract(text, field_name):
    patterns = {
        'name': [
            r'(?:Customer\s+Name|Full\s+Name|Name)\s*[:=]\s*([^\n\u2022<>@]+?)(?=\s*(?:[\u2022\n]|Email|Phone|$))',
            r'(?:From|Sender)\s*[:=]?\s*([^\n\u2022<>@]+?)(?=\s*(?:[\u2022\n]|Email|Phone|$))'
        ],
        'email': [
            r'(?:Email|E-Mail|Email\s+Address)\s*[:=]?\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
            r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
        ],
        'phone': [
            r'(?:Phone|Mobile|Contact|Phone\s+Number|Contact\s+Number|Tel|Telephone)\s*[:=]?\s*(\+?1?\s*\(?\d{3}\)?\s*\d{3}[-.\s]*\d{4})',
            r'(\+1\s*\(\d{3}\)\s*\d{3}\s*\d{4})'
        ]
    }

    for pattern in patterns.get(field_name.lower(), []):
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return re.sub(r'[<>"]', '', match.group(1).strip())
    return "''"

data = []
for msg in messages:
    try:
        msg_detail = service.users().messages().get(userId='me', id=msg['id'], format='full').execute()

        headers = {h['name'].lower(): h['value'] for h in msg_detail['payload']['headers']}

        date_str = headers.get('date', '')
        try:
            date_obj = datetime.strptime(date_str.split('(')[0].strip(), '%a, %d %b %Y %H:%M:%S %z')
            created_at = date_obj.strftime('%m/%d/%Y %H:%M:%S')
        except:
            created_at = date_str

        def get_body(payload):
            if 'parts' in payload:
                return ''.join(get_body(part) for part in payload['parts'])
            return base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8', errors='ignore') if 'data' in payload.get('body', {}) else ''

        body = get_body(msg_detail['payload'])
        clean_body = ' '.join(BeautifulSoup(body, 'html.parser').get_text().split())

        name = clean_extract(clean_body, 'name')
        email = clean_extract(clean_body, 'email')
        phone = clean_phone(clean_extract(clean_body, 'phone'))

        status = next(
            (label for label_id in msg_detail.get('labelIds', [])
             for label, lid in label_map.items()
             if lid == label_id and label in custom_labels),
            "Unknown")

        data.append([name, email, phone, created_at, status])

    except Exception as e:
        print(f"⚠️ Error processing message: {e}")

df = pd.DataFrame(data, columns=[
    'Customer Name',
    'Email',
    'Phone',
    'Created At',
    'Status'
])

print("\n\U0001F4BE Exporting to Excel...")
try:
    with pd.ExcelWriter('gmail_leads.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        worksheet.column_dimensions['A'].width = 25
        worksheet.column_dimensions['B'].width = 30
        worksheet.column_dimensions['C'].width = 20
        worksheet.column_dimensions['D'].width = 22
        worksheet.column_dimensions['E'].width = 15

    print("✅ Successfully exported to gmail_leads.xlsx")
    print("\nPreview of exported data:")
    print(df.head())
except Exception as e:
    print(f"❌ Export failed: {e}")
