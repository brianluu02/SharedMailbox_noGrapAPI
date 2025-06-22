import ssl
import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from O365 import Account
import io
import os

# ⚙️ 1. Disable SSL warnings (only in trusted internal networks)
ssl._create_default_https_context = ssl._create_unverified_context
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
requests.Session().verify = False

# 📤 2. Function to send email with optional attachment
def send_email_with_attachment(client_id, client_secret, tenant_id, shared_mailbox, to_email, subject, body, attachment=None):
    credentials = (client_id, client_secret)

    # Create an authenticated account
    account = Account(credentials, auth_flow_type='credentials', tenant_id=tenant_id, verify_ssl=False)

    if account.authenticate():
        try:
            mailbox = account.mailbox(shared_mailbox)
            message = mailbox.new_message()
            message.to.add(to_email)
            message.subject = subject
            message.body = body

            # Attach file if provided
            if attachment and os.path.exists(attachment):
                with open(attachment, 'rb') as f:
                    file_name = os.path.basename(attachment)
                    message.attach(io.BytesIO(f.read()), filename=file_name)

            message.send()
            print("✅ Email sent successfully!")
        except Exception as e:
            print(f"❌ Error sending email: {e}")
    else:
        print("❌ Authentication failed. Please check your credentials and permissions.")

# 🧪 3. Replace these with your actual credentials and test params
if __name__ == "__main__":
    send_email_with_attachment(
        client_id="j",
        client_secret="l",
        tenant_id="b",
        shared_mailbox="?", # send email to 
        to_email=",",
        subject="Test Email ",
        body="This is hôm nay qua tuyệt vời bạn ôi.",
        # attachment="C:/path/to/your/file.pdf"  # Uncomment to send attachment
    )
