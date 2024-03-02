import os
import requests
from msal import ConfidentialClientApplication
import time
from datetime import datetime, timedelta
import openai
import openpyxl
from dotenv import load_dotenv
import base64
from bs4 import BeautifulSoup

load_dotenv()

# Set OpenAI API key and other essential variables
openai.api_key = os.getenv('OPENAI_API_KEY')
TENANT_ID = os.environ.get('TENANT_ID')
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPE = ['https://graph.microsoft.com/.default']
USER_EMAIL = 'karl@karl.guru'
ENDPOINT = f'https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/mailFolders/inbox/messages'
assistant_id = os.environ.get('ASSISTANT_ID')
email_keyword = os.getenv('EMAIL_KEYWORD')  # The specific keyword to look for in email subjects
report_recipient = os.getenv('REPORT_RECIPIENT')
if not report_recipient:
    raise ValueError("The REPORT_RECIPIENT environment variable is missing.")

# Ensure essential environment variables are set
if not all([openai.api_key, CLIENT_ID, CLIENT_SECRET, TENANT_ID, USER_EMAIL, assistant_id, email_keyword]):
    raise ValueError("One or more required environment variables are missing.")

# Create a client instance for OpenAI
client = openai.Client()

# Global variable to store the access token and expiration time
access_token_info = {'token': None, 'expires_at': datetime.now()}

def get_access_token():
    global access_token_info
    app = ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    # Only acquire a new token if the current one is expired or about to expire
    if not access_token_info['token'] or datetime.now() >= access_token_info['expires_at']:
        token_response = app.acquire_token_for_client(scopes=SCOPE)
        if 'access_token' in token_response and 'expires_in' in token_response:
            access_token_info['token'] = token_response['access_token']
            # Set the expiration time to be a bit before the token actually expires to ensure we have a valid token
            expires_in = token_response['expires_in']
            access_token_info['expires_at'] = datetime.now() + timedelta(seconds=expires_in - 300)
    return access_token_info

def getAssistantResponse(prompt, file_ids=None):
    print("Generating response...")
    thread = client.beta.threads.create()
    print(thread)
    thread_id = thread.id

    if file_ids:
        message = client.beta.threads.messages.create(
            thread_id=thread_id,
            role="user",
            content=prompt,
            file_ids=file_ids
        )
    else:
        message = client.beta.threads.messages.create(
            thread_id=thread_id,
            role="user",
            content=prompt,
        )

    print(message)
    run = client.beta.threads.runs.create(
        thread_id=thread_id,
        assistant_id=assistant_id,
    )

    while True:  # Continually check for completion
        print("run")
        run = client.beta.threads.runs.retrieve(
            thread_id=thread_id,
            run_id=run.id
        )
        print(run.status)
        if run.status == "completed":
            break  # Exit the loop once the run is completed
        elif run.status == "failed":
            return "Failed to process your request."
        time.sleep(0.5)

    messages = client.beta.threads.messages.list(
        thread_id=thread_id
    )

    # Assuming the last message in the thread is the assistant's response
    last_message = messages.data[0].content[0].text.value

    return last_message

def send_file_to_openai(file_object):
    try:
        # Assuming the OpenAI API client can handle file-like objects directly
        response = client.files.create(
            file=file_object,
            purpose="assistants"
        )
        print(response)
        return response.id  # Assuming the response object has an 'id' attribute
    except Exception as e:
        print(f"Error: {str(e)}")
        return None

def format_message_history(messages):
    """Formats the messages history into a string for the assistant."""
    formatted_history = ""
    for msg in messages:
        formatted_history += f"From: {msg['from']['emailAddress']['name']}\nSubject: {msg['subject']}\nMessage: {msg['bodyPreview']}\n\n"
    return formatted_history

def check_and_handle_attachments(message_id, access_token):
    print("Checking for attachments...")
    file_ids = []
    attachments_endpoint = f'https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/messages/{message_id}/attachments'
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    response = requests.get(attachments_endpoint, headers=headers)
    attachments = response.json().get('value', [])
    print(f"Number of Attachments: {len(attachments)}")

    for attachment in attachments:
        print(f"Found attachment: {attachment['name']}")
        if attachment['@odata.type'] != '#microsoft.graph.fileAttachment':
            continue  # Skip non-file attachments

        # Download the attachment content
        attachment_content = requests.get(attachments_endpoint + f"/{attachment['id']}/$value", headers=headers).content

        # Save the attachment to a temporary file
        local_file_path = os.path.join(os.getcwd(), attachment['name'])
        with open(local_file_path, 'wb') as file:
            file.write(attachment_content)

        try:
            with open(local_file_path, "rb") as file:
                file_id = send_file_to_openai(file)
                if file_id:
                    file_ids.append(file_id)
                else:
                    print(f"Failed to process attachment: {attachment['name']}")
        finally:
            # Clean up the file after processing
            os.remove(local_file_path)

    return file_ids

def create_excel_report(received_emails, access_token, filename="Email_Report.xlsx"):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.append(["Date", "Question Sent", "Reply to Question", "Email Sent From"])

  for received_email in received_emails:
      date = datetime.strptime(received_email['receivedDateTime'], '%Y-%m-%dT%H:%M:%SZ').strftime('%Y-%m-%d %H:%M')
      question_sent = BeautifulSoup(received_email['body']['content'], 'html.parser').get_text()
      email_sent_from = received_email['from']['emailAddress']['address']

      # Fetch the reply based on conversationId
      conversation_id = received_email['conversationId']
      sent_folder_endpoint = f'https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/mailFolders/sentItems/messages'
      sent_query = f"{sent_folder_endpoint}?$filter=conversationId eq '{conversation_id}'"
      response = requests.get(sent_query, headers={'Authorization': f'Bearer {access_token}'})
      sent_emails = response.json().get('value', [])

      # Assuming the first email in the response is the reply to the received email
      reply_to_question = BeautifulSoup(sent_emails[0]['body']['content'], 'html.parser').get_text() if sent_emails else "No reply found"

      ws.append([date, question_sent, reply_to_question, email_sent_from])

  wb.save(filename)
  return filename


# Function to send an email with an attachment
def send_email_with_attachment(access_token, recipient, subject, body, attachment_path):
    message_endpoint = f'https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/sendMail'
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    with open(attachment_path, "rb") as file:
        attachment_content = file.read()
    attachment = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": os.path.basename(attachment_path),
        "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "contentBytes": base64.b64encode(attachment_content).decode('utf-8')
    }
    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient
                    }
                }
            ],
            "attachments": [attachment]
        }
    }
    response = requests.post(message_endpoint, headers=headers, json=payload)
    if response.status_code == 202:
        print("Email sent successfully.")
    else:
        print(f"Failed to send email: {response.text}")
    os.remove(attachment_path)  # Delete the report file after sending the email

def main():
    token_info = get_access_token()
    access_token = token_info['token']
    last_checked_time = datetime.now()
    report_interval = timedelta(days=7)
    last_report_time = datetime.now() - report_interval

    while True:
        if datetime.now() >= token_info['expires_at']:
            token_info = get_access_token()
            access_token = token_info['token']
        print(access_token)
        current_time = datetime.now()
        if current_time - last_report_time >= report_interval:
            print("Generating report...")
            start_time = current_time - report_interval
            received_query = f"{ENDPOINT}?$filter=receivedDateTime ge {start_time.strftime('%Y-%m-%dT%H:%M:%SZ')}&$top=999"

            response = requests.get(received_query, headers={'Authorization': f'Bearer {access_token}'})
            recent_received_emails = response.json().get('value', [])

            sent_folder_endpoint = f'https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/mailFolders/sentItems/messages'
            sent_query = f"{sent_folder_endpoint}?$filter=sentDateTime ge {start_time.strftime('%Y-%m-%dT%H:%M:%SZ')}"
            response = requests.get(sent_query, headers={'Authorization': f'Bearer {access_token}'})
            recent_sent_emails = response.json().get('value', [])

            report_filename = create_excel_report(recent_received_emails, access_token)

            send_email_with_attachment(access_token, report_recipient, "Email Report", "Here is the weekly email report.", report_filename)

            last_report_time = current_time

        print("Checking for new emails...")
        start_time = last_checked_time  # Look back a bit more to catch ongoing conversations and new emails
        last_checked_time = datetime.now()
        new_messages_query = f"{ENDPOINT}?$filter=receivedDateTime ge {start_time.strftime('%Y-%m-%dT%H:%M:%SZ')} and contains(subject, '{email_keyword}')"

        print(new_messages_query)
        response = requests.get(new_messages_query, headers={'Authorization': f'Bearer {access_token}'})
        new_messages = response.json().get('value', [])
        print(new_messages)

        for message in new_messages:
            message_id = message['id']
            conversation_id = message['conversationId']
            conversation_messages_query = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/messages?$filter=conversationId eq '{conversation_id}'&$top=10"
            response = requests.get(conversation_messages_query, headers={'Authorization': f'Bearer {access_token}'})
            conversation_messages_list = response.json().get('value', [])
            formatted_history = format_message_history(conversation_messages_list)

            file_ids = check_and_handle_attachments(message_id, access_token)

            if file_ids:
                prompt = f"Subject: {message['subject']}\nMessage: {message['bodyPreview']}\n\nEmail thread history: {formatted_history}\n\nPlease draft a reply considering the attachments."
                assistant_reply = getAssistantResponse(prompt, file_ids)
            else:
                if len(conversation_messages_list) > 1:
                    prompt = f"Continue the conversation based on this email thread:\n\n{formatted_history}\n\nPlease draft a reply."
                else:
                    prompt = f"Subject: {message['subject']}\nMessage: {message['bodyPreview']}\n\nPlease draft a reply."

                assistant_reply = getAssistantResponse(prompt)

            print(prompt)
            print(assistant_reply)

            # Reply within the same email thread
            reply_endpoint = f'https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/messages/{message_id}/reply'
            reply_payload = {
                "message": {
                    "body": {
                        "contentType": "HTML",
                        "content": assistant_reply.replace("\n", "<br>")
                    }
                }
            }
            response = requests.post(reply_endpoint, headers={'Authorization': f'Bearer {access_token}'}, json=reply_payload)
            if response.status_code == 202:
                print("Replied with assistant's response.")
            else:
                print(f"Failed to reply: {response.text}")

        time.sleep(60)  # Sleep before checking again

if __name__ == '__main__':
    main()
