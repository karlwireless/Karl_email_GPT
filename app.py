import os
from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
import time
from datetime import datetime, timedelta
import openai
import io
import openpyxl

load_dotenv()

# Set OpenAI API key and other essential variables
openai.api_key = os.getenv('OPENAI_API_KEY')
client_id = os.environ.get('O365_CLIENT_ID')
client_secret = os.environ.get('O365_CLIENT_SECRET')
assistant_id = os.environ.get('ASSISTANT_ID')
email_keyword = os.getenv('EMAIL_KEYWORD')  # The specific keyword to look for in email subjects
report_recipient = os.getenv('REPORT_RECIPIENT')
if not report_recipient:
    raise ValueError("The REPORT_RECIPIENT environment variable is missing.")

# Ensure essential environment variables are set
if not all([openai.api_key, client_id, client_secret, assistant_id, email_keyword]):
    raise ValueError("One or more required environment variables are missing.")

# Create a client instance for OpenAI
client = openai.Client()

def getAssistantResponse(prompt, file_ids = None):
    print("Generating response...")
    thread = client.beta.threads.create()
    print(thread)
    thread_id = thread.id

    if (file_ids):
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

# Setup token storage and initialize O365 account
token_backend = FileSystemTokenBackend(token_path='.', token_filename='o365_token.txt')
account = Account((client_id, client_secret), token_backend=token_backend)

# Authenticate if not already authenticated
if not account.is_authenticated:
    print("Authentication required. Redirecting to login page...")
    account.authenticate(scopes=['https://graph.microsoft.com/Mail.Read'])

print("Authenticated successfully. Proceeding with main logic...")

mailbox = account.mailbox()
inbox = mailbox.inbox_folder()
last_checked_time = datetime.now()

def format_message_history(messages):
    """Formats the messages history into a string for the assistant."""
    formatted_history = ""
    for msg in messages:
        formatted_history += f"From: {msg.sender.address}\nSubject: {msg.subject}\nMessage: {msg.body_preview}\n\n"
    return formatted_history


def check_and_handle_attachments(message):
    print("Checking for attachments...")
    file_ids = []
    # Explicitly fetch the attachment details
    message.attachments.download_attachments()
    
    attachments = [attachment for attachment in message.attachments if not attachment.is_inline]
    print(f"Number of Attachments: {len(attachments)}")

    for attachment in attachments:
        print(f"Found attachment: {attachment.name}")
        if attachment.is_inline:
            continue  # Skip inline attachments

        print(attachment.name)
        # Define the local file path where the attachment will be saved
        local_file_path = os.path.join(os.getcwd(), attachment.name)
        
        success = attachment.save()
        if success:
            print(f"Attachment {attachment.name} saved successfully.")
            try:
                with open(local_file_path, "rb") as file:
                    file_id = send_file_to_openai(file)
                    if file_id:
                        file_ids.append(file_id)
                    else:
                        print(f"Failed to process attachment: {attachment.name}")
            finally:
                # Clean up the file after processing
                os.remove(local_file_path)
        else:
            print(f"Failed to save attachment: {attachment.name}")
    
    return file_ids

def create_excel_report(received_emails, sent_emails, filename="Email_Report.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Type", "From/To", "Subject", "Received/Sent", "Body Preview"])

    # Add received emails to the report
    for email in received_emails:
        received = email.received.replace(tzinfo=None)  # Remove timezone info
        ws.append(["Received", email.sender.address, email.subject, received, email.body_preview])

    # Add sent emails to the report
    for email in sent_emails:
        sent = email.sent.replace(tzinfo=None)  # Remove timezone info
        recipient = ", ".join([recipient.address for recipient in email.to]) if email.to else "N/A"
        ws.append(["Sent", recipient, email.subject, sent, email.body_preview])

    wb.save(filename)
    return filename




# Function to send an email with an attachment
def send_email_with_attachment(account, recipient, subject, body, attachment_path):
    message = account.new_message()
    message.to.add(recipient)
    message.subject = subject
    message.body = body
    message.attachments.add(attachment_path)
    message.send()
    os.remove(attachment_path)  # Delete the report file after sending the email

report_interval = timedelta(days=7) 
last_report_time = datetime.now() - report_interval

while True:
    print("start...")
    current_time = datetime.now()
    if current_time - last_report_time >= report_interval:
        print("making report....")
        # Fetch emails from the last 7 days
        start_time = current_time - report_interval
        received_query = inbox.new_query().on_attribute('receivedDateTime').greater_equal(start_time)
        recent_received_emails = inbox.get_messages(limit=100, query=received_query)

        # Fetch emails from the last 7 days
        sent_folder = mailbox.sent_folder()
        sent_query = sent_folder.new_query().on_attribute('sentDateTime').greater_equal(start_time)
        recent_sent_emails = sent_folder.get_messages(limit=100, query=sent_query)

        # Generate and send the report
        report_filename = create_excel_report(recent_received_emails, recent_sent_emails)
        send_email_with_attachment(account, report_recipient, "Email Report", "Here is the weekly email report.", report_filename)

        # Update the last report time
        last_report_time = current_time
    
    print("Checking for new emails...")
    start_time = last_checked_time # Look back a bit more to catch ongoing conversations and new emails
    last_checked_time = datetime.now()

    # Fetch only messages with the specific keyword in the subject
    query = mailbox.new_query().on_attribute('receivedDateTime').greater_equal(start_time).chain('and').on_attribute('subject').contains(email_keyword)
    new_messages = inbox.get_messages(limit=25, query=query)

    for message in new_messages:

        # Fetch the conversation history
        conversation_id = message.conversation_id
        conversation_messages = mailbox.get_messages(query=mailbox.new_query().on_attribute('conversationId').equals(conversation_id), limit=10)
        conversation_messages_list = list(conversation_messages)
        formatted_history = format_message_history(conversation_messages_list)

        # Check for and process attachments
        file_ids = check_and_handle_attachments(message)

        if file_ids:
            # Emails with attachments
            prompt = f"Subject: {message.subject}\nMessage: {message.body_preview}\n\nEmail thread history: {formatted_history}\n\nPlease draft a reply considering the attachments."
            assistant_reply = getAssistantResponse(prompt, file_ids)
        else:
            # This part remains the same as before, handling emails without attachments
            if len(conversation_messages_list) > 1:
                prompt = f"Continue the conversation based on this email thread:\n\n{formatted_history}\n\nPlease draft a reply."
            else:
                prompt = f"Subject: {message.subject}\nMessage: {message.body_preview}\n\nPlease draft a reply."

            assistant_reply = getAssistantResponse(prompt)
        
        print(prompt)

        print(assistant_reply)
        
        # Reply within the same email thread
        reply = message.reply()
        reply.body = assistant_reply.replace("\n", "<br>")
        reply.body_type = 'HTML'
        reply.send()
        print("Replied with assistant's response.")

    time.sleep(60)  # Sleep before checking again
