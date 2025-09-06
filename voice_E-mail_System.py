import speech_recognition as sr
from gtts import gTTS
from playsound import playsound
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import imaplib
import email

# --- Function Definitions ---

def listen():
    """
    Listens for audio input from the microphone and converts it to text.
    Uses Google Web Speech API for recognition.
    """
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Say something!")
        # Adjust for ambient noise for better recognition
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

    try:
        # Use Google Web Speech API for speech-to-text conversion
        text = r.recognize_google(audio)
        print(f"You said: {text}")
        return text.lower() # Return text in lowercase for easier command matching
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
        return ""
    except sr.RequestError as e:
        print(f"Could not request results from Google Speech Recognition service; {e}")
        return ""
    except Exception as e:
        print(f"An unexpected error occurred during listening: {e}")
        return ""

def speak(text):
    """
    Converts text to speech and plays the audio.
    Uses Google Text-to-Speech (gTTS) for audio generation.
    """
    try:
        tts = gTTS(text=text, lang='en')
        filename = "response.mp3"
        # Save the generated audio to a temporary file
        tts.save(filename)
        # Play the audio file
        playsound(filename)
        # Clean up the temporary audio file
        os.remove(filename)
    except Exception as e:
        print(f"Error during text-to-speech: {e}")

def send_email(sender_email, sender_password, receiver_email, subject, body):
    """
    Sends an email using SMTP.
    Note: For Gmail, you might need to generate an 'App Password'
    and enable 'Less secure app access' (though this option is being phased out).
    Using environment variables for credentials is highly recommended for security.
    """
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        # Connect to Gmail's SMTP server on port 587 with TLS encryption
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls() # Enable TLS encryption
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver_email, text)
        server.quit()
        print("Email sent successfully!")
        return True
    except smtplib.SMTPAuthenticationError:
        print("SMTP Authentication Error: Check your email and app password.")
        print("For Gmail, ensure you've generated an App Password.")
        return False
    except Exception as e:
        print(f"Error sending email: {e}")
        return False

def read_inbox(email_address, app_password, num_emails=2):
    """
    Reads a specified number of the latest emails from the inbox using IMAP.
    Note: Similar to sending, you might need an App Password for Gmail.
    """
    try:
        # Connect to Gmail's IMAP server with SSL encryption
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(email_address, app_password)
        mail.select('inbox') # Select the inbox folder

        # Search for all emails
        status, email_ids = mail.search(None, 'ALL')
        email_id_list = email_ids[0].split()
        # Get the IDs of the latest N emails
        latest_emails = email_id_list[-num_emails:]

        emails_content = []
        for eid in latest_emails:
            # Fetch the full email content (RFC822)
            status, msg_data = mail.fetch(eid, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    # Parse the raw email content
                    msg = email.message_from_bytes(response_part[1])
                    subject = msg['subject']
                    from_address = msg['from']
                    body = ""
                    # Handle multipart emails to extract plain text body
                    if msg.is_multipart():
                        for part in msg.walk():
                            ctype = part.get_content_type()
                            cdisp = str(part.get('Content-Disposition'))
                            # Only consider plain text parts that are not attachments
                            if ctype == 'text/plain' and 'attachment' not in cdisp:
                                body = part.get_payload(decode=True).decode(errors='ignore')
                                break
                    else:
                        # For single-part emails
                        body = msg.get_payload(decode=True).decode(errors='ignore')

                    emails_content.append({
                        "from": from_address,
                        "subject": subject,
                        "body": body
                    })
        mail.logout()
        return emails_content

    except imaplib.IMAP4.error as e:
        print(f"IMAP Error: Check your email and app password. Details: {e}")
        return []
    except Exception as e:
        print(f"Error reading emails: {e}")
        return []

# --- Main Logic ---

def process_command(command):
    """
    Processes the user's voice command to perform email actions.
    """
    if "send email" in command:
        speak("Who is the recipient?")
        recipient = "reciver email adrress@gmail.com"
        if not recipient:
            speak("Could not get recipient. Please try again.")
            return

        speak("What is the subject?")
        subject = listen()
        if not subject:
            speak("Could not get subject. Please try again.")
            return

        speak("What is the message body?")
        body = listen()
        if not body:
            speak("Could not get message body. Please try again.")
            return

        # Confirm with the user before sending
        speak(f"You want to send an email to {recipient} with subject {subject} and body {body}. Confirm with yes or no.")
        confirmation = listen()
        if "yes" in confirmation:
            # IMPORTANT: Replace with your actual email and a generated App Password for security.
            # DO NOT hardcode your main email password here.
            sender_email = "your email id @gmail.com"
            app_password = "your app pass" # <<< REPLACE THIS WITH YOUR GENERATED APP PASSWORD
            if send_email(sender_email, app_password, recipient, subject, body):
                speak("Email sent successfully!")
            else:
                speak("Failed to send email. Please check your credentials and internet connection.")
        else:
            speak("Email sending cancelled.")

    elif "read emails" in command or "read inbox" in command:
        # IMPORTANT: Replace with your actual email and a generated App Password for security.
        # DO NOT hardcode your main email password here.
        email_address = "your email id @gmail.com"
        app_password = "your email app pass" # <<< REPLACE THIS WITH YOUR GENERATED APP PASSWORD
        inbox = read_inbox(email_address, app_password, num_emails=2) # Read the last 2 emails
        if inbox:
            speak(f"You have {len(inbox)} recent emails.")
            for i, msg in enumerate(inbox):
                speak(f"Email number {i+1}.")
                speak(f"From: {msg['from']}.")
                speak(f"Subject: {msg['subject']}.")
                # Read only a part of the body for brevity
                speak(f"Body starts with: {msg['body'][:150]}...")
        else:
            speak("Could not retrieve emails or no new emails found.")

    elif "exit" in command or "quit" in command:
        speak("Goodbye!")
        return "exit"
    else:
        speak("I didn't understand that. Please try again.")

def main_loop():
    """
    The main loop for the voice email system.
    """
    speak("Welcome to your voice email system. How can I help you?")
    while True:
        command = listen()
        if command: # Only process if a command was recognized
            if process_command(command) == "exit":
                break

if __name__ == "__main__":
    # Ensure you have the necessary libraries installed:
    # pip install SpeechRecognition gTTS playsound
    # (playsound might require an additional audio player on some systems)

    # Before running, make sure to:
    # 1. Replace "your_email@gmail.com" with your actual email.
    # 2. IMPORTANT: Generate an "App Password" for your Gmail account and
    #    replace "your_app_password" with it. Do NOT use your regular Gmail password.
    #    Steps to generate an App Password:
    #    - Go to your Google Account (myaccount.google.com).
    #    - Navigate to Security -> How you sign in to Google -> App passwords.
    #    - Follow the instructions to generate a new app password.
    #    - You might need to enable 2-Step Verification first.

    main_loop()
