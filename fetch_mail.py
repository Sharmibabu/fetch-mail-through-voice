import email
from email.base64mime import body_decode, body_encode
from email.message import EmailMessage
import os
import time
from bs4 import BeautifulSoup
from gtts import gTTS
import speech_recognition as sr
import imaplib
import pyglet
import pyttsx3
from win32com.client import Dispatch

# Create a recognizer object and microphone instance
r = sr.Recognizer()
mic = sr.Microphone()
engine=pyttsx3.init()
s = Dispatch("SAPI.SpVoice")

def speak(text): #defining the speaking function 
	s.Speak(text)

def talk(text):
    engine.say(text)
    engine.runAndWait()


# Get the email address of the person whose email you want to fetch through voice input
with mic as source:
    
    r.adjust_for_ambient_noise(source)
    print("Please say the email address you want to fetch without saying @ gmail.com.")
    speak("Please say the email address you want to fetch without saying @ gmail.com.")
    audio = r.listen(source)

email_address = r.recognize_google(audio)
rmaildomain = "@gmail.com"
text_without_space = email_address.replace(" ", "")
new=text_without_space + rmaildomain 



# Connect to the IMAP server using SSL
mail = imaplib.IMAP4_SSL('imap.gmail.com',993) 

# Login to the email account
mail.login('your_mail-id', 'your-password')

# Select the inbox folder
mail.select('inbox')

# Search for emails sent from the specified email address
result, data = mail.search(None, f'FROM "{new.lower()}"')

if not data[0]:
    print(f"No emails found from {new.lower()}.")
else:
    # Get the email IDs of the searched emails
    email_ids = data[0].split()

    # Fetch the first email in the search results
    result, data = mail.fetch(email_ids[0], "(RFC822)")

    # Parse the email message
    message = email.message_from_bytes(data[0][1])

    # Print the subject and body of the email
    print("Subject: ", message["Subject"])
    print("Body: ", message.get_payload())
    tts = gTTS(text="The subject of the mail is :"+ message["Subject"], lang='en')
    typ, msg_data = mail.search(None, f'(SUBJECT "{message["Subject"]}")')
    if msg_data[0]:
        email_id = msg_data[0].split()[-1]
        typ, msg_data = mail.fetch(email_id, '(RFC822)')
        raw_email = msg_data[0][1]
        email_message = email.message_from_bytes(raw_email)
        body = ''
        if email_message.is_multipart():
            for part in email_message.walk():
                if part.get_content_type() == 'text/plain':
                    body += part.get_payload()
        else:
            body = email_message.get_payload()
        # Use pyttsx3 to read the email body out loud
        engine = pyttsx3.init()
        engine.say(body)
        engine.runAndWait()





    
    