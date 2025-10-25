"""
assistant.py
A simple voice-enabled virtual assistant:
- speech recognition (microphone)
- text-to-speech (pyttsx3)
- open apps / websites
- get time/date
- weather via OpenWeatherMap
- send email via SMTP
- wikipedia search
Configure secrets in .env (EMAIL_USER, EMAIL_PASS, OPENWEATHER_APIKEY)
"""

import os
import webbrowser
import subprocess
import smtplib
import sys
import json
from datetime import datetime
from time import sleep
from email.message import EmailMessage

import requests
import wikipedia
import speech_recognition as sr
import pyttsx3
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
OPENWEATHER_APIKEY = os.getenv("OPENWEATHER_APIKEY")

# --- Text-to-speech setup ---
engine = pyttsx3.init()
engine.setProperty('rate', 165)  # speak rate (words/min)
voices = engine.getProperty('voices')
if voices:
    # use first voice available; user can change index if desired
    engine.setProperty('voice', voices[0].id)

def speak(text: str):
    """Speak text and also print to console."""
    print("Assistant:", text)
    engine.say(text)
    engine.runAndWait()

# --- Speech recognition ---
recognizer = sr.Recognizer()
MIC_INDEX = None  # Use default microphone. Set to integer to pick a different device.

def listen(timeout=5, phrase_time_limit=8):
    """Listen to microphone and return recognized text (or None)."""
    with sr.Microphone(device_index=MIC_INDEX) as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.6)
        print("Listening...")
        try:
            audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
        except sr.WaitTimeoutError:
            print("No speech detected (timeout).")
            return None
    try:
        text = recognizer.recognize_google(audio)
        print("You:", text)
        return text.lower()
    except sr.UnknownValueError:
        print("Couldn't understand audio.")
        return None
    except sr.RequestError as e:
        print("Speech recognition service error:", e)
        return None

# --- Helpers ---
def tell_time():
    now = datetime.now()
    t = now.strftime("%I:%M %p")
    speak(f"The time is {t}")

def tell_date():
    today = datetime.today()
    speak("Today is " + today.strftime("%A, %B %d, %Y"))

def search_wikipedia(query):
    try:
        speak(f"Searching Wikipedia for {query}")
        summary = wikipedia.summary(query, sentences=2, auto_suggest=True)
        speak(summary)
    except Exception as e:
        speak("Sorry, I couldn't find that on Wikipedia.")

def open_website(url):
    if not url.startswith("http"):
        url = "https://" + url
    speak(f"Opening {url}")
    webbrowser.open(url)

def open_app(app_key):
    """Map friendly names to commands. Modify paths per your OS."""
    apps = {
        "notepad": {"win": r"notepad.exe", "mac": None, "linux": "gedit"},
        "calculator": {"win": "calc.exe", "mac": "open -a Calculator", "linux": "gnome-calculator"},
        # add your own paths here
        "code": {"win": r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe",
                 "mac": "open -a 'Visual Studio Code'", "linux": "code"}
    }
    platform = sys.platform
    if app_key not in apps:
        speak(f"I don't have a mapping for {app_key}. You can add it to the apps dict.")
        return
    cmd = None
    if platform.startswith("win"):
        cmd = apps[app_key].get("win")
        if cmd:
            speak(f"Opening {app_key}")
            subprocess.Popen(cmd)  # Windows, can be exe path
            return
    elif platform.startswith("darwin"):
        cmd = apps[app_key].get("mac")
    else:
        cmd = apps[app_key].get("linux")
    if cmd:
        speak(f"Opening {app_key}")
        try:
            subprocess.Popen(cmd.split())  # naive split for mac/linux; adjust if needed
        except Exception as e:
            # fallback to open
            try:
                subprocess.call(cmd, shell=True)
            except Exception as e2:
                speak("Could not open the application: " + str(e2))
    else:
        speak("No command configured for your OS; please update the apps mapping in the script.")

def get_weather(city):
    if not OPENWEATHER_APIKEY:
        speak("Weather API key is not configured. Please set OPENWEATHER_APIKEY in your .env.")
        return
    url = "https://api.openweathermap.org/data/2.5/weather"
    params = {"q": city, "appid": OPENWEATHER_APIKEY, "units": "metric"}
    try:
        r = requests.get(url, params=params, timeout=8)
        r.raise_for_status()
        data = r.json()
        desc = data['weather'][0]['description']
        temp = data['main']['temp']
        feels = data['main']['feels_like']
        speak(f"Weather in {city}: {desc}. Temperature {temp} °C, feels like {feels} °C.")
    except Exception as e:
        speak("Failed to get weather: " + str(e))

def send_email(to_addr, subject, body):
    """Send email using SMTP. EMAIL_USER and EMAIL_PASS must be set in environment."""
    if not EMAIL_USER or not EMAIL_PASS:
        speak("Email credentials not configured. Set EMAIL_USER and EMAIL_PASS in .env.")
        return False
    try:
        msg = EmailMessage()
        msg['From'] = EMAIL_USER
        msg['To'] = to_addr
        msg['Subject'] = subject
        msg.set_content(body)
        # example for Gmail SMTP; change for other providers
        smtp_host = 'smtp.gmail.com'
        smtp_port = 587
        server = smtplib.SMTP(smtp_host, smtp_port)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.send_message(msg)
        server.quit()
        speak("Email sent successfully.")
        return True
    except Exception as e:
        speak("Failed to send email. " + str(e))
        return False

# --- Main loop and command parsing ---
def parse_and_execute(text):
    if text is None:
        return

    # greeting
    if any(kw in text for kw in ["hello", "hi", "hey"]):
        speak("Hello! How can I help you?")

    # time/date
    elif "time" in text:
        tell_time()
    elif "date" in text or "day" in text:
        tell_date()

    # open website
    elif text.startswith("open "):
        # "open youtube" -> open youtube.com
        target = text.replace("open ", "").strip()
        # common quick mapping
        quick = {"youtube":"youtube.com", "google":"google.com", "gmail":"mail.google.com"}
        if target in quick:
            open_website(quick[target])
        else:
            # if target looks like a domain or 'website' assume domain
            open_website(target)

    # open application
    elif "launch " in text or "open app " in text or "open application" in text:
        # try to extract app name
        for prefix in ["launch ", "open app ", "open application "]:
            if prefix in text:
                appname = text.split(prefix,1)[1].strip().split()[0]
                open_app(appname)
                return

    # wikipedia
    elif text.startswith("who is ") or text.startswith("what is ") or text.startswith("tell me about "):
        # extract topic
        topic = text.replace("who is ", "").replace("what is ", "").replace("tell me about ", "").strip()
        search_wikipedia(topic)

    # weather
    elif "weather" in text:
        # "weather in london" or "what's the weather in paris"
        words = text.split()
        if "in" in words:
            idx = words.index("in")
            city = " ".join(words[idx+1:])
            if city:
                get_weather(city)
                return
        # if no city, use default or ask user (we avoid asking per guidelines)
        speak("Please include the city after 'in', e.g. 'weather in Delhi'.")

    # send email
    elif "send email" in text or "send an email" in text:
        speak("Who is the recipient? Please say the email address.")
        to_addr = listen(timeout=8, phrase_time_limit=6)
        if not to_addr:
            speak("Recipient not provided. Cancelling.")
            return
        speak("What is the subject?")
        subject = listen(timeout=8, phrase_time_limit=8) or "No subject"
        speak("Tell me the message.")
        body = listen(timeout=12, phrase_time_limit=20) or ""
        # confirmation (simple)
        speak(f"Sending email to {to_addr} with subject {subject}. Confirm by saying yes.")
        conf = listen(timeout=5, phrase_time_limit=3)
        if conf and "yes" in conf:
            send_email(to_addr, subject, body)
        else:
            speak("Email cancelled.")

    # stop/quit
    elif any(kw in text for kw in ["quit", "exit", "shutdown", "stop assistant", "goodbye"]):
        speak("Goodbye!")
        sys.exit(0)

    else:
        # fallback: use web search
        speak("I didn't catch a command. I can search the web or Wikipedia. Searching the web for your phrase.")
        # open a search in default browser
        query = text.replace(" ", "+")
        webbrowser.open(f"https://www.google.com/search?q={query}")

def main_loop():
    speak("Assistant online. Say 'hello' to start, or say a command.")
    while True:
        text = listen()
        parse_and_execute(text)
        # small idle delay
        sleep(0.3)

if __name__ == "__main__":
    try:
        main_loop()
    except KeyboardInterrupt:
        speak("Shutting down. Bye!")
