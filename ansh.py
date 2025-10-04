"""
A.N.S.H. - Adaptive Natural Speech Helper
I mean, what could an 8th grader even do with an old laptop during the 2020 lockdown? Ah well… hopefully I’ll come back and see this again someday. I’m putting it on GitHub since I’ve stopped coding, so there’s no point keeping it on my laptop anymore. Anyway, I’m 17 now.
"""

import os
import sys
import json
import time
import random
import datetime
import logging
import webbrowser
import pyttsx3
import speech_recognition as sr
import wikipedia
import pyautogui
import psutil
from requests import get

# Optional imports (guarded)
try:
    import PyPDF2
except Exception:
    PyPDF2 = None
try:
    import pywhatkit as kit
except Exception:
    kit = None

# -------------------- Configuration --------------------
ASSISTANT_NAME = "A.N.S.H."
ASSISTANT_FULL = "Adaptive Natural Speech Helper"
MEMORY_FILE = "ansh_memory.json"
MUSIC_DIR_DEFAULT = os.path.expanduser(r"C:\Users\Dell\Music")
SCREENSHOT_DIR = os.path.expanduser(".")
LOG_FILE = "ansh.log"

# application paths 
APP_PATHS = {
    "vscode": r"D:\Microsoft Visual Studio Code\Microsoft VS Code\Code.exe",
    "notepad": r"C:\Windows\system32\notepad.exe",
    "explorer": r"C:\Windows\explorer.exe",
    "msedge": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "paint": r"C:\Windows\system32\mspaint.exe",
}

# -------------------- Logging --------------------
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("ansh")

# -------------------- TTS Setup --------------------
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# pick voice 0 by default; user can change to voices[1].id if desired
engine.setProperty('voice', voices[0].id)

def speak(text, block=True):
    """Speak aloud and print to console... picks small natural transitions automatically."""
    text = str(text)
    print(f"{ASSISTANT_NAME}: {text}")
    try:
        engine.say(text)
        if block:
            engine.runAndWait()
    except Exception as e:
        logger.exception("TTS error")
        print("TTS failed... continuing without spoken audio.")

# -------------------- Memory / Context --------------------
def load_memory():
    if not os.path.exists(MEMORY_FILE):
        default = {"user_name": None, "favorites": {}, "last_commands": []}
        save_memory(default)
        return default
    try:
        with open(MEMORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        logger.exception("Failed to load memory file... resetting.")
        default = {"user_name": None, "favorites": {}, "last_commands": []}
        save_memory(default)
        return default

def save_memory(data):
    try:
        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        logger.exception("Failed to save memory.")

memory = load_memory()

def remember(key, value):
    memory[key] = value
    save_memory(memory)

def push_last_command(cmd):
    memory.setdefault("last_commands", [])
    memory["last_commands"].insert(0, {"time": time.time(), "cmd": cmd})
    # keep last 50
    memory["last_commands"] = memory["last_commands"][:50]
    save_memory(memory)

# -------------------- Natural Response Helpers --------------------
def choose(*options):
    return random.choice(options)

GREETINGS = [
    "Hello... how can I help?",
    "Hi there... what's up?",
    "At your service... shoot your command.",
    "Ready when you are... what shall we do?"
]

CONFIRMATIONS = [
    "Done.",
    "All set.",
    "Consider it done.",
    "Okay... finished.",
    "On it."
]

ERROR_RESPONSES = [
    "I hit a snag... try again?",
    "Hmm... that didn't work. Wanna try another way?",
    "I couldn't do that... more details might help.",
    "That failed... I'm not proud, but it failed."
]

def natural_say_success(extra=""):
    speak(choose(*CONFIRMATIONS) + (f" {extra}" if extra else ""))

def natural_say_error(extra=""):
    speak(choose(*ERROR_RESPONSES) + (f" {extra}" if extra else ""))

# -------------------- Speech Recognition / Timing --------------------
recognizer = sr.Recognizer()
recognizer.dynamic_energy_threshold = True
recognizer.energy_threshold = 300  # base level; will auto-adjust
recognizer.pause_threshold = 0.9  # shorter pause to be snappier
recognizer.operation_timeout = None

def take_command(allow_text_fallback=True, retries=2):
    """Listen and return text... retries a few times and falls back to typed input if allowed."""
    mic = sr.Microphone()
    attempts = 0
    while attempts <= retries:
        attempts += 1
        with mic as source:
            try:
                # adjust slightly for ambient noise
                recognizer.adjust_for_ambient_noise(source, duration=0.9)
                print("Listening...")
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=8)
            except sr.WaitTimeoutError:
                print("No speech detected... timeout.")
                if attempts <= retries:
                    speak("I didn't catch that... say it again.")
                    continue
                break

        try:
            print("Recognizing...")
            query = recognizer.recognize_google(audio, language='en-in')
            print(f"Recognized: {query}")
            return query.lower().strip()
        except sr.UnknownValueError:
            print("Could not understand audio.")
            if attempts <= retries:
                speak("Sorry, didn't quite get that... please repeat.")
                continue
            break
        except sr.RequestError:
            logger.exception("Speech recognition request failed")
            natural_say_error("Network or recognition service issue.")
            break
        except Exception:
            logger.exception("Unexpected recognition error")
            break

    if allow_text_fallback:
        try:
            speak("You can type your command if voice isn't working. Type now:")
            text = input("Type command (or press Enter to skip): ").strip()
            if text:
                return text.lower()
        except Exception:
            pass

    return "none"

# -------------------- Utility / Small Features --------------------
def wish_me():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good morning... bright start.")
    elif 12 <= hour < 18:
        speak("Good afternoon... keep rolling.")
    else:
        speak("Good evening... time to unwind.")
    if memory.get("user_name"):
        speak(f"{memory['user_name']}, {choose(*GREETINGS)}")
    else:
        speak(choose(*GREETINGS))
        ask_and_store_name()

def ask_and_store_name():
    speak("By the way... what should I call you?")
    name = take_command()
    if name and name != "none":
        memory["user_name"] = name.title()
        save_memory(memory)
        speak(f"Nice to meet you, {memory['user_name']}. I'll remember that.")

def open_website(url, pretty_name=None):
    try:
        webbrowser.open(url)
        natural_say_success(f"Opening {pretty_name or url}.")
    except Exception:
        logger.exception("Failed to open website")
        natural_say_error("Could not open the website.")

def open_app(path, pretty_name=None):
    try:
        if not os.path.exists(path):
            raise FileNotFoundError(f"Path not found: {path}")
        os.startfile(path)
        natural_say_success(pretty_name or os.path.basename(path))
    except Exception as e:
        logger.exception("Failed to open app")
        natural_say_error(f"Could not open {pretty_name or path}...")

def play_music(music_dir=MUSIC_DIR_DEFAULT):
    try:
        files = os.listdir(music_dir)
        audio_files = [f for f in files if f.lower().endswith(('.mp3', '.wav', '.m4a', '.flac'))]
        if not audio_files:
            speak("I couldn't find music in the folder.")
            return
        chosen = audio_files[0]
        os.startfile(os.path.join(music_dir, chosen))
        speak(f"Playing {chosen} ... enjoy.")
        memory['last_music'] = chosen
        save_memory(memory)
    except Exception:
        logger.exception("Play music error")
        natural_say_error("Couldn't play music. Is the folder present?")

def take_screenshot_flow():
    speak("What should I name the screenshot?")
    name = take_command()
    if not name or name == "none":
        name = f"screenshot_{int(time.time())}"
    filename = os.path.join(SCREENSHOT_DIR, f"{name}.png")
    try:
        time.sleep(1.2)  # tiny buffer
        img = pyautogui.screenshot()
        img.save(filename)
        natural_say_success(f"Saved screenshot as {filename}")
        # remember last screenshot
        memory['last_screenshot'] = filename
        save_memory(memory)
    except Exception:
        logger.exception("Screenshot error")
        natural_say_error("Failed to take screenshot.")

def read_pdf(pdf_path=None):
    if PyPDF2 is None:
        speak("PDF reading requires PyPDF2... it's not installed.")
        return
    try:
        if not pdf_path:
            speak("Please type the full pdf path now.")
            pdf_path = input("PDF path: ").strip()
        if not pdf_path or not os.path.exists(pdf_path):
            speak("PDF file not found.")
            return
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfFileReader(f)
            pages = reader.numPages
            speak(f"The PDF has {pages} pages. Which page should I read?")
            page_num_txt = take_command()
            try:
                p = int(page_num_txt)
            except Exception:
                p = 0
            p = max(0, min(p, pages-1))
            page = reader.getPage(p)
            text = page.extractText()
            if text.strip():
                speak("Reading page now.")
                speak(text)
            else:
                speak("That page had no readable text.")
    except Exception:
        logger.exception("PDF reading error")
        natural_say_error("Could not read the PDF.")

def get_ip_address():
    try:
        ip = get('https://api.ipify.org').text
        speak(f"Your IP address appears to be {ip}")
    except Exception:
        logger.exception("IP fetch error")
        natural_say_error("Could not retrieve IP address.")

def send_whatsapp_message():
    if kit is None:
        speak("WhatsApp message sending requires pywhatkit which is not installed.")
        return
    try:
        speak("Please type the recipient number in international format (eg +911234567890)...")
        number = input("Number: ").strip()
        speak("What message should I send?")
        message = take_command()
        if number and message and message != "none":
            speak("When should I send it? Provide hour and minute in 24-hour format, separated by space (or press Enter to send in 1 minute).")
            when = input("Hour Minute (or enter): ").strip()
            if when:
                hour_s, minute_s = when.split()[:2]
                kit.sendwhatmsg(number, message, int(hour_s), int(minute_s))
            else:
                # send in 1 minute
                t = datetime.datetime.now() + datetime.timedelta(minutes=1)
                kit.sendwhatmsg(number, message, t.hour, t.minute)
            natural_say_success("WhatsApp message scheduled/sent.")
        else:
            speak("Message or number missing.")
    except Exception:
        logger.exception("WhatsApp send error")
        natural_say_error("Failed to send WhatsApp message.")

def check_battery():
    try:
        battery = psutil.sensors_battery()
        if battery is None:
            speak("Battery information is unavailable on this machine.")
            return
        percent = battery.percent
        plugged = battery.power_plugged
        speak(f"Battery is at {percent} percent.")
        if plugged:
            speak("Charger is connected.")
        else:
            if percent < 20:
                speak("Low battery... please plug in soon.")
            elif percent < 50:
                speak("Battery OK but not full.")
            else:
                speak("Battery healthy for now.")
    except Exception:
        logger.exception("Battery check error")
        natural_say_error("Couldn't get battery status.")

def adjust_volume(action):
    try:
        if action == "up":
            pyautogui.press("volumeup")
        elif action == "down":
            pyautogui.press("volumedown")
        elif action == "mute":
            pyautogui.press("volumemute")
        elif action == "unmute":
            pyautogui.press("volumeunmute")
        natural_say_success("Volume updated.")
    except Exception:
        logger.exception("Volume control error")
        natural_say_error("Volume control failed.")

# -------------------- Main Loop --------------------
def main_loop():
    wish_me()
    while True:
        try:
            query = take_command()
            if not query or query == "none":
                continue
            push_last_command(query)
            # --- Wikipedia ---
            if 'wikipedia' in query:
                try:
                    speak("Searching Wikipedia...")
                    search_term = query.replace("wikipedia", "").strip()
                    if not search_term:
                        speak("What should I search on Wikipedia?")
                        search_term = take_command()
                    results = wikipedia.summary(search_term, sentences=2)
                    speak("According to Wikipedia...")
                    speak(results)
                except Exception:
                    logger.exception("Wikipedia error")
                    natural_say_error("Wikipedia lookup failed.")
            # --- websites ---
            elif 'open youtube' in query or 'youtube' in query:
                open_website("https://www.youtube.com/", "YouTube")
            elif 'open google' in query or 'google' in query:
                open_website("https://www.google.co.in/webhp?tab=rw&authuser=0", "Google")
            elif 'open classroom' in query or 'classroom' in query:
                open_website("https://classroom.google.com/u/3/h", "Google Classroom")
            elif 'open meet' in query or 'meet' in query:
                open_website("https://meet.google.com/?authuser=3", "Google Meet")
            elif 'open whatsapp' in query or 'whatsapp' in query:
                open_website("https://web.whatsapp.com/", "WhatsApp")
            elif 'open gmail' in query or 'gmail' in query:
                open_website("https://mail.google.com/mail/u/2/#inbox", "Gmail")
            elif 'open drive' in query or 'google drive' in query:
                open_website("https://drive.google.com/drive/u/2/my-drive", "Google Drive")
            # --- Search helpers ---
            elif 'search google' in query or query.startswith("search "):
                speak("What should I search for?")
                what = take_command()
                if what and what != "none":
                    open_website(f"https://www.google.com/search?q={what}", f"Google search for {what}")
            elif 'search youtube' in query or 'youtube search' in query:
                speak("What should I search on YouTube?")
                what = take_command()
                if what and what != "none":
                    open_website(f"https://www.youtube.com/results?search_query={what}", f"YouTube search for {what}")
            # --- media and files ---
            elif 'play music' in query or 'play some music' in query:
                play_music()
            elif 'screenshot' in query or 'take screenshot' in query:
                take_screenshot_flow()
            elif 'open visual studio code' in query or 'visual studio code' in query:
                open_app(APP_PATHS.get("vscode"), "Visual Studio Code")
            elif 'notepad' in query:
                open_app(APP_PATHS.get("notepad"), "Notepad")
            elif 'command prompt' in query or 'cmd' in query:
                try:
                    os.system("start cmd")
                    natural_say_success("Command prompt opened.")
                except Exception:
                    logger.exception("cmd open error")
                    natural_say_error("Couldn't open command prompt.")
            elif 'files' in query or 'explorer' in query:
                open_app(APP_PATHS.get("explorer"), "File Explorer")
            elif 'paint' in query:
                open_app(APP_PATHS.get("paint"), "Paint")
            elif 'open edge' in query or 'microsoft edge' in query:
                open_app(APP_PATHS.get("msedge"), "Microsoft Edge")
            # --- system info ---
            elif 'system information' in query or 'winver' in query:
                try:
                    os.system("start winver.exe")
                    natural_say_success("Opened system information.")
                except Exception:
                    logger.exception("winver error")
                    natural_say_error("Could not open system information.")
            # --- pdf read ---
            elif 'pdf' in query or 'read pdf' in query:
                read_pdf()
            # --- ip address ---
            elif 'ip address' in query or 'my ip' in query:
                get_ip_address()
            # --- whatsapp message ---
            elif 'message on whatsapp' in query or 'send whatsapp' in query:
                send_whatsapp_message()
            # --- battery ---
            elif 'battery' in query:
                check_battery()
            # --- volume controls ---
            elif 'volume up' in query:
                adjust_volume("up")
            elif 'volume down' in query:
                adjust_volume("down")
            elif 'mute volume' in query or 'mute' in query:
                adjust_volume("mute")
            elif 'unmute volume' in query or 'unmute' in query:
                adjust_volume("unmute")
            # --- time ---
            elif 'time' in query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"The time is {strTime}")
            # --- quitting ---
            elif any(x in query for x in ['quit', 'exit', 'you can quit now', 'shutdown', 'stop assistant']):
                speak("Thank you... I'll be here if you need me.")
                sys.exit()
            else:
                speak(choose("Sorry... I didn't understand that.", "Could you rephrase that?", "I don't have a handler for that yet."))
        except KeyboardInterrupt:
            speak("Keyboard interrupt received... exiting.")
            sys.exit()
        except Exception:
            logger.exception("Main loop unexpected error")
            natural_say_error("Something unexpected occurred... continuing.")

if __name__ == "__main__":
    try:
        logger.info("A.N.S.H started")
        speak(f"Booting {ASSISTANT_NAME}... {ASSISTANT_FULL}", block=False)
        main_loop()
    except Exception:
        logger.exception("Fatal error on startup")
        speak("Critical failure... check the logs.")
