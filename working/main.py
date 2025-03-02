from flask import Flask, render_template, jsonify
from comtypes import CoInitialize, CoUninitialize
import os
import webbrowser
import threading
import datetime
import pyttsx3
import speech_recognition as sr
import subprocess
import pyautogui
import time
import google.generativeai as genai  # Gemini AI library

app = Flask(__name__)

class VoiceAssistant:
    def __init__(self):
        self.engine = pyttsx3.init()
        voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', voices[1].id if len(voices) > 1 else voices[0].id)
        self.silent_mode = False
        self.output_dir = "output"
        os.makedirs(self.output_dir, exist_ok=True)
        self.app_keywords = {
            "chrome": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            "notepad": "notepad.exe",
            "calculator": "calc.exe"
        }
        self.app_names = ["chrome", "notepad", "calculator"]

    def speak(self, audio):
        if not self.silent_mode:
            self.engine.say(audio)
            self.engine.runAndWait()

    def take_command(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1
            try:
                audio = r.listen(source, timeout=4, phrase_time_limit=3)
            except sr.WaitTimeoutError:
                self.speak("I didn't hear anything.")
                return "none"
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}")
            return query.lower()
        except sr.UnknownValueError:
            self.speak("Sorry, I didn't catch that.")
            return "none"
        except sr.RequestError:
            self.speak("Couldn't connect to the recognition service.")
            return "none"

    def handle_silent_mode(self, query):
        silent_commands = ["chup raho", "shut up", "bandh", "chup", "stay silent", "silent"]
        wake_commands = ["system", "hello system", "hello assistant"]

        if any(word in query for word in silent_commands):
            self.speak("Sorry to interrupt, going silent.")
            self.silent_mode = True
            return True
        elif any(word in query for word in wake_commands):
            self.silent_mode = False
            self.speak("Yeah, I'm here. How can I help you?")
            time.sleep(2)
            return False
        return False

    def take_screenshot(self):
        screenshot = pyautogui.screenshot()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filepath = os.path.join(self.output_dir, f"screenshot_{timestamp}.png")
        screenshot.save(filepath)
        self.speak(f"Screenshot saved as {filepath}")

    def open_app_or_website(self, query):
        for app_name, path in self.app_keywords.items():
            if app_name in query:
                if "close" in query:
                    os.system(f"taskkill /f /im {app_name}.exe")
                    self.silent_mode = False
                    self.speak(f"{app_name.capitalize()} closed.")
                else:
                    if path.startswith("http"):
                        webbrowser.open(path)
                    else:
                        os.startfile(path)
                    self.silent_mode = True
                    self.speak(f"{app_name.capitalize()} opened.")
                return

        for app_name in self.app_names:
            if app_name in query:
                try:
                    app_path = os.path.join(os.getenv('ProgramData'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', app_name + '.lnk')
                    subprocess.Popen(app_path)
                    self.silent_mode = True
                    self.speak(f"{app_name.capitalize()} opened.")
                    return
                except Exception as e:
                    self.speak(f"Failed to open {app_name}.")

        site_name = query.replace("open ", "").strip()
        default_url = f"http://www.{site_name}.com"
        self.speak(f"Opening {site_name}.com.")
        webbrowser.open(default_url)

    def youtube_search(self):
        self.speak("Youtube is opened. What would you like to search?")
        for attempt in range(3):
            search_query = self.take_command()
            if search_query and search_query != "none":
                youtube_search_url = f"https://www.youtube.com/results?search_query={search_query.replace(' ', '+')}"
                self.speak(f"Searching YouTube for {search_query}.")
                webbrowser.open(youtube_search_url)
                return
            elif attempt < 2:
                self.speak("I didn't catch that. Please tell me what you'd like to search on YouTube.")
        self.speak("I couldn't hear your search query. Please try again later.")
    
    def wikipedia_search(self, query):
        topic = query.replace("search", "").strip()
        wikipedia_url = f"https://en.wikipedia.org/wiki/{topic.replace(' ', '_')}"
        self.speak(f"Searching Wikipedia for {topic}.")
        webbrowser.open(wikipedia_url)

    def anime(self, query):
        topic = query.replace("anime", "").strip()
        if topic:
            search_url = f"https://hianime.to/search?keyword={topic}"
            self.speak(f"Searching HiAnime for {topic}.")
            webbrowser.open(search_url)
        else:
            self.speak("Please specify the anime title you want to search for.")

    #def ask(self, query):
        topic = query.replace("", "").strip()
        if topic:
            search_url = f"https://in.search.yahoo.com/search?fr=mcafee&type=E210IN826G0&p={topic}"
            self.speak(f"let me trhin {topic}.")
            webbrowser.open(search_url)
        else:
            self.speak("Please specify the anime title you want to search for.")    

    def execute_query(self, query):
        if self.handle_silent_mode(query):
            return
        elif "anime" in query:
            self.anime(query)
        elif "youtube" in query:
            self.youtube_search()
        elif "search" in query:
            self.wikipedia_search(query)
        elif "open" in query:
            self.open_app_or_website(query)
        elif "screenshot" in query:
            self.take_screenshot()
        if "stop" in query or "exit" in query:
            self.speak("Shutting down. Have a great day!")
            exit()
        if "close" in query:
            self.speak("Please specify which app to close.")

    def run(self):
        CoInitialize()  # Initialize COM library
        self.speak("Starting voice assistant. How can I assist you?")
        while True:
            query = self.take_command()
            if query != "none":
                self.execute_query(query)
        CoUninitialize()  # Uninitialize COM library when done

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/start_assistant', methods=['POST'])
def start_assistant():
    def run_assistant():
        assistant = VoiceAssistant()
        assistant.run()

    assistant_thread = threading.Thread(target=run_assistant)
    assistant_thread.daemon = True
    assistant_thread.start()

    return jsonify({'response': 'Voice Assistant started'})

if __name__ == "__main__":
    app.run(debug=True)