import pyttsx3
import wikipedia
import webbrowser
import os
import datetime
import keyboard
import speech_recognition as sr
import nltk
import logging
import tkinter as tk
from youtubesearchpython import VideosSearch
from pynput.keyboard import Controller  # Import pynput for keyboard control
import subprocess  # Import subprocess to run external commands

nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('maxent_ne_chunker')
nltk.download('words')

# Initialize text-to-speech engine
engine = pyttsx3.init()
keyboard_controller = Controller()  # Initialize keyboard controller for pynput

# Set up logging
logging.basicConfig(filename='assistant.log', level=logging.INFO)

def speak(text):
    engine.say(text)
    engine.runAndWait()

def greet_user():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good morning!")
    elif 12 <= hour < 18:
        speak("Good afternoon!")
    else:
        speak("Good evening!")

def search_wikipedia(query):
    try:
        search_results = wikipedia.search(query)
        if not search_results:
            speak("Sorry, I couldn't find information on that topic")
            return

        top_result = search_results[0]  # Get the top search result
        page = wikipedia.page(top_result)
        results = wikipedia.summary(page.title, sentences=3)
        speak("Here is what I found on Wikipedia")
        speak(results)
        display_results("Wikipedia Search Results", results)
    except wikipedia.exceptions.PageError:
        speak("Sorry, I couldn't find information on that topic")
    except wikipedia.exceptions.DisambiguationError:
        speak("There are multiple matches for that query. Please be more specific.")


def open_application(app_name):
    try:
        applications = {
            "excel": "start excel",
            "chrome": "start chrome",
            "word": "start winword",
            "visual studio code": "code",
            "task manager": "start taskmgr",
            "youtube": "https://www.youtube.com/",
            "google": "https://www.google.com/",
            "this pc": "start explorer",
            "settings": "start ms-settings:",
            "documents": "start documents",
            "control panel": "start control",
            "local disk d": "start D:",
            "local disk c": "start C:",
            "local disk e": "start E:",
        }
        if app_name.lower() in applications:
            if app_name.lower() in ["youtube", "google"]:
                webbrowser.open(applications[app_name.lower()])
            else:
                os.system(applications[app_name.lower()])
            speak(f"Opening {app_name}")
        else:
            speak(f"Sorry, I couldn't find the application {app_name}")
    except Exception as e:
        logging.error(f"Error opening application {app_name}: {e}")
        speak("Sorry, I couldn't open the application")

def close_application(app_name):
    try:
        app_processes = {
            "excel": "excel.exe",
            "chrome": "chrome.exe",
            "word": "winword.exe",
            "visual studio code": "code.exe",
            "task manager": "taskmgr.exe"
        }
        if app_name.lower() in app_processes:
            os.system(f"taskkill /f /im {app_processes[app_name.lower()]}")
            speak(f"Closing {app_name}")
        else:
            speak(f"Sorry, I couldn't find the application {app_name} to close")
    except Exception as e:
        logging.error(f"Error closing application {app_name}: {e}")
        speak("Sorry, I couldn't close the application")

def shutdown_system():
    try:
        speak("Shutting down the system")
        os.system("shutdown /s /t 1")
    except Exception as e:
        logging.error(f"Error shutting down the system: {e}")
        speak("Sorry, I couldn't shut down the system")

def restart_system():
    try:
        speak("Restarting the system")
        os.system("shutdown /r /t 1")
    except Exception as e:
        logging.error(f"Error restarting the system: {e}")
        speak("Sorry, I couldn't restart the system")

def sleep_system():
    try:
        speak("Putting the system to sleep")
        os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
    except Exception as e:
        logging.error(f"Error putting the system to sleep: {e}")
        speak("Sorry, I couldn't put the system to sleep")

def play_youtube_video(query):
    try:
        videos_search = VideosSearch(query, limit=1)
        result = videos_search.result()
        if result and result['result']:
            first_result = result['result'][0]
            video_url = first_result['link']
            webbrowser.open(video_url)
            speak(f"Playing {first_result['title']} on YouTube")
        else:
            speak("Sorry, I couldn't find any videos matching your query")
    except Exception as e:
        logging.error(f"Error playing YouTube video: {e}")
        speak("Sorry, I couldn't play the video")

def type_in_notepad(text):
    subprocess.Popen(["notepad.exe"])  # Open Notepad
    keyboard_controller.type(text)  # Type the text in Notepad

def process_command():
    speak("How can I assist you?")
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    while True:
        if keyboard.is_pressed('ctrl'):
            with mic as source:
                recognizer.adjust_for_ambient_noise(source)
                print("Listening...")
                audio = recognizer.listen(source)
            
            try:
                text = recognizer.recognize_google(audio).lower()
                logging.info(f"User command: {text}")
                print("You said:", text)
                
                tokens = nltk.word_tokenize(text)
                tagged = nltk.pos_tag(tokens)
                entities = nltk.chunk.ne_chunk(tagged)

                for subtree in entities.subtrees():
                    if subtree.label() == 'PERSON':
                        person_name = ' '.join([word for word, tag in subtree.leaves()])
                        speak(f"Hello {person_name}, how can I help you?")
                        break

                if "search" in text:
                    query = text.replace("search", "").strip()
                    search_wikipedia(query)
                elif "open" in text:
                    target = text.replace("open", "").strip()
                    open_application(target)
                elif "close" in text:
                    target = text.replace("close", "").strip()
                    close_application(target)
                elif "shutdown" in text:
                    shutdown_system()
                    break
                elif "restart" in text:
                    restart_system()
                    break
                elif "sleep" in text:
                    sleep_system()
                    break
                elif "play" in text and "youtube" in text:
                    query = text.replace("play", "").replace("youtube", "").strip()
                    play_youtube_video(query)
                elif "type" in text or "write" in text:
                    message = text.replace("type", "").replace("write", "").strip()
                    type_in_notepad(message)  # Call the new function to type in Notepad
                    speak("Text typed successfully")
                elif "bye" in text or "goodbye" in text:
                    speak("Goodbye! Have a great day!")
                    break
                else:
                    speak("Please provide a valid command")
                
                speak("How else can I assist you?")
            
            except sr.UnknownValueError:
                logging.error("Could not understand audio")
                print("Could not understand audio")
            except sr.RequestError as e:
                logging.error(f"Could not request results; check your internet connection: {e}")
                print("Could not request results; check your internet connection")

def display_results(title, message):
    result_window = tk.Tk()
    result_window.title(title)
    result_window.geometry("400x200")
    
    label = tk.Label(result_window, text=message, wraplength=300)
    label.pack(pady=20)
    
    button = tk.Button(result_window, text="OK", command=result_window.destroy)
    button.pack(pady=10)
    
    result_window.mainloop()

if __name__ == "__main__":
    greet_user()
    process_command()
