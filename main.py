import subprocess
import os
import pygame

from bs4 import BeautifulSoup
import requests
import speech_recognition as sr
import win32com.client as wincom
import datetime
import webbrowser
import wikipedia
import sys
from pywikihow import search_wikihow
from PyDictionary import PyDictionary
import googletrans
from googletrans import Translator
from gtts import gTTS

print("1.To use wikipedia: simple say for examole Shah Rukh khan wikipedia.....")
print("2.To use website: say there name for eg: youtube........although it wont work on your pc because the "
      "chrome_path is differen..!")
print("3.To set alarm simply say set alarm or set or set reminder.......and i dont know whether it will work or not")
print("4.If you want a step by step module for eg: say activate or process mode and say how to drive a car....it will "
      "show the step")
print("5.For Dictionary: say dictionary")
print("6.For sleeping: say sleep....it wont stop the program....if you want to use it again say wake up")
print("7.To shutdown the program or to exit: say exit")
print("8.To shutdown your pc: say shut down")
print("=============IF YOU WANT MANUAL SIMPLY SAY MANUAL THIS WHOLE WILL SHOW AGAIN==================")
print("AND THIS WHOLE PROGRAM IS CALL:'CHINCHI'------")
print("----------THIS IS DEVELOP BY ROHAN KHAWAS")
print("------------------------------------------TO START THIS JUST SAY 'WAKE UP' OR "
      "'START'-------------------------------------------------------------------")


def speak(audio):
    speech = wincom.Dispatch("SAPI.SpVoice")
    speech.Speak(audio)


def wishMe():
    time = datetime.datetime.now().strftime("%I:%M:%p")
    day = datetime.datetime.now().strftime("%B:%d:%Y:%A")

    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        print("Good Morning...")
        speak("Good Morning...")
    elif 12 <= hour < 18:
        print("Good Afternoon.....")
        speak("Good Afternoon.....")
    else:
        print("Good Evening......")
        speak("Good Evening......")
    print(f"Today's {day} and its {time}")
    speak(f"Today's {day} and its {time}")


def weather():
    search = "weather of kathmandu"
    url = f"https://www.google.com/search?q={search}"
    r = requests.get(url)
    soup = BeautifulSoup(r.text, 'html.parser')
    temperature_element = soup.find("div", class_='BNeawe')
    condition_element = soup.find("div", class_="BNeawe tAd8D AP7Wnd")

    if temperature_element:
        temperature = temperature_element.text
        condition = condition_element.text
        print(f"The current temperature is {temperature} and its {condition}")
        speak(f"The current temperature is {temperature} and its {condition}")
        print(f"Hello I am Chinchi, What can I do for you....?")
        speak(f"Hello I am Chinchi, what can I do for you.....?")

    else:
        print("Failed to retrieve weather information.")


def manual():
    print("1.To use wikipedia: simple say for examole Shah Rukh khan wikipedia.....")
    print(
        "2.To use website: say there name for eg: youtube........although it wont work on your pc because the chrome_path is differen..!")
    print(
        "3.To set alarm simply say set alarm or set or set reminder.......and i dont know whether it will work or not")
    print(
        "4.If you want a step by step module for eg: say activate or process mode and say how to drive a car....it will show the step")
    print("5.For Dictionary: say dictionary")
    print("6.For sleeping: say sleep....it wont stop the program....if you want to use it again say wake up")
    print("To Remind just say remember that-------------")
    print("7.To shutdown the program or to exit: say exit")
    print("8.To shutdown your pc: say shut down")
    print("=============IF YOU WANT MANUAL SIMPLY SAY MANUAL THIS WHOLE WILL SHOW AGAIN==================")


def lookup_word(word):
    dictionary = PyDictionary()
    definition = dictionary.meaning(word)
    if definition:
        print(f"Word: {word}")
        for part_of_speech, meanings in definition.items():
            print(f"{part_of_speech.capitalize()}:")
            for meaning in meanings:
                print(f"- {meaning}")
        print()
    else:
        print(f"Word '{word}' not found in the dictionary.\n")


def set_alarm():
    try:
        speak("please enter the time...")
        alarm_time = input("Please enter the time (format HH:MM:AM/PM):")
        reminder_text = input("Enter what you want me to remind you of, when you wake up :")
        current_time = datetime.datetime.now().strftime("%I:%M:%p")
        print(f"setting alarm for {alarm_time}")
        print(f"current time :{current_time}")
        speak(f"setting alarm for :{alarm_time}")
        while current_time != alarm_time:
            current_time = datetime.datetime.now().strftime("%I:%M:%p")
        print(f"reminder: {reminder_text}")
        speak(f"reminder: {reminder_text}")
        os.startfile("music.mp3")
    except:
        print("Developer needs to learn more coding.....!")


def chrome():
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))

def translate_text(query):
    speak("Sure sir")
    translator = Translator()
    speak("choose the language from below...")
    print(googletrans.LANGUAGES)
    dest_lang = input("Enter destination language: ")
    try:
        translation = translator.translate(query, dest=dest_lang)
        translated_text = translation.text
        print(f"Translated text: {translated_text}")
        #speak(f"Translated text: {translated_text}")
        text_to_speech = gTTS(text=translated_text, lang=dest_lang, slow=False)
        text_to_speech.save('voice.mp3')
        pygame.mixer.init()
        pygame.mixer.music.load('voice.mp3')
        pygame.mixer.music.play()

        while pygame.mixer.music.get_busy():
            continue

        pygame.mixer.music.stop()
        pygame.mixer.quit()

        os.remove('voice.mp3')

        #os.remove('voice.mp3')
    except:
        print("An error occurred. Please try again.")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        print("Listening.....")
        audio_text = r.listen(source)
    try:
        print("Recognizing.....")
        query = r.recognize_google(audio_text, language="en")
        print(f"User said:{query}")
        # speak(query)
    except Exception as e:
        print("Sorry I didn't catch that, please try again........")
        return "None"
    return query


def Execution():
    wishMe()
    weather()
    while True:
        query = takeCommand().lower()
        if 'time' in query:
            time = datetime.datetime.now().strftime("%I:%M:%p")
            print(f"The current time is  {time}")
            speak(f"The current time is  {time}")

        elif 'wikipedia' in query:
            speak("Searching from wikipedia...")
            try:
                q = query.replace("search from wikipedia", "")
                q = query.replace("wikipedia", "")
                speak("According to wikipedia......")
                search = wikipedia.summary(q, 4)
                print(search)
                speak(search)
            except:
                print("Sorry, The developer needs to learn more coding...... ")


        elif 'weather' in query:
            weather()

        elif 'youtube' in query:
            chrome()
            try:
                query = query.replace("search from youtube", "")
                webbrowser.get('chrome').open_new_tab(url=f"https://www.youtube.com/results?search_query={query}")
            except:
                print("your chrome path is different from developer...........")
        elif 'google' in query:
            chrome()
            try:
                query = query.replace("search from google", "")
                webbrowser.get('chrome').open_new_tab(url=f"https://www.google.com/search?q={query}")
            except:
                print("your chrome path is different from developer...........")

        elif 'anime' in query:
            chrome()
            try:
                webbrowser.get('chrome').open_new_tab(url="https://9anime.gs/home")
            except:
                print("Sorry, The developer needs to learn more coding...... ")

        elif 'drama' in query:
            chrome()
            try:
                webbrowser.get('chrome').open_new_tab(url="https://kissasian.pe/drama/king-the-land-2023-episode-1")
            except:
                print("Sorry, The developer needs to learn more coding...... ")


        elif 'remember that' in query:
            RememberMessage = query.replace("remember that", "")
            speak(f"you told me to remember that {RememberMessage}")
            remember = open("Remember.txt", "w")
            remember.write(RememberMessage)
            remember.close()

        elif 'what i told you to remember ' in query or 'remember' in query:
            remember = open("Remember.txt", "r")
            speak(f"you told me to remember {remember.read()}")

        elif 'translate' in query or 'T' in query:
            trans = query.replace("translate", "")
            translate_text(trans)


        elif 'set alarm' in query or 'set' in query or 'set reminder' in query:
            set_alarm()

        elif 'activate' in query or 'process mode' in query:
            speak("Activating the process mode")
            try:
                print("What you want to search:")
                speak("What you want to search")
                q = takeCommand().lower()
                max_result = 1
                process = search_wikihow(q, max_result)
                assert len(process) == 1
                process[0].print()
                speak(process[0].summary)
            except:
                print("Unable to search because I lack coding....")

        elif 'dictionary' in query:
            speak("Activating the simple dictionary")
            speak("enter what you want to search....")
            search_word = input("Enter a word to look up: ")
            lookup_word(search_word)

        elif 'manual' in query:
            manual()
        elif 'notepad' in query or 'note' in query:
            speak("opening notepad")
            subprocess.Popen(['notepad.exe'])

        elif 'sleep' in query:
            speak("Sleeping mode activated")
            break

        elif 'exit' in query:
            speak("shutting down the program....")
            sys.exit()

        elif 'shut down' in query or 'shutdown  system' in query or 'shutdown' in query:
            speak("Are you sure you want to shut down your computer..")
            shut_down = input("Do you wish to shutdown ? (yes/no):")
            if shut_down == "yes":
                os.system("shutdown /s /t 1")
            elif shut_down == 'no':
                break


if __name__ == "__main__":
    while True:
        query = takeCommand().lower()
        if 'wake up' in query or 'start' in query:
            Execution()
