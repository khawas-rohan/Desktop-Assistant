import subprocess
import os
from time import sleep

import pyautogui
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
from pynput.keyboard import Key, Controller
import speedtest
from gtts import gTTS
import random
from plyer import notification

print("To use Desktop Assistant say<-------Wake Up or Start-------->")


def speak(audio):
    speech = wincom.Dispatch("SAPI.SpVoice")
    speech.Speak(audio)


def volumeup():
    keyboard = Controller()
    for i in range(5):
        keyboard.press(Key.media_volume_up)
        keyboard.release(Key.media_volume_up)
        sleep(0.1)


def volumedowm():
    keyboard = Controller()
    for i in range(5):
        keyboard.press(Key.media_volume_down)
        keyboard.release(Key.media_volume_down)
        sleep(0.1)


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


def quotes():
    quotes = ["The greatest glory in living lies not in never falling, but in rising every time we fall.",
              "The future belongs to those who believe in the beauty of their dreams.",
              "Success is not final, failure is not fatal: It is the courage to continue that counts.",
              "The only limit to our realization of tomorrow will be our doubts of today.",
              "The way to get started is to quit talking and begin doing.",
              "In the end, we will remember not the words of our enemies, but the silence of our friends.",
              "Don't watch the clock; do what it does. Keep going.",
              "The only thing we have to fear is fear itself.",
              "Life is really simple, but we insist on making it complicated.",
              "The journey of a thousand miles begins with one step.",
              "It's so hard to forget someone who gave you so much to remember.",
              "Sometimes, the person you'd take a bullet for ends up being behind the trigger.",
              "The saddest thing about love is that not only the love cannot last forever, but even the heartbreak is soon forgotten.",
              "Pain is inevitable. Suffering is optional.",
              "It hurts to breathe because every breath I take proves I can't live without you.",
              "The saddest thing about love is that not only the love cannot last forever, but even the heartbreak is soon forgotten.",
              "Tears come from the heart and not from the brain.",
              "The emotion that can break your heart is sometimes the very one that heals it.",
              "There is one pain I often feel, which you will never know. It's caused by the absence of you.",
              "It's amazing how someone can break your heart and you can still love them with all the little pieces.",
              "You can change what you do, but you can't change what you want.",
              "Sometimes, you can't see what's right in front of you, until it's too late.",
              "You don't parley with the Devil. The Devil parleys with you.",
              "The problem with ambition is it's like setting a watch. It never stops.",
              "To be in hell is to drift. To be in heaven is to steer."
              ]
    random_quotes = random.choice(quotes)
    print(random_quotes)
    notification.notify(title="Quote of the Day",
                        message=random_quotes,
                        app_icon="reminder.ico",
                        timeout=10)
    speak(random_quotes)


def weather():
    search = "temperature of kathmandu"
    url = f"https://www.google.com/search?q={search}"
    r = requests.get(url)
    soup = BeautifulSoup(r.text, 'html.parser')
    temperature_element = soup.find("div", class_="BNeawe")
    condition_element = soup.find("div", class_="BNeawe tAd8D AP7Wnd")

    if temperature_element:
        temperature = temperature_element.text
        condition = condition_element.text
        print(f"The current temperature is {temperature} and its {condition}")
        speak(f"The current temperature is {temperature} and its {condition}")
        print(f"Hello I am ChinchiMario, What can I do for you....?")
        speak(f"Hello I am ChinchiMario, what can I do for you.....?")

    else:
        print("Failed to retrieve weather information.")


def manual():
    print("1.To use wikipedia: simple say for examole Shah Rukh khan wikipedia.....")
    print(
        "2.To use website: say there name for eg: youtube........although it wont work on your pc because the "
        "chrome_path is differen..!")
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
        # speak(f"Translated text: {translated_text}")
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

    except:
        print("An error occurred. Please try again.")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 2
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
        elif 'quote of the day' in query or 'quote' in query:
            quotes()

        elif 'wikipedia' in query:
            speak("Searching from wikipedia...")
            try:
                q = query.replace("search from wikipedia", "")
                q = query.replace("wikipedia", "")
                speak("According to wikipedia......")
                search = wikipedia.summary(q, 2)
                print(search)
                speak(search)
            except:
                print("Sorry, The developer needs to learn more coding...... ")

        elif 'question' in query or 'query' in query:
            chrome()
            print("Sure, Sir")
            speak("Sure, Sir")
            print("Please write down your query...")
            speak("Please write down your query...")
            search = input("Write your query: ")
            webbrowser.get('chrome').open_new_tab(url="https://chat.openai.com/")
            sleep(15)
            pyautogui.typewrite(search)
            pyautogui.press("enter")
            sleep(2)
            pyautogui.typewrite(search)
            pyautogui.press("enter")

        elif 'weather' in query:
            weather()

        elif 'youtube' in query:
            speak("Opening youtube..")
            chrome()
            try:
                webbrowser.get('chrome').open_new_tab(url="https://www.youtube.com/@khawas_rohan/videos")
                sleep(15)
                pyautogui.click(650, 112, duration=1)
                speak("Sir, what should I search on youtube...")
                query = query.replace("search", "")
                query = takeCommand().lower()
                sleep(3)
                pyautogui.typewrite(query)
                sleep(3)
                pyautogui.press("enter")
                sleep(10)
                #pyautogui.scroll(-800)
                pyautogui.click(600, 515, duration=1)
            except:
                print("your chrome path is different from developer...........")

        elif 'search'in query:
            pyautogui.click(650, 112, duration=1)
            #query = takeCommand().lower()
            sleep(2)
            pyautogui.hotkey("ctrl", "a")
            sleep(1)
            pyautogui.press("delete")
            sleep(1)
            query = query.replace("search", "")
            sleep(3)
            pyautogui.typewrite(query)
            sleep(3)
            pyautogui.press("enter")
            sleep(10)
            pyautogui.click(600, 515, duration=1)

        elif 'scroll down' in query:
            pyautogui.moveTo(750, 240, duration=1)
            pyautogui.scroll(-600)

        elif 'scroll up' in query:
            pyautogui.moveTo(750, 240, duration=1)
            pyautogui.scroll(600)

        elif 'send a message in facebook' in query:
            speak("Sir, Who should I send message to, Please write the name of the Recipient's")
            person = input("Recipient's name: ")
            speak("Sir, Please Write your message....")
            message = input("What's the message: ")
            speak("Opening Messenger....")
            sleep(2)
            pyautogui.click(150, 750, duration=1)
            # pyautogui.press("super")
            sleep(10)
            pyautogui.typewrite("messenger")
            sleep(5)
            pyautogui.press("enter")
            sleep(10)
            pyautogui.hotkey("Ctrl", "K")
            pyautogui.click(300, 100, duration=1)
            sleep(10)
            pyautogui.typewrite(person)
            sleep(5)
            pyautogui.hotkey("Ctrl", "1")
            speak("Sending message...")
            sleep(10)
            pyautogui.typewrite(message)
            sleep(5)
            pyautogui.press("enter")
            sleep(10)
            pyautogui.hotkey("alt", "f4")
            speak("Task completed..")


        elif 'send message in Instragram' in query:
            speak("Opening Instragram....")
            person = input("User_id: ")
            message = input("Enter your message: ")
            sleep(2)
            pyautogui.click(150, 750, duration=1)
            # pyautogui.press("win")
            sleep(5)
            pyautogui.typewrite("instagram")
            sleep(5)
            pyautogui.press("enter")
            sleep(10)
            pyautogui.click(357, 240, duration=1)
            sleep(5)
            pyautogui.typewrite(person)
            sleep(5)
            pyautogui.click(500, 300, duration=1)
            sleep(10)
            pyautogui.click(1100, 100, duration=1)
            sleep(5)
            speak("Sending message...")
            pyautogui.typewrite(message)
            sleep(5)
            pyautogui.press("enter")
            sleep(10)
            pyautogui.hotkey("alt", "f4")
            speak("Task completed..")

        elif 'play' in query:
            pyautogui.press("k")
            speak("playing the video")
        elif 'pause' in query or 'stop' in query:
            pyautogui.press("k")
            speak("Video paused")
        elif 'mute' in query:
            pyautogui.press("m")
            speak("video muted")
        elif 'unmute' in query:
            pyautogui.press("m")
            speak("video unmuted")
        elif 'volume up' in query or 'increase the volume' in query:
            print("Turning the volume up...")
            speak("Turning the volume up...")
            volumeup()
        elif 'volume down' in query or 'decrease the volume' in query:
            print("Turning the volume down..")
            speak("Turning the volume down...")
            volumedowm()

        elif '1 tab close' in query or 'one tab close' in query or 'one tab' in query:
            pyautogui.hotkey("ctrl", "w")
            speak("One Tab closed..")
        elif '2 tab close' in query or "two tab close" in query or 'two tab ' in query:
            pyautogui.hotkey("ctrl", "w")
            sleep(0.5)
            pyautogui.hotkey("ctrl", "w")
            speak("Two Tab closed..")
        elif '3 tab close' in query or "three tab close" in query or 'three tab' in query:
            pyautogui.hotkey("ctrl", "w")
            sleep(0.5)
            pyautogui.hotkey("ctrl", "w")
            sleep(0.5)
            pyautogui.hotkey("ctrl", "w")
            speak("Three Tab closed..")

        elif 'close' in query:
            speak("Closing sir..")
            pyautogui.hotkey("alt", "f4")

        elif 'open' in query or 'launch' in query:
            speak("Launching.....")
            query = query.replace("open", "")
            query = query.replace("open the", "")
            query = query.replace("launch", "")
            query = query.replace("Chinchimario open", "")
            pyautogui.click(150, 750, duration=1)
            # pyautogui.press("super")
            pyautogui.sleep(5)
            pyautogui.typewrite(query)
            pyautogui.sleep(2)
            pyautogui.press("enter")


        elif "internet speed" in query:
            wifi = speedtest.Speedtest()
            wifi.get_best_server()
            print("Testing uploading speed...")
            speak("Testing uploading speed...")
            upload_net = wifi.upload() / 10 ** 6
            print("Testing downloading  speed...")
            speak("Testing downloading  speed...")
            download_net = wifi.download() / 10 ** 6
            print(f"Upload Speed: {upload_net:.2f} Mbps")
            speak(f"Upload Speed: {upload_net:.2f} Mbps")
            print(f"Download Speed: {download_net:.2f} Mbps")
            speak(f"Download Speed: {download_net:.2f} Mbps")


        elif 'google' in query:
            chrome()
            speak("Searching from google")
            try:
                query = query.replace("search in google", "")
                webbrowser.get('chrome').open_new_tab(url=f"https://www.google.com/search?q={query}")
            except:
                print("your chrome path is different from developer...........")

        elif 'anime' in query:
            speak("Opening 9anime.to")
            chrome()
            try:
                webbrowser.get('chrome').open_new_tab(url="https://aniwave.to/watch/one-piece.ov8/ep-1072")
            except:
                print("Sorry, The developer needs to learn more coding...... ")

        elif 'drama' in query:
            speak("Sure, sir..")
            chrome()
            try:
                webbrowser.get('chrome').open_new_tab(url="https://kissasian.pe/drama/king-the-land-2023-episode-1")
            except:
                print("Sorry, The developer needs to learn more coding...... ")

        elif 'remember that' in query:
            RememberMessage = query.replace("remember that", "")
            speak(f"you told me to remember that {RememberMessage}")
            speak("Storing your Reminder...")
            print("Storing your Reminder...")
            remember = open("Remember.txt", "w")
            remember.write(RememberMessage)
            remember.close()

        elif 'what i told you to remember ' in query or 'what' in query  :
            try:
                remember = open("Remember.txt", "r")
                speak(f"you told me to remember {remember.read()}")
            except:
                print("You haven't put data to remember")
                speak("You haven't put data to remember")

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
