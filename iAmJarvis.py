#JARVIS V4.0 10:20AM 19/11/21

import time
import random
from win32com.client import Dispatch
import urllib #2
import requests#15
import speech_recognition as sr # Voice_listner
import os #4
import wikipedia#19
import webbrowser#13
import google #13
import datetime
from googlesearch import search#13


#Text To Speech Engine
def speak(text):
    speak =Dispatch("SAPI.Spvoice")
    speak.Speak(text)

# 1. Wish According to time
def wishMe():
    mor = ["Good Morning Sir , Let's Go To Work", "A very good morning  sir", "Morning Sir, Have A good day.", "Ready to rock for the new day sir", "Hi Sir, Good Morning"]
    aft = ["Good Afternoon Sir", "Good Afternoon Sir", "A very good afternoon sir from your JARVIS", "A very good afternoon sir from jarvis"]
    eve = ["Good evening sir", "Good evening sir", "A very good evening sir", "Hello sir, Good evening"]

    time_hr = int(datetime.datetime.now().strftime("%H"))

    if time_hr > 8 and time_hr < 10:
        speak(random.choice(mor))
    elif time_hr > 22 and time_hr < 7:
        print("SSSHHHHHHHH.......Somebody is sleeping")
    elif time_hr > 10 and time_hr < 16:
        speak(random.choice(aft))
    elif time_hr > 16 and time_hr < 19:
        speak(random.choice(eve))

# 2. Check Internet Connection
def checkInternet(host='http://google.com'):
    try:
        urllib.request.urlopen(host)
        return True
    except:
        return False

# Voice listner
def listen():
    r = sr.Recognizer()
    r.energy_threshold = 5000.00
    with sr.Microphone() as mp:
        audio = r.listen(mp)
        try:
            voice_data = r.recognize_google(audio)
        except:
            voice_data = ""
    return voice_data

#14 Shut Down Computer
def shutdownCom():
    speak("Ok sir! Computer is shutting Down Now!")
    os.system("shutdown /s /t 1")

def sleepCom():
    speak("Ok sir! I going to sleep")
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

def JarvisIntro():
    speak('Sir I am your Just a rather very intelligent system....You can call me JARVIS')
    speak("Do you Want More Information About Me")
    query3 = input("Yes/No:\n")
    query3 = query3.lower()
    if query3 == 'yes':
        print("Here is my full detail")
        speak("Here is my full detail")
        print("I am Jarvis version 2 point 0 . I am created by Dhruv Sir")
        print("I am Jarvis version 2 point 0 . I am created by Dhruv Sir")
        speak("I am written in python")
        print("I am written in python")
        speak("My BirthDate is 1 May 2020")
        print("My BirthDate is 1 May 2020")
        speak("I am Written in Python in 300+ lines of Code")
        print("I am Written in Python in 300+ lines of Code")
        speak("I am currently in version 1 point 0 ")
        print("I am currently in version 21 point 0 ")
        speak("That's all ")
        print("That's all ")
        speak("And a last thing i am your Friend")
        print("And a last thing i am your Freind")
    else:
        speak("Ok sir.")



#15 Tells News
def news_of_india():
    """This Function Fetch Top Live News headlines from Google News"""
    try:
        # Google news api
        main_url = "http://newsapi.org/v2/top-headlines?sources=google-news-in&apiKey=2fe749e4f1b542fbb92b085933804760"

        # fetching data in json format
        open_google_page = requests.get(main_url).json()

        # getting all articles in a string article
        article = open_google_page["articles"]

        # empty list which will
        # contain all trending news
        results = []

        for ar in article:
            results.append(ar["title"])
        speak("Ok sir, Getting News from Web..Please Wait")
        speak("Sir I got the News.. Here is the top headlines...")
        for i in range(len(results)):
            # printing all trending news
            Headline = i + 1, results[i]
            print(Headline)
            speak(Headline)
    except Exception as ConnectionError:
        speak("Sorry sir, i got some error on fetching news from web")
        print(ConnectionError)

#19 Wikipedia Info Fetcher
def search_on_wikipedia(topic):
    """This function is fecthing data from wikipedia on any topic. ex:- Ramamyan on wikipedia"""
    
    try:
        speak(f'Searcing {topic} in wikipedia...')
        results = wikipedia.summary(topic, sentences=2)
        speak("Sir...According To wikipedia")
        print(results)
        speak(results)
    except Exception as ConnectionError:
         speak("Sorry Sir.....Can't Reach wikipedia,Maybe your connection was lost")
         print(ConnectionError)

#20 Security Function (RUN AT YOUR OWN RISK)
def security():
    """This function is used to delete all the code if and only if 'Computer Name' & 'Face' did not match to it's
    requirements """
    speak("You are not authorized to use JARVIS.")
    speak("Entering Self Destruction Mode....")
    speak("Earsing JARVIS v2.0 in...3....2......1.....")
    os.remove(sys.argv[0])
    quit()
    

#13
def TellTime():
    strTime = datetime.datetime.now().strftime("%H:%M %p")
    print(f"Sir,The time is {strTime}")
    speak(f"Sir,The time is {strTime}")
