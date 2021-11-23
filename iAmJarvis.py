#JARVIS V4.0 10:20AM 19/11/21

import time
import random
from win32com.client import Dispatch


#Text To Speech Engine
def speak(text):
    speak =Dispatch("SAPI.Spvoice")
    speak.Speak(text)

# 1. Wish According to time
def wishMe():
    mor = ["Good Morning Sir", "A very good morning  sir", "Morning Sir, Have A good day.", "Ready to rock for the new day sir", "Hi Sir, Good Morning"]
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

