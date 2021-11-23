# PYWAKE_FRIEND
from win32com.client import Dispatch
import datetime
import random
# TTS
def speak(text):
    speak =Dispatch("SAPI.Spvoice")
    speak.Speak(text)


mor = ["Good Morning Sir", "A very good morning dhruv sir", "Morning Sir, Have A good day.", "Ready to rock for the new day sir", "Hi Dhruv Sir, Good Morning"]
aft = ["Good Afternoon Dhruv Sir", "Good Afternoon Sir", "A very good afternoon Dhruv sir from your JARVIS", "A very good afternoon sir from jarvis"]
eve = ["Good evening sir", "Good evening dhruv sir", "A very good evening sir", "Hello sir, Good evening"]

time_hr = int(datetime.datetime.now().strftime("%H"))

if time_hr > 8 and time_hr < 10:
    speak(random.choice(mor))
elif time_hr > 22 and time_hr < 7:
    print("SSSHHHHHHHH.......Somebody is sleeping")
elif time_hr > 10 and time_hr < 16:
    speak(random.choice(aft))
elif time_hr > 16 and time_hr < 19:
    speak(random.choice(eve))

