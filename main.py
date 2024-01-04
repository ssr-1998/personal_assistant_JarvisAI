import speech_recognition as sr
import winsound
import os
import webbrowser
import pyaudio as pa
import sys
import re
import win32com.client
import warnings
import datetime

coder = "Shubham"
warnings.filterwarnings("ignore")
speaker = win32com.client.Dispatch("SAPI.SpVoice")
sites = {
    "youtube music": "https://music.youtube.com/",
    "youtube": "https://www.youtube.com", 
    "wikipedia": "https://www.wikipedia.com",
    "google": "https://www.google.com",
    "upgrad": "https://learn.upgrad.com/courses",
    "upgrade": "https://learn.upgrad.com/courses"
    }
apps = {
    "telegram": "D:\S_S_R\Softwares Installed\Telegram Desktop\Telegram.exe",
    "powershell": "C:/Users/hp/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Windows PowerShell/Windows PowerShell.lnk",
    "sublime text": "C:\Program Files\Sublime Text\sublime_text.exe",
    "anaconda prompt": "C:/Users/hp/Desktop/Anaconda Prompt for S.S.R.lnk",
    "control panel": "C:\Windows\System32\control.exe",
    "control panel": "C:\Windows\System32\control.exe",
    "cmd": "C:\Windows\system32\cmd.exe",
    }


def say(text):
    speaker.Speak(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.energy_threshold = 500
        r.pause_threshold = 1
        audio = r.listen(source)
        say("Recognizing Audio Input...")
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            e_traceback = sys.exc_info()[2]
            line_no = str(e_traceback.tb_lineno)
            error_string = " ".join(re.findall('[A-Z][^A-Z]*', str(type(e).__name__)))
            print("{}: {}. Error Occurred at line number ({})".format(error_string, e, line_no))
            if error_string == "Unknown Value Error":
                say("Sorry, I didn't get that. Can you please repeat!")
                return None
            else:
                error_string = "Mr. {}, I appologise for leaving you in the middle but I am encountering an {} at line number {} which is forcing me to close. I'll take your leave now Sir. Good Day!".format(coder, error_string, line_no)
                say(error_string)
                sys.exit(1)


if __name__ == "__main__":
    print("Jarvis is Up & Running!")
    say("Hello, I'm Jarvis, an A.I. powered tool which is still learning & is under development by "
        "Mr. {}.\nHow can I help you Sir?".format(coder))
    try:
        bye = "on"
        while bye == "on":
            winsound.Beep(4000, 200)
            print("Listening...")
            _input = takeCommand()

            if _input == None:
                continue

            _input = _input.lower()

            """
            --- Code for Greetings from Jarvis ---
            """
            if ("hello" in _input or "hey" in _input or "hi" in _input) and "jarvis" in _input:
                if "hi" in _input:
                    if "hi" in list(_input.split()):  # Confirming that `Hi` is an individual word or not.
                        say("Hello, Mr. {}".format(coder))
                elif "hey" in _input:
                    if "hey" in list(_input.split()):  # Confirming that `Hey` is an individual word or not.
                        say("Hello, Mr. {}".format(coder))
                else:
                    say("Hello, Mr. {}".format(coder))

            if "how are you" in _input:
                say("I'm good. Thank you for asking.\nHow are you today?")
            
            if "thank you" in _input or "thank u" in _input:
                say("I'm always happy to help Sir.")

            if "goodbye" in _input or "good bye" in _input or ("goodbye" in _input and "jarvis" in _input):
                say("Goodbye Mr. {}\nHave a Great Day ahead.".format(coder))
                bye = "off"

            """
            --- Code for Opening Sites with the help of Jarvis ---
            """
            if "open" in _input:
                open_site = False
                for i in list(sites.keys()):
                    if i in _input:
                        open_site = True
                        site_name = i
                        break
                open_app = False
                for i in list(apps.keys()):
                    if i in _input:
                        open_app = True
                        app_name = i
                        break
                if open_site:
                    say(f"Opening {site_name} Sir")
                    webbrowser.open_new(sites.get(site_name))
                elif open_app:
                    os.startfile(apps.get(app_name))
                else:
                    print("Open_Site - {}".format(open_site))
                    print("Open_App - {}".format(open_app))
            
            if ("what" in _input or "what's" in _input) and "time" in _input:
                nowtime = datetime.datetime.now().strftime("%I : %M %p")
                say("Sir, it's {}".format(nowtime))
    except Exception as e:
        e_traceback = sys.exc_info()[2]
        line_no = str(e_traceback.tb_lineno)
        error_string = " ".join(re.findall('[A-Z][^A-Z]*', str(type(e).__name__)))
        print("{}: {}. Error Occurred at line number ({})".format(error_string, e, line_no))
        error_string = "Mr. {}, I appologise for leaving you in the middle but I am encountering an {} at line number {} which is forcing me to close. I'll take your leave now Sir. Good Day!".format(coder, error_string, line_no)
        say(error_string)
        sys.exit(1)
