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
from openai import OpenAI

coder = "Shubham"
bot_name = "Jarvis"
warnings.simplefilter("ignore")
speaker = win32com.client.Dispatch("SAPI.SpVoice")
sites = {
    "youtube music": "https://music.youtube.com/",
    "youtube": "https://www.youtube.com", 
    "wikipedia": "https://www.wikipedia.com",
    "google": "https://www.google.com",
    "upgrad": "https://learn.upgrad.com/courses",
    "upgrade": "https://learn.upgrad.com/courses",
    "your development sources": [
        "https://www.youtube.com/watch?v=s_8b5iq4Rvk&list=WL&index=3&t=2610s",
        "https://www.youtube.com/watch?v=7RsS2o7DpK0&list=WL&index=2&t=1254s",
        "https://github.com/BolisettySujith/J.A.R.V.I.S/tree/main"
        ],
    "your github repo": "https://github.com/ssr-1998/personal_assistant_JarvisAI/tree/dev"
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


def open_ai(_prompt):
    client = OpenAI()  # It will automatically search for API Key in Environment Variables as "OPENAI_API_KEY"
    text = f"Prompt : {_prompt}\nResponse : "

    response = client.completions.create(
    model="gpt-3.5-turbo-instruct",
    prompt=_prompt,
    temperature=1,
    max_tokens=256,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=0
    )
    try:
        print(response["choices"][0]["text"])
        text += response["choices"][0]["text"]
    except Exception as e:
        print(f"Unable to understand! Please check the error.\n{e}")
        sys.exit()

    if not os.path.exists("OpenAI_Responses"):
        os.mkdir("OpenAI_Responses")
        n = 1
    else:
        res_no_lst = re.findall("r(\d+).txt", "__".join(os.listdir("OpenAI_Responses")))
        res_no_lst = [int(i) for i in res_no_lst]
        n = max(res_no_lst) + 1

    with open("OpenAI_Responses/r{}.txt".format(n), "w") as f:
        f.write(text)


def say(text):
    speaker.Speak(text)


def takeCommand(_timeout):
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.energy_threshold = 500
        r.pause_threshold = 1
        try:
            audio = r.listen(source, timeout = _timeout)
            say("Recognizing Audio Input...")
            query = r.recognize_google(audio, language = "en-in")
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
            elif error_string == "Wait Timeout Error":
                return None
            else:
                error_string = "Mr. {}, I appologise for leaving you in the middle but I am encountering an {} at line number {} which is forcing me to close. I'll take your leave now Sir. Good Day!".format(coder, error_string, line_no)
                say(error_string)
                sys.exit(1)


if __name__ == "__main__":
    print("{} is Up & Running!".format(bot_name))
    say("Hello, I'm {}, an A.I. powered tool which is still learning & is under development by "
        "Mr. {}.\nHow can I help you Sir?".format(bot_name, coder))
    try:
        bye = "on"
        audio_input_timeout = 60
        while bye == "on":
            winsound.Beep(4000, 200)
            print("Listening... (Audio Input Timeout - {})".format(audio_input_timeout))
            _input = takeCommand(_timeout = audio_input_timeout)

            if _input == None:  # If no input received, then bot will jump to the next iteration of the While Loop i.e. will again start Listening for any possible command.
                continue

            _input = _input.lower()

            if "set" in _input and "audio input timeout" in _input:  # Code to change Audio Input Timeout Seconds using bot.
                sec = re.findall("\s(\d+)\s?\w*", _input)
                if len(sec) > 0:
                    if sec[0].isnumeric():
                        audio_input_timeout = int(sec[0])

            """
            --- Code for Greetings from Bot ---
            """
            if ("hello" in _input or "hey" in _input or "hi" in _input) and bot_name in _input:
                if "hi" in _input:
                    if "hi" in list(_input.split()):  # Confirming that `Hi` is an individual word or not.
                        say("Hello Mr. {}".format(coder))
                elif "hey" in _input:
                    if "hey" in list(_input.split()):  # Confirming that `Hey` is an individual word or not.
                        say("Hello Mr. {}".format(coder))
                else:
                    say("Hello Mr. {}".format(coder))

            if "how are you" in _input:
                say("I'm good. Thank you for asking.\nHow are you today?")
            
            if "thank you" in _input or "thank u" in _input:
                say("I'm always happy to help Sir.")

            if "goodbye" in _input or "good bye" in _input or ("goodbye" in _input and bot_name in _input):
                say("Goodbye Mr. {}\nHave a Great Day ahead.".format(coder))
                bye = "off"

            """
            --- Code for Opening Sites with the help of Bot ---
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
                    if site_name == "your development sources":
                        say("Opening all my development sources")
                        for dSources in sites.get(site_name):
                            webbrowser.open_new(dSources)
                    else:
                        if site_name == "your github repo":
                            say("Opening my Github Repo")
                        else:
                            say(f"Opening {site_name} Sir")
                        webbrowser.open_new(sites.get(site_name))
                elif open_app:
                    say(f"Opening {app_name} Sir")
                    os.startfile(apps.get(app_name))
                else:
                    print("Open_Site - {}".format(open_site))
                    print("Open_App - {}".format(open_app))
            
            if ("what" in _input or "what's" in _input) and "time" in _input:
                nowtime = datetime.datetime.now().strftime("%I : %M %p")
                say("Sir, it's {}".format(nowtime))
            
            if ("use" in _input or "using" in _input) and "ai" in _input:
                ai_res = open_ai(_input)
                say(ai_res)

    except Exception as e:
        e_traceback = sys.exc_info()[2]
        line_no = str(e_traceback.tb_lineno)
        error_string = " ".join(re.findall('[A-Z][^A-Z]*', str(type(e).__name__)))
        print("{}: {}. Error Occurred at line number ({})".format(error_string, e, line_no))
        error_string = "Mr. {}, I appologise for leaving you in the middle but I am encountering an {} at line number {} which is forcing me to close. I'll take your leave now Sir. Good Day!".format(coder, error_string, line_no)
        say(error_string)
        sys.exit(1)
