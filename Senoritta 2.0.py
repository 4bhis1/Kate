import pyttsx3
import datetime
import socket
import win32api,win32con
import speech_recognition as sr

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

def wikipedia_search(nkb):
    try:
        results = wikipedia.summary(query, sentences=2)
        speak("According to Wikipedia")
        print(results)
        speak(results)
    except:
        speak("Bad request. Error 401")

    

def is_connected():
    try:
        # connect to the host -- tells us if the host is actually
        # reachable
        socket.create_connection(("www.google.com", 80))
        speak("Connected to internet.")
        global internet
        internet = True
    except OSError:
        speak("Sorry sir system is offline.")

k=0

def news(urlo):
    news=requests.get(urlo).text
    news=json.loads(news)
    #print(news["articles"])
    #print(news)
    speak('fetching the news from server')
    print('fetching the news from server')
    arts=news['articles']
    for arti in arts:
        print(arti["title"])
        #tranlation=translator.translate(arti["title"])
        #print(tranlation)
        speak(arti["title"])
        time.sleep(2)
    speak("Over")
    print("khtm")


def play_music():
    music_dir = 'E:\\Music'
    songs = os.listdir(music_dir)
    os.startfile(os.path.join(music_dir, songs[random.randint(0,51)]))

llopk=False

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)    
        print("Say that again please")
        #speak("Say that again please...")  
        return "None"
    return query


    
if __name__ == "__main__":

    while True:
        jklm = takeCommand().lower()
        
        if 'senorita' in jklm and 'on' in jklm:
            #print("i am on");
            llopk=True
            wishMe()

            speak("Checking intenet connection")
            internet = False
            is_connected()  
                
            speak("Connecting to server sir. Wait for a while")
                
            
            from win32com.client import Dispatch
            #from googlesearch import search
            import requests
            import json
            import time
            import speech_recognition as sr 
            from datetime import date as dt
            from datetime import datetime as ddt
            import wikipedia
            import webbrowser
            import os
            import smtplib
            import random

            speak("Conection established.")
                        
            while llopk:
                
                query = takeCommand().lower()
                query=query.replace("senoritta","")
                
                if 'wikipedia' in query:
                    query = query.replace("search", "")
                    query = query.replace("wikipedia", "")
                    wikipedia_search(query)

                elif 'play music'in query:
                    play_music()
                    k=1

                elif 'next' in query and k==1:
                    play_music()
                    
                elif 'time' in query:
                    strTime = datetime.datetime.now().strftime("%H:%M:%S")
                    speak(f"Sir, the time is {strTime}")

                elif 'date' in query:
                    today=dt.today()
                    speak(f"todays date is {today}")

                elif 'day' in query:
                    gk=ddt.today().strftime("%A")
                    speak(f"todays date is {gk}")

                elif 'speak' in query:
                    query = query.replace("speak", "")
                    speak(query)

                elif 'write' in query and 'line' in query:
                    speak("Working on this project.. It will take some time")

                elif 'money' in query and 'write' in query:
                    speak("Working on this project... It will take some time")
                                 

                elif 'spell' in query:
                    query=query.replace("spell","")
                    kl=""
                    for i in range(0,len(query)) :
                        kl=kl+query[i]+" "
                    speak(kl)
                    print(kl)

                elif 'news' in query:
                    try:
                        if 'sports' in query:
                            url="https://newsapi.org/v2/top-headlines?country=in&category=sports&apiKey=d841d6dce86044d6bda182f600f91e73"
                        elif "techno" in query:
                            url="https://newsapi.org/v2/top-headlines?country=in&category=technology&apiKey=d841d6dce86044d6bda182f600f91e73"
                        elif "science" in query:
                            url="https://newsapi.org/v2/top-headlines?country=in&category=science&apiKey=d841d6dce86044d6bda182f600f91e73"
                        else:
                            url="https://newsapi.org/v2/top-headlines?sources=google-news-in&apiKey=d841d6dce86044d6bda182f600f91e73"
                        news(url)
                    except OSError:
                        speak("Sorry sir system is offline.")
                    

                elif 'close' in query:
                    break

                elif query!='':
                    speak('should i search it on wikipedia sir?' )
                    bkl=takeCommand().lower()
                    if 'yes' in bkl:
                        if internet:
                            wikipedia_search(bkl)
                        else:
                            speak("System is offline")
                    #else:
                        #speak('not programmed to do this work, building up myself')
        elif "over" in jklm:
            os.system("shutdown /s /t 1")
