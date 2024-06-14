import pyttsx3
import speech_recognition as sr
import pyaudio
import wikipedia
import webbrowser
import os
import datetime
import smtplib
import subprocess as sp
import ctypes
import wolframalpha
import tkinter
import pyjokes
import cv2
import mediapipe as mp
from twilio.rest import Client
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import requests
import json
import openai
import calendar
import time

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')

engine.setProperty('voice' , voices[0].id)

openai.api_key = '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'
messages = [ {"role": "system", "content": "You are a intelligent assistant."} ]

def speak(audio):
    
    engine.say(audio)
    engine.runAndWait()

paths = {
    'notepad': "C:\\Program Files\\Notepad\\notepad.exe",
    'calculator': "C:\\Windows\\System32\\calc.exe"
}

def open_camera():
    sp.run('start microsoft.windows.camera:', shell=True)

def open_calculator():
    sp.Popen(paths['calculator'])

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    #server.ehlo()
    server.starttls()
     
    # Enable low security in gmail
    server.login('sg.22u@btech.nitdgp.ac.in', 'P00j@.s@nj@y')
    server.sendmail('sg.22u@btech.nitdgp.ac.in', to, content)
    server.quit()

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1.1
        audio = r.listen(source)

def countdown(t):
	
	while t:
		mins, secs = divmod(t, 60)
		timer = '{:02d}:{:02d}'.format(mins, secs)
		print(timer, end="\r")
		time.sleep(1)
		t -= 1
	speak("times up!")
	print('times up!')

def open_notepad():
    sp.Popen(paths['notepad'])

def open_cmd():
    os.system('start cmd')


def wishMe():
    hour = int(datetime.datetime.now().hour)
    
    if hour>=4 and hour< 12:
        speak("Good Morning sir")
    elif hour>= 12 and hour< 18:
        speak("Good Afternoon sir")
    elif hour>= 18 and hour< 20:
        speak("Good Evening sir")    
    else:
    
        speak("Good Night sir")
        
    speak(f"I am arjun. How may I assist you ?")
  
if __name__ == "__main__":
    clear = lambda: os.system('cls')
     
    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    while True:
        
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1.1
            audio = r.listen(source)
        
        try: 
            print("Recognising...")
            query = r.recognize_google(audio, language='en-in') 
            print(f"You said: {query}\n")
            query=query.lower()
        
            if 'stop' in query:
                speak("Okay sir")
                break
        
            elif 'tell me something about' in query:
                speak("Ok sir")
                speak("Showing results.........")
                query = query.replace("tell me something about", "")
                results = wikipedia.summary(query, sentences=3)
                speak("According to Wikipedia")
                print(results)
                speak(results)
            
            elif 'what is the time' in query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")    
                speak(f"The time is {strTime}")
            
            elif 'open youtube' in query:
                speak("Opening youtube")
                webbrowser.open('https://www.youtube.com/') 

            elif "take" in query:
                speak("taking your snap")
                ec.capture(0, "arjun Camera ", "img.jpg")
               
            elif 'open google' in query:
                speak("cool, opening google")
                webbrowser.open('https://www.google.com/')

            elif 'tracker' in query:
                speak("Initialising augmented tracking interface")
                cap=cv2.VideoCapture(0)

                mpHands= mp.solutions.hands
                hands=mpHands.Hands()
                mpDraw = mp.solutions.drawing_utils
                mpDraw = mp.solutions.drawing_utils
                mpFaceMesh = mp.solutions.face_mesh
                faceMesh = mpFaceMesh.FaceMesh(max_num_faces=2)
                drawSpec = mpDraw.DrawingSpec(thickness=1, circle_radius=1)

                pTime=0
                cTime=0

                while True:
                    success, img= cap.read()
                    imgRGB=cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                    results = hands.process(imgRGB)
                    results1 = faceMesh.process(imgRGB)
    #print(results.multi_hand_landmarks)
                    if results1.multi_face_landmarks:
                        for faceLms in results1.multi_face_landmarks:
                            mpDraw.draw_landmarks(img, faceLms, mpFaceMesh.FACEMESH_CONTOURS,drawSpec,drawSpec)
                            
                    if results.multi_hand_landmarks:
                        for handLms in results.multi_hand_landmarks:
                            for id, lm in enumerate(handLms.landmark):
                #print(id,lm)
                                h,w,c=img.shape
                                cx,cy=int(lm.x*w), int(lm.y*h)
                                print(id, cx, cy)
                                if id==4:
                                    cv2.circle(img, (cx,cy),10 ,(95,0,255),cv2.FILLED)
                            mpDraw.draw_landmarks(img, handLms, mpHands.HAND_CONNECTIONS)

                    cTime=time.time()
                    fps=1/(cTime-pTime)
                    pTime=cTime

                    cv2.putText(img,str(int(fps)),(10,70), cv2.FONT_HERSHEY_PLAIN,3,(95,10,255),3)



                    cv2.imshow("Image",img)
                    cv2.waitKey(1)

            elif 'email to bhai' in query:
                try:
                    print("What should I say?")
                    speak("What should I say?")
                    content = takeCommand()
                    to = "sanchaygupta.72@gmail.com"   
                    sendEmail(to, content)
                    speak("Email has been sent !")
                except Exception as e:
                    print(e)
                    speak("I am not able to send this email")
                
            elif 'news' in query:
             
                try:
                    query_params = {
                        "source": "bbc-news",
                        "sortBy": "top",
                        "apiKey": "feb3fa7ed27c4836b9dbdbdfcf2de905"
                    }
                    main_url = " https://newsapi.org/v1/articles"
 
    # fetching data in json format
                    res = requests.get(main_url, params=query_params)
                    open_bbc_page = res.json()
                     
                        # getting all articles in a string article
                    article = open_bbc_page["articles"]
                     
                
                    results = []
                         
                    for ar in article:
                        results.append(ar["title"])
                             
                    for i in range(len(results)):
                             
                            # printing all trending news
                        print(i + 1, results[i])
                     
                        #to read the news out loud for us
                    from win32com.client import Dispatch
                    speak = Dispatch("SAPI.Spvoice")
                    speak.Speak(results)
                except Exception as e:
                 
                    print(str(e))
 
            elif 'send a mail' in query:
                try:
                    speak("What should I say?")
                    content = takeCommand()
                    speak("to whom should i send")
                    to = input()   
                    sendEmail(to, content)
                    speak("Email has been sent !")
                except Exception as e:
                    print(e)
                    speak("I am not able to send this email")

            elif 'joke' in query:
                speak(pyjokes.get_joke())

            elif 'show me the calendar of year' in query:
            
                query = query.replace("show me the calendar of year", "")
                speak('Loading calendar of year' + query)
                print (calendar.calendar(query))
             
            elif "calculate" in query:
             
                app_id = "XQG6QJ-28EW8E75RR"
                client = wolframalpha.Client(app_id)
                indx = query.lower().split().index('calculate')
                query = query.split()[indx + 1:]
                res = client.query(' '.join(query))
                answer = next(res.results).text
                print("The answer is " + answer)
                speak("The answer is " + answer)
            
            elif 'chat gpt' in query:
                speak("connecting you to open-ai")
                while True:
	                message = input("User : ")
	                try:
		                messages.append({"role": "user", "content": message},)
		                chat = openai.ChatCompletion.create(model="gpt-3.5-turbo", messages=messages)
		                reply = chat.choices[0].message.content
		                print(f"ChatGPT: {reply}")
		                messages.append({"role": "assistant", "content": reply})
	                except Exception as e:print(e)


            elif 'lock windows' in query or "lock my pc" in query or "lock the device" in query:
                speak("locking the device master")
                ctypes.windll.user32.LockWorkStation()  

            elif "hibernate" in query or "sleep" in query:
                speak("Hibernating")
                sp.call("shutdown / h")

            elif "restart" in query or "restart the device" in query:
                speak("Your device will be restarting within a minute!")
                sp.call(["shutdown", "/r"])
            
            elif 'open Matlab' in query:
                speak("Opening matlab")
                codePath = "C:\\Program Files\\MATLAB\\R2022b\\bin\\matlab.exe"   
                os.startfile(codePath) 

            elif 'search' in query or 'play' in query:
                query = query.replace("search", "")
                query = query.replace("play", "")
                speak("Showing results")         
                webbrowser.open(webbrowser.open("https://www.google.com/search?q=" + query + ""))
        
            elif 'change background' in query:
                ctypes.windll.user32.SystemParametersInfoW(20,0,"Location of wallpaper",0)
                speak("Background changed successfully")

            elif 'empty recycle bin' in query:
                winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
                speak("Recycle Bin emptied")

            elif "what is the weather" in query or "how is the weather" in query:
                         
                    try:    # Google Open weather website
                        # to get API of Open weather
                        query=query.replace("What is the weather of" + "")
                        query=query.replace("How is the weather of" + "")
                        api_key = "d3971a59e2d2415e6be8b06e93932f4b"
                        base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
                        speak("The weather at" + query)
                        print("The weather at" + query)
                        complete_url = base_url + "appid =" + api_key + "&q =" + city_name
                        response = requests.get(complete_url)
                        x = response.json()
                         
                        if x["code"] != "404":
                            y = x["main"]
                            current_temperature = y["temp"]
                            current_pressure = y["pressure"]
                            current_humidiy = y["humidity"]
                            z = x["weather"]
                            weather_description = z[0]["description"]
                            print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
                         
                        else:
                            speak(" City Not Found ")

                    except Exception as e:
                 
                        print(str(e))    
           
            elif "what's your name" in query or "What is your name" in query:
                speak("My sir call me")
                speak("arjun")
                print("My master call me arjun")

            elif "who made you" in query or "who created you" in query:
                speak("I have been created by Sanyam.")

            elif 'open aio' in query:
                speak("Opening A I O")
                codePath = "C:\demo\AIO.py"
                os.startfile(codePath)

            elif 'make a folder with name' in query or 'create a folder with name' in query:
        
                query = query.replace("make a folder with name", "")
                query = query.replace("create a folder with name", "")
                speak("creating a folder named" + query + "")
                os.mkdir(query)
                speak("folder created")

            elif "write a note" in query:
                print("What should I write, sir")
                speak("What should i write, sir")
                note = takeCommand()
                file = open('master.txt', 'w')
                print("master, Should I include date and time")
                speak("master, Should i include date and time")
                snfm = takeCommand()
                if 'yes' in snfm or 'sure' in snfm:
                    strTime = datetime.datetime.now().strftime("%H:%M:%S")
                    file.write(strTime)
                    file.write(" :- ")
                    file.write(note)
                else:
                    file.write(note)
         
            elif "show notes" in query:
                speak("Showing Notes")
                file = open("master.txt", "r")
                print(file.read())
                t=file.read()
                speak(t)
        
            elif 'open camera' in query:
                speak("Cool, opening camera.")
                open_camera()

            elif 'open calculator' in query:
                speak("Cool, I'm opening calculator.")
                open_calculator() 

            elif "who i am" in query:
                speak("If you talk then definitely your human.")
 
            elif "why you came to world" in query:
                speak("Thanks to my master. further It's a secret")

            elif 'open cmd' in query:
                speak("opening command center.")
                open_cmd()
        
            elif "where is" in query:
                query = query.replace("where is", "")
                location = query
                speak("locating" + query)
                speak(query + "located")
                webbrowser.open("https://www.google.com/maps/place/" + location + "")

            elif 'open my linkedin profile' in query:
                speak("cool,Opening your linkedin profile")
                webbrowser.open('https://www.linkedin.com/in/sanyam-gupta-7b7624258/')

            elif 'open my github profile' in query:
                speak("Just a second sir.")
                speak("Opening github")
                webbrowser.open('https://github.com/mehisanyam/codeit')

            elif 'start a timer for'in query:
                query=query.replace("start a timer for", "")
                query=query.replace("seconds", "")
                t = query
                speak("starting the timer now!")
                countdown(int(t))
        
            elif 'speak what i write' in query:
                speak("Okay, go on")
                while True:
                    inp=input("SAY: ")
                    speak(inp)
                    if (inp=="stop now"):
                        speak("Okay sir")
                        break

            else:
                engine.say("Sorry, I didn't understand that.")
                engine.runAndWait()
        
        except Exception as e:
            print("Sorry, I could not understand. Could you please say that again?")  
            engine.runAndWait()
            
        