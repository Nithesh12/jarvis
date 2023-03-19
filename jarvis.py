from email import message 
from logging import exception, shutdown #used in email
import webbrowser #to import all this packages first we have to install all this modules using cammanand "pip install xyz"
from numpy import number #numpy used for storing arrays 
import pyttsx3   #Converts text to speech
from random import choice
import requests #used in api requests
import speech_recognition as speechRecog  #Recognizes the speech
from datetime import datetime
import os #used in opening desktop application
import subprocess as da  
import wikipedia
import pywhatkit # Used in whatsapp 
import pyjokes
from mailer import Mailer
from bs4 import BeautifulSoup
import warnings

#Code for Jarvis voice
engine=pyttsx3.init('sapi5')  #init: get engine instance for speech synthesis , sapi5:microsoft driver
voice = engine.getProperty('voices') #getting voices of microsft default narrators
rate = engine.getProperty('rate') #speed of voice
volume = engine.getProperty('volume') 
engine.setProperty('voice',voice[0].id)#setting male voice
engine.setProperty('rate',150)
engine.setProperty('volume',1)

#Function to output audio
def speak(audio):
    engine.say(audio)#asking engine to convert given to speech
    engine.runAndWait()#process the voice command


# Function to greet
def greet():
    time=datetime.now().hour #getting exact current time
    if(time>=6 and time<12 ):
        print("Good Morning Sir")
        speak("Good Morning Sir")
    elif(time>=12 and time<=15):
        print("Good Afternoon sir")
        speak("Good Afternoon sir")
    else:
        print("Good Evening Sir")
        speak("Good Evening Sir")
    print("I'm your Jarvis")
    speak("I am your Jarvis how may i assist you?")
    print("Sir Im a beginner :} please speak after you see (listening...) in the display")
    # speak("Sir speak after you see listening in the display")

#function to convert given command to text so that we can use this text for processing
def userCommand():
    r=speechRecog.Recognizer()
    with speechRecog.Microphone() as microphone:#microphone as source
        print("Listening to you Sir!!!.....")
        r.adjust_for_ambient_noise(microphone,0.5)
        r.pause_threshold=1 #pausing 
        audio = r.listen(microphone) #listens users input and store it in audio
        print("recognizing...")
    try:
            query=r.recognize_google(audio) #converting speech to text and storing it in query
            if not "exit" in query: #if we donot command exit it will enter here
                
                 print(" \n",query)
                #  speak(query) #text that we converted with a help of speechrecognizer is converted to speech here
            else:
                print("Have a wonderful day Sir :-)") #if the command is exit it enters here
                speak("Have a wonderful day Sir")
                exit() #exit from code
            
           
    except Exception:  #if speech recognizer unable to recognize
        print("Can You Please Repeat That :-)")
        # speak("Can You Please Repeat That sir")
        return "None"

    return query

#PowerPoint
def powerpoint():
    path="C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"  #path of TARGET #two back slashes because one backslash we act as escape in python
    os.startfile(path) #opening the extension file using path (opening web application)
 
#VisualStudio
def visualstudio():
    path="C:\\Users\\neeth\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
    os.startfile(path)

#MSword
def msword():
    path="C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
    os.startfile(path)

#Brave
def brave():
    path="C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"
    os.startfile(path)

#wikipedia

def wiki():
    print("What should I search for you sir?")
    speak("What should i search for you sir")
    query=userCommand().lower() #command is converted to text which will be having all lower case(this is because in main function we used query==lowercaseof command (you can refer that later))
    while(query=="none"):
        query=userCommand().lower()
    print("working on it...")
    results=wikipedia.summary(query,sentences=2,auto_suggest=False)#summary function in wekipedia module helps in providing short summary
    print(results)
    speak(results)
        

def youtube():
    print("What should I search for you sir?")
    speak("What should i search for you sir")
    query=userCommand().lower()
    while(query=="none"):
     query=userCommand().lower()
    print("working on it...")
    pywhatkit.playonyt(query)#playonyt function in  pywhatkit module helps in opening youtube and playing command 

def google():
    print("What should I search for you sir?")
    speak("What should i search for you sir")
    query=userCommand().lower()
    while(query=="none"):
        query=userCommand().lower()
    print("working on it...")
    pywhatkit.search(query)


def weather():
    print("Weather of which city sir")
    speak("Weather of which city sir")
    query=userCommand().lower()
    while(query=="none"):
     query=userCommand().lower()   
    soup=BeautifulSoup(requests.get(f"https://www.google.com/search?q=weather+in+{query}").text,"html.parser")  #webscrapping
    region=soup.find("span",class_="BNeawe tAd8D AP7Wnd")
    day=soup.find("div",class_="BNeawe tAd8D AP7Wnd")
    temp=soup.find("div",class_="BNeawe iBp4i AP7Wnd")
    weathe=day.text.split("m",1)
    temper=temp.text.split("C",1)
    print("its"+weathe[1]+temper[0]+"celcius"+"\tin\t"+""+region.text)
    speak("its"+weathe[1]+temper[0]+"celcius"+"in"+region.text)
  
    

def news():
    print("Here is the latest news sir")
    speak("Here is the latest news sir")
    result=pywhatkit.search("latest news")

def whatsapp(number,message):
    print("Working on it..")
    pywhatkit.sendwhatmsg_instantly(f"+91{number}",message)#sendwhatmsg_instantly function insantly sends message

def joke():
 jok = pyjokes.get_joke(language="en", category="neutral") #pyjokes to get jokes
 print(jok)
 speak(jok)
 path="C:\\Users\\neeth\\Downloads\\Crowd Laughing Notification.mp3"
 os.startfile(path)

def advice():
    result = requests.get("https://api.adviceslip.com/advice").json() #api used to get random advice from given address
    adv= result['slip']['advice'] #referring random advice
    print(adv)
    speak(adv)


def email():
    mail = Mailer(email="neethueditbucks@gmail.com", password="zezycddxaorcrhfb") #function helps in stroing address and password and processing it
    speak("Enter recievers email address")
    recievers=input("Enter recievers email address")
    print("subject\n")
    speak("What is the subject sir")
    subjects=userCommand().capitalize()
    while(subjects=="None"):
        subjects=userCommand().capitalize()
    print("message\n")
    speak("what is the message sir")
    messages=userCommand().capitalize()
    while(messages=="None"):
        messages=userCommand().capitalize()
    print("working on it...")
    try:
     mail.send(receiver=recievers,subject=subjects,message=messages) #sending the message
     print("email sent")
     speak("Email sent")
    except Exception: 
        print("Error occured")

#Main function
if __name__=="__main__": # its a main function
    greet()

    while True:
        query = userCommand().lower()
        if "open powerpoint" in query: #if query (text) is equal to open power point it executes powerpoint function
            speak("sure sir")
            print("Opening PowerPoint")
            powerpoint()
            speak("Opening PowerPoint")
        elif "open visual" in query:
            speak("definetly")
            print("Opening Visual Studio")
            visualstudio()
            speak("Opening Visual Studio")
        elif "open ms" in query:
            speak("sure")
            print("Opening MSword")
            msword()
            speak("Opening MSword")
        elif "open browser" in query:
            speak("why not")
            print("Opening Brave")
            brave()
            speak("Opening Brave")
        elif "wikipedia" in query: #query == wikipedia or query == Wkipedia
            speak("almost there")
            wiki()
        elif "youtube" in query:
            speak("alright")
            youtube()
        elif "google" in query:
            speak("sure")
            google()
        elif "weather" in query:
            weather()
        elif "news" in query:
            news()
        elif "whatsapp" in query:
            speak("Please enter the number sir")
            number=input("Enter the number\n")
            print("what's the message sir?")             
            speak("what's the message sir")             
            message=userCommand().capitalize()
            while(message=="None"):
                message=userCommand().capitalize()
            whatsapp(number,message)
            print("Sent the message sir")
            speak("Sent the message sir")
  
        elif "joke" in query:
            speak("you may like this")
            joke()
        elif "advice" in query:
            speak("alright")
            advice()
        elif "email" in query:
            email()
        elif "shutdown" in query:
            shutdown=input("Should I shutdown your pc sir ?  (y/n)")
            if shutdown=='n':
                speak("Thank You Sir")
                exit()
            else:
             speak("Thank you sir have a wonderful day")
             os.system("C:\Windows\SysWOW64\shutdown.exe /s")
        elif "log out" in query:
            speak("logging out sir")
            os.system("C:\Windows\SysWOW64\shutdown.exe /l")
        
    
    
    