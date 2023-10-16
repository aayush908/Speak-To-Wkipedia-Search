from win32com.client import Dispatch
import datetime 
import speech_recognition as sr
import pyaudio
import wikipedia
import webbrowser
import os
import smtplib
#import sounddevice
#voices=engine.getProperty("voices")
#print(voices[0].id)
#engine.setProperty("voices",voices[0].id)

def spk(audio):
    dis=Dispatch("SAPI.SpVoice")
    dis.Speak(audio)
  
def wish():
    hr=int(datetime.datetime.now().hour)
    if hr>=0 and hr<12:
        spk("Good MORNING SIR!")
    elif hr>=12 and hr<18:
        spk("Good afternoon SIR!")
    else:
        spk("Good evening sir!")
        
    spk("I AM Aayush kapil AI !HOW MAY I HELP buddy")
    
    
    
def take():
    """it takes audio from microphone and returns string """
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source,duration=1)
        audio=r.listen(source)
    try:
        print("Recognizing...")
        query=r.recognize_google(audio,language="en-in")
        print(f"USER SAID:{query}")
    except Exception as e:
        print("Say that again plz")
        return "NONE"
    return query

#calling functions
wish()

while 1:
    query=take().lower()
    
    
    if "wikipedia" in query:
        spk("searching please wait sir")
        query=query.replace("wikipedia","")
        results=wikipedia.summary(query,sentences=3)
        print(results)
        spk("According to wikipedia")
        spk(results)
        
    elif "open youtube" in query:
        webbrowser.open("youtube.com")
        
    elif "open google" in query:
        webbrowser.open("google.com")
        
  
    elif "time" in query:
        strt=datetime.datetime.now().strftime("%H:%M:%S")
        print(strt)
        spk(strt)
        
    elif "made" in query:
        audio = "Aayush kapil ! created  me    let me show you  his  profile "
        spk(audio)
        webbrowser.open("file:///C:/Users/hp/Desktop/python/aayush's%20Resume.pdf")
        
    elif "banaya hai" in query:
        audio = "Aayush kapil ! created  me    let me show you  his  profile  "
        spk(audio)
        webbrowser.open("file:///C:/Users/hp/Desktop/python/aayush's%20Resume.pdf") 

    elif "open website" in query:
        webbrowser.open("tinyurl.com/herobots")


    elif "close" in query:
        spk("THANKS SIR! FOR USING ME NOW I AM GOING to my home back!!")
        break
        
    elif "how are you" in query:
        spk("i am fine sir u tell how may i help u!!")
    
    elif "bad" in query:
        spk("sorry sir for it!I WILL IMPROVE IT FROM NEXT TIME")