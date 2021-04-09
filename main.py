import win32com.client
import speech_recognition as sr
import wikipedia as wiki
import random
import webbrowser as browser
import sys
import os
MY_NAME = "ashraful"
AI_NAME = "jarvis"
EXIT_NOW=""
def speak(myinput):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(myinput)
#Conditions Start
def con():
    try:
        if(AI_NAME == mycommand):
            speak("yes, sir.")
            speak("how can i help you?")
        elif("hi" in mycommand or "hello" in mycommand or "hey" in mycommand):
            speak(random.choice(["Hello, how can i help you?", "Hi, how can i help you?", "Hey, how can i help you?"]))
        elif("how" in mycommand and "are" in mycommand and "you" in mycommand):
            speak("i'm fine. thank you.")
        elif("who" in mycommand and "are" in mycommand and "you" in mycommand):
            speak("I'm your virtual assistant")
        elif("your" in mycommand and "name" in mycommand):
            speak("my name is "+AI_NAME)
        elif("who" in mycommand and AI_NAME in mycommand):
            speak("i'm "+AI_NAME)
        elif("my" in mycommand and "name" in mycommand):
            speak("your name is "+MY_NAME)
        elif("who" in mycommand and MY_NAME in mycommand):
            speak("you are "+MY_NAME)
        elif("what" in mycommand and "can" in mycommand and "you" in mycommand and "do" in mycommand):
            speak("i can do anything what you like.")
        elif("you" in mycommand and "are" in mycommand and "human" in mycommand):
            speak("no.")
            speak("I am an artificial intelligence.")
        elif("who" in mycommand and ("made" in mycommand or "created" in mycommand) and "you" in mycommand):
            speak(MY_NAME+" has Created Me. So, he is my inventor and i'm his virtual assistant.")
        elif(("open" in mycommand or "visit" in mycommand or "lunch" in mycommand) and "facebook" in mycommand):
            speak("Sure sir, Just a momment please.")
            browser.open_new_tab("facebook.com")
        elif(("open" in mycommand or "visit" in mycommand or "lunch" in mycommand) and "google" in mycommand):
            speak("Sure sir, Just a momment please.")
            browser.open_new_tab("google.com")
        elif(("open" in mycommand or "visit" in mycommand or "lunch" in mycommand) and "twitter" in mycommand):
            speak("Sure sir, Just a momment please.")
            browser.open_new_tab("twitter.com")
        elif(("open" in mycommand or "visit" in mycommand or "lunch" in mycommand) and "youtube" in mycommand):
            speak("Sure sir, Just a momment please.")
            browser.open_new_tab("youtube.com")
        elif("exit now" in mycommand or "goodbye jarvis" in mycommand):
            speak("As your wish, sir. Goodbye. Have a nice day")
            global EXIT_NOW
            EXIT_NOW = True
        elif("complex"==mycommand):
            speak("Please ask me clearly.")
        else:
            w = wiki.summary(mycommand, sentences=2)
            speak(w)
    except:
        pass
#Conditions End
def hearme():
    try:
        print("Listing...")
        with sr.Microphone() as source:
            l = r.listen(source)
        try:
            global mycommand
            mycommand = r.recognize_google(l, language="en-GB")
            mycommand = mycommand.lower()
            print(mycommand)
            con()
        except:
            pass
    except:
        print("I haven't understand.")
if __name__=="__main__":
    speak("Hello sir, This is your personal assistant Jarvis.")
    while True:
        if(EXIT_NOW==True):
            exit()
        r = sr.Recognizer()
        r.energy_threshold = 9000
        hearme()
