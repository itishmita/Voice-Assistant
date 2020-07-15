import pyttsx3
import webbrowser
import smtplib
import random
import speech_recognition as sr
import wikipedia
import datetime
import wolframalpha
import os
import sys
import pandas as pd

engine = pyttsx3.init('sapi5')

client = wolframalpha.Client('Your_App_ID')

voices = engine.getProperty('voices')
engine.setProperty('voice', voices[len(voices)-1].id)

# To get audio which is present inside the computer
def speak(audio):
    print('Computer: ' + audio)
    engine.say(audio)
    engine.runAndWait()

# To wish with respect to current time in the computer
def greetMe():
    currentH = int(datetime.datetime.now().hour)
    if currentH >= 0 and currentH < 12:
        speak('Good Morning!')

    if currentH >= 12 and currentH < 18:
        speak('Good Afternoon!')

    if currentH >= 18 and currentH !=0:
        speak('Good Evening!')

greetMe()

speak('Hello Sir, I am your  assistant!')
speak('How may I help you?')

# To take the command of the user in the form of query
def myCommand():
   
    r = sr.Recognizer()                                                                                   
    with sr.Microphone() as source:                                                                       
        print("Listening...")
        r.pause_threshold =  1
        audio = r.listen(source)
    try:
        query = r.recognize_google(audio, language='en-in')
        print('User: ' + query + '\n')
        
    except sr.UnknownValueError:
        speak('Sorry sir! I didn\'t get that! Try typing the command!')
        query = str(input('Command: '))

    return query
        

if __name__ == '__main__':

    while True:
    
        query = myCommand()
        query = query.lower()
        
        if 'open youtube' in query:
            speak('okay')
            webbrowser.open('www.youtube.com')

        elif 'open google' in query:
            speak('okay')
            webbrowser.open('www.google.co.in')

        elif 'open gmail' in query:
            speak('okay')
            webbrowser.open('www.gmail.com')

        elif "what\'s up" in query or 'how are you' in query:
            stMsgs = ['Just doing my thing!', 'I am fine!', 'Nice!', 'I am nice and full of energy']
            speak(random.choice(stMsgs))

        elif 'email' in query:
            speak('Who is the recipient? ')
            recipient = myCommand()
            SenderAddress = "itishmita.mishra98@gmail.com"
            password = "sakuntalamishra"

            if 'me' in recipient:
                try:
                    speak('What is the message? ')
                    msg = myCommand()
                    speak('What is the subject? ')
                    subject = myCommand()
                    body = "Subject: {}\n\n{}".format(subject,msg)
                    e = pd.read_excel("Email.xlsx")
                    emails = e['Emails'].values
        
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.ehlo()
                    server.starttls()
                    server.login(SenderAddress, password)
                    for email in emails:
                       server.sendmail(SenderAddress, email, body)
                    server.close()
                    speak('Email sent!')

                except:
                    speak('Sorry Sir! I am unable to send your message at this moment!')


        elif 'nothing' in query or 'abort' in query or 'stop'in query or 'quit'in query or 'end' in query:
            speak('okay')
            speak('Bye Sir, have a good day.')
            sys.exit()
           
        elif 'hello' in query:
            speak('Hello Sir')

        elif 'bye' in query:
            speak('Bye Sir, have a good day.')
            sys.exit()
                                    
        elif 'play music' in query:
            music_folder = 'E:\\songs'
            music = ['\Dhoom']
            random_music = music_folder + random.choice(music) + '.mp3' 
            os.system(random_music)
            speak('Okay, here is your music! Enjoy!')

        elif 'play movie' in query:
            movie_folder = 'F:\\movies'
            movie = ['\sairat']
            random_movie = movie_folder + random.choice(movie) + '.mkv' 
            os.system(random_movie)
            speak('Okay, here is your movie! Enjoy!')

        elif 'open code' in query:
            codePath = "C:\\Users\\DELL\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'open word' in query:
            codePath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            os.startfile(codePath)

        

        elif 'open notepad' in query:
            codePath = "C:\\Program Files\\Notepad++\\notepad++.exe"
            os.startfile(codePath)

        
        elif 'open powerpoint' in query:
            codePath = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
            os.startfile(codePath)

        elif 'open excel' in query:
            codePath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
            os.startfile(codePath)

        



        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'open facebook' in query:
            webbrowser.open("facebook.com")


            

        else:
            query = query
            speak('Searching...')
            try:
                try:
                    res = client.query(query)
                    results = next(res.results).text
                    speak('WOLFRAM-ALPHA says - ')
                    speak('Got it.')
                    speak(results)
                    
                except:
                    results = wikipedia.summary(query, sentences=2)
                    speak('Got it.')
                    speak('WIKIPEDIA says - ')
                    speak(results)
        
            except:
                webbrowser.open('www.google.com')
        
        speak('Next Command! Sir!')