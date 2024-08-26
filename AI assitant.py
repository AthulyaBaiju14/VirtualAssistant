#Virtual AI assistant.
from sys import setswitchinterval
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import smtplib
import ctypes
import pywhatkit as kit
import win32com.client as winc1

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak(" Hello! Good Morning!")

    elif hour >= 12 and hour < 18:
        speak(" Hello! Good Afternoon!")

    else:
        speak(" Hello! Good Evening!")

    speak("I am Friday! Please tell me your authentication code for your identification.")
      

def takeCommand():

    r = sr.Recognizer()

    try:
        with sr.Microphone() as source:
            print('Listening.....')
            r.pause_threshold = 1
            r.energy_threshold = 300

            audio = r.listen(source)

        print("Recognizing....")
        query = r.recognize_google(audio, language="en-in")
        print(f"User said: {query}\n") 
    except sr.RequestError as e:
         print(f"Could not request results; {e}")
         return "None"
    except sr.UnknownValueError:
         print("Unknown error occured during recognition; please try again.")
         return "None"
    except Exception as e:
        print(f"An error has occured: {e}")
        return "None"
    return query.lower()
    

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('athulyabaijuacharya@gmail.com','Aimld@01')
    server.sendmail('athulyabaijuacharya@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
    wishMe()
    query = takeCommand()

    if '0101' in query:
            speak("Authentication successful. Please tell me how can I help you?")
    else:
      speak("Sorry!Invalid code.")
      speak("I will have to Shutdown the Laptop according to the Security Protocols.")
      try:
        ctypes.windll.user32.LockWorkStation()
      except Exception as e:
        print(f"Could not lock the workstation: {e}")
      exit()               
             
    while True:
        query = takeCommand()

                
        #wikipedia information.
        if 'wikipedia' in query:
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        #To open Youtube.
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.co.in")  
             
        #To open email inbox.
        elif 'open my mailbox' in query:
            webbrowser.open("gmail.com")

        #Any random youtube video. 
        elif 'ronaldo bicycle kick'in query:
            kit.playonyt("CRISTIANO RONALDO OVERHEAD KICK FROM ALL ANGLES!! #GoalOfTheSeason")
        #To play music.  
        elif 'play music' in query:
            music_dir = r"C:\\Users\\hp\\Music\\music_dir"
            if os.path.exists(music_dir):
                songs = os.listdir(music_dir)
                if songs:
                     os.startfile(os.path.join(music_dir, songs[0]))
                else:
                     speak("No music files found in the directory")
            else:
                 speak("Music directory not found")
            

        elif 'open code' in query:
            codePath = "C:\\Users\\hp\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            if os.path.exists(codePath):
                 os.startfile(codePath)
            else:
                 speak("VS Code is not installed in the specified path")
       
       #To know the current time.
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Mam, the time is {strTime}")

        #To run any system application/software.
        
        elif 'open fifa' in query:
            fifaPath = "D:\\DOCUMENTS\\FIFA 19\\FIFA 19\\FIFA19.exe"
            if os.path.exists(fifaPath):
                 os.startfile(fifaPath)
            else:
                 speak("FIFA is not installed in the specified path")
      
       #Sending email.  
        elif 'send an email' in query:
            try:
                speak("What should I send?")
                content = takeCommand()
                to = "sijesh7achari@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent successfully!")
            except Exception as e:
                print (e)
                speak("Sorry Mam! I am not able to send the email due to network issue! Please check your connection! and try again!")
       
       #Conversation with Friday.
        elif 'hello' in query:
            speak('Hello! Mam')
        elif 'how are you' in query:
            speak("I am doing great! As always! What about you Mam?")
        elif any (word in query for word in ['sad', 'not ok']):
                speak("What happened Mam? Tell me how can I solve your problem?!I can tell you a joke to swing your mood!")
        elif 'yes' in query:
                speak("During a job interview at the 99 Cents store, my son was asked, “Where do you see yourself in five years?” My son’s reply: “At the Dollar Store!”HA!HA!HA!!I hope I tried my level best to help you mam!")        
        elif 'ok' in query:
               speak("During a job interview at the 99 Cents store, my son was asked, “Where do you see yourself in five years?” My son’s reply: “At the Dollar Store!”HA!HA!HA!!I hope I tried my level best to help you mam!")  
        elif 'no friday' in query:
                speak("ok mam, I hope I tried my level best to help you!")
        elif 'great' in query:
                speak("Good to hear that you are doing great! mam")
        elif 'i am fine' in query:
                speak("Good to hear that you are doing great! mam")
        elif 'thank you' in query:
             speak("Your Welcome! Mam")
        elif 'tell me about' in query:
            speak("Shelda  is a bullshit nothing else!")     
        

# about or infomation
        elif "about artificial intelligence" in query:
            try:
                 with open("defAI.txt") as f:
                  content = f.read()
                  speak(content)
            except FileNotFoundError:
                 speak("I couldn't fint the file containing the information.")

                  
        elif "about yourself" in query:
            try:
                 with open("about friday.txt") as f:
                  content = f.read()
                  speak(content) 
            except FileNotFoundError:
                speak("I couldn't fint the file containing the information.")

       #Joke 
        elif 'joke' in query:
            speak("During a job interview at the 99 Cents store, my son was asked, “Where do you see yourself in five years?” My son’s reply: “At the Dollar Store!”HA!HA!HA!!")         
       #To end or deactivate friday
        elif 'deactivate' in query:
            speak("Mam, once again confirm your permission to deactivate!")
        elif '2000' in query:
            speak("Deactivating! Goodbye Mam!")
            break
