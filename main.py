import win32com.client
import webbrowser
import speech_recognition as sr
import time
timestamp=time.strftime('%H:%M:%S')
hour=int(time.strftime('%H'))
your_voice_index = 1 
speaker=win32com.client.Dispatch("SAPI.SpVoice")
speaker.Voice = speaker.GetVoices().Item(your_voice_index)
def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        audio=r.listen(source)
        try:
            print("Recognizing.....")
            query=r.recognize_google(audio,language="en-in")
            print(f"User said: {query}")
            speaker.speak(query)
            return query
        except Exception as e:
            return "some error occured. Sorry from Friday"



if __name__=='__main__':
    speaker.speak("Hello, I am Friday! Your virtual A I assistant. How can I help you sir?")
    while True:
        print("Listening.....")
        query=takeCommand()
        sites=[["Youtube","https://www.youtube.com"],["Instagram","https://www.instagram.com"],["Github","https://www.github.com"],["Whatsapp","https://www.whatsapp.com"]]
        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                speaker.speak(f"opening {site[0]} sir")
                webbrowser.open(site[1])
        if query.lower()==f"good morning Friday" or query.lower()==f"good afternoon Friday" or query.lower()==f"good evening Friday":
            if(hour>=6 and hour<=12):
                speaker.speak(f"good morning sir, how can I help you")
            elif(hour>=12 and hour<=17):
                speaker.speak(f"good afternoon sir, how can I help you")
            elif(hour>=17 and hour<=20):
                speaker.speak(f"good evening sir, how can I help you")
            elif(hour>=20 and hour<=6):
                speaker.speak(f"good evening sir,how can I help you")
            
        if "good night Friday" in query.lower():
            speaker.speak(f"good night sir, have a nice day ahead")
            break
        
        if "what is the time" in query.lower():
            speaker.speak(f"time is:{timestamp}")
        
        if "hello Friday" in query.lower():
            speaker.speak(f"I'm fine sir, how are you")
        
        
                
        
