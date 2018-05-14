import win32com.client as wincl
import speech_recognition as sr
import webbrowser as wb

speak = wincl.Dispatch("SAPI.SpVoice")

r = sr.Recognizer()
with sr.Microphone() as source:
    speak.Speak("Hi Comp Sci A period class, what video should we watch?")
    print("Listening...")
    audio = r.listen(source)
    print("Thinking...")


try:
    words = r.recognize_google(audio)
    speak.Speak("Ok, Michaela, let's look for " + r.recognize_google(audio))
    wb.open("https://www.youtube.com/watch?v=3OYtjXVilsA" + words)

except sr.UnknownValueError:
    print("Google speech recognition dose not understand audio")
except sr.RequestError as e:
    print("Couldn't connect to internet.")
