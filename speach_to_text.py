import speech_recognition as sr
from win32com.client import constants, Dispatch

r = sr.Recognizer()

with sr.Microphone() as mp:
    print('say')
    audio = r.listen(mp)
    
try:
    print(r.recognize_google(audio))

except:
    pass

Msg = r.recognize_google(audio)
speaker = Dispatch("SAPI.SpVoice")
speaker.Speak(Msg)
del speaker