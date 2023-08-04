import os
import webbrowser
import openai
import win32com.client
import speech_recognition as sr
from config import apikey
import datetime

speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Voice = speaker.GetVoices().item(1)

chatStr = ""
def chat(query):
    openai.api_key = apikey
    global chatStr
    chatStr += f"Nitin: {query}\n April: "
    response = openai.ChatCompletion.create(
        model="text-davinci-001",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # todo: Wrap inside try catch block
    speaker.Speak(response["choices"][0]["text"])
    chatStr += response["choices"][0]["text"]
    return response["choices"][0]["text"]


    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)



def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n ****************************\n\n"

    response = openai.ChatCompletion.create(
        model="text-davinci-001",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # todo: Wrap inside try catch block
    print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip() }.txt", "w") as f:
        f.write(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold = 0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from April"

speaker.Speak("Hello I am April AI")
while 1:
    print("Listening...")
    query = takeCommand()
    sites = [["youtube", "https://youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"]]
    for site in sites:
        if f"Open {site[0]}".lower() in query.lower():
            speaker.Speak(f"Opening {site[0]} sir...")
            webbrowser.open(site[1])

    if "open music" in query:
        musicPath = "C:/Users/Nitin Kadyan/Downloads/test.mp3"
        os.startfile(musicPath)

    elif "the time" in query:
        strfTime = datetime.datetime.now().strftime("%H:%M:%S")
        speaker.Speak(f"Sir the time is {strfTime}")

    elif "Using artificial intelligence".lower() in query.lower():
        ai(prompt=query)

    elif "April Quit".lower() in query.lower():
        exit()

    elif "reset chat ".lower() in query.lower():
        chatStr = ""

    else:
        print("Chatting...")
        chat(query)
