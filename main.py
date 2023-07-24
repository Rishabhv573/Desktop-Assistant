import speech_recognition as sr
import subprocess
import pyjokes
import win32com.client
import webbrowser
import openai
from config import apikey_openai, app_id_wolframalpha, api_key_weather, passWord_mail, self_mail
import datetime as dt
import time
import requests
import pyautogui
import wolframalpha
import ctypes
import generator
import takenote
from ecapture import ecapture as ec
from AppOpener import open, close
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from googletrans import Translator, LANGUAGES
import threading
import customtkinter
from PIL import ImageGrab


customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

root = customtkinter.CTk()
root.title("Desktop Assistant")
root.geometry("800x700")



def print_output(text):
    if text[-1] == 'r':
        text = text[:-1]
        output_text.insert(customtkinter.END, f"\n{text}",'right')
        output_text.see(customtkinter.END)
    else:
        output_text.insert(customtkinter.END, f"\n{text}")
        output_text.see(customtkinter.END)


waiting_for_input = False
def on_enter(callback):
    global waiting_for_input
    if waiting_for_input:
        waiting_for_input = False
        callback(query)
    else:
        waiting_for_input = True

def get_input():
    input_text = entry.get()
    output_text.insert(customtkinter.END, "\nUser: " + input_text, 'right')
    entry.delete(0, customtkinter.END)  # Clear the input field after inserting text
    return input_text


chatStr = ""
def chat(query):
    global chatStr
    openai.api_key = apikey_openai
    chatStr += f"User: {query}\n Bot: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]

def ai(prompt):
    generator.func(prompt)

def note(query):
    text = get_input()
    takenote.func(text)

def take_screenshot(filename):
    screenshot = ImageGrab.grab()
    screenshot.save(filename)

def weather(query):
    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    api_key = api_key_weather
    city = f"{query}"

    url = base_url + "appid=" + api_key + "&q=" + city
    response = requests.get(url).json()

    temperature_celcius = response['main']['temp'] - 273.15
    humidity = response['main']['humidity']
    windspeed = response['wind']['speed']
    description = response['weather'][0]['description']
    sunrise_time = dt.datetime.utcfromtimestamp(response['sys']['sunrise']+response['timezone'])
    sunset_time = dt.datetime.utcfromtimestamp(response['sys']['sunset']+response['timezone'])
    print_output(f"{city} weather report\n************************************************************\n")
    print_output(f"Temperature: {temperature_celcius:.2f}")
    print_output(f"Humidity: {humidity}%")
    print_output(f"Windspeed: {windspeed}m/s")
    print_output(f"Sunrise Time: {sunrise_time}")
    print_output(f"Sunset Time: {sunset_time}")
    print_output(f"General description: {description}")

def send_email(self_mail, passWord_mail, to, subject, content):
    # Set up the SMTP server
    smtp_server = 'smtp.office365.com'
    smtp_port = 587

    # Create a secure connection to the SMTP server
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(self_mail, passWord_mail)
    msg = MIMEMultipart()

    msg['From'] = self_mail
    msg['To'] = to
    msg['Subject'] = subject

    msg.attach(MIMEText(content, 'plain'))

    # Send the email
    server.send_message(msg)
    server.quit()

def translate(query):
    trans_lang_from = query.split()[1]
    trans_lang_to = query.split()[-1]
    text = get_input()
    translator = Translator()
    # Specify the text and the language in which you want to translate

    target_lang_code = ''
    for code, language in LANGUAGES.items():
        if trans_lang_to.lower() == language:
            target_lang_code = code
            break

    translation = translator.translate(text, dest=target_lang_code)

    # Print the translated text
    translated = translation.text
    print_output(f"{trans_lang_to}: {translated}")
    speak(translated)

def keyboardControl():
    pyautogui.typewrite(['volumemute'])
def speak(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def takecmd():
    global r, query
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = .6
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print_output(f"User: {query} r")
            return query
        except Exception as e:
            return "Some error occurred"


def start_cmd():
    global waiting_for_input
    waiting_for_input = False
    speak("I'm Desktop Assistant. How can I help you")

    def run_command():
        run = True
        while run:
            query = takecmd()
            sites = [["Youtube", "https://youtube.com"], ["google", "https://google.com"],
                     ["wikipedia", "https://wikipedia.com"], ["leetcode", "https://leetcode.com"]]
            for site in sites:
                if f"Open {site[0]}".lower() in query.lower():
                    webbrowser.open(f"{site[1]}")

            apps = ["Whatsapp", "Spotify", "Chrome", "Proton", "Visual Studio Code", "Word", "Excel", "Powerpoint",
                    "QuickDrop", "Brave", "Edge", "Snipping Tool", "Settings"]
            for app in apps:
                if f"Open {app}".lower() in query.lower():
                    open(f"{app}", match_closest=True)

            for app in apps:
                if f"Close {app}".lower() in query.lower():
                    close(f"{app}", match_closest=True)

            if 'joke'.lower() in query.lower():
                output_joke = pyjokes.get_joke()
                print_output(output_joke)
                speak(output_joke)

            elif 'search'.lower() in query.lower():
                query = query.replace("search", "")
                speak("Ok")
                webbrowser.open(query)

            elif 'lock window'.lower() in query.lower():
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

            elif "restart".lower() in query.lower():
                subprocess.call(["shutdown", "/r"])

            elif "hibernate".lower() in query.lower() or "sleep".lower() in query.lower():
                speak("Hibernating")
                subprocess.call("shutdown / h")

            elif "Quit".lower() in query.lower():
                quit()
                root.destroy()

            elif "weather".lower() in query.lower():
                speak("Here's the weather report")
                words = query.split()
                text = words[-1]
                weather(text)

            elif "time now".lower() in query.lower():
                hour = dt.datetime.now().strftime("%H")
                minute = dt.datetime.now().strftime("%M")
                speak(f"The time is {hour}:{minute}")

            elif "Keyboard Control".lower() in query.lower():
                keyboardControl()

            elif "camera".lower() in query.lower() or "take a photo".lower() in query.lower():
                date = dt.date.today()
                ec.capture(0, "Camera ", f"img-{date}.jpg")

            elif "screenshot".lower() in query.lower():
                date = dt.date.today()
                timenow = dt.datetime.now().strftime("%H%M")
                path = f"Openai/screenshot-{date}-{timenow}.png"
                take_screenshot(path)

            elif "Calculate".lower() in query.lower():
                app_id = app_id_wolframalpha
                client = wolframalpha.Client(app_id)
                index = query.lower().split().index('calculate')
                query = query.split()[index + 1:]
                res = client.query(' '.join(query))
                answer = next(res.results).text
                print_output("The answer is " + answer)
                speak("The answer is " + answer)

            elif "don't listen".lower() in query.lower() or "stop listening".lower() in query.lower() or \
                    "not listen".lower() in query.lower():
                print_output("OK")
                speak("Ok")
                words = query.split()
                last = words[-1]
                a = int(words[-2])
                b = a
                if last == "hours" or last == "hour":
                    b = a * 3600
                elif last == "minutes" or last == "minute":
                    b = a * 60
                time.sleep(b)
                print_output(f"{a} {last} are over now")
                speak(f"{a} {last} are over now")

            elif "location".lower() in query.lower() or "locate".lower() in query.lower():
                words = query.split()
                location = words[-1]
                speak(f"Location of {location}")
                webbrowser.open("https://www.google.com/maps/place/" + location + "")

            elif 'mail'.lower() in query.lower():
                try:
                    speak("Write the receiver's mail address")
                    to = input("To: ")
                    speak("You can write the subject and afterwards the content")
                    subject = input("Subject: ")
                    content = input("**********Content**********\n")
                    send_email(self_mail, passWord_mail, to, subject, content)
                    speak("Email has been sent successfully !")
                except Exception as e:
                    print(e)
                    speak("I am sorry, not able to send this email")

            elif "generate".lower() in query.lower():
                prompt = query[1:]
                ai(prompt)

            elif "note".lower() in query.lower():
                speak("What should I write, You can type it")
                entry.bind("<Return>", lambda event: on_enter(note))

            elif "translate".lower() in query.lower():
                speak("Write what you want to translate")
                entry.bind("<Return>", lambda event: on_enter(translate))

            elif "recycle bin".lower() in query.lower():
                SHEmptyRecycleBin = ctypes.windll.shell32.SHEmptyRecycleBinW
                SHEmptyRecycleBin(None, None, True)
                speak("Recycle bin is empty now")

            elif "chat".lower() in query.lower():
                print_output(chat(query))

            else:
                run = True

    cmd_thread = threading.Thread(target=run_command)
    cmd_thread.start()


frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

# Create a label
label = customtkinter.CTkLabel(master=frame, text="Desktop Assistant", font=("Arial", 26))
label.pack(pady=12, padx=10)

# Output Box
output_text = customtkinter.CTkTextbox(master=frame, font=("Arial", 20), height=200, width=500)
output_text.pack(pady=12, padx=10, fill=customtkinter.BOTH, expand=True)

# Input box
entry = customtkinter.CTkEntry(master=frame, font=("Arial", 20), width=780, height=50)
entry.pack(pady=12, padx=10)

# Create Start button
button = customtkinter.CTkButton(master=frame, text="Start", font=("Arial", 20),
                                 command=start_cmd)
button.pack(side="left", fill=customtkinter.BOTH, expand=True, pady=12, padx=70)

# Create an exit button
exit_button = customtkinter.CTkButton(master=frame, text="Exit", font=("Arial", 20),
                                      command=root.destroy)
exit_button.pack(side="left", fill=customtkinter.BOTH, expand=True, pady=12, padx=70)

# Configure the 'right' tag for right alignment
output_text.tag_config('right', justify='right')

# Start the main event loop
root.mainloop()
