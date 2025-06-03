import speech_recognition as sr
import pyttsx3
import webbrowser
import os
import pywhatkit 
import psutil
import screen_brightness_control as sbc
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import asyncio
from bleak import BleakClient 
    

# Initialize the recognizer and the engine for speech
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Set properties for pyttsx3 voice (optional)
engine.setProperty('rate', 150)  # Speed of speech
engine.setProperty('volume', 0.9)  # Volume level (0.0 to 1.0)

# Bluetooth MAC Address or UUID of the Smartwatch 
SMARTWATCH_MAC = "26:E6:BB:12:C7:62"  # Example: "A4:C1:38:XX:XX:XX"

# BLE Characteristic UUIDs 
SPO2_UUID = "00002a35-0000-1000-8000-00805f9b34fb"  # Example SPO2 UUID
HEARTRATE_UUID = "00002a37-0000-1000-8000-00805f9b34fb"  # Example HEARTRATE UUID



# Function to make the assistant speak
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Function to listen to user's voice
def take_command():
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source,duration=0.2)  # Reduce background noise
        audio = recognizer.listen(source,timeout=5,phrase_time_limit=10)

    try:
        print("Recognizing...")
        command = recognizer.recognize_google(audio)
        print(f"User said: {command}\n")
    except Exception as e:
        print("Sorry, I didn't catch that. Could you say that again?")
        
        return "None"
    return command.lower()

# Function to open websites and system application 
def open_application(command):
    if 'youtube' in command:
        speak("Opening YouTube")
        webbrowser.open("https://www.youtube.com")
    
    elif 'google' in command:
        speak("Opening google")
        webbrowser.open("https://www.google.com")
    
    
    elif 'facebook' in command:
        speak("Opening Facebook")
        webbrowser.open("https://www.facebook.com")
    
    elif 'google play music' in command:
        # If you have Google Play Music installed locally
        speak("Opening google Play Music")
        os.startfile("path_to_google_play_music")  # Replace with the correct path if needed
    
    elif 'today\'s news' in command or 'news' in command:
        speak("Opening today's news")
        webbrowser.open("https://news.google.com")  # Open Google News
    
    elif 'notepad' in command:
        speak("Opening Notepad")
        os.system("notepad.exe")  # Windows

    elif 'calculator' in command:
        speak("Opening Calculator")
        os.system("calc.exe")  # Windows

    elif 'command prompt' in command or 'cmd' in command:
        speak("Opening Command Prompt")
        os.system("start cmd")  # Windows

    elif 'file explorer' in command:
        speak("Opening File Explorer")
        os.system("explorer")  # Windows

    elif 'word' in command:
        speak("Opening Microsoft Word")
        os.system("start winword")  # Windows

    elif 'excel' in command:
        speak("Opening Microsoft Excel")
        os.system("start excel")  # Windows

    elif 'powerpoint' in command:
        speak("Opening Microsoft PowerPoint")
        os.system("start powerpnt")  # Windows

    elif 'paint' in command:
        speak("Opening Paint")
        os.system("mspaint")  # Windows

    elif 'task manager' in command:
        speak("Opening Task Manager")
        os.system("taskmgr")  # Windows

    elif 'settings' in command:
        speak("Opening Settings")
        os.system("start ms-settings:")  # Windows

    else:
        speak("Sorry, I can’t do that right now.")


def search_online(query):
    """Search query on Google and speak it out"""
    search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
    speak(f"Searching {query} on Google")
    webbrowser.open(search_url)

def play_music(command):
    if "play" in command:
        song = command.replace("play", "").strip()
        speak(f"Playing {song} on YouTube")
        pywhatkit.playonyt(song)
        
def get_battery_status():
    battery = psutil.sensors_battery()
    return f"Battery is at {battery.percent}% and charging: {battery.power_plugged}"

def increase_brightness(amount=10):
    current = sbc.get_brightness()[0]  # Get current brightness
    new_level = min(current + amount, 100)
    sbc.set_brightness(new_level)
    speak(f"Increasing brightness to {new_level} percent")

def decrease_brightness(amount=10):
    current = sbc.get_brightness()[0]
    new_level = max(current - amount, 0)
    sbc.set_brightness(new_level)
    speak(f"Decreasing brightness to {new_level} percent")
    
def get_volume_interface():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return cast(interface, POINTER(IAudioEndpointVolume))

def set_volume(percent):
    volume = get_volume_interface()
    volume.SetMasterVolumeLevelScalar(percent / 100, None)
    speak(f"Volume set to {percent} percent")

def increase_volume(step=10):
    volume = get_volume_interface()
    current = volume.GetMasterVolumeLevelScalar()
    new = min(current + step / 100, 1.0)
    volume.SetMasterVolumeLevelScalar(new, None)
    speak(f"Increasing volume to {int(new * 100)} percent")

def decrease_volume(step=10):
    volume = get_volume_interface()
    current = volume.GetMasterVolumeLevelScalar()
    new = max(current - step / 100, 0.0)
    volume.SetMasterVolumeLevelScalar(new, None)
    speak(f"Decreasing volume to {int(new * 100)} percent")

def mute_volume():
    volume = get_volume_interface()
    volume.SetMute(1, None)
    speak("Volume muted")

def unmute_volume():
    volume = get_volume_interface()
    volume.SetMute(0, None)
    speak("Volume unmuted")

def shutdown_system():
    os.system("shutdown /s /t 1")

def restart_system():
    os.system("shutdown /r /t 1")
        

# Function to Fetch BP and Pulse Data from Smartwatch
async def fetch_health_data():
    try:
        async with BleakClient(SMARTWATCH_MAC) as client:
            if await client.is_connected():
                print("Connected to Smartwatch!")

                # Read SPO2
                spo2_data = await client.read_gatt_char(SPO2_UUID)
                spo2= int.from_bytes(spo2_data[:2], byteorder='little')

                # Read HEART RATE 
                heart_data = await client.read_gatt_char(HEARTRATE_UUID)
                heart_rate = int.from_bytes(heart_data[:2], byteorder='little')

                # Announce the results
                result = f"Your spo2 is {spo2} and your pulse rate is {heart_rate}."
                print(result)
                speak(result)
            else:
                speak("Failed to connect to the smartwatch.")

    except Exception as e:
        print("Error:", e)
        speak("I couldn't fetch the health data. Please check the smartwatch connection.")



# Main function to execute the assistant
def jarvis():
    speak("Hello, how can I assist you?")
    
    while True:
        command = take_command()

        # Respond to "Jarvis"
        if 'jarvis' in command:
            speak("Yes, I'm Jarvis, how can I assist you?")
        
        # Check for commands to open apps/websites
        elif 'open' in command:
            open_application(command)

        elif 'search'in command or 'find' in command:
            search_online(command)

    
        elif "play" in command:
                play_music(command)
                
        elif "shutdown" in command:
              shutdown_system() 
              
        elif "restart" in command:
              restart_system()
              
        elif "battery" in command:
              speak(get_battery_status())
              
        elif "increase brightness" in command:
              increase_brightness()

        elif "decrease brightness" in command:
              decrease_brightness()

        elif "increase volume" in command:
             increase_volume()

        elif "decrease volume" in command:
              decrease_volume()

        elif "mute volume" in command:
             mute_volume()

        elif "unmute volume" in command:
             unmute_volume()


        elif "check my spo2" in command or "check my heart rate" in command:
            speak("Fetching your spo2 and heart rate.")
            asyncio.run(fetch_health_data())


        # Stop or exit the assistant
        elif 'stop' in command or 'exit' in command:
            speak("Goodbye!")
            break

# Start the virtual assistant
if __name__ == "__main__":
    jarvis()
