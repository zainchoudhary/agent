import threading
import time
import subprocess
import pyperclip
import pyautogui
from pynput import keyboard
import speech_recognition as sr

HOTKEY = {keyboard.Key.ctrl_l, keyboard.Key.alt_l, keyboard.KeyCode(char='s')}
current_keys = set()


def paste_text(text):
    pyperclip.copy(text)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')


def open_notepad_and_paste(text):
    subprocess.Popen(['notepad.exe'])
    time.sleep(1.5)
    paste_text(text)


def listen_and_paste():
    recognizer = sr.Recognizer()
    pyautogui.alert(text='Agent Active! Speak now...', title='Voice Agent', button='OK')
    with sr.Microphone() as source:
        print("sListening...")
        audio = recognizer.listen(source)
    try:
        text = recognizer.recognize_google(audio)
        print(f"Recognized: {text}")
        try:
            paste_text(text)
        except Exception:
            open_notepad_and_paste(text)
    except Exception as e:
        print(f"Error: {e}")


def on_press(key):
    if key in HOTKEY:
        current_keys.add(key)
        if all(k in current_keys for k in HOTKEY):
            threading.Thread(target=listen_and_paste).start()

def on_release(key):
    if key in current_keys:
        current_keys.remove(key)


def start_hotkey_listener():
    with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()

if __name__ == "__main__":
    start_hotkey_listener()
