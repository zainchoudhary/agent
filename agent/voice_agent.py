
import threading
import time
import subprocess
import pyperclip
import pyautogui
import requests
import speech_recognition as sr
from config import GROQ_API_KEY


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
    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)
    try:
        user_text = recognizer.recognize_google(audio)
        print(f"Recognized: {user_text}")
        # Send to Groq API for intent/action
        action = get_groq_action(user_text)
        print(f"Groq action: {action}")
        perform_action(action, user_text)
    except Exception as e:
        print(f"Error: {e}")

# --- Groq API integration ---
def get_groq_action(user_text):
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json"
    }
    prompt = (
        "You are a desktop automation agent. "
        "Given a user command, respond with a JSON object describing the action. "
        "Supported actions: open_app, search_web, paste_text, type_text, system_command. "
        "For example: 'open chrome and search google.com' => {\"action\": \"open_app\", \"app\": \"chrome\", \"search\": \"google.com\"}. "
        "'open notepad' => {\"action\": \"open_app\", \"app\": \"notepad\"}. "
        "'shutdown the computer' => {\"action\": \"system_command\", \"command\": \"shutdown\"}. "
        "'hello how are you' => {\"action\": \"paste_text\", \"text\": \"hello how are you\"}. "
        f"User command: {user_text}\nRespond with JSON only."
    )
    data = {
        "model": "llama-3.1-8b-instant",
        "messages": [
            {"role": "system", "content": "You are a desktop automation agent."},
            {"role": "user", "content": prompt}
        ],
        "max_tokens": 256,
        "temperature": 0.2
    }
    try:
        response = requests.post(url, headers=headers, json=data, timeout=15)
        response.raise_for_status()
        content = response.json()["choices"][0]["message"]["content"]
        # Try to extract JSON from response
        import json
        start = content.find('{')
        end = content.rfind('}') + 1
        if start != -1 and end != -1:
            return json.loads(content[start:end])
        else:
            return {"action": "paste_text", "text": user_text}
    except Exception as e:
        print(f"Groq API error: {e}")
        return {"action": "paste_text", "text": user_text}

# --- Action performer ---
def perform_action(action, user_text):
    act = action.get("action", "paste_text")
    if act == "open_app":
        app = action.get("app", "")
        if app == "chrome":
            url = action.get("search") or "https://www.google.com"
            try:
                import webbrowser
                webbrowser.get(using='chrome').open(url)
            except Exception:
                import webbrowser
                webbrowser.open(url)
        elif app == "notepad":
            subprocess.Popen(['notepad.exe'])
        else:
            # Try to open any app by name
            try:
                subprocess.Popen([app])
            except Exception as e:
                print(f"App open error: {e}")
    elif act == "search_web":
        query = action.get("query", user_text)
        url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
        import webbrowser
        webbrowser.open(url)
    elif act == "system_command":
        cmd = action.get("command", "")
        if cmd == "shutdown":
            print("Shutdown command detected (not executing for safety)")
            # subprocess.Popen(['shutdown', '/s', '/t', '1'])
        elif cmd == "lock":
            subprocess.Popen(['rundll32.exe', 'user32.dll,LockWorkStation'])
        else:
            print(f"Unknown system command: {cmd}")
    elif act == "type_text":
        text = action.get("text", user_text)
        pyautogui.typewrite(text)
    else:  # Default: paste text
        text = action.get("text", user_text)
        try:
            paste_text(text)
        except Exception:
            open_notepad_and_paste(text)




if __name__ == "__main__":
    pyautogui.alert(text='Agent Active! Always listening...', title='Voice Agent', button='OK')
    print("Agent started. Always listening for voice input...")
    while True:
        listen_and_paste()
