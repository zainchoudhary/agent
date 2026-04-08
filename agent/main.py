
import pyautogui
from .voice_agent import listen_and_paste

if __name__ == "__main__":
    pyautogui.alert(text='Agent Active! Always listening...', title='Voice Agent', button='OK')
    print("Agent started. Always listening for voice input...")
    while True:
        listen_and_paste()
