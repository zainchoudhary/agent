# Desktop Voice Agent with FastAPI

This project is a Python-based desktop agent that:
- Activates with a global keyboard shortcut
- Listens for voice input and converts it to text
- Pastes the recognized text into the currently focused application
- If no app is focused, opens Notepad and pastes the text there

## Tech Stack
- Python 3.10+
- FastAPI (for backend and future extensibility)
- PyInstaller (for packaging)
- pynput (for global hotkey)
- SpeechRecognition (for voice to text)
- pyttsx3 or gTTS (for TTS, optional)
- pyperclip (for clipboard)
- pyautogui (for automation)
- pyaudio (for microphone input)

## Setup
1. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
2. Install dependencies:
   ```bash
   pip install fastapi uvicorn pynput SpeechRecognition pyttsx3 pyperclip pyautogui pyaudio
   ```
3. Run the agent:
   ```bash
   uvicorn agent.main:app --reload
   ```

## Usage
- Press the configured hotkey (e.g., Ctrl+Shift+V) to activate voice input.
- Speak your text; it will be pasted into the active window.
- If no window is focused, Notepad will open and the text will be pasted there.

## Packaging
- Use PyInstaller to create a standalone executable for Windows.

## Notes
- Ensure your microphone is set up and working.
- Some features may require running as administrator for full automation.
