import os
import sys
import json
import math
import time
import logging
import threading
import subprocess
import webbrowser
import queue
from datetime import datetime
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Dict, Any, List

import pyperclip
import pyautogui
import requests
import speech_recognition as sr
import pyttsx3
import psutil
import winreg

from .config import AGENT_CONFIG

# Load .env and get GROQ_API_KEY
from dotenv import load_dotenv
load_dotenv()  # Load .env file for secrets (GROQ_API_KEY)
load_dotenv(dotenv_path=Path(__file__).parent.parent / '.env')
GROQ_API_KEY = os.getenv('GROQ_API_KEY')

# ══════════════════════════════════════════════════════════
# LOGGING
# ══════════════════════════════════════════════════════════
LOG_DIR = Path.home() / ".provoiceagent" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",  # ASCII pipe to avoid UnicodeEncodeError
    handlers=[
        logging.FileHandler(LOG_DIR / f"agent_{datetime.now().strftime('%Y%m%d')}.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("ProVoiceAgent")


# ══════════════════════════════════════════════════════════
# DATA CLASSES
# ══════════════════════════════════════════════════════════
@dataclass
class ActionResult:
    success: bool
    message: str
    data: Optional[Any] = None
    speak: Optional[str] = None   # what the agent should say aloud


@dataclass
class ConversationTurn:
    role: str
    content: str
    timestamp: datetime = field(default_factory=datetime.now)


# ══════════════════════════════════════════════════════════
# TTS ENGINE
# ══════════════════════════════════════════════════════════
class TTSEngine:
    def __init__(self):
        self._engine = pyttsx3.init()
        self._engine.setProperty("rate", AGENT_CONFIG.get("tts_rate", 185))
        self._engine.setProperty("volume", 0.92)
        self._lock = threading.Lock()
        self._prefer_female_voice()

    def _prefer_female_voice(self):
        voices = self._engine.getProperty("voices")
        for v in voices:
            if any(k in v.name.lower() for k in ("zira", "female", "samantha", "hazel")):
                self._engine.setProperty("voice", v.id)
                break

    def speak(self, text: str):
        """Non-blocking TTS."""
        threading.Thread(target=self._say, args=(text,), daemon=True).start()

    def speak_sync(self, text: str):
        """Blocking TTS — use for startup/shutdown messages."""
        self._say(text)

    def _say(self, text: str):
        with self._lock:
            try:
                self._engine.say(text)
                self._engine.runAndWait()
            except Exception as e:
                log.error(f"TTS error: {e}")


# ══════════════════════════════════════════════════════════
# VOICE ENGINE
# ══════════════════════════════════════════════════════════
class VoiceEngine:
    def __init__(self, tts: TTSEngine):
        self.tts = tts
        self.recognizer = sr.Recognizer()
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.pause_threshold = AGENT_CONFIG.get("pause_threshold", 0.8)
        self.mic = sr.Microphone()
        self._calibrate()

    def _calibrate(self):
        log.info("Calibrating microphone for ambient noise…")
        with self.mic as src:
            self.recognizer.adjust_for_ambient_noise(src, duration=1.5)
        log.info(f"Energy threshold → {self.recognizer.energy_threshold:.0f}")

    def listen(self, timeout: int = 15, phrase_limit: int = 20) -> Optional[str]:
        with self.mic as src:
            log.info("🎤  Listening…")
            try:
                audio = self.recognizer.listen(
                    src, timeout=timeout, phrase_time_limit=phrase_limit
                )
            except sr.WaitTimeoutError:
                return None

        try:
            text = self.recognizer.recognize_google(
                audio, language=AGENT_CONFIG.get("language", "en-US")
            )
            log.info(f"Recognized → \"{text}\"")
            return text.strip()
        except sr.UnknownValueError:
            return None
        except sr.RequestError as e:
            log.error(f"Speech API error: {e}")
            return None


# ══════════════════════════════════════════════════════════
# ACTION EXECUTOR  (30+ actions)
# ══════════════════════════════════════════════════════════
class ActionExecutor:

    # ── App launch table ─────────────────────────────────
    APP_MAP: Dict[str, List[str]] = {
        "chrome":        ["chrome",   r"C:\Program Files\Google\Chrome\Application\chrome.exe"],
        "firefox":       ["firefox",  r"C:\Program Files\Mozilla Firefox\firefox.exe"],
        "edge":          ["msedge",   r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"],
        "brave":         ["brave",    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"],
        "notepad":       ["notepad"],
        "calculator":    ["calc"],
        "explorer":      ["explorer"],
        "paint":         ["mspaint"],
        "cmd":           ["cmd"],
        "powershell":    ["powershell"],
        "task manager":  ["taskmgr"],
        "word":          ["winword"],
        "excel":         ["excel"],
        "powerpoint":    ["powerpnt"],
        "outlook":       ["outlook"],
        "onenote":       ["onenote"],
        "vscode":        ["code"],
        "discord":       ["discord"],
        "spotify":       ["spotify"],
        "vlc":           ["vlc"],
        "zoom":          ["zoom"],
        "teams":         ["teams"],
        "slack":         ["slack"],
        "telegram":      ["telegram"],
        "whatsapp":      ["WhatsApp"],
        "snipping tool": ["snippingtool"],
        "control panel": ["control"],
        "regedit":       ["regedit"],
        "device manager":["devmgmt.msc"],
        "event viewer":  ["eventvwr"],
    }

    # ── App name aliases (normalize common variations) ────
    APP_ALIASES: Dict[str, str] = {
        "file explorer": "explorer",
        "file_explorer": "explorer",
        "windows explorer": "explorer",
        "explorer app": "explorer",
        "file manager": "explorer",
        "task mgr": "task manager",
        "tasks": "task manager",
        "taskmgr": "task manager",
        "visual studio code": "vscode",
        "vs code": "vscode",
        "vscode": "vscode",
        "code editor": "vscode",
        "notepad app": "notepad",
        "notes": "notepad",
        "text editor": "notepad",
        "calc": "calculator",
        "calc app": "calculator",
        "calculator app": "calculator",
        "paint app": "paint",
        "paint tool": "paint",
        "cmd prompt": "cmd",
        "command prompt": "cmd",
        "powershell app": "powershell",
        "ps": "powershell",
        "settings app": "settings",
        "system settings": "settings",
        "windows settings": "settings",
        "chrome browser": "chrome",
        "google chrome": "chrome",
        "firefox browser": "firefox",
        "mozilla firefox": "firefox",
        "edge browser": "edge",
        "microsoft edge": "edge",
        "telegram app": "telegram",
        "teams app": "teams",
        "microsoft teams": "teams",
        "slack app": "slack",
        "discord app": "discord",
        "spotify app": "spotify",
        "zoom meeting": "zoom",
        "onenote app": "onenote",
        "outlook email": "outlook",
        "microsoft outlook": "outlook",
        "word doc": "word",
        "microsoft word": "word",
        "excel sheet": "excel",
        "microsoft excel": "excel",
        "powerpoint ppt": "powerpoint",
        "microsoft powerpoint": "powerpoint",
        "vlc media": "vlc",
        "vlc player": "vlc",
        "snip": "snipping tool",
        "screenshot tool": "snipping tool",
        "control panel app": "control panel",
        "settings": "settings",
        "camera app": "camera",
        "webcam": "camera",
    }

    # ── Search engine table ───────────────────────────────
    SEARCH_ENGINES: Dict[str, str] = {
        "google":       "https://www.google.com/search?q={}",
        "youtube":      "https://www.youtube.com/results?search_query={}",
        "bing":         "https://www.bing.com/search?q={}",
        "github":       "https://github.com/search?q={}",
        "stackoverflow":"https://stackoverflow.com/search?q={}",
        "reddit":       "https://www.reddit.com/search/?q={}",
        "amazon":       "https://www.amazon.com/s?k={}",
        "maps":         "https://www.google.com/maps/search/{}",
        "wikipedia":    "https://en.wikipedia.org/wiki/Special:Search?search={}",
    }

    def __init__(self, tts: TTSEngine):
        self.tts = tts
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.05

        # Register handlers
        self._handlers = {
            "open_app":         self._open_app,
            "close_app":        self._close_app,
            "restart_app":      self._restart_app,
            "list_apps":        self._list_apps,
            "app_status":       self._app_status,
            "search_web":       self._search_web,
            "open_url":         self._open_url,
            "type_text":        self._type_text,
            "paste_text":       self._paste_text,
            "write_notepad":    self._write_notepad,
            "screenshot":       self._screenshot,
            "system_command":   self._system_command,
            "volume_control":   self._volume_control,
            "media_control":    self._media_control,
            "window_control":   self._window_control,
            "open_file":        self._open_file,
            "find_file":        self._find_file,
            "create_file":      self._create_file,
            "clipboard":        self._clipboard,
            "keyboard_shortcut":self._keyboard_shortcut,
            "get_info":         self._get_info,
            "calculate":        self._calculate,
            "scroll":           self._scroll,
            "click":            self._click,
            "answer":           self._answer,
            "hotkey":           self._hotkey,
            "run_command":      self._run_command,
            "open_path":        self._open_path,
            "set_reminder":     self._set_reminder,
            "toggle_wifi":      self._toggle_wifi,
            "brightness":       self._brightness,
        }

    def execute(self, action: Dict[str, Any], raw_text: str) -> ActionResult:
        act = action.get("action", "paste_text")
        log.info(f"⚡  Action: {act}  |  params: {action}")
        handler = self._handlers.get(act, self._paste_text)
        try:
            return handler(action, raw_text)
        except Exception as e:
            log.error(f"Action error [{act}]: {e}", exc_info=True)
            return ActionResult(False, f"Error in {act}: {e}")

    # ── INTELLIGENT APP NAME MATCHING ──────────────────────
    def _normalize_and_match_app(self, app_input: str) -> str:
        """Intelligently match app name using aliases and fuzzy matching."""
        if not app_input:
            return app_input
        
        app = app_input.lower().strip()
        
        # Step 1: Direct alias match
        if app in self.APP_ALIASES:
            matched = self.APP_ALIASES[app]
            log.info(f"App alias match: '{app_input}' → '{matched}'")
            return matched
        
        # Step 2: Fuzzy match against known apps
        from difflib import SequenceMatcher
        all_known_apps = list(self.APP_MAP.keys()) + list(self.APP_ALIASES.keys())
        best_match = None
        best_ratio = 0.6  # Threshold for fuzzy match
        
        for known_app in all_known_apps:
            ratio = SequenceMatcher(None, app, known_app).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = known_app
        
        if best_match:
            final = self.APP_ALIASES.get(best_match, best_match)
            log.info(f"App fuzzy match: '{app_input}' → '{final}' (confidence: {best_ratio:.0%})")
            return final
        
        # Step 3: Try to match just the main word
        words = app.split()
        if len(words) > 1:
            # If name has multiple words, try matching the longest/most descriptive word
            for word in sorted(words, key=len, reverse=True):
                if word in self.APP_ALIASES:
                    matched = self.APP_ALIASES[word]
                    log.info(f"App word match: '{app_input}' → '{matched}'")
                    return matched
        
        # Step 4: Return normalized version
        return app

    # ── INTELLIGENT PROCESS MANAGEMENT ──────────────────────
    def _find_processes_by_name(self, app: str) -> List[int]:
        """Find all process IDs matching the app name using intelligent matching."""
        from difflib import SequenceMatcher
        
        pids = []
        app_clean = app.replace(" ", "").replace(".exe", "").lower()
        
        try:
            for p in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    name = p.info['name']
                    name_lower = name.lower()
                    name_clean = name_lower.replace(" ", "").replace(".exe", "")
                    
                    # Direct matches
                    if app == name_lower or app_clean == name_clean:
                        pids.append(p.info['pid'])
                        continue
                    
                    # Substring matches
                    if app in name_lower or name_clean in app_clean:
                        pids.append(p.info['pid'])
                        continue
                    
                    # Fuzzy match for typos/variations
                    ratio = SequenceMatcher(None, app_clean, name_clean).ratio()
                    if ratio > 0.7:  # 70% match threshold
                        pids.append(p.info['pid'])
                        log.info(f"Fuzzy matched process: {name} ({ratio:.0%})")
                        continue
                    
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
        except Exception as e:
            log.error(f"Error scanning processes: {e}")
        
        return pids

    def _kill_process_by_pid(self, pid: int, force: bool = True) -> bool:
        """Kill a process by PID safely."""
        try:
            p = psutil.Process(pid)
            if force:
                p.kill()
            else:
                p.terminate()
            # Wait a bit for process to die
            try:
                p.wait(timeout=2)
            except psutil.TimeoutExpired:
                if not force:
                    p.kill()  # Force kill if normal termination times out
            log.info(f"Killed process PID {pid}: {p.name()}")
            return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired) as e:
            log.warning(f"Could not kill process {pid}: {e}")
            return False

    # ── APP DISCOVERY & MANAGEMENT ──────────────────────
    def _find_app_executable(self, app: str) -> Optional[str]:
        """Find app executable in common locations and registry."""
        # Try registered apps first
        cmds = self.APP_MAP.get(app)
        if cmds:
            for cmd in cmds:
                if Path(cmd).exists():
                    return cmd
                try:
                    result = subprocess.run(["where", cmd], capture_output=True, text=True, timeout=3)
                    if result.returncode == 0:
                        return result.stdout.strip().split('\n')[0]
                except Exception:
                    pass

        # Search in Program Files
        search_paths = [
            os.environ.get("ProgramFiles", r"C:\Program Files"),
            os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
            os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")),
        ]
        exe_name = app + ".exe"
        for base in search_paths:
            try:
                for root, dirs, files in os.walk(base):
                    if exe_name in files:
                        return os.path.join(root, exe_name)
            except (PermissionError, OSError):
                continue

        # Try Windows Registry
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall")
            for i in range(winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                try:
                    display_name = winreg.QueryValueEx(subkey, "DisplayName")[0].lower()
                    if app in display_name:
                        try:
                            exe_path = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                            if exe_path and os.path.exists(exe_path):
                                return exe_path
                        except WindowsError:
                            pass
                except WindowsError:
                    pass
        except Exception:
            pass
        return None

    # ── BROWSER / APP LAUNCH ─────────────────────────────
    def _open_app(self, a: dict, raw: str) -> ActionResult:
        app = a.get("app", "").lower().strip()
        url  = a.get("url") or a.get("search") or ""
        args = a.get("args", "")

        # Intelligent app name matching and normalization
        app = self._normalize_and_match_app(app)

        # Browser with optional URL
        browsers = {"chrome", "firefox", "edge", "brave"}
        if app in browsers:
            target = url if url else "https://www.google.com"
            if target and not target.startswith("http"):
                target = "https://" + target
            try:
                webbrowser.get(app).open(target)
            except Exception:
                webbrowser.open(target)
            return ActionResult(True, f"Opened {app}" + (f" → {target}" if url else ""))

        # System settings shortcuts
        settings_map = {
            "settings":          "ms-settings:",
            "wifi settings":     "ms-settings:network-wifi",
            "bluetooth settings":"ms-settings:bluetooth",
            "display settings":  "ms-settings:display",
            "sound settings":    "ms-settings:sound",
            "apps settings":     "ms-settings:appsfeatures",
            "camera":            "microsoft.windows.camera:",
            "calculator":        "calculator:",
        }
        if app in settings_map:
            try:
                os.startfile(settings_map[app])
                return ActionResult(True, f"Opened {app}")
            except Exception as e:
                return ActionResult(False, f"Could not open {app}: {e}")

        # Try to find and open the app
        exe_path = self._find_app_executable(app)
        if exe_path:
            try:
                if args:
                    subprocess.Popen([exe_path, args], shell=True)
                else:
                    subprocess.Popen([exe_path], shell=True)
                return ActionResult(True, f"Opened {app}")
            except Exception as e:
                return ActionResult(False, f"Could not launch {app}: {e}")

        # Last resort: shell
        try:
            subprocess.Popen(app, shell=True)
            return ActionResult(True, f"Launched: {app}")
        except Exception as e:
            return ActionResult(False, f"Could not open '{app}': {e}")

    def _close_app(self, a: dict, raw: str) -> ActionResult:
        app = a.get("app", "").lower()
        force = a.get("force", True)  # Force close by default

        # Intelligent app name matching and normalization
        app = self._normalize_and_match_app(app)
        log.info(f"Attempting to close: {app}")

        # Apps that truly cannot be closed
        uncloseables = set()  # Actually, most apps CAN be closed!
        if app in uncloseables:
            return ActionResult(False, f"System app '{app}' cannot be closed.")

        # Step 1: Try specific exe names from our maps
        special_map = {
            "settings": ["SystemSettings.exe", "SettingsApp.exe", "Settings.exe"],
            "camera": ["cameracapture.exe", "WindowsCamera.exe", "Camera.exe"],
            "notepad": ["notepad.exe", "notepad++.exe"],
            "calculator": ["calculator.exe", "calc.exe"],
            "explorer": ["explorer.exe"],
        }

        proc_map = {
            "chrome": ["chrome.exe"],  "firefox": ["firefox.exe"],
            "edge": ["msedge.exe"],    "notepad": ["notepad.exe"],
            "calculator": ["calculator.exe"], "vlc": ["vlc.exe"],
            "discord": ["discord.exe"], "spotify": ["spotify.exe"],
            "code": ["Code.exe"],      "teams": ["Teams.exe"],
        }

        procs_to_try = special_map.get(app) or proc_map.get(app) or [app + ".exe", app]
        flag = "/f" if force else ""
        
        for proc in procs_to_try:
            result = subprocess.run(["taskkill", flag, "/im", proc],
                                   capture_output=True, text=True, check=False)
            if result.returncode == 0:
                time.sleep(0.3)
                log.info(f"Taskkill succeeded for: {proc}")
                return ActionResult(True, f"Closed {app}", speak=f"Closed {app}")

        # Step 2: Use intelligent process finder
        log.info(f"Using intelligent process finder for: {app}")
        pids = self._find_processes_by_name(app)
        
        if pids:
            killed = 0
            for pid in pids:
                if self._kill_process_by_pid(pid, force):
                    killed += 1
            
            if killed > 0:
                time.sleep(0.3)
                log.info(f"Successfully killed {killed} process(es)")
                return ActionResult(True, f"Closed {app}", speak=f"Closed {app}")

        # Step 3: Last resort - PowerShell for stubborn apps
        log.info(f"Trying PowerShell for: {app}")
        try:
            # Try to kill by process name pattern
            ps_commands = [
                f"Get-Process | Where-Object {{$_.Name -like '*{app}*'}} | Stop-Process -Force -ErrorAction SilentlyContinue",
                f"Get-Process *{app}* -ErrorAction SilentlyContinue | Stop-Process -Force",
            ]
            
            for ps_cmd in ps_commands:
                result = subprocess.run(
                    ["PowerShell", "-NoProfile", "-Command", ps_cmd],
                    capture_output=True, text=True, check=False, timeout=5
                )
                if result.returncode == 0 or "Stop-Process" in result.stdout:
                    time.sleep(0.5)
                    log.info(f"PowerShell close succeeded for: {app}")
                    return ActionResult(True, f"Closed {app}", speak=f"Closed {app}")
        except Exception as e:
            log.warning(f"PowerShell close failed: {e}")

        log.warning(f"Could not close {app}: no running process found")
        return ActionResult(False, f"Could not close {app}: app not running or access denied.")

    def _restart_app(self, a: dict, raw: str) -> ActionResult:
        """Close and reopen an app."""
        app = a.get("app", "").lower().strip()
        app = self._normalize_and_match_app(app)  # Intelligent matching
        delay = int(a.get("delay", 1))  # Delay in seconds before reopening

        # Close
        close_result = self._close_app({"app": app}, raw)
        if not close_result.success:
            return ActionResult(False, f"Could not restart {app}: close failed")

        # Wait and reopen
        time.sleep(delay)
        return self._open_app({"app": app}, raw)

    def _list_apps(self, a: dict, raw: str) -> ActionResult:
        """List all running processes/apps."""
        try:
            running = []
            for p in psutil.process_iter(['name']):
                try:
                    running.append(p.info['name'])
                except psutil.NoSuchProcess:
                    pass
            running = sorted(set(running))[:15]  # Top 15 unique apps
            text = ", ".join(running[:10])
            return ActionResult(True, f"Running: {text}",
                              speak=f"Found {len(running)} running processes")
        except Exception as e:
            return ActionResult(False, f"Could not list apps: {e}")

    def _app_status(self, a: dict, raw: str) -> ActionResult:
        """Check if an app is running."""
        app = a.get("app", "").lower()
        app = self._normalize_and_match_app(app)  # Intelligent matching
        app_clean = app.replace(" ", "").lower()
        try:
            for p in psutil.process_iter(['name']):
                try:
                    name = p.info['name'].lower()
                    if app in name or app_clean in name.replace(" ", ""):
                        return ActionResult(True, f"{app} is running",
                                          speak=f"{app} is currently running")
                except psutil.NoSuchProcess:
                    continue
            return ActionResult(False, f"{app} is not running",
                              speak=f"{app} is not currently running")
        except Exception as e:
            return ActionResult(False, f"Could not check status: {e}")

    # ── WEB ──────────────────────────────────────────────
    def _search_web(self, a: dict, raw: str) -> ActionResult:
        query  = a.get("query", raw)
        engine = a.get("engine", "google").lower()
        tpl    = self.SEARCH_ENGINES.get(engine, self.SEARCH_ENGINES["google"])
        webbrowser.open(tpl.format(query.replace(" ", "+")))
        return ActionResult(True, f"Searched '{query}' on {engine}")

    def _open_url(self, a: dict, raw: str) -> ActionResult:
        url = a.get("url", "")
        if not url.startswith("http"):
            url = "https://" + url
        webbrowser.open(url)
        return ActionResult(True, f"Opened {url}")

    # ── TEXT INPUT ────────────────────────────────────────
    def _type_text(self, a: dict, raw: str) -> ActionResult:
        text = a.get("text", raw)
        time.sleep(0.3)
        pyautogui.typewrite(text, interval=0.03)
        return ActionResult(True, f"Typed: {text[:60]}")

    def _paste_text(self, a: dict, raw: str) -> ActionResult:
        text = a.get("text", raw)
        pyperclip.copy(text)
        time.sleep(0.15)
        pyautogui.hotkey("ctrl", "v")
        return ActionResult(True, f"Pasted text")

    def _write_notepad(self, a: dict, raw: str) -> ActionResult:
        text = a.get("text", raw)
        subprocess.Popen(["notepad.exe"])
        time.sleep(1.8)
        self._paste_text({"text": text}, raw)
        return ActionResult(True, "Wrote to Notepad")

    def _answer(self, a: dict, raw: str) -> ActionResult:
        """LLM answered a factual question — speak the answer."""
        answer = a.get("text", "")
        return ActionResult(True, answer, speak=answer)

    # ── SCREENSHOT ────────────────────────────────────────
    def _screenshot(self, a: dict, raw: str) -> ActionResult:
        dest = Path(a.get("path", str(Path.home() / "Pictures" / "Screenshots")))
        dest.mkdir(parents=True, exist_ok=True)
        name = f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        path = dest / name
        time.sleep(0.6)
        pyautogui.screenshot().save(str(path))
        log.info(f"Screenshot → {path}")
        return ActionResult(True, f"Screenshot saved: {path}",
                            speak=f"Screenshot saved to {dest.name} folder")

    # ── SYSTEM COMMANDS ───────────────────────────────────
    def _system_command(self, a: dict, raw: str) -> ActionResult:
        cmd = a.get("command", "").lower()
        delay = a.get("delay", 30)

        dispatch = {
            "shutdown":          lambda: subprocess.run(["shutdown", "/s", "/t", str(delay)]),
            "shutdown_now":      lambda: subprocess.run(["shutdown", "/s", "/t", "0"]),
            "restart":           lambda: subprocess.run(["shutdown", "/r", "/t", str(delay)]),
            "restart_now":       lambda: subprocess.run(["shutdown", "/r", "/t", "0"]),
            "sleep":             lambda: subprocess.run(
                                     ["rundll32.exe", "powrprof.dll,SetSuspendState", "0,1,0"]),
            "hibernate":         lambda: subprocess.run(["shutdown", "/h"]),
            "lock":              lambda: subprocess.run(
                                     ["rundll32.exe", "user32.dll,LockWorkStation"]),
            "logoff":            lambda: subprocess.run(["shutdown", "/l"]),
            "cancel_shutdown":   lambda: subprocess.run(["shutdown", "/a"]),
            "empty_recycle_bin": lambda: subprocess.run(
                                     ["PowerShell", "-NoProfile", "-Command",
                                      "Clear-RecycleBin -Force"],
                                     capture_output=True),
            "open_task_manager": lambda: subprocess.Popen(["taskmgr"]),
            "open_settings":     lambda: os.startfile("ms-settings:"),
            "open_control_panel":lambda: subprocess.Popen(["control"]),
            "check_updates":     lambda: os.startfile("ms-settings:windowsupdate"),
        }

        fn = dispatch.get(cmd)
        if fn:
            fn()
            speak_map = {
                "shutdown":  f"System will shut down in {delay} seconds.",
                "restart":   f"System will restart in {delay} seconds.",
                "sleep":     "Going to sleep.",
                "lock":      "Locking the screen.",
                "logoff":    "Logging off.",
            }
            return ActionResult(True, f"System: {cmd}",
                                speak=speak_map.get(cmd, f"Done: {cmd}"))
        return ActionResult(False, f"Unknown system command: {cmd}")

    # ── VOLUME ────────────────────────────────────────────
    def _volume_control(self, a: dict, raw: str) -> ActionResult:
        cmd   = a.get("command", "up").lower()
        steps = int(a.get("steps", 5))

        if cmd in ("up", "increase"):
            for _ in range(steps): pyautogui.press("volumeup")
        elif cmd in ("down", "decrease"):
            for _ in range(steps): pyautogui.press("volumedown")
        elif cmd in ("mute", "unmute", "toggle"):
            pyautogui.press("volumemute")
        elif cmd == "max":
            for _ in range(50): pyautogui.press("volumeup")
        elif cmd == "min":
            for _ in range(50): pyautogui.press("volumedown")

        return ActionResult(True, f"Volume {cmd}")

    # ── MEDIA ────────────────────────────────────────────
    def _media_control(self, a: dict, raw: str) -> ActionResult:
        cmd = a.get("command", "").lower()
        key_map = {
            "play":       "playpause",
            "pause":      "playpause",
            "play_pause": "playpause",
            "next":       "nexttrack",
            "previous":   "prevtrack",
            "prev":       "prevtrack",
            "stop":       "stop",
        }
        key = key_map.get(cmd)
        if key:
            pyautogui.press(key)
            return ActionResult(True, f"Media: {cmd}")
        return ActionResult(False, f"Unknown media command: {cmd}")

    # ── WINDOW MANAGEMENT ────────────────────────────────
    def _window_control(self, a: dict, raw: str) -> ActionResult:
        cmd = a.get("command", "").lower()
        actions_map = {
            "minimize":     lambda: pyautogui.hotkey("win", "down"),
            "maximize":     lambda: pyautogui.hotkey("win", "up"),
            "close":        lambda: pyautogui.hotkey("alt", "f4"),
            "fullscreen":   lambda: pyautogui.press("f11"),
            "switch":       lambda: pyautogui.hotkey("alt", "tab"),
            "show_desktop": lambda: pyautogui.hotkey("win", "d"),
            "split_left":   lambda: pyautogui.hotkey("win", "left"),
            "split_right":  lambda: pyautogui.hotkey("win", "right"),
            "new_tab":      lambda: pyautogui.hotkey("ctrl", "t"),
            "close_tab":    lambda: pyautogui.hotkey("ctrl", "w"),
            "next_tab":     lambda: pyautogui.hotkey("ctrl", "tab"),
            "prev_tab":     lambda: pyautogui.hotkey("ctrl", "shift", "tab"),
            "reopen_tab":   lambda: pyautogui.hotkey("ctrl", "shift", "t"),
            "incognito":    lambda: pyautogui.hotkey("ctrl", "shift", "n"),
            "zoom_in":      lambda: pyautogui.hotkey("ctrl", "="),
            "zoom_out":     lambda: pyautogui.hotkey("ctrl", "-"),
            "zoom_reset":   lambda: pyautogui.hotkey("ctrl", "0"),
        }
        fn = actions_map.get(cmd)
        if fn:
            time.sleep(0.2)
            fn()
            return ActionResult(True, f"Window: {cmd}")
        return ActionResult(False, f"Unknown window command: {cmd}")

    # ── FILE OPERATIONS ───────────────────────────────────
    def _open_file(self, a: dict, raw: str) -> ActionResult:
        path = a.get("path", "")
        if not path:
            return ActionResult(False, "No file path given")
        try:
            os.startfile(path)
            return ActionResult(True, f"Opened: {path}")
        except Exception as e:
            return ActionResult(False, f"Cannot open {path}: {e}")

    def _open_path(self, a: dict, raw: str) -> ActionResult:
        path = a.get("path", str(Path.home() / "Desktop"))
        try:
            subprocess.Popen(["explorer", path])
            return ActionResult(True, f"Opened folder: {path}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _find_file(self, a: dict, raw: str) -> ActionResult:
        name       = a.get("name", "")
        search_dir = Path(a.get("dir", str(Path.home())))
        results    = list(search_dir.rglob(f"*{name}*"))[:8]
        if results:
            for r in results:
                log.info(f"  Found: {r}")
            return ActionResult(True, f"Found {len(results)} match(es)",
                                speak=f"Found {len(results)} files matching {name}")
        return ActionResult(False, f"No files found for '{name}'",
                            speak=f"No files found matching {name}")

    def _create_file(self, a: dict, raw: str) -> ActionResult:
        name    = a.get("name", f"note_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        content = a.get("content", "")
        loc     = a.get("location", str(Path.home() / "Desktop"))
        path    = Path(loc) / name
        path.write_text(content, encoding="utf-8")
        os.startfile(str(path))
        return ActionResult(True, f"Created: {path}",
                            speak=f"File {name} created on Desktop")

    # ── CLIPBOARD ─────────────────────────────────────────
    def _clipboard(self, a: dict, raw: str) -> ActionResult:
        cmd = a.get("command", "copy").lower()
        if cmd == "copy":
            pyautogui.hotkey("ctrl", "c")
            return ActionResult(True, "Copied")
        elif cmd == "paste":
            pyautogui.hotkey("ctrl", "v")
            return ActionResult(True, "Pasted")
        elif cmd == "clear":
            pyperclip.copy("")
            return ActionResult(True, "Clipboard cleared")
        elif cmd == "get":
            content = pyperclip.paste()
            return ActionResult(True, "Got clipboard", speak=f"Clipboard says: {content[:120]}")
        return ActionResult(False, f"Unknown clipboard command: {cmd}")

    # ── KEYBOARD SHORTCUTS ────────────────────────────────
    def _keyboard_shortcut(self, a: dict, raw: str) -> ActionResult:
        shortcut = a.get("shortcut", "")
        keys     = a.get("keys", [])

        builtin = {
            "copy":       ["ctrl", "c"],    "paste":    ["ctrl", "v"],
            "cut":        ["ctrl", "x"],    "undo":     ["ctrl", "z"],
            "redo":       ["ctrl", "y"],    "save":     ["ctrl", "s"],
            "save_all":   ["ctrl", "shift", "s"],
            "select_all": ["ctrl", "a"],    "find":     ["ctrl", "f"],
            "new":        ["ctrl", "n"],    "open":     ["ctrl", "o"],
            "print":      ["ctrl", "p"],    "refresh":  ["f5"],
            "hard_refresh":["ctrl", "shift", "r"],
            "address_bar":["alt", "d"],
            "zoom_in":    ["ctrl", "="],    "zoom_out": ["ctrl", "-"],
            "devtools":   ["f12"],          "screenshot":["win", "shift", "s"],
            "emoji":      ["win", "."],     "settings": ["win", "i"],
            "task_view":  ["win", "tab"],
        }

        final_keys = builtin.get(shortcut, keys)
        if isinstance(final_keys, str):
            final_keys = final_keys.split("+")

        if final_keys:
            time.sleep(0.2)
            pyautogui.hotkey(*final_keys)
            return ActionResult(True, f"Hotkey: {'+'.join(final_keys)}")
        return ActionResult(False, "No keys to press")

    def _hotkey(self, a: dict, raw: str) -> ActionResult:
        """Raw hotkey from action."""
        keys = a.get("keys", [])
        if isinstance(keys, str):
            keys = keys.split("+")
        if keys:
            time.sleep(0.2)
            pyautogui.hotkey(*keys)
            return ActionResult(True, f"Pressed: {'+'.join(keys)}")
        return ActionResult(False, "No keys given")

    # ── RUN SHELL COMMAND ────────────────────────────────
    def _run_command(self, a: dict, raw: str) -> ActionResult:
        cmd = a.get("command", "")
        if not cmd:
            return ActionResult(False, "No command given")
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True,
                                    text=True, timeout=15)
            output = result.stdout.strip() or result.stderr.strip() or "Done"
            log.info(f"Run command output: {output[:200]}")
            return ActionResult(True, output[:200], speak=f"Command executed")
        except subprocess.TimeoutExpired:
            return ActionResult(False, "Command timed out")
        except Exception as e:
            return ActionResult(False, str(e))

    # ── INFO QUERIES ──────────────────────────────────────
    def _get_info(self, a: dict, raw: str) -> ActionResult:
        info_type = a.get("type", "time").lower()

        if info_type in ("time", "clock"):
            text = datetime.now().strftime("It is %I:%M %p")
            return ActionResult(True, text, speak=text)

        elif info_type == "date":
            text = datetime.now().strftime("Today is %A, %B %d, %Y")
            return ActionResult(True, text, speak=text)

        elif info_type == "datetime":
            text = datetime.now().strftime("It is %I:%M %p on %A, %B %d, %Y")
            return ActionResult(True, text, speak=text)

        elif info_type == "battery":
            try:
                bat = psutil.sensors_battery()
                if bat:
                    status = "charging" if bat.power_plugged else "on battery"
                    text = f"Battery is at {bat.percent:.0f} percent, {status}"
                    return ActionResult(True, text, speak=text)
            except Exception:
                pass
            return ActionResult(False, "Battery info unavailable",
                                speak="Battery info is not available on this device")

        elif info_type == "system":
            cpu = psutil.cpu_percent(interval=1)
            ram = psutil.virtual_memory()
            disk = psutil.disk_usage("/")
            text = (f"CPU at {cpu}%, RAM {ram.percent}% used "
                    f"({ram.used // 1024**2} MB of {ram.total // 1024**2} MB), "
                    f"Disk {disk.percent}% used")
            log.info(text)
            return ActionResult(True, text, speak=text)

        elif info_type == "ip":
            try:
                resp = requests.get("https://api.ipify.org", timeout=6)
                text = f"Your public IP address is {resp.text}"
                return ActionResult(True, text, speak=text)
            except Exception:
                return ActionResult(False, "Could not fetch IP",
                                    speak="Could not get your IP address")

        elif info_type == "network":
            stats = psutil.net_if_stats()
            active = [iface for iface, s in stats.items() if s.isup]
            text = f"Active interfaces: {', '.join(active)}"
            return ActionResult(True, text, speak=text)

        elif info_type == "uptime":
            boot  = datetime.fromtimestamp(psutil.boot_time())
            delta = datetime.now() - boot
            hours, rem = divmod(int(delta.total_seconds()), 3600)
            mins  = rem // 60
            text  = f"System has been up for {hours} hours and {mins} minutes"
            return ActionResult(True, text, speak=text)

        return ActionResult(False, f"Unknown info type: {info_type}")

    # ── CALCULATOR ────────────────────────────────────────
    def _calculate(self, a: dict, raw: str) -> ActionResult:
        expr = a.get("expression", "")
        if not expr:
            return ActionResult(False, "No expression given")
        try:
            safe_env = {k: getattr(math, k) for k in dir(math) if not k.startswith("_")}
            safe_env["abs"] = abs
            result = eval(expr, {"__builtins__": {}}, safe_env)
            if isinstance(result, float) and result.is_integer():
                result = int(result)
            text = f"{expr} equals {result}"
            log.info(f"Calculation: {text}")
            return ActionResult(True, text, speak=text)
        except Exception as e:
            return ActionResult(False, f"Calculation error: {e}",
                                speak="Sorry, I couldn't calculate that")

    # ── SCROLL ───────────────────────────────────────────
    def _scroll(self, a: dict, raw: str) -> ActionResult:
        direction = a.get("direction", "down").lower()
        amount    = int(a.get("amount", 3))
        factor    = 3 if direction == "up" else -3
        pyautogui.scroll(factor * amount)
        return ActionResult(True, f"Scrolled {direction}")

    def _click(self, a: dict, raw: str) -> ActionResult:
        button = a.get("button", "left").lower()
        double = a.get("double", False)
        if double:
            pyautogui.doubleClick(button=button)
        else:
            pyautogui.click(button=button)
        return ActionResult(True, f"Clicked {button}")

    # ── REMINDER ─────────────────────────────────────────
    def _set_reminder(self, a: dict, raw: str) -> ActionResult:
        message = a.get("message", "Reminder!")
        seconds = int(a.get("seconds", 60))

        def _remind():
            time.sleep(seconds)
            pyautogui.alert(text=message, title="⏰ ProVoiceAgent Reminder", button="OK")

        threading.Thread(target=_remind, daemon=True).start()
        mins = seconds // 60
        text = f"Reminder set for {mins} minute{'s' if mins != 1 else ''}"
        return ActionResult(True, text, speak=text)

    # ── MISC ─────────────────────────────────────────────
    def _toggle_wifi(self, a: dict, raw: str) -> ActionResult:
        # Windows: toggle via netsh (requires elevated prompt ideally)
        try:
            # This works for some Wi-Fi adapters
            subprocess.run(
                ["netsh", "interface", "set", "interface", "Wi-Fi", "enable"],
                capture_output=True, check=False
            )
            return ActionResult(True, "Wi-Fi toggled")
        except Exception as e:
            return ActionResult(False, str(e))

    def _brightness(self, a: dict, raw: str) -> ActionResult:
        level = a.get("level", 70)
        try:
            subprocess.run(
                ["PowerShell", "-NoProfile", "-Command",
                 f"(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods)"
                 f".WmiSetBrightness(1,{level})"],
                capture_output=True, check=False
            )
            return ActionResult(True, f"Brightness set to {level}%")
        except Exception as e:
            return ActionResult(False, str(e))


# ══════════════════════════════════════════════════════════
# GROQ BRAIN  (LLaMA 3.3 70B — most capable free model)
# ══════════════════════════════════════════════════════════
class AgentBrain:

    SYSTEM_PROMPT = """You are ProVoiceAgent — an elite desktop automation AI.
Parse the user's voice command and return a single JSON object describing the action.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SUPPORTED ACTIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

open_app       → {"action":"open_app","app":"chrome","url":"youtube.com"}
               → {"action":"open_app","app":"notepad"}
               → {"action":"open_app","app":"spotify"}

close_app      → {"action":"close_app","app":"chrome"}

search_web     → {"action":"search_web","query":"cats","engine":"youtube"}
               engines: google | youtube | bing | github | stackoverflow | reddit | amazon | maps | wikipedia

open_url       → {"action":"open_url","url":"github.com"}

type_text      → {"action":"type_text","text":"Hello World"}

paste_text     → {"action":"paste_text","text":"<exact speech>"}

write_notepad  → {"action":"write_notepad","text":"shopping list: eggs, milk"}

screenshot     → {"action":"screenshot"}

system_command → {"action":"system_command","command":"shutdown|restart|sleep|hibernate|lock|logoff|cancel_shutdown|empty_recycle_bin|open_task_manager|open_settings|check_updates"}
               → optionally add "delay":60 (seconds) for shutdown/restart

volume_control → {"action":"volume_control","command":"up|down|mute|unmute|max|min","steps":5}

media_control  → {"action":"media_control","command":"play|pause|next|previous|stop"}

window_control → {"action":"window_control","command":"minimize|maximize|close|fullscreen|switch|show_desktop|split_left|split_right|new_tab|close_tab|next_tab|prev_tab|reopen_tab|incognito|zoom_in|zoom_out|zoom_reset|devtools"}

open_file      → {"action":"open_file","path":"C:\\Users\\User\\Desktop\\report.pdf"}

open_path      → {"action":"open_path","path":"C:\\Users\\User\\Downloads"}

find_file      → {"action":"find_file","name":"resume","dir":"C:\\Users\\User\\Documents"}

create_file    → {"action":"create_file","name":"todo.txt","content":"buy groceries","location":"C:\\Users\\User\\Desktop"}

clipboard      → {"action":"clipboard","command":"copy|paste|clear|get"}

keyboard_shortcut → {"action":"keyboard_shortcut","shortcut":"copy|paste|cut|undo|redo|save|select_all|find|refresh|hard_refresh|devtools|screenshot|emoji|settings|task_view"}

hotkey         → {"action":"hotkey","keys":["ctrl","shift","esc"]}

run_command    → {"action":"run_command","command":"ipconfig /all"}

get_info       → {"action":"get_info","type":"time|date|datetime|battery|system|ip|network|uptime"}

calculate      → {"action":"calculate","expression":"sqrt(144) + 5**2"}

scroll         → {"action":"scroll","direction":"up|down","amount":3}

answer         → {"action":"answer","text":"<your direct answer to factual question>"}

set_reminder   → {"action":"set_reminder","message":"Take a break","seconds":300}

volume + brightness:
               → {"action":"volume_control","command":"up","steps":10}
               → {"action":"brightness","level":70}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DECISION RULES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. "open chrome and search youtube.com"   → open_app with url
2. "search python tutorials on youtube"   → search_web with engine=youtube
3. "go to github.com"                     → open_url
4. "what time is it"                      → get_info type=time
5. "take screenshot"                      → screenshot
6. "calculate 25 times 4"                 → calculate expression="25*4"
7. "what is the capital of France"        → answer with text="The capital of France is Paris."
8. "shutdown the computer"                → system_command shutdown
9. "play next song"                       → media_control next
10. "volume up"                           → volume_control up
11. For pure dictation with no clear action → paste_text

Return ONLY valid JSON. No explanation. No markdown.
"""

    def __init__(self):
        self.api_url = "https://api.groq.com/openai/v1/chat/completions"
        self.headers = {
            "Authorization": f"Bearer {GROQ_API_KEY}",
            "Content-Type":  "application/json",
        }
        self.history:  List[ConversationTurn] = []
        self.max_hist = 6

    def parse(self, user_text: str) -> Dict[str, Any]:
        context = ""
        if self.history:
            recent  = self.history[-4:]
            context = "\n\nRecent context:\n" + "\n".join(
                f"  {t.role}: {t.content}" for t in recent
            )

        payload = {
            "model":   AGENT_CONFIG.get("model", "llama-3.3-70b-versatile"),
            "messages": [
                {"role": "system", "content": self.SYSTEM_PROMPT},
                {"role": "user",   "content": f"Voice command: {user_text}{context}\n\nJSON:"},
            ],
            "max_tokens":     350,
            "temperature":    0.05,
            "response_format": {"type": "json_object"},
        }

        try:
            resp = requests.post(self.api_url, headers=self.headers,
                                 json=payload, timeout=15)
            resp.raise_for_status()
            raw = resp.json()["choices"][0]["message"]["content"].strip()
            action = json.loads(raw)
            log.info(f"Brain → {action}")

            # Store turns
            self.history.append(ConversationTurn("user",      user_text))
            self.history.append(ConversationTurn("assistant", json.dumps(action)))
            if len(self.history) > self.max_hist * 2:
                self.history = self.history[-self.max_hist * 2:]

            return action

        except json.JSONDecodeError as e:
            log.error(f"JSON parse fail: {e}")
        except requests.HTTPError as e:
            log.error(f"Groq HTTP error {e.response.status_code}: {e.response.text[:200]}")
        except requests.RequestException as e:
            log.error(f"Groq request error: {e}")
        except Exception as e:
            log.error(f"Brain unexpected error: {e}", exc_info=True)

        return {"action": "paste_text", "text": user_text}


# ══════════════════════════════════════════════════════════
# FLOATING STATUS HUD  (Tkinter borderless overlay)
# ══════════════════════════════════════════════════════════
class StatusHUD:
    COLORS = {
        "ready":      "#00ff88",
        "listening":  "#00ff88",
        "processing": "#ffaa00",
        "executing":  "#00aaff",
        "success":    "#00ff88",
        "error":      "#ff4455",
        "speaking":   "#cc88ff",
    }

    def __init__(self):
        self._q       = queue.Queue()
        self._running = False
        self._thread  = threading.Thread(target=self._run, daemon=True)

    def start(self):
        self._thread.start()
        time.sleep(0.4)          # let Tk initialize

    def show(self, msg: str, state: str = "ready"):
        self._q.put(("show", msg, self.COLORS.get(state, "#ffffff")))

    def _run(self):
        import tkinter as tk
        root = tk.Tk()
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0.94)
        root.configure(bg="#0d0d1a")
        root.resizable(False, False)

        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        w, h = 380, 54
        root.geometry(f"{w}x{h}+{sw - w - 16}+{sh - h - 50}")

        # Dot indicator
        dot = tk.Label(root, text="◉", font=("Segoe UI", 16),
                       fg="#00ff88", bg="#0d0d1a")
        dot.pack(side="left", padx=(10, 6), pady=6)

        # Title
        tk.Label(root, text="ProVoiceAgent", font=("Segoe UI", 9, "bold"),
                 fg="#5566ff", bg="#0d0d1a").place(x=40, y=4)

        # Status label
        lbl = tk.Label(root, text="Initializing…",
                       font=("Segoe UI", 10), fg="white",
                       bg="#0d0d1a", anchor="w")
        lbl.place(x=40, y=24, width=320)

        self._running = True

        def poll():
            try:
                while True:
                    msg = self._q.get_nowait()
                    if msg[0] == "show":
                        lbl.config(text=msg[1])
                        dot.config(fg=msg[2])
            except queue.Empty:
                pass
            if self._running:
                root.after(80, poll)

        poll()
        root.mainloop()


# ══════════════════════════════════════════════════════════
# MAIN ORCHESTRATOR
# ══════════════════════════════════════════════════════════
class ProVoiceAgent:

    # Agent states
    STATES = {
        "STOPPED": 0,
        "IDLE": 1,
        "ACTIVE": 2,
        "EXECUTING": 3,
    }

    EXIT_PHRASES = {
        "goodbye lucifer",
        "stop lucifer",
        "lucifer stop",
        "lucifer goodbye",
        "lucifer quit",
        "quit lucifer",
        "lucifer exit",
        "exit lucifer",
    }

    START_PHRASES = {
        "start lucifer",
        "lucifer start",
        "activate lucifer",
        "lucifer activate",
        "launch lucifer",
        "lucifer launch",
    }

    ACTIVATION_PHRASES = {
        "hello lucifer",
        "lucifer",
        "activate agent",
        "agent activate",
        "wake up lucifer",
        "lucifer wake up",
    }

    def __init__(self):
        log.info("╔══════════════════════════════════════════╗")
        log.info("║     ProVoiceAgent - LUCIFER MODE        ║")
        log.info("╚══════════════════════════════════════════╝")

        self.tts      = TTSEngine()
        self.voice    = VoiceEngine(self.tts)
        self.brain    = AgentBrain()
        self.executor = ActionExecutor(self.tts)
        self.hud      = StatusHUD()
        
        # State management
        self._state = self.STATES["STOPPED"]
        self._running = False
        self._activated = False
        
        log.info(f"Agent initialized in STOPPED state")

    def _get_state_name(self) -> str:
        """Get human-readable state name."""
        for name, value in self.STATES.items():
            if value == self._state:
                return name
        return "UNKNOWN"

    def start(self):
        """Main entry point - wait for startup command."""
        self.hud.start()
        self._state = self.STATES["IDLE"]
        
        self.hud.show("🛑  STOPPED - Say 'Start Lucifer'", "ready")
        log.info(f"Agent state: {self._get_state_name()}")

        self.tts.speak_sync(
            "Lucifer is offline. Say 'Start Lucifer' to activate."
        )

        self._running = True
        log.info("Waiting for startup command...")

        while self._running:
            try:
                if self._state == self.STATES["IDLE"]:
                    self._wait_for_startup()
                elif self._state == self.STATES["ACTIVE"]:
                    self._wait_for_activation()
                elif self._state == self.STATES["EXECUTING"]:
                    self._execute_single_command()
                else:
                    time.sleep(0.5)
            except KeyboardInterrupt:
                self._shutdown()
            except Exception as e:
                log.error(f"Cycle error: {e}", exc_info=True)
                self.hud.show(f"❌ Error: {str(e)[:40]}", "error")
                time.sleep(1)

    def _wait_for_startup(self):
        """Wait for 'Start Lucifer' command to boot the agent."""
        self.hud.show("🛑  STOPPED - Say 'Start Lucifer'", "ready")
        log.info(f"State: {self._get_state_name()} - Waiting for startup...")

        text = self.voice.listen(
            timeout=AGENT_CONFIG.get("listen_timeout", 30),
            phrase_limit=AGENT_CONFIG.get("phrase_limit", 10),
        )

        if not text:
            return

        text_lower = text.lower()
        
        # Check for startup command
        if text_lower in self.START_PHRASES or "start" in text_lower.lower():
            self._boot_agent()
            return

        # Check for exit command
        if text_lower in self.EXIT_PHRASES or "stop" in text_lower:
            log.info("Received stop command in STOPPED state")
            self._shutdown()
            return

        log.info(f"Ignored (not startup): \"{text}\"")

    def _boot_agent(self):
        """Boot the agent - transition from IDLE to ACTIVE."""
        self._state = self.STATES["ACTIVE"]
        self.hud.show("⚡  ONLINE - Lucifer ready!", "processing")
        self.tts.speak("Lucifer is now online. Say 'Hello Lucifer' for commands.")
        log.info(f"Agent BOOTED → State: {self._get_state_name()}")
        time.sleep(0.8)

    def _wait_for_activation(self):
        """Active state: Wait for 'Hello Lucifer' to execute commands."""
        self.hud.show("👀  ONLINE (say 'Hello Lucifer')", "ready")

        text = self.voice.listen(
            timeout=AGENT_CONFIG.get("listen_timeout", 30),
            phrase_limit=AGENT_CONFIG.get("phrase_limit", 10),
        )

        if not text:
            return

        text_lower = text.lower()
        
        # Check for exit command
        if text_lower in self.EXIT_PHRASES or "stop" in text_lower:
            self._shutdown()
            return

        # Check for activation (command execution)
        if text_lower in self.ACTIVATION_PHRASES or "hello" in text_lower or "activate" in text_lower:
            self._activate_for_command()
            return

        log.info(f"Ignored (not activation): \"{text}\"")
        self.hud.show("👀  ONLINE (not recognized)", "ready")

    def _activate_for_command(self):
        """Activate for ONE command execution."""
        self._state = self.STATES["EXECUTING"]
        self.hud.show("🎤  LISTENING for command", "processing")
        self.tts.speak("Ready. What do you need?")
        log.info(f"Agent ACTIVATED → State: {self._get_state_name()}")

    def _execute_single_command(self):
        """Listen for ONE command and execute it."""
        self.hud.show("🎤  LISTENING for command…", "listening")

        text = self.voice.listen(
            timeout=AGENT_CONFIG.get("listen_timeout", 15),
            phrase_limit=AGENT_CONFIG.get("phrase_limit", 20),
        )

        if not text:
            self._return_to_active("No command detected")
            return

        # Exit check
        if text.lower() in self.EXIT_PHRASES or "stop" in text.lower():
            self._shutdown()
            return

        log.info(f"📝  Command: \"{text}\"")
        self.hud.show(f"🧠  {text[:42]}…", "processing")

        # Parse intent
        action = self.brain.parse(text)

        act_name = action.get("action", "?")
        self.hud.show(f"⚡  {act_name}: {self._action_summary(action)}", "executing")

        # Execute
        result = self.executor.execute(action, text)

        # Handle TTS feedback
        speech = result.speak or (result.message if not result.success else None)
        if speech:
            self.hud.show(f"🔊  {speech[:42]}", "speaking")
            self.tts.speak(speech)

        # Update HUD
        icon   = "✅" if result.success else "❌"
        state  = "success" if result.success else "error"
        self.hud.show(f"{icon}  {result.message[:48]}", state)
        time.sleep(0.8)
        
        # Return to active after command execution
        self._return_to_active("Command executed")

    def _return_to_active(self, reason: str = ""):
        """Return from EXECUTING to ACTIVE state."""
        self._state = self.STATES["ACTIVE"]
        msg = f"Returned to active. ({reason})" if reason else "Returned to active."
        log.info(f"Agent returned to ACTIVE: {reason}")
        self.hud.show("👀  ONLINE (say 'Hello Lucifer')", "ready")
        time.sleep(0.5)

    def _shutdown(self):
        """Graceful shutdown - transition back to IDLE (waiting for restart)."""
        self._state = self.STATES["IDLE"]
        log.info(f"Agent SHUTDOWN → State: {self._get_state_name()}")
        self.hud.show("🛑  OFFLINE - Say 'Start Lucifer' to restart", "ready")
        self.tts.speak_sync("Lucifer going offline. Say 'Start Lucifer' to restart.")
        time.sleep(0.5)

    def _action_summary(self, action: dict) -> str:
        a = action.get("action", "")
        if a == "open_app":
            return action.get("app", "") + (" → " + action.get("url", "") if action.get("url") else "")
        if a == "search_web":
            return action.get("query", "")[:30]
        if a == "open_url":
            return action.get("url", "")[:30]
        if a in ("type_text", "paste_text", "write_notepad"):
            return action.get("text", "")[:30]
        if a == "system_command":
            return action.get("command", "")
        if a == "calculate":
            return action.get("expression", "")
        return ""


# ══════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════

# For main.py entry point
def listen_and_paste():
    agent = ProVoiceAgent()
    agent.start()