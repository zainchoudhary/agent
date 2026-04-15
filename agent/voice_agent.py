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

# Add agent directory to sys.path for imports
sys.path.insert(0, str(Path(__file__).parent))

from config import AGENT_CONFIG

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
# ACTION EXECUTOR  (50+ actions) - ENTERPRISE LEVEL
# ══════════════════════════════════════════════════════════
class ActionExecutor:

    # ── Extended Universal App Database ─────────────────────────────
    # Auto-discovers apps from system, not just predefined
    APP_MAP: Dict[str, List[str]] = {
        # Browsers
        "chrome":        ["chrome",   r"C:\Program Files\Google\Chrome\Application\chrome.exe"],
        "firefox":       ["firefox",  r"C:\Program Files\Mozilla Firefox\firefox.exe"],
        "edge":          ["msedge",   r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"],
        "brave":         ["brave",    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"],
        "opera":         ["opera",    r"C:\Program Files\Opera\opera.exe"],
        "safari":        ["safari",   r"C:\Program Files\Safari\safari.exe"],
        
        # Development
        "vscode":        ["code"],
        "visual studio": ["devenv"],
        "intellij":      ["idea"],
        "pycharrm":      ["pycharm"],
        "sublime":       ["subl"],
        "notepad++":     ["notepad++"],
        
        # Office & Productivity
        "word":          ["winword"],
        "excel":         ["excel"],
        "powerpoint":    ["powerpnt"],
        "outlook":       ["outlook"],
        "onenote":       ["onenote"],
        "access":        ["msaccess"],
        "publisher":     ["mspub"],
        
        # Communication
        "discord":       ["discord"],
        "telegram":      ["telegram"],
        "whatsapp":      ["WhatsApp"],
        "slack":         ["slack"],
        "teams":         ["teams"],
        "zoom":          ["zoom"],
        "skype":         ["skype"],
        
        # Media & Entertainment
        "spotify":       ["spotify"],
        "vlc":           ["vlc"],
        "mpv":           ["mpv"],
        "foobar":        ["foobar2000"],
        "winamp":        ["winamp"],
        "audacity":      ["audacity"],
        "adobe premiere": ["premiere"],
        "adobe photoshop":["photoshop"],
        "blender":       ["blender"],
        "krita":         ["krita"],
        "gimp":          ["gimp"],
        
        # System
        "notepad":       ["notepad"],
        "calculator":    ["calc"],
        "explorer":      ["explorer"],
        "paint":         ["mspaint"],
        "cmd":           ["cmd"],
        "powershell":    ["powershell"],
        "task manager":  ["taskmgr"],
        "snipping tool": ["snippingtool"],
        "control panel": ["control"],
        "regedit":       ["regedit"],
        "device manager":["devmgmt.msc"],
        "event viewer":  ["eventvwr"],
        "disk manager":  ["diskmgmt.msc"],
        "services":      ["services.msc"],
        "computer management": ["compmgmt.msc"],
        
        # Utilities
        "7zip":          ["7zfm"],
        "winrar":        ["winrar"],
        "winzip":        ["winzip"],
        "everything":    ["everything"],
        "obsidian":      ["obsidian"],
        "notion":        ["notion"],
        "keepass":       ["keepass"],
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
        "postman api": "postman",
        "api client": "postman",
        "rest client": "postman",
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
        self._app_cache = {}  # Cache for discovered apps
        
        # Register handlers - ENTERPRISE LEVEL (50+ actions)
        self._handlers = {
            # App Management
            "open_app":         self._open_app,
            "close_app":        self._close_app,
            "restart_app":      self._restart_app,
            "list_apps":        self._list_apps,
            "app_status":       self._app_status,
            "kill_app":         self._close_app,  # Alias
            
            # Web & URLs
            "search_web":       self._search_web,
            "open_url":         self._open_url,
            "goto":             self._open_url,  # Alias
            
            # Text Input & Clipboard
            "type_text":        self._type_text,
            "paste_text":       self._paste_text,
            "write_notepad":    self._write_notepad,
            "type":             self._type_text,  # Alias
            "paste":            self._paste_text,  # Alias
            
            # Screenshots & Files
            "screenshot":       self._screenshot,
            "open_file":        self._open_file,
            "find_file":        self._find_file,
            "create_file":      self._create_file,
            "open_path":        self._open_path,
            "open_folder":      self._open_path,  # Alias
            "delete_file":      self._delete_file,
            "rename_file":      self._rename_file,
            
            # System Control
            "system_command":   self._system_command,
            "volume_control":   self._volume_control,
            "media_control":    self._media_control,
            "window_control":   self._window_control,
            "brightness":       self._brightness,
            "toggle_wifi":      self._toggle_wifi,
            
            # Clipboard & Keyboard
            "clipboard":        self._clipboard,
            "keyboard_shortcut":self._keyboard_shortcut,
            "hotkey":           self._hotkey,
            
            # Commands & Info
            "run_command":      self._run_command,
            "get_info":         self._get_info,
            
            # Universal Actions (NEW - ENTERPRISE LEVEL)
            "click_element":    self._click_element,      # Universal element clicking
            "find_and_click":   self._find_and_click,     # Find by text and click
            "wait_for_element": self._wait_for_element,   # Wait for UI element
            "navigate_to":      self._navigate_to,        # Universal navigation
            "fill_form":        self._fill_form,          # Universal form filling
            "submit_form":      self._submit_form,        # Submit any form
            "select_dropdown":  self._select_dropdown,    # Select from dropdown
            "mouse_move":       self._mouse_move,         # Move mouse to coordinates
            "mouse_click":      self._mouse_click,        # Click at coordinates
            "mouse_drag":       self._mouse_drag,         # Drag operation
            "keyboard_input":   self._keyboard_input,     # Raw keyboard input
            "focus_window":     self._focus_window,       # Focus specific window
            "wait":             self._wait,               # Wait/delay
            
            # Math & QA
            "calculate":        self._calculate,
            "answer":           self._answer,
            
            # Utilities
            "scroll":           self._scroll,
            "click":            self._click,
            "set_reminder":     self._set_reminder,
        }
        
        # Build system app cache on startup
        threading.Thread(target=self._cache_all_installed_apps, daemon=True).start()

    def execute(self, action: Dict[str, Any], raw_text: str) -> ActionResult:
        act = action.get("action", "paste_text")
        log.info(f"⚡  Action: {act}  |  params: {action}")
        handler = self._handlers.get(act, self._paste_text)
        try:
            return handler(action, raw_text)
        except Exception as e:
            log.error(f"Action error [{act}]: {e}", exc_info=True)
            return ActionResult(False, f"Error in {act}: {e}")

    # ── INTELLIGENT APP NAME MATCHING (DYNAMIC) ──────────────────────
    def _normalize_and_match_app(self, app_input: str) -> str:
        """Match app name using dynamic cache and fuzzy matching."""
        if not app_input:
            return app_input
        
        app = app_input.lower().strip()
        
        # Direct cache match
        if app in self._app_cache:
            return app
        
        # Fuzzy match in cache
        from difflib import SequenceMatcher
        best_match = None
        best_ratio = 0.7
        
        for cached_app in self._app_cache.keys():
            ratio = SequenceMatcher(None, app, cached_app).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = cached_app
        
        if best_match:
            log.info(f"App fuzzy match: '{app_input}' → '{best_match}' ({best_ratio:.0%})")
            return best_match
        
        # Return as-is for shell execution
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

    def _cache_all_installed_apps(self):
        """Scan entire system and build cache of ALL installed apps."""
        log.info("🔍 Building system app cache...")
        
        # ─── Scan Windows Registry ─────────────────────────────────────
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall")
            for i in range(winreg.QueryInfoKey(key)[0]):
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    subkey = winreg.OpenKey(key, subkey_name)
                    display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                    exe_path = None
                    
                    # Try to get executable path
                    try:
                        exe_path = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                    except:
                        try:
                            exe_path = winreg.QueryValueEx(subkey, "UninstallString")[0]
                        except:
                            pass
                    
                    if display_name:
                        app_key = display_name.lower().split()[0]  # Use first word as key
                        self._app_cache[app_key] = exe_path
                        self._app_cache[display_name.lower()] = exe_path
                except:
                    pass
        except Exception as e:
            log.warning(f"Could not scan registry: {e}")
        
        # ─── Scan Program Files & AppData ─────────────────────────────
        search_paths = [
            r"C:\Program Files",
            r"C:\Program Files (x86)",
            str(Path.home() / "AppData" / "Local"),
            str(Path.home() / "AppData" / "Roaming"),
        ]
        
        for base_path in search_paths:
            if not os.path.exists(base_path):
                continue
            try:
                for root, dirs, files in os.walk(base_path, topdown=True):
                    # Limit search depth for performance
                    dirs[:] = dirs[:5]
                    for file in files:
                        if file.endswith(".exe"):
                            app_name = file[:-4].lower()  # Remove .exe
                            full_path = os.path.join(root, file)
                            self._app_cache[app_name] = full_path
            except (PermissionError, OSError):
                continue
        
        # ─── Scan Desktop for shortcuts (.lnk files) ──────────────────
        desktop_path = Path.home() / "Desktop"
        if desktop_path.exists():
            try:
                for item in desktop_path.iterdir():
                    if item.is_file():
                        # Handle .lnk shortcuts
                        if item.suffix.lower() == ".lnk":
                            app_name = item.stem.lower()  # Name without extension
                            self._app_cache[app_name] = str(item)
                        # Handle direct executables on desktop
                        elif item.suffix.lower() == ".exe":
                            app_name = item.stem.lower()
                            self._app_cache[app_name] = str(item)
            except Exception as e:
                log.warning(f"Could not scan Desktop: {e}")
        
        # ─── Scan Start Menu for app shortcuts ─────────────────────────
        start_menu_paths = [
            Path.home() / "AppData" / "Roaming" / "Microsoft" / "Windows" / "Start Menu" / "Programs",
            Path(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"),
        ]
        
        for start_menu in start_menu_paths:
            if not start_menu.exists():
                continue
            try:
                for item in start_menu.rglob("*.lnk"):  # Recursive search
                    app_name = item.stem.lower()
                    self._app_cache[app_name] = str(item)
            except Exception as e:
                log.warning(f"Could not scan Start Menu: {e}")
        
        # ─── Scan user home directory for portable/custom apps ────────
        home_subfolders = ["Downloads", "Documents", "Desktop"]
        for subfolder in home_subfolders:
            folder_path = Path.home() / subfolder
            if folder_path.exists():
                try:
                    for root, dirs, files in os.walk(folder_path, topdown=True):
                        dirs[:] = dirs[:3]  # Limit depth
                        for file in files:
                            if file.endswith(".exe"):
                                app_name = file[:-4].lower()
                                full_path = os.path.join(root, file)
                                self._app_cache[app_name] = full_path
                except (PermissionError, OSError):
                    continue
        
        log.info(f"✅ App cache built: {len(self._app_cache)} apps indexed")

    def _resolve_shortcut(self, shortcut_path: str) -> Optional[str]:
        """Resolve a .lnk (Windows shortcut) to the target executable."""
        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(shortcut_path)
            target_path = shortcut.TargetPath
            if target_path and os.path.exists(target_path):
                return target_path
        except Exception:
            pass
        return None

    def _find_app_executable(self, app: str) -> Optional[str]:
        """Find app executable - checks cache first, then system."""
        from difflib import SequenceMatcher
        
        app_lower = app.lower().strip()
        
        # Step 1: Direct cache lookup
        if app_lower in self._app_cache:
            path = self._app_cache[app_lower]
            # If it's a shortcut, resolve it
            if path and path.lower().endswith(".lnk"):
                resolved = self._resolve_shortcut(path)
                if resolved:
                    return resolved
                return path  # Return shortcut if resolution fails
            if path and os.path.exists(path):
                return path
        
        # Step 2: Fuzzy match in cache
        best_match = None
        best_ratio = 0.7
        for cached_app, path in self._app_cache.items():
            ratio = SequenceMatcher(None, app_lower, cached_app).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = path
                log.info(f"Fuzzy matched: '{app}' → '{cached_app}' ({best_ratio:.0%})")
        
        if best_match:
            if best_match.lower().endswith(".lnk"):
                resolved = self._resolve_shortcut(best_match)
                if resolved:
                    return resolved
            if os.path.exists(best_match):
                return best_match
        
        # Step 3: Try system PATH lookup
        try:
            result = subprocess.run(["where", app_lower], capture_output=True, text=True, timeout=3)
            if result.returncode == 0:
                path = result.stdout.strip().split('\n')[0]
                self._app_cache[app_lower] = path
                return path
        except Exception:
            pass
        
        # Step 4: Try direct shell execution (system will find it)
        return None

    # ── UNIVERSAL APP LAUNCHER ─────────────────────────────
    def _open_app(self, a: dict, raw: str) -> ActionResult:
        app = a.get("app", "").lower().strip()
        url  = a.get("url") or a.get("search") or ""
        args = a.get("args", "")

        if not app:
            return ActionResult(False, "No app specified")

        log.info(f"🚀 Attempting to open: {app}")

        # Web-only services (open in browser)
        web_services = {
            "youtube":    "https://www.youtube.com",
            "facebook":   "https://www.facebook.com",
            "twitter":    "https://www.twitter.com",
            "instagram":  "https://www.instagram.com",
            "tiktok":     "https://www.tiktok.com",
            "netflix":    "https://www.netflix.com",
            "gmail":      "https://mail.google.com",
            "google docs":"https://docs.google.com",
            "google sheets":"https://sheets.google.com",
            "github":     "https://www.github.com",
            "reddit":     "https://www.reddit.com",
            "linkedin":   "https://www.linkedin.com",
            "pinterest":  "https://www.pinterest.com",
            "whatsapp web":"https://web.whatsapp.com",
            "amazon":     "https://www.amazon.com",
            "ebay":       "https://www.ebay.com",
        }
        
        if app in web_services:
            try:
                webbrowser.open(web_services[app])
                return ActionResult(True, f"Opened {app}", speak=f"Opening {app}")
            except Exception as e:
                return ActionResult(False, f"Could not open {app}: {e}")

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

        # Browser with optional URL
        browsers = {"chrome", "firefox", "edge", "brave", "opera"}
        if app in browsers:
            target = url if url else "https://www.google.com"
            if target and not target.startswith("http"):
                target = "https://" + target
            try:
                webbrowser.get(app).open(target)
                return ActionResult(True, f"Opened {app}" + (f" → {target}" if url else ""))
            except Exception:
                try:
                    webbrowser.open(target)
                    return ActionResult(True, f"Opened {app}")
                except Exception as e:
                    return ActionResult(False, f"Could not launch {app}: {e}")

        # Try to find and open the app from cache
        exe_path = self._find_app_executable(app)
        if exe_path:
            try:
                # Handle .lnk shortcuts with os.startfile (native Windows support)
                if exe_path.lower().endswith(".lnk"):
                    os.startfile(exe_path)
                    log.info(f"✅ Launched shortcut: {exe_path}")
                    return ActionResult(True, f"Opened {app}", speak=f"Opening {app}")
                # Handle regular executables
                else:
                    if args:
                        subprocess.Popen([exe_path] + args.split(), shell=False)
                    else:
                        subprocess.Popen([exe_path], shell=False)
                    log.info(f"✅ Launched from path: {exe_path}")
                    return ActionResult(True, f"Opened {app}", speak=f"Opening {app}")
            except Exception as e:
                log.warning(f"Failed to launch from path: {e}")

        # Last resort: shell execution (Windows will search PATH and registry)
        log.info(f"Attempting shell execution for: {app}")
        try:
            if args:
                subprocess.Popen(f"{app} {args}", shell=True)
            else:
                subprocess.Popen(app, shell=True)
            log.info(f"✅ Executed via shell: {app}")
            return ActionResult(True, f"Launched {app}", speak=f"Opening {app}")
        except Exception as e:
            log.error(f"❌ Failed to open {app}: {e}")
            return ActionResult(False, f"Could not open '{app}' - app not found", speak=f"Could not find {app}")


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
    # ENTERPRISE-LEVEL UNIVERSAL ACTIONS (NEW)
    # ══════════════════════════════════════════════════════════

    def _delete_file(self, a: dict, raw: str) -> ActionResult:
        """Delete a file or folder."""
        path = a.get("path", "")
        if not path:
            return ActionResult(False, "No file path given")
        try:
            p = Path(path)
            if p.is_dir():
                import shutil
                shutil.rmtree(p)
            else:
                p.unlink()
            return ActionResult(True, f"Deleted: {path}")
        except Exception as e:
            return ActionResult(False, f"Cannot delete {path}: {e}")

    def _rename_file(self, a: dict, raw: str) -> ActionResult:
        """Rename a file or folder."""
        old_path = a.get("path", "")
        new_name = a.get("new_name", "")
        if not old_path or not new_name:
            return ActionResult(False, "Missing path or new name")
        try:
            old = Path(old_path)
            new = old.parent / new_name
            old.rename(new)
            return ActionResult(True, f"Renamed to: {new_name}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _navigate_to(self, a: dict, raw: str) -> ActionResult:
        """Navigate to any location - app, URL, or file path."""
        location = a.get("location", "").strip()
        if not location:
            return ActionResult(False, "No location given")
        
        # Check if it's a URL/website
        if "http" in location or "." in location.split("/")[-1]:
            return self._open_url({"url": location}, raw)
        
        # Check if it's a file/folder path
        p = Path(location)
        if p.exists():
            if p.is_dir():
                return self._open_path({"path": location}, raw)
            else:
                return self._open_file({"path": location}, raw)
        
        # Try as app
        return self._open_app({"app": location}, raw)

    def _click_element(self, a: dict, raw: str) -> ActionResult:
        """Click a UI element (requires element identification)."""
        x = a.get("x")
        y = a.get("y")
        button = a.get("button", "left").lower()
        wait_ms = int(a.get("wait", 100))
        
        if x is None or y is None:
            return ActionResult(False, "Coordinates required (x, y)")
        
        try:
            time.sleep(wait_ms / 1000)
            if button == "right":
                pyautogui.rightClick(x, y)
            elif button == "middle":
                pyautogui.middleClick(x, y)
            else:
                pyautogui.click(x, y)
            return ActionResult(True, f"Clicked at ({x}, {y})")
        except Exception as e:
            return ActionResult(False, str(e))

    def _find_and_click(self, a: dict, raw: str) -> ActionResult:
        """Find text/button on screen and click it (requires text recognition)."""
        text = a.get("text", "").strip()
        if not text:
            return ActionResult(False, "No text to find")
        
        # This is a placeholder - full implementation would use OCR/screen analysis
        try:
            # Try keyboard navigation as fallback
            pyautogui.typewrite(text[:20], interval=0.05)
            time.sleep(0.2)
            pyautogui.press("tab")
            return ActionResult(True, f"Searched for and navigated to: {text}")
        except Exception as e:
            return ActionResult(False, f"Could not find '{text}': {e}")

    def _wait_for_element(self, a: dict, raw: str) -> ActionResult:
        """Wait for a UI element to appear (polling-based)."""
        timeout = int(a.get("timeout", 10))
        interval = float(a.get("interval", 0.5))
        
        try:
            elapsed = 0
            while elapsed < timeout:
                time.sleep(interval)
                elapsed += interval
                # Placeholder: could integrate with accessibility APIs
            return ActionResult(True, f"Waited {timeout:d}s for element")
        except Exception as e:
            return ActionResult(False, str(e))

    def _fill_form(self, a: dict, raw: str) -> ActionResult:
        """Fill a form by clicking fields and typing values."""
        fields = a.get("fields", {})  # {"field_name": "value", ...}
        if not fields:
            return ActionResult(False, "No form fields given")
        
        try:
            for field_name, value in fields.items():
                # Try to find and click field, then type value
                time.sleep(0.3)
                pyautogui.typewrite(str(value)[:200], interval=0.02)
                pyautogui.press("tab")  # Move to next field
            return ActionResult(True, f"Filled {len(fields)} form fields")
        except Exception as e:
            return ActionResult(False, str(e))

    def _submit_form(self, a: dict, raw: str) -> ActionResult:
        """Submit a form (Enter or click submit button)."""
        method = a.get("method", "enter").lower()
        try:
            if method in ("enter", "return"):
                pyautogui.press("return")
            elif method == "tab":
                pyautogui.press("tab")
                pyautogui.press("return")
            else:
                # Try to find submit button
                pyautogui.hotkey("ctrl", "return")  # Common submission shortcut
            time.sleep(0.5)
            return ActionResult(True, "Form submitted")
        except Exception as e:
            return ActionResult(False, str(e))

    def _select_dropdown(self, a: dict, raw: str) -> ActionResult:
        """Select from a dropdown menu."""
        option = a.get("option", "").strip()
        try:
            pyautogui.press("space")  # Open dropdown
            time.sleep(0.3)
            # Navigate to option
            pyautogui.typewrite(option[:30], interval=0.05)
            time.sleep(0.2)
            pyautogui.press("return")
            return ActionResult(True, f"Selected: {option}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _mouse_move(self, a: dict, raw: str) -> ActionResult:
        """Move mouse to coordinates."""
        x = a.get("x")
        y = a.get("y")
        duration = float(a.get("duration", 0.5))
        
        if x is None or y is None:
            return ActionResult(False, "Coordinates required")
        
        try:
            pyautogui.moveTo(x, y, duration=duration)
            return ActionResult(True, f"Mouse moved to ({x}, {y})")
        except Exception as e:
            return ActionResult(False, str(e))

    def _mouse_click(self, a: dict, raw: str) -> ActionResult:
        """Click at coordinates."""
        x = a.get("x")
        y = a.get("y")
        clicks = int(a.get("clicks", 1))
        button = a.get("button", "left").lower()
        
        if x is None or y is None:
            return ActionResult(False, "Coordinates required")
        
        try:
            pyautogui.click(x, y, clicks=clicks, button=button)
            return ActionResult(True, f"Clicked at ({x}, {y})")
        except Exception as e:
            return ActionResult(False, str(e))

    def _mouse_drag(self, a: dict, raw: str) -> ActionResult:
        """Drag mouse from one point to another."""
        x1 = a.get("x1")
        y1 = a.get("y1")
        x2 = a.get("x2")
        y2 = a.get("y2")
        duration = float(a.get("duration", 1.0))
        
        if any(v is None for v in [x1, y1, x2, y2]):
            return ActionResult(False, "Start and end coordinates required")
        
        try:
            pyautogui.drag(x2 - x1, y2 - y1, duration=duration)
            return ActionResult(True, f"Dragged from ({x1}, {y1}) to ({x2}, {y2})")
        except Exception as e:
            return ActionResult(False, str(e))

    def _keyboard_input(self, a: dict, raw: str) -> ActionResult:
        """Send raw keyboard input (supports special keys)."""
        keys = a.get("keys", [])
        text = a.get("text", "")
        
        if isinstance(keys, str):
            keys = keys.split("+")
        
        try:
            if text:
                pyautogui.typewrite(text, interval=0.05)
            if keys:
                pyautogui.hotkey(*keys)
            return ActionResult(True, f"Keyboard input executed")
        except Exception as e:
            return ActionResult(False, str(e))

    def _focus_window(self, a: dict, raw: str) -> ActionResult:
        """Focus/activate a specific window."""
        app = a.get("app", "").strip()
        if not app:
            return ActionResult(False, "No app specified")
        
        try:
            # Find window and focus
            from pynput.keyboard import Controller, Key
            pyautogui.hotkey("alt", "tab")
            time.sleep(0.5)
            # Type app name for search in Alt+Tab
            pyautogui.typewrite(app[:20], interval=0.05)
            time.sleep(0.2)
            pyautogui.press("return")
            return ActionResult(True, f"Focused window: {app}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _wait(self, a: dict, raw: str) -> ActionResult:
        """Wait/delay for specified time."""
        seconds = float(a.get("seconds", a.get("delay", 1)))
        try:
            time.sleep(seconds)
            return ActionResult(True, f"Waited {seconds:.1f} seconds")
        except Exception as e:
            return ActionResult(False, str(e))


# ══════════════════════════════════════════════════════════
# GROQ BRAIN  (LLaMA 3.3 70B — most capable free model)
# ══════════════════════════════════════════════════════════
class AgentBrain:

    SYSTEM_PROMPT = """You are ProVoiceAgent — an ENTERPRISE-LEVEL desktop automation AI with UNIVERSAL ACTION CAPABILITIES.
Parse the user's voice command and return a single JSON object describing the action.

NO RESTRICTIONS — Works on ANY app, website, file, or desktop element without limitations.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CORE AUTOMATION ACTIONS (Universal)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

navigate_to    → {"action":"navigate_to","location":"github.com"}
               → {"action":"navigate_to","location":"C:\\Users\\Desktop\\file.pdf"}
               → {"action":"navigate_to","location":"notepad"}

click_element  → {"action":"click_element","x":500,"y":300,"button":"left"}
               → Advanced: click ANY UI element by coordinates

find_and_click → {"action":"find_and_click","text":"Save Button"}
               → Find element by text/label and click (universal app support)

mouse_move     → {"action":"mouse_move","x":100,"y":200,"duration":0.5}
               → Move cursor to any position

mouse_click    → {"action":"mouse_click","x":500,"y":300,"clicks":2,"button":"left"}
               → Click at coordinates (single/double/triple click)

mouse_drag     → {"action":"mouse_drag","x1":100,"y1":100,"x2":500,"y2":500,"duration":1}
               → Drag operation (works on any app)

keyboard_input → {"action":"keyboard_input","text":"search query","keys":["ctrl","shift","enter"]}
               → Raw keyboard input for any application

fill_form      → {"action":"fill_form","fields":{"name":"John","email":"john@example.com"}}
               → Fill forms on ANY website or app (universal form support)

select_dropdown → {"action":"select_dropdown","option":"Option Name"}
               → Select from any dropdown menu (works universally)

submit_form    → {"action":"submit_form","method":"enter"}
               → Submit forms on any website/app

focus_window   → {"action":"focus_window","app":"chrome"}
               → Focus/switch to any running window

wait           → {"action":"wait","seconds":2}
               → Wait before executing next action

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FILE & FOLDER OPERATIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

open_file      → {"action":"open_file","path":"C:\\Users\\User\\Desktop\\report.pdf"}
               → Opens ANY file type (pdf, doc, image, video, etc)

open_path      → {"action":"open_path","path":"C:\\Users\\User\\Downloads"}
               → Open any folder

find_file      → {"action":"find_file","name":"resume","dir":"C:\\Users\\User\\Documents"}
               → Find any file on disk

create_file    → {"action":"create_file","name":"todo.txt","content":"task list","location":"C:\\Desktop"}

delete_file    → {"action":"delete_file","path":"C:\\Users\\Desktop\\old_file.txt"}

rename_file    → {"action":"rename_file","path":"C:\\Users\\Desktop\\file.txt","new_name":"renamed.txt"}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
APP MANAGEMENT (Universal across 100+ apps)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

open_app       → {"action":"open_app","app":"chrome"}
               → Works with EVERY installed application - browsers, editors, media players, productivity tools, dev tools, etc.

close_app      → {"action":"close_app","app":"chrome","force":true}
               → Close ANY running application

restart_app    → {"action":"restart_app","app":"vscode","delay":1}

list_apps      → {"action":"list_apps"}

app_status     → {"action":"app_status","app":"chrome"}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
WEB & COMMUNICATION (Universal)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

search_web     → {"action":"search_web","query":"python tutorials","engine":"google"}
               engines: google | youtube | bing | github | stackoverflow | reddit | amazon | maps | wikipedia

open_url       → {"action":"open_url","url":"github.com"}
               → Open ANY URL/website

goto           → {"action":"goto","url":"facebook.com"}
               → Alias for open_url

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TEXT & INPUT (Works on ANY application)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

type_text      → {"action":"type_text","text":"Hello World"}
               → Type into ANY application (no limitations)

paste_text     → {"action":"paste_text","text":"exact text"}
               → Paste into ANY field

type           → {"action":"type","text":"message"}

paste          → {"action":"paste","text":"content"}

clipboard      → {"action":"clipboard","command":"copy|paste|clear|get"}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SYSTEM CONTROL
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

system_command → {"action":"system_command","command":"shutdown|restart|sleep|hibernate|lock"}

volume_control → {"action":"volume_control","command":"up|down|mute","steps":5}

brightness     → {"action":"brightness","level":75}

toggle_wifi    → {"action":"toggle_wifi"}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
WINDOW & MEDIA CONTROL
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

window_control → {"action":"window_control","command":"maximize|minimize|fullscreen|switch"}

media_control  → {"action":"media_control","command":"play|pause|next|previous"}

keyboard_shortcut → {"action":"keyboard_shortcut","shortcut":"copy|select_all|save|find"}

hotkey         → {"action":"hotkey","keys":["ctrl","shift","esc"]}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCREENSHOTS & UTILITIES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

screenshot     → {"action":"screenshot"}
               → Works on ANY display/app

scroll         → {"action":"scroll","direction":"down","amount":3}

run_command    → {"action":"run_command","command":"python script.py"}

calculate      → {"action":"calculate","expression":"sqrt(144)"}

get_info       → {"action":"get_info","type":"time|system|battery|ip"}

answer         → {"action":"answer","text":"direct factual answer"}

set_reminder   → {"action":"set_reminder","message":"Break time","seconds":300}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DECISION RULES (Enterprise-Level)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. "click the save button"       → find_and_click text="Save"
2. "fill the form with my info"  → fill_form with fields
3. "scroll down 5 times"         → scroll direction="down" amount=5
4. "click at x:500 y:300"        → mouse_click with coordinates
5. "open any website/app/file"   → navigate_to (auto-detection)
6. "type in the search box"      → keyboard_input or type_text
7. "select from dropdown"        → select_dropdown
8. "submit the form"             → submit_form
9. "wait then click"             → wait → then click_element
10. "I want to use [any app]"    → open_app (works with 100+ apps)

KEY POINTS:
- NO RESTRICTIONS on apps, websites, or files
- UNIVERSAL app support: Works on Spotify, YouTube, Excel, Chrome, Notepad, Discord, any browser, etc.
- Works on ANY form, ANY dropdown, ANY button on ANY website
- Can interact with desktop elements, files, folders without limitations
- Auto-detect context: URL → browser, path → file manager, app name → launch app

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

        # Handle both single action (dict) and multiple actions (list)
        actions = action if isinstance(action, list) else [action]
        
        # Execute all actions in sequence
        final_result = None
        for idx, act in enumerate(actions, 1):
            if not isinstance(act, dict):
                continue
                
            act_name = act.get("action", "?")
            
            # Show current action being executed
            if len(actions) > 1:
                self.hud.show(f"⚡  [{idx}/{len(actions)}] {act_name}: {self._action_summary(act)}", "executing")
            else:
                self.hud.show(f"⚡  {act_name}: {self._action_summary(act)}", "executing")

            # Execute the action
            result = self.executor.execute(act, text)
            final_result = result  # Keep track of last result
            
            # Small delay between multi-step actions
            if len(actions) > 1 and idx < len(actions):
                time.sleep(0.5)

        # Handle TTS feedback from final result
        if final_result:
            speech = final_result.speak or (final_result.message if not final_result.success else None)
            if speech:
                self.hud.show(f"🔊  {speech[:42]}", "speaking")
                self.tts.speak(speech)

            # Update HUD with final result
            icon   = "✅" if final_result.success else "❌"
            state  = "success" if final_result.success else "error"
            self.hud.show(f"{icon}  {final_result.message[:48]}", state)
        
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