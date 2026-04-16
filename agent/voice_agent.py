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
import psutil
import winreg
import ctypes
import socket
import hashlib
from datetime import datetime, timedelta
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
from collections import defaultdict
import shutil

import pyperclip
import pyautogui
import requests
import speech_recognition as sr
import pyttsx3

# Add agent directory to sys.path for imports
sys.path.insert(0, str(Path(__file__).parent))

from config import AGENT_CONFIG

# Load .env and get GROQ_API_KEY
from dotenv import load_dotenv
load_dotenv()
load_dotenv(dotenv_path=Path(__file__).parent.parent / '.env')
GROQ_API_KEY = os.getenv('GROQ_API_KEY')

# ══════════════════════════════════════════════════════════
# LOGGING CONFIGURATION
# ══════════════════════════════════════════════════════════
LOG_DIR = Path.home() / ".provoiceagent" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
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
    speak: Optional[str] = None

@dataclass
class ConversationTurn:
    role: str
    content: str
    timestamp: datetime = field(default_factory=datetime.now)

@dataclass
class SystemState:
    """Complete snapshot of system state"""
    timestamp: datetime
    cpu_percent: float
    ram_percent: float
    disk_percent: float
    running_apps: List[str]
    recent_files: List[str]
    network_status: Dict[str, Any]
    battery_info: Optional[Dict[str, Any]]


# ══════════════════════════════════════════════════════════
# ADVANCED SYSTEM INTELLIGENCE
# ══════════════════════════════════════════════════════════
class SystemIntelligence:
    """Deep system analysis and context awareness"""
    
    def __init__(self):
        self.system_state_history = []
        self.app_usage_tracker = defaultdict(lambda: {"launches": 0, "total_time": 0, "last_used": None})
        self.file_access_log = []
        self.network_log = []
        self._cache_invalidate_time = 30  # seconds
        self._last_system_scan = None
        
    def get_complete_system_info(self) -> Dict[str, Any]:
        """Get comprehensive system information"""
        try:
            # CPU & Memory
            cpu_percent = psutil.cpu_percent(interval=0.5, percpu=True)
            ram = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            
            # Network info
            net_io = psutil.net_io_counters()
            network_interfaces = psutil.net_if_addrs()
            
            # Battery
            battery = None
            try:
                bat = psutil.sensors_battery()
                if bat:
                    battery = {
                        "percent": bat.percent,
                        "charging": bat.power_plugged,
                        "time_left": str(bat.secsleft) if bat.secsleft != psutil.POWER_TIME_UNLIMITED else "Plugged in"
                    }
            except:
                pass
            
            # Processes
            running_processes = {}
            for proc in psutil.process_iter(['pid', 'name', 'memory_percent']):
                try:
                    pinfo = proc.as_dict(attrs=['pid', 'name', 'memory_percent'])
                    if pinfo['memory_percent'] > 0.5:  # Only significant consumers
                        running_processes[pinfo['name']] = {
                            "pid": pinfo['pid'],
                            "memory_mb": round(pinfo['memory_percent'] * ram.total / 100 / 1024**2, 1)
                        }
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            return {
                "timestamp": datetime.now().isoformat(),
                "cpu": {
                    "percent_per_core": cpu_percent,
                    "count": psutil.cpu_count(),
                    "freq": psutil.cpu_freq().current if psutil.cpu_freq() else 0
                },
                "memory": {
                    "total_gb": round(ram.total / 1024**3, 1),
                    "used_gb": round(ram.used / 1024**3, 1),
                    "percent": ram.percent,
                    "available_gb": round(ram.available / 1024**3, 1)
                },
                "disk": {
                    "total_gb": round(disk.total / 1024**3, 1),
                    "used_gb": round(disk.used / 1024**3, 1),
                    "free_gb": round(disk.free / 1024**3, 1),
                    "percent": disk.percent
                },
                "network": {
                    "bytes_sent": net_io.bytes_sent,
                    "bytes_recv": net_io.bytes_recv,
                    "active_interfaces": list(network_interfaces.keys())
                },
                "battery": battery,
                "top_processes": running_processes,
                "process_count": len(list(psutil.process_iter())),
                "boot_time": datetime.fromtimestamp(psutil.boot_time()).isoformat()
            }
        except Exception as e:
            log.error(f"System info error: {e}")
            return {}
    
    def find_files_by_pattern(self, pattern: str, search_depth: int = 3) -> List[Dict[str, Any]]:
        """Find files matching pattern with metadata"""
        results = []
        try:
            for root, dirs, files in os.walk(Path.home(), topdown=True):
                # Limit depth
                if root.count(os.sep) - Path.home().as_posix().count(os.sep) > search_depth:
                    dirs.clear()
                    continue
                
                # Skip system folders
                dirs[:] = [d for d in dirs if d not in ['.git', '__pycache__', 'node_modules', '.venv']]
                
                for file in files:
                    if pattern.lower() in file.lower():
                        filepath = Path(root) / file
                        try:
                            stat = filepath.stat()
                            results.append({
                                "path": str(filepath),
                                "name": file,
                                "size_mb": round(stat.st_size / 1024**2, 2),
                                "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                                "type": filepath.suffix
                            })
                        except (PermissionError, OSError):
                            pass
                
                if len(results) >= 20:  # Limit results
                    break
        except Exception as e:
            log.warning(f"File search error: {e}")
        
        return sorted(results, key=lambda x: x['modified'], reverse=True)
    
    def get_recent_files(self, limit: int = 10) -> List[Dict[str, Any]]:
        """Get recently accessed/modified files"""
        try:
            # Check common directories
            search_dirs = [
                Path.home() / "Desktop",
                Path.home() / "Documents",
                Path.home() / "Downloads",
                Path.home() / "Pictures",
            ]
            
            recent = []
            for search_dir in search_dirs:
                if search_dir.exists():
                    for item in search_dir.rglob("*"):
                        if item.is_file():
                            try:
                                stat = item.stat()
                                recent.append({
                                    "path": str(item),
                                    "name": item.name,
                                    "modified": stat.st_mtime,
                                    "size_mb": round(stat.st_size / 1024**2, 2)
                                })
                            except (PermissionError, OSError):
                                pass
            
            # Sort by modification time
            recent.sort(key=lambda x: x['modified'], reverse=True)
            return recent[:limit]
        except Exception as e:
            log.warning(f"Recent files error: {e}")
            return []
    
    def analyze_disk_usage(self) -> Dict[str, Any]:
        """Analyze what's consuming disk space"""
        try:
            disk_analysis = {}
            
            # Check main directories
            check_dirs = {
                "Downloads": Path.home() / "Downloads",
                "Documents": Path.home() / "Documents",
                "Desktop": Path.home() / "Desktop",
                "AppData": Path.home() / "AppData",
            }
            
            for name, path in check_dirs.items():
                if path.exists():
                    total_size = sum(f.stat().st_size for f in path.rglob("*") if f.is_file())
                    disk_analysis[name] = round(total_size / 1024**3, 2)
            
            return disk_analysis
        except Exception as e:
            log.warning(f"Disk analysis error: {e}")
            return {}


# ══════════════════════════════════════════════════════════
# TTS ENGINE (Text-to-Speech)
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
        """Non-blocking TTS"""
        threading.Thread(target=self._say, args=(text,), daemon=True).start()

    def speak_sync(self, text: str):
        """Blocking TTS"""
        self._say(text)

    def _say(self, text: str):
        with self._lock:
            try:
                self._engine.say(text)
                self._engine.runAndWait()
            except Exception as e:
                log.error(f"TTS error: {e}")


# ══════════════════════════════════════════════════════════
# VOICE ENGINE (Speech Recognition)
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
# ADVANCED ACTION EXECUTOR (70+ Enterprise Actions)
# ══════════════════════════════════════════════════════════
class ActionExecutor:
    """Enterprise-level action executor with universal app support and advanced capabilities"""

    # Comprehensive app database
    APP_MAP: Dict[str, List[str]] = {
        # Browsers
        "chrome": ["chrome", r"C:\Program Files\Google\Chrome\Application\chrome.exe"],
        "firefox": ["firefox", r"C:\Program Files\Mozilla Firefox\firefox.exe"],
        "edge": ["msedge", r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"],
        "brave": ["brave", r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"],
        
        # Development
        "vscode": ["code"],
        "visual studio": ["devenv"],
        "intellij": ["idea"],
        "pycharm": ["pycharm"],
        "sublime": ["subl"],
        "notepad++": ["notepad++"],
        "git": ["git"],
        "postman": ["postman"],
        
        # Office & Productivity
        "word": ["winword"],
        "excel": ["excel"],
        "powerpoint": ["powerpnt"],
        "outlook": ["outlook"],
        "onenote": ["onenote"],
        "access": ["msaccess"],
        
        # Communication
        "discord": ["discord"],
        "telegram": ["telegram"],
        "whatsapp": ["WhatsApp"],
        "slack": ["slack"],
        "teams": ["teams"],
        "zoom": ["zoom"],
        "skype": ["skype"],
        
        # Media
        "spotify": ["spotify"],
        "vlc": ["vlc"],
        "adobe premiere": ["premiere"],
        "adobe photoshop": ["photoshop"],
        "blender": ["blender"],
        "gimp": ["gimp"],
        
        # System
        "notepad": ["notepad"],
        "calculator": ["calc"],
        "explorer": ["explorer"],
        "cmd": ["cmd"],
        "powershell": ["powershell"],
        "task manager": ["taskmgr"],
        "regedit": ["regedit"],
    }

    SEARCH_ENGINES: Dict[str, str] = {
        "google": "https://www.google.com/search?q={}",
        "youtube": "https://www.youtube.com/results?search_query={}",
        "bing": "https://www.bing.com/search?q={}",
        "github": "https://github.com/search?q={}",
        "stackoverflow": "https://stackoverflow.com/search?q={}",
        "reddit": "https://www.reddit.com/search/?q={}",
        "wikipedia": "https://en.wikipedia.org/wiki/Special:Search?search={}",
    }

    def __init__(self, tts: TTSEngine, system_intel: SystemIntelligence):
        self.tts = tts
        self.system_intel = system_intel
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.05
        self._app_cache = {}
        
        self._handlers = {
            # App Management (Advanced)
            "open_app": self._open_app,
            "close_app": self._close_app,
            "restart_app": self._restart_app,
            "list_apps": self._list_apps,
            "app_status": self._app_status,
            "switch_app": self._switch_app,
            "kill_app": self._close_app,
            
            # Web & Navigation
            "search_web": self._search_web,
            "open_url": self._open_url,
            "goto": self._open_url,
            "browser_action": self._browser_action,
            
            # Text Input & Clipboard
            "type_text": self._type_text,
            "paste_text": self._paste_text,
            "write_notepad": self._write_notepad,
            "clipboard": self._clipboard,
            
            # File Operations (Advanced)
            "open_file": self._open_file,
            "find_file": self._find_file,
            "create_file": self._create_file,
            "delete_file": self._delete_file,
            "rename_file": self._rename_file,
            "copy_file": self._copy_file,
            "move_file": self._move_file,
            "open_path": self._open_path,
            "file_info": self._file_info,
            "recent_files": self._recent_files,
            
            # System Control (Advanced)
            "system_command": self._system_command,
            "system_info": self._system_info,
            "disk_analysis": self._disk_analysis,
            "running_processes": self._running_processes,
            "process_control": self._process_control,
            
            # Volume & Media
            "volume_control": self._volume_control,
            "media_control": self._media_control,
            "brightness": self._brightness,
            
            # Window Management
            "window_control": self._window_control,
            "keyboard_shortcut": self._keyboard_shortcut,
            "hotkey": self._hotkey,
            
            # Advanced UI Automation
            "click_element": self._click_element,
            "find_and_click": self._find_and_click,
            "fill_form": self._fill_form,
            "submit_form": self._submit_form,
            "select_dropdown": self._select_dropdown,
            "scroll": self._scroll,
            "mouse_move": self._mouse_move,
            "mouse_click": self._mouse_click,
            "mouse_drag": self._mouse_drag,
            
            # Screenshots & Capture
            "screenshot": self._screenshot,
            "screenshot_region": self._screenshot_region,
            
            # Info & Queries
            "get_info": self._get_info,
            "calculate": self._calculate,
            "answer": self._answer,
            "analyze_system": self._analyze_system,
            
            # Utilities
            "run_command": self._run_command,
            "set_reminder": self._set_reminder,
            "toggle_wifi": self._toggle_wifi,
            "wait": self._wait,
        }
        
        # Build app cache
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

    def _cache_all_installed_apps(self):
        """Build comprehensive system app cache"""
        log.info("🔍 Building system app cache...")
        
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall")
            for i in range(winreg.QueryInfoKey(key)[0]):
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    subkey = winreg.OpenKey(key, subkey_name)
                    display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                    if display_name:
                        self._app_cache[display_name.lower()] = None
                except:
                    pass
        except Exception as e:
            log.warning(f"Registry scan error: {e}")
        
        # Scan file system
        search_paths = [
            r"C:\Program Files",
            r"C:\Program Files (x86)",
            str(Path.home() / "AppData" / "Local"),
        ]
        
        for base_path in search_paths:
            if not os.path.exists(base_path):
                continue
            try:
                for root, dirs, files in os.walk(base_path, topdown=True):
                    dirs[:] = dirs[:5]
                    for file in files:
                        if file.endswith(".exe"):
                            app_name = file[:-4].lower()
                            self._app_cache[app_name] = os.path.join(root, file)
            except (PermissionError, OSError):
                continue
        
        log.info(f"✅ App cache built: {len(self._app_cache)} apps indexed")

    def _find_processes_by_name(self, app: str) -> List[int]:
        """Intelligent process finder with fuzzy matching"""
        from difflib import SequenceMatcher
        
        pids = []
        app_clean = app.replace(" ", "").replace(".exe", "").lower()
        
        try:
            for p in psutil.process_iter(['pid', 'name']):
                try:
                    name = p.info['name'].lower()
                    name_clean = name.replace(" ", "").replace(".exe", "")
                    
                    if app == name or app_clean == name_clean:
                        pids.append(p.info['pid'])
                    elif app in name or name_clean in app_clean:
                        pids.append(p.info['pid'])
                    else:
                        ratio = SequenceMatcher(None, app_clean, name_clean).ratio()
                        if ratio > 0.7:
                            pids.append(p.info['pid'])
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
        except Exception as e:
            log.error(f"Process scan error: {e}")
        
        return pids

    def _kill_process_by_pid(self, pid: int, force: bool = True) -> bool:
        """Kill process by PID"""
        try:
            p = psutil.Process(pid)
            if force:
                p.kill()
            else:
                p.terminate()
            try:
                p.wait(timeout=2)
            except psutil.TimeoutExpired:
                if not force:
                    p.kill()
            return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
            return False

    def _find_app_executable(self, app: str) -> Optional[str]:
        """Find app executable from cache"""
        from difflib import SequenceMatcher
        
        if not app:
            return None
        
        app_lower = app.lower().strip()
        
        # Direct cache lookup
        if app_lower in self._app_cache:
            path = self._app_cache[app_lower]
            if path and os.path.exists(path):
                return path
        
        # Fuzzy match in cache
        best_match = None
        best_ratio = 0.7
        for cached_app, path in self._app_cache.items():
            ratio = SequenceMatcher(None, app_lower, cached_app).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = path
                log.info(f"Fuzzy matched: '{app}' → '{cached_app}' ({best_ratio:.0%})")
        
        if best_match and os.path.exists(best_match):
            return best_match
        
        return None

    # ════════════════════════════════════════════════════════
    # APP MANAGEMENT HANDLERS
    # ════════════════════════════════════════════════════════

    def _open_app(self, a: dict, raw: str) -> ActionResult:
        """Open any application"""
        # Accept both 'app' and 'app_name' from Brain
        app = (a.get("app") or a.get("app_name") or "").lower().strip()
        app_path = a.get("app_path", "")  # Brain may provide full path
        url = a.get("url", "")
        
        log.info(f"DEBUG _open_app: app='{app}', app_path='{app_path}'")
        
        if not app:
            return ActionResult(False, "No app specified")

        log.info(f"🚀 Opening: {app}")

        # ── PRIORITY 1: Use app_path if Brain provided it ───────────────
        if app_path:
            log.info(f"Using Brain-provided path: {app_path}")
            try:
                # Use os.startfile for Windows executables (native Windows way)
                os.startfile(app_path)
                log.info(f"✅ Launched via app_path with os.startfile()")
                return ActionResult(True, f"Opened {app}", speak=f"Opening {app}")
            except FileNotFoundError:
                log.warning(f"app_path not found: {app_path}")
            except Exception as e:
                log.warning(f"Failed to launch via app_path: {e}")

        # ── Web services (don't need executable) ────────────────────────
        web_services = {
            "youtube": "https://www.youtube.com",
            "facebook": "https://www.facebook.com",
            "twitter": "https://www.twitter.com",
            "instagram": "https://www.instagram.com",
            "tiktok": "https://www.tiktok.com",
            "netflix": "https://www.netflix.com",
            "gmail": "https://mail.google.com",
            "github": "https://www.github.com",
            "reddit": "https://www.reddit.com",
            "linkedin": "https://www.linkedin.com",
        }
        
        if app in web_services:
            try:
                webbrowser.open(web_services[app])
                return ActionResult(True, f"Opened {app}", speak=f"Opening {app}")
            except Exception as e:
                return ActionResult(False, str(e))

        # ── System shortcuts using os.startfile ──────────────────────────
        settings_shortcuts = {
            "settings": "ms-settings:",
            "wifi": "ms-settings:network-wifi",
            "bluetooth": "ms-settings:bluetooth",
            "display": "ms-settings:display",
            "sound": "ms-settings:sound",
            "camera": "microsoft.windows.camera:",
            "calculator": "calculator:",
        }
        
        if app in settings_shortcuts:
            try:
                os.startfile(settings_shortcuts[app])
                log.info(f"✅ Opened {app} via os.startfile()")
                return ActionResult(True, f"Opened {app}", speak=f"Opening {app}")
            except Exception as e:
                log.warning(f"Failed to open {app}: {e}")

        # ── PRIORITY 2: Search in cache ────────────────────────────────
        exe_path = None
        if app in self._app_cache:
            exe_path = self._app_cache[app]
        
        if exe_path and os.path.exists(exe_path):
            try:
                # Handle .lnk shortcuts
                if exe_path.lower().endswith(".lnk"):
                    os.startfile(exe_path)
                else:
                    # Use os.startfile for Windows apps (more reliable)
                    os.startfile(exe_path)
                log.info(f"✅ Launched from cache via os.startfile()")
                return ActionResult(True, f"Opened {app}", speak=f"Opening {app}")
            except Exception as e:
                log.warning(f"Failed to launch from cache: {e}")
        
        # ── PRIORITY 3: Shell execution (Windows searches PATH & registry) ─
        log.info(f"Attempting shell execution for: {app}")
        try:
            os.startfile(app)
            log.info(f"✅ Opened {app} via os.startfile()")
            return ActionResult(True, f"Launched {app}", speak=f"Opening {app}")
        except Exception as e:
            log.warning(f"os.startfile failed, trying shell: {e}")
            try:
                subprocess.Popen(app, shell=True)
                return ActionResult(True, f"Launched {app}", speak=f"Opening {app}")
            except Exception as e:
                return ActionResult(False, f"Could not open '{app}'")

    def _close_app(self, a: dict, raw: str) -> ActionResult:
        """Close any running application"""
        # Accept both 'app' and 'app_name' from Brain
        app = (a.get("app") or a.get("app_name") or "").lower()
        force = a.get("force", True)
        
        if not app:
            return ActionResult(False, "No app specified")

        log.info(f"Closing: {app}")
        
        pids = self._find_processes_by_name(app)
        
        if pids:
            killed = 0
            for pid in pids:
                if self._kill_process_by_pid(pid, force):
                    killed += 1
            
            if killed > 0:
                time.sleep(0.3)
                return ActionResult(True, f"Closed {app}", speak=f"Closed {app}")

        return ActionResult(False, f"Could not close {app}: app not running")

    def _restart_app(self, a: dict, raw: str) -> ActionResult:
        """Restart an application"""
        # Accept both 'app' and 'app_name' from Brain
        app = (a.get("app") or a.get("app_name") or "").lower().strip()
        delay = int(a.get("delay", 1))

        self._close_app({"app": app}, raw)
        time.sleep(delay)
        return self._open_app({"app": app}, raw)

    def _list_apps(self, a: dict, raw: str) -> ActionResult:
        """List running applications"""
        try:
            running = set()
            for p in psutil.process_iter(['name']):
                try:
                    running.add(p.info['name'])
                except psutil.NoSuchProcess:
                    pass
            
            top_apps = sorted(list(running))[:15]
            text = ", ".join(top_apps)
            return ActionResult(True, f"Running: {text}", speak=f"Found {len(running)} running apps")
        except Exception as e:
            return ActionResult(False, str(e))

    def _app_status(self, a: dict, raw: str) -> ActionResult:
        """Check if app is running"""
        # Accept both 'app' and 'app_name' from Brain
        app = (a.get("app") or a.get("app_name") or "").lower()
        
        pids = self._find_processes_by_name(app)
        if pids:
            return ActionResult(True, f"{app} is running", speak=f"{app} is currently running")
        
        return ActionResult(False, f"{app} is not running", speak=f"{app} is not currently running")

    def _switch_app(self, a: dict, raw: str) -> ActionResult:
        """Switch to another app"""
        # Accept both 'app' and 'app_name' from Brain
        app = a.get("app") or a.get("app_name") or ""
        try:
            pyautogui.hotkey("alt", "tab")
            time.sleep(0.3)
            pyautogui.typewrite(app[:15], interval=0.05)
            time.sleep(0.2)
            pyautogui.press("return")
            return ActionResult(True, f"Switched to {app}")
        except Exception as e:
            return ActionResult(False, str(e))

    # ════════════════════════════════════════════════════════
    # WEB & NAVIGATION HANDLERS
    # ════════════════════════════════════════════════════════

    def _search_web(self, a: dict, raw: str) -> ActionResult:
        """Search the web - handles multiple parameter formats from Brain"""
        # Brain provides multiple possible parameter names
        url = a.get("url") or a.get("navigation", "")
        browser = a.get("browser", "")
        app_path = a.get("app_path", "")
        app_name = a.get("app_name", "")
        query = a.get("query", raw)
        
        log.info(f"DEBUG search_web: url='{url}', browser='{browser}', app_path='{app_path}'")
        
        # Priority 1: Use Brain's direct URL (highest priority)
        if url:
            try:
                # Try to find browser path if only executable name was provided
                if browser and not app_path:
                    app_path = self._find_app_executable(browser.replace(".exe", ""))
                
                if app_path:
                    # Open URL in specific browser
                    log.info(f"Opening '{url}' in browser: {browser}")
                    subprocess.Popen([app_path, url], shell=False)
                    return ActionResult(True, f"Searching in {browser}", speak=f"Searching")
                else:
                    # Open in default browser
                    log.info(f"Opening URL in default browser: {url}")
                    webbrowser.open(url)
                    return ActionResult(True, f"Searching", speak=f"Searching")
            except Exception as e:
                log.warning(f"Failed to open URL: {e}")
        
        # Fallback: Use query-based search
        if not query:
            query = raw
        
        engine = a.get("engine", "google").lower()
        tpl = self.SEARCH_ENGINES.get(engine, self.SEARCH_ENGINES["google"])
        search_url = tpl.format(query.replace(" ", "+"))
        
        try:
            if app_path:
                # Open search URL in specific browser
                subprocess.Popen([app_path, search_url], shell=False)
                return ActionResult(True, f"Searching for '{query}'", speak=f"Searching")
            else:
                webbrowser.open(search_url)
                return ActionResult(True, f"Searched '{query}'", speak=f"Searching")
        except Exception as e:
            return ActionResult(False, str(e))

    def _open_url(self, a: dict, raw: str) -> ActionResult:
        """Open any URL - now supports opening in specific browser"""
        url = a.get("url", a.get("navigation", ""))
        app_path = a.get("app_path", "")
        app_name = a.get("app_name", "")
        
        if not url.startswith("http"):
            url = "https://" + url
        
        try:
            if app_path:
                # Open URL in specific browser
                log.info(f"Opening '{url}' in {app_name}")
                subprocess.Popen([app_path, url], shell=False)
                return ActionResult(True, f"Opened in {app_name}", speak=f"Opening in {app_name}")
            else:
                # Open in default browser
                webbrowser.open(url)
                return ActionResult(True, f"Opened {url}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _browser_action(self, a: dict, raw: str) -> ActionResult:
        """Browser automation"""
        action = a.get("action", "").lower()
        
        actions_map = {
            "back": ["alt", "left"],
            "forward": ["alt", "right"],
            "reload": ["f5"],
            "hard_reload": ["ctrl", "shift", "r"],
            "new_tab": ["ctrl", "t"],
            "close_tab": ["ctrl", "w"],
            "next_tab": ["ctrl", "tab"],
            "prev_tab": ["ctrl", "shift", "tab"],
            "fullscreen": ["f11"],
            "devtools": ["f12"],
        }
        
        keys = actions_map.get(action)
        if keys:
            pyautogui.hotkey(*keys)
            return ActionResult(True, f"Browser: {action}")
        
        return ActionResult(False, f"Unknown browser action: {action}")

    # ════════════════════════════════════════════════════════
    # FILE OPERATIONS HANDLERS
    # ════════════════════════════════════════════════════════

    def _open_file(self, a: dict, raw: str) -> ActionResult:
        """Open any file"""
        path = a.get("path", "")
        if not path:
            return ActionResult(False, "No file path given")
        
        try:
            os.startfile(path)
            return ActionResult(True, f"Opened: {path}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _find_file(self, a: dict, raw: str) -> ActionResult:
        """Find files by pattern"""
        name = a.get("name", "")
        search_dir = Path(a.get("dir", str(Path.home())))
        
        results = self.system_intel.find_files_by_pattern(name)
        
        if results:
            for r in results:
                log.info(f"  Found: {r['path']}")
            return ActionResult(True, f"Found {len(results)} match(es)")
        
        return ActionResult(False, f"No files found for '{name}'")

    def _create_file(self, a: dict, raw: str) -> ActionResult:
        """Create a new file"""
        name = a.get("name", f"note_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        content = a.get("content", "")
        loc = a.get("location", str(Path.home() / "Desktop"))
        
        try:
            path = Path(loc) / name
            path.write_text(content, encoding="utf-8")
            os.startfile(str(path))
            return ActionResult(True, f"Created: {path}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _delete_file(self, a: dict, raw: str) -> ActionResult:
        """Delete file or folder"""
        path = a.get("path", "")
        if not path:
            return ActionResult(False, "No path given")
        
        try:
            p = Path(path)
            if p.is_dir():
                shutil.rmtree(p)
            else:
                p.unlink()
            return ActionResult(True, f"Deleted: {path}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _rename_file(self, a: dict, raw: str) -> ActionResult:
        """Rename file or folder"""
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

    def _copy_file(self, a: dict, raw: str) -> ActionResult:
        """Copy file"""
        src = a.get("source", "")
        dst = a.get("destination", "")
        
        if not src or not dst:
            return ActionResult(False, "Source and destination required")
        
        try:
            shutil.copy2(src, dst)
            return ActionResult(True, f"Copied to: {dst}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _move_file(self, a: dict, raw: str) -> ActionResult:
        """Move file"""
        src = a.get("source", "")
        dst = a.get("destination", "")
        
        try:
            shutil.move(src, dst)
            return ActionResult(True, f"Moved to: {dst}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _open_path(self, a: dict, raw: str) -> ActionResult:
        """Open folder in explorer"""
        path = a.get("path", str(Path.home() / "Desktop"))
        
        try:
            subprocess.Popen(["explorer", path])
            return ActionResult(True, f"Opened: {path}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _file_info(self, a: dict, raw: str) -> ActionResult:
        """Get file information"""
        path = a.get("path", "")
        
        try:
            p = Path(path)
            if not p.exists():
                return ActionResult(False, "File not found")
            
            stat = p.stat()
            info = {
                "name": p.name,
                "size_mb": round(stat.st_size / 1024**2, 2),
                "created": datetime.fromtimestamp(stat.st_ctime).isoformat(),
                "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                "type": p.suffix
            }
            
            return ActionResult(True, json.dumps(info), data=info)
        except Exception as e:
            return ActionResult(False, str(e))

    def _recent_files(self, a: dict, raw: str) -> ActionResult:
        """Get recently accessed files"""
        limit = int(a.get("limit", 10))
        recent = self.system_intel.get_recent_files(limit)
        
        return ActionResult(True, f"Found {len(recent)} recent files", data=recent)

    # ════════════════════════════════════════════════════════
    # SYSTEM CONTROL HANDLERS
    # ════════════════════════════════════════════════════════

    def _system_command(self, a: dict, raw: str) -> ActionResult:
        """Execute system commands"""
        cmd = a.get("command", "").lower()
        delay = a.get("delay", 30)

        dispatch = {
            "shutdown": lambda: subprocess.run(["shutdown", "/s", "/t", str(delay)]),
            "shutdown_now": lambda: subprocess.run(["shutdown", "/s", "/t", "0"]),
            "restart": lambda: subprocess.run(["shutdown", "/r", "/t", str(delay)]),
            "restart_now": lambda: subprocess.run(["shutdown", "/r", "/t", "0"]),
            "sleep": lambda: subprocess.run(["rundll32.exe", "powrprof.dll,SetSuspendState", "0,1,0"]),
            "lock": lambda: subprocess.run(["rundll32.exe", "user32.dll,LockWorkStation"]),
            "logoff": lambda: subprocess.run(["shutdown", "/l"]),
        }

        fn = dispatch.get(cmd)
        if fn:
            fn()
            return ActionResult(True, f"System: {cmd}", speak=f"System will {cmd}")
        
        return ActionResult(False, f"Unknown command: {cmd}")

    def _system_info(self, a: dict, raw: str) -> ActionResult:
        """Get detailed system information"""
        info = self.system_intel.get_complete_system_info()
        
        # Format for speech
        speech = (f"CPU at {info['cpu']['percent_per_core'][0]:.0f}%, "
                 f"RAM {info['memory']['percent']:.0f}% used of {info['memory']['total_gb']} GB, "
                 f"Disk {info['disk']['percent']:.0f}% full")
        
        return ActionResult(True, json.dumps(info), speak=speech, data=info)

    def _disk_analysis(self, a: dict, raw: str) -> ActionResult:
        """Analyze disk usage"""
        analysis = self.system_intel.analyze_disk_usage()
        
        return ActionResult(True, json.dumps(analysis), data=analysis)

    def _running_processes(self, a: dict, raw: str) -> ActionResult:
        """Get running processes"""
        try:
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'memory_percent']):
                try:
                    if proc.info['memory_percent'] > 0.5:
                        processes.append({
                            "name": proc.info['name'],
                            "pid": proc.info['pid'],
                            "memory_percent": proc.info['memory_percent']
                        })
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            top_processes = sorted(processes, key=lambda x: x['memory_percent'], reverse=True)[:10]
            
            return ActionResult(True, f"Found {len(processes)} processes", data=top_processes)
        except Exception as e:
            return ActionResult(False, str(e))

    def _process_control(self, a: dict, raw: str) -> ActionResult:
        """Control specific process"""
        pid = a.get("pid")
        action = a.get("action", "kill").lower()
        
        try:
            p = psutil.Process(pid)
            if action == "kill":
                p.kill()
            elif action == "terminate":
                p.terminate()
            elif action == "suspend":
                p.suspend()
            elif action == "resume":
                p.resume()
            
            return ActionResult(True, f"Process {action}ed: {p.name()}")
        except Exception as e:
            return ActionResult(False, str(e))

    # ════════════════════════════════════════════════════════
    # TEXT & INPUT HANDLERS
    # ════════════════════════════════════════════════════════

    def _type_text(self, a: dict, raw: str) -> ActionResult:
        """Type text into any application"""
        text = a.get("text", raw)
        time.sleep(0.3)
        pyautogui.typewrite(text, interval=0.03)
        return ActionResult(True, f"Typed: {text[:60]}")

    def _paste_text(self, a: dict, raw: str) -> ActionResult:
        """Paste text"""
        text = a.get("text", raw)
        pyperclip.copy(text)
        time.sleep(0.15)
        pyautogui.hotkey("ctrl", "v")
        return ActionResult(True, "Pasted text")

    def _write_notepad(self, a: dict, raw: str) -> ActionResult:
        """Write to Notepad"""
        text = a.get("text", raw)
        subprocess.Popen(["notepad.exe"])
        time.sleep(1.8)
        self._paste_text({"text": text}, raw)
        return ActionResult(True, "Wrote to Notepad")

    def _clipboard(self, a: dict, raw: str) -> ActionResult:
        """Manage clipboard"""
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
            return ActionResult(True, "Got clipboard", speak=f"Clipboard contains: {content[:100]}")
        
        return ActionResult(False, f"Unknown clipboard command: {cmd}")

    # ════════════════════════════════════════════════════════
    # MULTIMEDIA & CONTROL HANDLERS
    # ════════════════════════════════════════════════════════

    def _volume_control(self, a: dict, raw: str) -> ActionResult:
        """Control system volume using nircmd or fallback to pyautogui"""
        # Brain sends command in parameters.action, fall back to command or action keys
        params = a.get("parameters", {})
        cmd = params.get("action") or a.get("command") or "up"
        cmd = cmd.lower() if cmd else "up"
        steps = int(a.get("steps", 5))

        log.info(f"Volume control: {cmd} ({steps} steps)")
        
        # Method 1: Try nircmd (more reliable for Windows)
        try:
            if cmd in ("mute", "toggle"):
                result = subprocess.run(
                    ["nircmd", "mutesysvolume", "1"], 
                    timeout=2, 
                    capture_output=True
                )
                if result.returncode == 0:
                    log.info("[OK] Muted via nircmd")
                    return ActionResult(True, "Muted", speak="Muted volume")
            elif cmd in ("unmute",):
                result = subprocess.run(
                    ["nircmd", "mutesysvolume", "0"], 
                    timeout=2, 
                    capture_output=True
                )
                if result.returncode == 0:
                    log.info("[OK] Unmuted via nircmd")
                    return ActionResult(True, "Unmuted", speak="Unmuted volume")
            elif cmd in ("up", "increase"):
                for i in range(steps):
                    subprocess.run(
                        ["nircmd", "changesysvolume", "5000"], 
                        timeout=2, 
                        capture_output=True
                    )
                log.info(f"[OK] Volume up: +{steps} steps via nircmd")
                return ActionResult(True, f"Volume +{steps}", speak="Volume increased")
            elif cmd in ("down", "decrease"):
                for i in range(steps):
                    subprocess.run(
                        ["nircmd", "changesysvolume", "-5000"], 
                        timeout=2, 
                        capture_output=True
                    )
                log.info(f"[OK] Volume down: -{steps} steps via nircmd")
                return ActionResult(True, f"Volume -{steps}", speak="Volume decreased")
        except FileNotFoundError:
            log.debug("nircmd not found, falling back to pyautogui")
        except Exception as e:
            log.debug(f"nircmd failed: {e}, falling back to pyautogui")
        
        # Fallback: pyautogui keyboard events
        try:
            if cmd in ("up", "increase"):
                for _ in range(steps):
                    pyautogui.press("volumeup")
                log.warning(f"[WARN] Volume up via pyautogui (fallback)")
            elif cmd in ("down", "decrease"):
                for _ in range(steps):
                    pyautogui.press("volumedown")
                log.warning(f"[WARN] Volume down via pyautogui (fallback)")
            elif cmd in ("mute", "toggle"):
                pyautogui.press("volumemute")
                log.warning(f"[WARN] Volume mute via pyautogui (fallback)")
            
            time.sleep(0.3)
            return ActionResult(True, f"Volume {cmd}", speak=f"Volume {cmd}")
        except Exception as e:
            log.error(f"[ERROR] All volume methods failed: {e}")
            return ActionResult(False, f"Volume control failed: {e}")

    def _media_control(self, a: dict, raw: str) -> ActionResult:
        """Control media playback"""
        cmd = a.get("command", "").lower()
        
        key_map = {
            "play": "playpause",
            "pause": "playpause",
            "play_pause": "playpause",
            "next": "nexttrack",
            "previous": "prevtrack",
            "prev": "prevtrack",
        }
        
        key = key_map.get(cmd)
        if key:
            pyautogui.press(key)
            return ActionResult(True, f"Media: {cmd}")
        
        return ActionResult(False, f"Unknown media command: {cmd}")

    def _brightness(self, a: dict, raw: str) -> ActionResult:
        """Control brightness"""
        level = a.get("level", 70)
        
        try:
            subprocess.run([
                "PowerShell", "-NoProfile", "-Command",
                f"(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods)"
                f".WmiSetBrightness(1,{level})"
            ], capture_output=True, check=False)
            
            return ActionResult(True, f"Brightness set to {level}%")
        except Exception as e:
            return ActionResult(False, str(e))

    # ════════════════════════════════════════════════════════
    # WINDOW & KEYBOARD HANDLERS
    # ════════════════════════════════════════════════════════

    def _window_control(self, a: dict, raw: str) -> ActionResult:
        """Control windows"""
        cmd = a.get("command", "").lower()
        
        actions_map = {
            "minimize": lambda: pyautogui.hotkey("win", "down"),
            "maximize": lambda: pyautogui.hotkey("win", "up"),
            "close": lambda: pyautogui.hotkey("alt", "f4"),
            "fullscreen": lambda: pyautogui.press("f11"),
            "switch": lambda: pyautogui.hotkey("alt", "tab"),
        }
        
        fn = actions_map.get(cmd)
        if fn:
            time.sleep(0.2)
            fn()
            return ActionResult(True, f"Window: {cmd}")
        
        return ActionResult(False, f"Unknown window command: {cmd}")

    def _keyboard_shortcut(self, a: dict, raw: str) -> ActionResult:
        """Execute keyboard shortcuts - supports Win+L, Ctrl+C, Alt+Tab, etc."""
        # Brain sends shortcut in parameters, fall back to shortcut key
        params = a.get("parameters", {})
        shortcut = params.get("action") or params.get("shortcut") or a.get("shortcut", "")
        shortcut = shortcut.strip() if shortcut else ""
        
        if not shortcut:
            return ActionResult(False, "No shortcut specified")
        
        log.info(f"Executing keyboard shortcut: {shortcut}")
        
        # Hardcoded shortcuts with alternative names
        builtin = {
            "copy": ["ctrl", "c"],
            "paste": ["ctrl", "v"],
            "cut": ["ctrl", "x"],
            "undo": ["ctrl", "z"],
            "save": ["ctrl", "s"],
            "select_all": ["ctrl", "a"],
            "select all": ["ctrl", "a"],
            "find": ["ctrl", "f"],
            "new": ["ctrl", "n"],
            "open": ["ctrl", "o"],
            "print": ["ctrl", "p"],
            "refresh": ["f5"],
            "devtools": ["f12"],
            "lock": ["win", "l"],
            "lock screen": ["win", "l"],
            "lock_screen": ["win", "l"],
            "win+l": ["win", "l"],
            "window+l": ["win", "l"],
        }
        
        # Check builtin names first
        shortcut_lower = shortcut.lower()
        keys = builtin.get(shortcut_lower)
        if keys:
            try:
                # Special handling for screen lock - use Windows API
                if shortcut_lower in ("lock", "lock screen", "lock_screen", "win+l", "window+l"):
                    try:
                        user32 = ctypes.windll.user32
                        user32.LockWorkStation()
                        log.info(f"[OK] Locked screen via Windows API")
                        return ActionResult(True, "Screen locked", speak="Screen locked")
                    except Exception as e:
                        log.debug(f"Windows API lock failed: {e}, trying hotkey")
                
                time.sleep(0.2)
                pyautogui.hotkey(*keys)
                time.sleep(0.5)  # Wait for action to complete
                log.info(f"[OK] Executed shortcut: {shortcut}")
                return ActionResult(True, f"Shortcut: {shortcut}", speak=f"Executed")
            except Exception as e:
                log.error(f"Failed to execute: {e}")
                return ActionResult(False, f"Failed: {e}")
        
        # Parse Win+L, Ctrl+C, Alt+Tab format
        try:
            parts = shortcut.replace(" ", "").split("+")
            keys_to_press = []
            
            key_map = {
                "win": "win", "windows": "win", "⊞": "win",
                "ctrl": "ctrl", "control": "ctrl",
                "alt": "alt",
                "shift": "shift",
            }
            
            for part in parts:
                part_lower = part.lower()
                if part_lower in key_map:
                    keys_to_press.append(key_map[part_lower])
                else:
                    keys_to_press.append(part_lower)
            
            if not keys_to_press:
                return ActionResult(False, f"Invalid format: {shortcut}")
            
            log.info(f"Pressing keys: {keys_to_press}")
            
            # Special case: Win+L lock screen via Windows API
            if len(keys_to_press) == 2 and "win" in keys_to_press and "l" in keys_to_press:
                try:
                    user32 = ctypes.windll.user32
                    user32.LockWorkStation()
                    log.info(f"[OK] Screen locked via Windows API")
                    return ActionResult(True, "Screen locked", speak="Screen locked")
                except Exception as e:
                    log.debug(f"Windows API lock failed: {e}, trying hotkey")
            
            time.sleep(0.2)
            pyautogui.hotkey(*keys_to_press)
            time.sleep(0.5)
            log.info(f"[OK] Executed: {shortcut}")
            return ActionResult(True, f"Executed: {shortcut}", speak=f"Done")
        
        except Exception as e:
            log.error(f"Shortcut parsing error: {e}")
            return ActionResult(False, f"Error: {e}")

    def _hotkey(self, a: dict, raw: str) -> ActionResult:
        """Execute raw hotkey"""
        keys = a.get("keys", [])
        
        if isinstance(keys, str):
            keys = keys.split("+")
        
        if keys:
            time.sleep(0.2)
            pyautogui.hotkey(*keys)
            return ActionResult(True, f"Pressed: {'+'.join(keys)}")
        
        return ActionResult(False, "No keys given")

    # ════════════════════════════════════════════════════════
    # UI AUTOMATION HANDLERS
    # ════════════════════════════════════════════════════════

    def _click_element(self, a: dict, raw: str) -> ActionResult:
        """Click at coordinates"""
        x = a.get("x")
        y = a.get("y")
        button = a.get("button", "left").lower()
        
        if x is None or y is None:
            return ActionResult(False, "Coordinates required")
        
        try:
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
        """Find text and click"""
        text = a.get("text", "").strip()
        
        try:
            pyautogui.typewrite(text[:20], interval=0.05)
            time.sleep(0.2)
            pyautogui.press("tab")
            return ActionResult(True, f"Found and navigated to: {text}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _fill_form(self, a: dict, raw: str) -> ActionResult:
        """Fill form fields"""
        fields = a.get("fields", {})
        
        if not fields:
            return ActionResult(False, "No fields given")
        
        try:
            for field_name, value in fields.items():
                time.sleep(0.3)
                pyautogui.typewrite(str(value)[:200], interval=0.02)
                pyautogui.press("tab")
            
            return ActionResult(True, f"Filled {len(fields)} fields")
        except Exception as e:
            return ActionResult(False, str(e))

    def _submit_form(self, a: dict, raw: str) -> ActionResult:
        """Submit form"""
        method = a.get("method", "enter").lower()
        
        try:
            if method in ("enter", "return"):
                pyautogui.press("return")
            elif method == "tab":
                pyautogui.press("tab")
                pyautogui.press("return")
            
            time.sleep(0.5)
            return ActionResult(True, "Form submitted")
        except Exception as e:
            return ActionResult(False, str(e))

    def _select_dropdown(self, a: dict, raw: str) -> ActionResult:
        """Select from dropdown"""
        option = a.get("option", "").strip()
        
        try:
            pyautogui.press("space")
            time.sleep(0.3)
            pyautogui.typewrite(option[:30], interval=0.05)
            time.sleep(0.2)
            pyautogui.press("return")
            return ActionResult(True, f"Selected: {option}")
        except Exception as e:
            return ActionResult(False, str(e))

    def _scroll(self, a: dict, raw: str) -> ActionResult:
        """Scroll"""
        direction = a.get("direction", "down").lower()
        amount = int(a.get("amount", 3))
        
        factor = 3 if direction == "up" else -3
        pyautogui.scroll(factor * amount)
        
        return ActionResult(True, f"Scrolled {direction}")

    def _mouse_move(self, a: dict, raw: str) -> ActionResult:
        """Move mouse"""
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
        """Click at coordinates"""
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
        """Drag mouse"""
        x1 = a.get("x1")
        y1 = a.get("y1")
        x2 = a.get("x2")
        y2 = a.get("y2")
        duration = float(a.get("duration", 1.0))
        
        if any(v is None for v in [x1, y1, x2, y2]):
            return ActionResult(False, "Coordinates required")
        
        try:
            pyautogui.drag(x2 - x1, y2 - y1, duration=duration)
            return ActionResult(True, f"Dragged from ({x1}, {y1}) to ({x2}, {y2})")
        except Exception as e:
            return ActionResult(False, str(e))

    # ════════════════════════════════════════════════════════
    # SCREENSHOT HANDLERS
    # ════════════════════════════════════════════════════════

    def _screenshot(self, a: dict, raw: str) -> ActionResult:
        """Take screenshot"""
        dest = Path(a.get("path", str(Path.home() / "Pictures" / "Screenshots")))
        dest.mkdir(parents=True, exist_ok=True)
        
        name = f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        path = dest / name
        
        time.sleep(0.6)
        pyautogui.screenshot().save(str(path))
        
        return ActionResult(True, f"Screenshot saved: {path}", speak="Screenshot saved")

    def _screenshot_region(self, a: dict, raw: str) -> ActionResult:
        """Screenshot specific region"""
        x = a.get("x", 0)
        y = a.get("y", 0)
        width = a.get("width", 400)
        height = a.get("height", 300)
        
        dest = Path(a.get("path", str(Path.home() / "Pictures")))
        dest.mkdir(parents=True, exist_ok=True)
        
        name = f"region_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        path = dest / name
        
        img = pyautogui.screenshot(region=(x, y, width, height))
        img.save(str(path))
        
        return ActionResult(True, f"Region screenshot saved: {path}")

    # ════════════════════════════════════════════════════════
    # INFO & UTILITY HANDLERS
    # ════════════════════════════════════════════════════════

    def _get_info(self, a: dict, raw: str) -> ActionResult:
        """Get system information"""
        info_type = a.get("type", "time").lower()

        if info_type in ("time", "clock"):
            text = datetime.now().strftime("It is %I:%M %p")
            return ActionResult(True, text, speak=text)

        elif info_type == "date":
            text = datetime.now().strftime("Today is %A, %B %d, %Y")
            return ActionResult(True, text, speak=text)

        elif info_type == "battery":
            try:
                bat = psutil.sensors_battery()
                if bat:
                    status = "charging" if bat.power_plugged else "on battery"
                    text = f"Battery at {bat.percent:.0f}%, {status}"
                    return ActionResult(True, text, speak=text)
            except:
                pass
            return ActionResult(False, "Battery info unavailable")

        return ActionResult(False, f"Unknown info type: {info_type}")

    def _calculate(self, a: dict, raw: str) -> ActionResult:
        """Calculate mathematical expression"""
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
            return ActionResult(True, text, speak=text)
        except Exception as e:
            return ActionResult(False, f"Calculation error: {e}")

    def _answer(self, a: dict, raw: str) -> ActionResult:
        """Speak answer"""
        answer = a.get("text", "")
        return ActionResult(True, answer, speak=answer)

    def _analyze_system(self, a: dict, raw: str) -> ActionResult:
        """Deep system analysis"""
        info = self.system_intel.get_complete_system_info()
        analysis = {
            "health": "Good" if info['cpu']['percent_per_core'][0] < 80 else "Warning",
            "recommendations": []
        }
        
        if info['memory']['percent'] > 85:
            analysis['recommendations'].append("Close some applications to free memory")
        if info['disk']['percent'] > 90:
            analysis['recommendations'].append("Consider freeing up disk space")
        
        return ActionResult(True, json.dumps(analysis), data=analysis)

    def _run_command(self, a: dict, raw: str) -> ActionResult:
        """Run shell command"""
        cmd = a.get("command", "")
        
        if not cmd:
            return ActionResult(False, "No command given")
        
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=15)
            output = result.stdout.strip() or result.stderr.strip() or "Done"
            return ActionResult(True, output[:200], speak="Command executed")
        except subprocess.TimeoutExpired:
            return ActionResult(False, "Command timed out")
        except Exception as e:
            return ActionResult(False, str(e))

    def _set_reminder(self, a: dict, raw: str) -> ActionResult:
        """Set system reminder"""
        message = a.get("message", "Reminder!")
        seconds = int(a.get("seconds", 60))

        def _remind():
            time.sleep(seconds)
            pyautogui.alert(text=message, title="⏰ ProVoiceAgent Reminder", button="OK")

        threading.Thread(target=_remind, daemon=True).start()
        
        mins = seconds // 60
        text = f"Reminder set for {mins} minute{'s' if mins != 1 else ''}"
        return ActionResult(True, text, speak=text)

    def _toggle_wifi(self, a: dict, raw: str) -> ActionResult:
        """Toggle Wi-Fi"""
        try:
            subprocess.run(
                ["netsh", "interface", "set", "interface", "Wi-Fi", "enable"],
                capture_output=True, check=False
            )
            return ActionResult(True, "Wi-Fi toggled")
        except Exception as e:
            return ActionResult(False, str(e))

    def _wait(self, a: dict, raw: str) -> ActionResult:
        """Wait/delay"""
        seconds = float(a.get("seconds", a.get("delay", 1)))
        try:
            time.sleep(seconds)
            return ActionResult(True, f"Waited {seconds:.1f} seconds")
        except Exception as e:
            return ActionResult(False, str(e))


# ══════════════════════════════════════════════════════════
# ADVANCED GROQ BRAIN (LLaMA 3.3 70B)
# ══════════════════════════════════════════════════════════
class AgentBrain:
    """Enterprise-level AI brain with advanced reasoning"""

    SYSTEM_PROMPT = """You are ProVoiceAgent™ — an ENTERPRISE-GRADE AI desktop automation assistant with unrestricted capabilities.

You have COMPLETE access to:
• User's entire filesystem, folders, and files
• All installed applications and system processes
• Network, system information, and resources
• Deep file searching and analysis
• Advanced web integration and automation
• Complex multi-step workflows

CORE CAPABILITIES:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. FILE MANAGEMENT
   open_file → Open any file (docs, videos, images, etc.)
   find_file → Deep search for files with pattern matching
   create_file → Create files with content
   delete_file, rename_file, copy_file, move_file → Full file control
   file_info → Detailed file metadata
   recent_files → Access recently modified files
   open_path → Open folders in explorer

2. APPLICATION MANAGEMENT
   open_app → Launch ANY installed app (100+ supported)
   close_app → Close running applications
   restart_app → Restart applications
   switch_app → Switch between running apps
   list_apps → Show all running applications
   app_status → Check if app is running

3. WEB & INTERNET
   search_web → Search Google, YouTube, GitHub, etc.
   open_url → Open any website
   browser_action → Control browser (back, forward, reload, etc.)

4. TEXT & INPUT
   type_text → Type into any application
   paste_text → Paste content
   write_notepad → Create text in Notepad
   clipboard → Manage clipboard

5. SYSTEM INTELLIGENCE
   system_info → Complete system snapshot (CPU, RAM, disk, processes)
   disk_analysis → See what's consuming disk space
   running_processes → List all running processes
   process_control → Kill, suspend, or resume processes
   analyze_system → Deep system analysis with recommendations

6. MEDIA & CONTROL
   volume_control → Adjust volume, mute
   media_control → Play, pause, next, previous
   brightness → Control display brightness
   window_control → Minimize, maximize, close windows
   keyboard_shortcut → Execute shortcuts (copy, paste, save, etc.)
   hotkey → Custom keyboard combinations

7. UI AUTOMATION (Universal - works on any app/website)
   click_element → Click at exact coordinates
   find_and_click → Find text/element and click
   fill_form → Fill form fields automatically
   submit_form → Submit forms
   select_dropdown → Choose from dropdowns
   scroll → Scroll up/down
   mouse_move, mouse_click, mouse_drag → Full mouse control

8. UTILITIES
   screenshot → Capture full screen
   screenshot_region → Capture specific area
   calculate → Solve mathematical expressions
   set_reminder → Set system reminders
   get_info → Get time, date, battery status
   run_command → Execute PowerShell/CMD commands
   wait → Add delays between actions

DECISION RULES (Enterprise Logic):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• "Open [any file]" → open_file with path
• "Find [anything]" → find_file with search pattern
• "Launch [app name]" → open_app
• "Search for [query]" → search_web
• "Click [button/text]" → click_element or find_and_click
• "Fill out form" → fill_form with field mapping
• "Show me [data]" → system_info or analyze_system
• "What's using disk" → disk_analysis
• "Restart [app]" → restart_app
• Multiple actions → Return array of action objects

IMPORTANT NOTES:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✓ NO RESTRICTIONS - works on any desktop element
✓ INTELLIGENT MATCHING - understands context and variations
✓ MULTI-STEP WORKFLOWS - can chain actions together
✓ DEEP SYSTEM ACCESS - full visibility into user's desktop
✓ COMPLETE INFORMATION - knows user's files, apps, system state
✓ ENTERPRISE QUALITY - professional-grade automation

Return ONLY valid JSON (single object or array of objects). No explanation or markdown.
"""

    def __init__(self):
        self.api_url = "https://api.groq.com/openai/v1/chat/completions"
        self.headers = {
            "Authorization": f"Bearer {GROQ_API_KEY}",
            "Content-Type": "application/json",
        }
        self.history: List[ConversationTurn] = []
        self.max_hist = 8

    def parse(self, user_text: str, system_context: Dict[str, Any] = None) -> Dict[str, Any]:
        """Parse user command with system context awareness"""
        context = ""
        
        if system_context:
            context = f"\n\nSystem Context: CPU {system_context.get('cpu', 'N/A')}%, "
            context += f"RAM {system_context.get('ram', 'N/A')}%, "
            context += f"Running Apps: {system_context.get('apps', 'N/A')}"
        
        if self.history:
            recent = self.history[-4:]
            context += "\n\nRecent Context:\n" + "\n".join(
                f"  {t.role}: {t.content}" for t in recent
            )

        payload = {
            "model": AGENT_CONFIG.get("model", "llama-3.3-70b-versatile"),
            "messages": [
                {"role": "system", "content": self.SYSTEM_PROMPT},
                {"role": "user", "content": f"Voice command: {user_text}{context}\n\nJSON:"},
            ],
            "max_tokens": 500,
            "temperature": 0.05,
            "response_format": {"type": "json_object"},
        }

        try:
            resp = requests.post(self.api_url, headers=self.headers, json=payload, timeout=15)
            resp.raise_for_status()
            raw = resp.json()["choices"][0]["message"]["content"].strip()
            action = json.loads(raw)
            log.info(f"Brain → {action}")

            # Store turns
            self.history.append(ConversationTurn("user", user_text))
            self.history.append(ConversationTurn("assistant", json.dumps(action)))
            if len(self.history) > self.max_hist * 2:
                self.history = self.history[-self.max_hist * 2:]

            return action

        except json.JSONDecodeError as e:
            log.error(f"JSON parse fail: {e}")
        except requests.HTTPError as e:
            log.error(f"Groq HTTP error: {e.response.status_code}")
        except requests.RequestException as e:
            log.error(f"Groq request error: {e}")
        except Exception as e:
            log.error(f"Brain error: {e}", exc_info=True)

        return {"action": "paste_text", "text": user_text}


# ══════════════════════════════════════════════════════════
# FLOATING STATUS HUD
# ══════════════════════════════════════════════════════════
class StatusHUD:
    """Modern floating status display"""
    
    COLORS = {
        "ready": "#00ff88",
        "listening": "#00ff88",
        "processing": "#ffaa00",
        "executing": "#00aaff",
        "success": "#00ff88",
        "error": "#ff4455",
        "speaking": "#cc88ff",
    }

    def __init__(self):
        self._q = queue.Queue()
        self._running = False
        self._thread = threading.Thread(target=self._run, daemon=True)

    def start(self):
        self._thread.start()
        time.sleep(0.4)

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
        w, h = 420, 60
        root.geometry(f"{w}x{h}+{sw - w - 16}+{sh - h - 50}")

        dot = tk.Label(root, text="◉", font=("Segoe UI", 16), fg="#00ff88", bg="#0d0d1a")
        dot.pack(side="left", padx=(10, 6), pady=6)

        tk.Label(root, text="ProVoiceAgent™", font=("Segoe UI", 9, "bold"), fg="#5566ff", bg="#0d0d1a").place(x=40, y=4)

        lbl = tk.Label(root, text="Initializing…", font=("Segoe UI", 10), fg="white", bg="#0d0d1a", anchor="w")
        lbl.place(x=40, y=24, width=360)

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
    """Enterprise-grade voice automation agent"""

    STATES = {
        "STOPPED": 0,
        "IDLE": 1,
        "ACTIVE": 2,
        "EXECUTING": 3,
    }

    EXIT_PHRASES = {"goodbye", "stop", "quit", "exit", "logout", "shutdown", "off"}
    ACTIVATION_PHRASES = {"hello", "activate", "wake", "start", "ready", "listen"}

    def __init__(self):
        log.info("╔════════════════════════════════════════════╗")
        log.info("║   ProVoiceAgent™ ENTERPRISE EDITION       ║")
        log.info("║   Advanced Desktop Automation AI           ║")
        log.info("╚════════════════════════════════════════════╝")

        self.system_intel = SystemIntelligence()
        self.tts = TTSEngine()
        self.voice = VoiceEngine(self.tts)
        self.brain = AgentBrain()
        self.executor = ActionExecutor(self.tts, self.system_intel)
        self.hud = StatusHUD()
        
        self._state = self.STATES["STOPPED"]
        self._running = False
        
        log.info(f"Agent initialized - State: STOPPED")

    def _get_state_name(self) -> str:
        for name, value in self.STATES.items():
            if value == self._state:
                return name
        return "UNKNOWN"

    def start(self):
        """Main agent entry point"""
        self.hud.start()
        self._state = self.STATES["IDLE"]
        
        self.hud.show("🛑  OFFLINE - Say 'Start Lucifer'", "ready")
        self.tts.speak_sync("Lucifer is offline. Say 'Start Lucifer' to activate.")

        self._running = True
        log.info("Waiting for startup command...")

        while self._running:
            try:
                if self._state == self.STATES["IDLE"]:
                    self._wait_for_startup()
                elif self._state == self.STATES["ACTIVE"]:
                    self._wait_for_activation()
                elif self._state == self.STATES["EXECUTING"]:
                    self._execute_command()
                else:
                    time.sleep(0.5)
            except KeyboardInterrupt:
                self._shutdown()
            except Exception as e:
                log.error(f"Cycle error: {e}", exc_info=True)
                self.hud.show(f"❌ Error: {str(e)[:40]}", "error")
                time.sleep(1)

    def _wait_for_startup(self):
        """Wait for startup command"""
        self.hud.show("🛑  OFFLINE - Say 'Start Lucifer'", "ready")
        
        text = self.voice.listen(timeout=30, phrase_limit=10)
        if not text:
            return

        text_lower = text.lower()
        if "start" in text_lower:
            self._boot_agent()
        elif any(phrase in text_lower for phrase in self.EXIT_PHRASES):
            self._shutdown()

    def _boot_agent(self):
        """Boot the agent"""
        self._state = self.STATES["ACTIVE"]
        self.hud.show("⚡  ONLINE - Lucifer ready!", "processing")
        self.tts.speak("Lucifer is now online. Say 'Hello Lucifer' for commands.")
        log.info("Agent BOOTED")
        time.sleep(0.8)

    def _wait_for_activation(self):
        """Wait for command activation"""
        self.hud.show("👀  ONLINE (say 'Hello Lucifer')", "ready")
        
        text = self.voice.listen(timeout=30, phrase_limit=10)
        if not text:
            return

        text_lower = text.lower()
        if any(phrase in text_lower for phrase in self.EXIT_PHRASES):
            self._shutdown()
        elif any(phrase in text_lower for phrase in self.ACTIVATION_PHRASES):
            self._activate_for_command()

    def _activate_for_command(self):
        """Activate for command execution"""
        self._state = self.STATES["EXECUTING"]
        self.hud.show("🎤  LISTENING for command", "processing")
        self.tts.speak("Ready. What do you need?")

    def _execute_command(self):
        """Execute voice command"""
        self.hud.show("🎤  LISTENING for command…", "listening")
        
        text = self.voice.listen(timeout=15, phrase_limit=20)
        if not text:
            self._return_to_active("No command")
            return

        text_lower = text.lower()
        if any(phrase in text_lower for phrase in self.EXIT_PHRASES):
            self._shutdown()
            return

        log.info(f"📝  Command: \"{text}\"")
        self.hud.show(f"🧠  Processing…", "processing")

        # Get system context
        sys_info = self.system_intel.get_complete_system_info()
        sys_context = {
            "cpu": f"{sys_info.get('cpu', {}).get('percent_per_core', [0])[0]:.0f}",
            "ram": f"{sys_info.get('memory', {}).get('percent', 0):.0f}",
            "apps": ", ".join(list(sys_info.get('top_processes', {}).keys())[:3])
        }

        # Parse with context
        action = self.brain.parse(text, sys_context)
        
        # Execute action(s)
        actions = action if isinstance(action, list) else [action]
        final_result = None

        for idx, act in enumerate(actions, 1):
            if not isinstance(act, dict):
                continue

            act_name = act.get("action", "?")
            if len(actions) > 1:
                self.hud.show(f"⚡  [{idx}/{len(actions)}] {act_name}", "executing")
            else:
                self.hud.show(f"⚡  {act_name}", "executing")

            result = self.executor.execute(act, text)
            final_result = result

            if len(actions) > 1 and idx < len(actions):
                time.sleep(0.5)

        # Handle result
        if final_result:
            speech = final_result.speak or (final_result.message if not final_result.success else None)
            if speech:
                self.hud.show(f"🔊  Speaking…", "speaking")
                self.tts.speak(speech)

            icon = "✅" if final_result.success else "❌"
            state = "success" if final_result.success else "error"
            self.hud.show(f"{icon}  {final_result.message[:48]}", state)

        time.sleep(0.8)
        self._return_to_active("Command done")

    def _return_to_active(self, reason: str = ""):
        """Return to active state"""
        self._state = self.STATES["ACTIVE"]
        log.info(f"Returned to ACTIVE ({reason})")
        self.hud.show("👀  ONLINE (say 'Hello Lucifer')", "ready")
        time.sleep(0.5)

    def _shutdown(self):
        """Shutdown agent"""
        self._state = self.STATES["IDLE"]
        log.info("Agent SHUTDOWN")
        self.hud.show("🛑  OFFLINE - Say 'Start Lucifer'", "ready")
        self.tts.speak_sync("Lucifer going offline.")
        time.sleep(0.5)


# ══════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════
def listen_and_paste():
    agent = ProVoiceAgent()
    agent.start()