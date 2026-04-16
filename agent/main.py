#!/usr/bin/env python3
"""
ProVoiceAgent™ - Enterprise Desktop Automation
Advanced voice-controlled desktop agent with full system access
"""

import os
import sys
import logging
import psutil
import platform
from pathlib import Path
from datetime import datetime

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent))

from voice_agent import ProVoiceAgent
from config import AGENT_CONFIG

# ══════════════════════════════════════════════════════════
# CONFIGURATION & LOGGING
# ══════════════════════════════════════════════════════════
LOG_DIR = Path.home() / ".provoiceagent" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / f"agent_{datetime.now().strftime('%Y%m%d')}.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
)

log = logging.getLogger("ProVoiceAgent-Main")


# ══════════════════════════════════════════════════════════
# SYSTEM HEALTH CHECK
# ══════════════════════════════════════════════════════════
class SystemHealthCheck:
    """Verify system requirements and health before startup"""
    
    @staticmethod
    def check_python_version() -> bool:
        """Check Python version"""
        if sys.version_info < (3, 8):
            log.error("[ERROR] Python 3.8+ required")
            return False
        log.info(f"[OK] Python {sys.version.split()[0]} OK")
        return True
    
    @staticmethod
    def check_os() -> bool:
        """Check if running on Windows"""
        if platform.system() != "Windows":
            log.error("[ERROR] Windows required")
            return False
        log.info(f"[OK] OS: Windows {platform.release()} OK")
        return True
    
    @staticmethod
    def check_system_resources() -> bool:
        """Check minimum system resources"""
        ram = psutil.virtual_memory()
        cpu = psutil.cpu_count()
        
        if ram.total < 2 * 1024**3:  # 2GB
            log.warning(f"[WARN] Low RAM: {ram.total / 1024**3:.1f}GB")
            return False
        
        if cpu < 2:
            log.warning(f"[WARN] Low CPU cores: {cpu}")
            return False
        
        log.info(f"[OK] Resources OK - CPU: {cpu} cores, RAM: {ram.total / 1024**3:.1f}GB")
        return True
    
    @staticmethod
    def check_microphone() -> bool:
        """Check microphone availability"""
        try:
            import speech_recognition as sr
            mic = sr.Microphone()
            # Try to use microphone
            with mic as source:
                pass
            log.info("[OK] Microphone OK")
            return True
        except Exception as e:
            log.error(f"[ERROR] Microphone error: {e}")
            return False
    
    @staticmethod
    def check_audio_output() -> bool:
        """Check audio output"""
        try:
            import pyttsx3
            engine = pyttsx3.init()
            log.info("[OK] Audio output OK")
            return True
        except Exception as e:
            log.error(f"[ERROR] Audio error: {e}")
            return False
    
    @staticmethod
    def check_api_key() -> bool:
        """Check if API key is configured"""
        from dotenv import load_dotenv
        load_dotenv()
        
        api_key = os.getenv('GROQ_API_KEY')
        if not api_key:
            log.error("[ERROR] GROQ_API_KEY not found in .env")
            return False
        
        log.info("[OK] API Key configured")
        return True
    
    @staticmethod
    def check_internet() -> bool:
        """Check internet connection"""
        try:
            import socket
            socket.create_connection(("8.8.8.8", 53), timeout=3)
            log.info("[OK] Internet connection OK")
            return True
        except (socket.timeout, socket.error):
            log.warning("[WARN] Internet connection issue")
            return False
    
    @classmethod
    def run_all_checks(cls) -> bool:
        """Run all health checks"""
        log.info("\n" + "="*50)
        log.info("  SYSTEM HEALTH CHECK")
        log.info("="*50)
        
        checks = [
            ("Python Version", cls.check_python_version),
            ("Operating System", cls.check_os),
            ("System Resources", cls.check_system_resources),
            ("Microphone", cls.check_microphone),
            ("Audio Output", cls.check_audio_output),
            ("API Key", cls.check_api_key),
            ("Internet Connection", cls.check_internet),
        ]
        
        results = []
        for name, check_fn in checks:
            try:
                result = check_fn()
                results.append((name, result))
            except Exception as e:
                log.error(f"[ERROR] {name} check error: {e}")
                results.append((name, False))
        
        passed = sum(1 for _, r in results if r)
        total = len(results)
        
        log.info(f"\nResults: {passed}/{total} checks passed")
        log.info("="*50 + "\n")
        
        return passed >= 5  # At least 5 critical checks


# ══════════════════════════════════════════════════════════
# STARTUP & CONFIGURATION
# ══════════════════════════════════════════════════════════
class AgentStartup:
    """Handle agent initialization and startup"""
    
    @staticmethod
    def print_banner():
        """Print startup banner"""
        banner = """
╔════════════════════════════════════════════════════════════╗
║                                                            ║
║              ProVoiceAgent™ ENTERPRISE EDITION             ║
║                                                            ║
║         Advanced Desktop Automation with AI                ║
║                                                            ║
║  • Full system access and control                          ║
║  • Deep file & app management                              ║
║  • Intelligent voice command processing                    ║
║  • Web automation and search                               ║
║  • Multi-step workflow execution                           ║
║  • Real-time system monitoring                             ║
║                                                            ║
╚════════════════════════════════════════════════════════════╝
        """
        log.info(banner)
        # Use a simple print for the banner to avoid encoding issues in some terminals
        try:
            print(banner)
        except UnicodeEncodeError:
            print("ProVoiceAgent - Enterprise Edition")
    
    @staticmethod
    def print_config():
        """Print active configuration"""
        log.info("\nActive Configuration:")
        log.info(f"  Model: {AGENT_CONFIG.get('model')}")
        log.info(f"  Language: {AGENT_CONFIG.get('language')}")
        log.info(f"  TTS Rate: {AGENT_CONFIG.get('tts_rate')} wpm")
        log.info(f"  Microphone Sensitivity: {AGENT_CONFIG.get('pause_threshold')}")
        log.info(f"  Log Directory: {LOG_DIR}")
        log.info("")
    
    @staticmethod
    def print_usage():
        """Print usage instructions"""
        instructions = """
USAGE INSTRUCTIONS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. STARTUP
   • Say: "Start Lucifer" to activate the agent
   • Listen for confirmation beep

2. COMMAND MODE
   • Say: "Hello Lucifer" to enter command mode
   • Agent will say "Ready. What do you need?"
   • Give your voice command

3. EXAMPLES
   • "Open Chrome"
   • "Search for Python tutorials"
   • "Create a file named todo.txt"
   • "Find files matching report"
   • "Close all applications except Chrome"
   • "Take a screenshot"
   • "Show system information"
   • "Restart the computer in 60 seconds"
   • "Set a reminder for 5 minutes"
   • "What's the current time?"

4. SHUTDOWN
   • Say: "Stop Lucifer" or "Goodbye"
   • Agent will shut down gracefully

ADVANCED FEATURES:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Deep file searching and analysis
• System process monitoring
• Disk usage analysis
• Multi-application automation
• Form filling and web automation
• Screenshot capture
• System command execution
• Reminder management

For logs: {0}
For issues: Check agent.log in current directory
        """.format(LOG_DIR)
        
        log.info(instructions)
        # Use a simple print for instructions to avoid encoding issues
        try:
            print(instructions)
        except UnicodeEncodeError:
            print("See logs for usage instructions.")


# ══════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ══════════════════════════════════════════════════════════
def main():
    """Main entry point with error handling"""
    
    try:
        # Print startup information
        AgentStartup.print_banner()
        AgentStartup.print_config()
        
        # Run health checks
        if not SystemHealthCheck.run_all_checks():
            log.warning("[WARN] Some system checks failed, but continuing...")
        
        # Print usage
        AgentStartup.print_usage()
        
        # Initialize and start agent
        log.info("Initializing ProVoiceAgent...")
        agent = ProVoiceAgent()
        
        log.info("Starting agent main loop...")
        agent.start()
        
    except KeyboardInterrupt:
        log.info("\n[INFO] Agent interrupted by user")
        print("\n\nShutting down gracefully...")
    
    except ImportError as e:
        log.error(f"[ERROR] Missing dependency: {e}")
        print(f"\n[ERROR] Missing Python package: {e}")
        print("\nInstall required packages:")
        print("  pip install pyperclip pyautogui speech-recognition pyttsx3 psutil requests python-dotenv")
        sys.exit(1)
    
    except PermissionError as e:
        log.error(f"[ERROR] Permission denied: {e}")
        print(f"\n[ERROR] Permission Error: {e}")
        print("Try running as Administrator")
        sys.exit(1)
    
    except Exception as e:
        log.error(f"[ERROR] Fatal error: {e}", exc_info=True)
        print(f"\n[ERROR] Fatal Error: {e}")
        print(f"Check logs at: {LOG_DIR}")
        sys.exit(1)
    
    finally:
        log.info("ProVoiceAgent shut down")
        print("\n[OK] ProVoiceAgent shut down successfully")


if __name__ == "__main__":
    main()