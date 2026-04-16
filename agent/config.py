"""
ProVoiceAgent™ Configuration
Enterprise-grade desktop automation agent settings
"""

# ══════════════════════════════════════════════════════════
# CORE AGENT CONFIGURATION
# ══════════════════════════════════════════════════════════
AGENT_CONFIG = {
    # ────────────────────────────────────────────────────────
    # AI MODEL SETTINGS
    # ────────────────────────────────────────────────────────
    "model": "llama-3.3-70b-versatile",  # Best free LLM for automation
    "temperature": 0.05,                 # Very low for consistent behavior
    "max_tokens": 500,                   # Token limit per request
    
    
    # ────────────────────────────────────────────────────────
    # VOICE & SPEECH RECOGNITION
    # ────────────────────────────────────────────────────────
    "language": "en-US",                 # Speech recognition language
                                         # Options: "en-US", "ur-PK", "en-GB", etc.
    "listen_timeout": 15,                # Max seconds to wait for speech
    "phrase_limit": 20,                  # Max seconds per spoken phrase
    "pause_threshold": 0.8,              # Silence duration to end phrase (seconds)
    
    # Microphone calibration
    "calibration_duration": 1.5,         # Seconds for noise calibration
    "energy_threshold_min": 300,         # Minimum energy threshold
    
    
    # ────────────────────────────────────────────────────────
    # TEXT-TO-SPEECH (TTS) SETTINGS
    # ────────────────────────────────────────────────────────
    "tts_rate": 185,                     # Words per minute (150-220)
    "tts_volume": 0.92,                  # Volume level (0.0-1.0)
    "tts_voice_preference": "female",    # Preferred voice gender
    
    # Verbosity control
    "speak_on_success": False,           # Speak confirmation for every success
    "speak_on_error": True,              # Always speak errors
    "speak_on_startup": True,            # Speak startup messages
    "speak_status": True,                # Speak status updates
    
    
    # ────────────────────────────────────────────────────────
    # AGENT BEHAVIOR
    # ────────────────────────────────────────────────────────
    # Command activation
    "require_activation": True,          # Require "Hello Lucifer" to execute
    "multi_command_support": True,       # Support multiple actions per request
    "context_awareness": True,           # Use system context in decisions
    
    # Command execution
    "auto_retry_on_failure": True,      # Retry failed actions
    "max_retries": 2,                    # Maximum retry attempts
    "action_timeout": 30,                # Timeout per action (seconds)
    
    # Safety features
    "confirm_destructive": False,        # Confirm delete/shutdown operations
    "rate_limit_actions": False,         # Prevent action spam
    
    
    # ────────────────────────────────────────────────────────
    # SYSTEM MONITORING & INTELLIGENCE
    # ────────────────────────────────────────────────────────
    "system_monitoring": True,           # Monitor system resources
    "track_app_usage": True,             # Track which apps are used
    "log_file_access": False,            # Log file operations (privacy)
    "network_monitoring": True,          # Monitor network stats
    
    # Performance tuning
    "app_cache_ttl": 300,                # App cache refresh time (seconds)
    "system_scan_interval": 60,          # System scan interval (seconds)
    
    
    # ────────────────────────────────────────────────────────
    # FILE & STORAGE MANAGEMENT
    # ────────────────────────────────────────────────────────
    "default_file_location": None,       # Default save location (None = Desktop)
    "auto_organize_downloads": False,    # Auto-organize downloads
    "search_depth": 3,                   # File search depth levels
    "max_search_results": 20,            # Maximum file search results
    
    # File operation settings
    "safe_delete": False,                # Move to recycle bin instead of delete
    "backup_before_modify": False,       # Create backups before modifying files
    
    
    # ────────────────────────────────────────────────────────
    # WEB & BROWSER SETTINGS
    # ────────────────────────────────────────────────────────
    "default_browser": "chrome",         # Default browser to use
    "search_engine": "google",           # Default search engine
    "auto_focus_browser": True,          # Auto-focus browser window
    
    # Security
    "verify_urls": True,                 # Verify URLs before opening
    "block_malicious_sites": False,      # Basic malicious site blocking
    
    
    # ────────────────────────────────────────────────────────
    # UI AUTOMATION SETTINGS
    # ────────────────────────────────────────────────────────
    "automation_speed": "normal",        # Automation speed: "fast", "normal", "slow"
    "pyautogui_pause": 0.05,            # Pause between actions (seconds)
    "failsafe": True,                    # Enable fail-safe (move to corner to exit)
    
    # Screenshot settings
    "screenshot_quality": 85,            # JPEG quality (1-100)
    "screenshot_format": "png",          # Format: "png", "jpg"
    
    
    # ────────────────────────────────────────────────────────
    # ADVANCED FEATURES
    # ────────────────────────────────────────────────────────
    # Automation & scheduling
    "enable_scheduling": False,          # Enable scheduled tasks
    "enable_macros": False,              # Enable macro recording/playback
    "enable_workflow_templates": True,  # Enable pre-built workflow templates
    
    # Learning & personalization
    "learn_user_patterns": False,        # ML-based pattern learning
    "personalize_responses": True,       # Personalize agent responses
    "remember_preferences": True,        # Remember user preferences
    
    # Advanced AI
    "enable_intent_prediction": True,    # Predict user intent from context
    "enable_anomaly_detection": False,   # Detect unusual activity
    "enable_smart_scheduling": False,    # Intelligent task scheduling
    
    
    # ────────────────────────────────────────────────────────
    # LOGGING & DEBUGGING
    # ────────────────────────────────────────────────────────
    "log_level": "INFO",                 # Log level: DEBUG, INFO, WARNING, ERROR
    "log_commands": True,                # Log all voice commands
    "log_actions": True,                 # Log all executed actions
    "log_system_metrics": False,         # Log system metrics periodically
    
    # Performance logging
    "track_performance": True,           # Track action performance
    "report_timing": False,              # Report action timing
    
    # Privacy
    "encrypt_logs": False,               # Encrypt sensitive logs
    "anonymize_paths": False,            # Anonymize file paths in logs
    
    
    # ────────────────────────────────────────────────────────
    # INTEGRATION & PLUGINS
    # ────────────────────────────────────────────────────────
    "enable_plugins": False,             # Enable custom plugins
    "plugin_dir": None,                  # Plugin directory (None = default)
    "enable_webhooks": False,            # Enable webhook integrations
    
    
    # ────────────────────────────────────────────────────────
    # DEVELOPMENT & TESTING
    # ────────────────────────────────────────────────────────
    "debug_mode": False,                 # Enable debug mode
    "verbose_logging": False,            # Verbose logging output
    "test_mode": False,                  # Test mode (no actual execution)
    "mock_microphone": False,            # Mock microphone input (testing)
    "mock_tts": False,                   # Mock TTS output (testing)
}


# ══════════════════════════════════════════════════════════
# ACTIVATION PHRASES (Customizable)
# ══════════════════════════════════════════════════════════
ACTIVATION_PHRASES = {
    "hello lucifer",
    "hello agent",
    "activate agent",
    "agent activate",
    "hey lucifer",
    "lucifer listen",
    "lucifer activate",
    "wake up",
    "are you there",
    "ready",
}

# ══════════════════════════════════════════════════════════
# STARTUP PHRASES (Customizable)
# ══════════════════════════════════════════════════════════
STARTUP_PHRASES = {
    "start lucifer",
    "start agent",
    "activate lucifer",
    "lucifer start",
    "begin",
    "initialize",
    "go online",
    "boot up",
}

# ══════════════════════════════════════════════════════════
# SHUTDOWN PHRASES (Customizable)
# ══════════════════════════════════════════════════════════
SHUTDOWN_PHRASES = {
    "stop lucifer",
    "stop agent",
    "goodbye",
    "exit",
    "quit",
    "shutdown",
    "go offline",
    "goodbye lucifer",
    "see you later",
}


# ══════════════════════════════════════════════════════════
# AUTOMATION PROFILES (Pre-configured sets)
# ══════════════════════════════════════════════════════════
AUTOMATION_PROFILES = {
    "productivity": {
        "description": "Optimized for office work and document processing",
        "speak_on_success": True,
        "auto_retry_on_failure": True,
        "context_awareness": True,
    },
    
    "entertainment": {
        "description": "Optimized for media and entertainment control",
        "tts_rate": 200,
        "auto_focus_browser": True,
        "speak_on_error": True,
    },
    
    "development": {
        "description": "Optimized for coding and development tasks",
        "log_level": "DEBUG",
        "verbose_logging": True,
        "context_awareness": True,
        "enable_intent_prediction": True,
    },
    
    "stealth": {
        "description": "Minimal output and notifications",
        "speak_on_success": False,
        "speak_on_error": False,
        "log_level": "WARNING",
    },
    
    "learning": {
        "description": "Educational and tutorial mode with explanations",
        "speak_on_success": True,
        "learn_user_patterns": True,
        "personalize_responses": True,
        "remember_preferences": True,
    },
}


# ══════════════════════════════════════════════════════════
# ADVANCED ACTIONS DATABASE
# ══════════════════════════════════════════════════════════
ADVANCED_ACTIONS = {
    # System optimization
    "optimize_system": {
        "description": "Optimize system performance",
        "actions": ["clear_temp", "close_background_apps", "optimize_memory"]
    },
    
    # Backup operations
    "backup_desktop": {
        "description": "Backup all Desktop files",
        "source": "Desktop",
        "destination": "Desktop_Backup"
    },
    
    # Cleanup
    "cleanup_downloads": {
        "description": "Organize Downloads folder",
        "target": "Downloads",
        "organize_by": "type"
    },
}


# ══════════════════════════════════════════════════════════
# SEARCH ENGINE TEMPLATES
# ══════════════════════════════════════════════════════════
SEARCH_ENGINES_CUSTOM = {
    "google": "https://www.google.com/search?q={}",
    "youtube": "https://www.youtube.com/results?search_query={}",
    "github": "https://github.com/search?q={}",
    "stackoverflow": "https://stackoverflow.com/search?q={}",
    "reddit": "https://www.reddit.com/search/?q={}",
    "wikipedia": "https://en.wikipedia.org/wiki/Special:Search?search={}",
    "twitter": "https://twitter.com/search?q={}",
    "amazon": "https://www.amazon.com/s?k={}",
}


# ══════════════════════════════════════════════════════════
# SYSTEM ALERTS & THRESHOLDS
# ══════════════════════════════════════════════════════════
SYSTEM_THRESHOLDS = {
    "cpu_warning": 80,                   # CPU % to trigger warning
    "cpu_critical": 95,                  # CPU % to trigger critical alert
    
    "ram_warning": 85,                   # RAM % to trigger warning
    "ram_critical": 95,                  # RAM % to trigger critical alert
    
    "disk_warning": 80,                  # Disk % to trigger warning
    "disk_critical": 95,                 # Disk % to trigger critical alert
    
    "temp_warning": 80,                  # Temperature warning (Celsius)
    "temp_critical": 95,                 # Temperature critical (Celsius)
}


# ══════════════════════════════════════════════════════════
# FUNCTION TO GET ACTIVE CONFIG
# ══════════════════════════════════════════════════════════
def get_active_profile(profile_name: str = None) -> dict:
    """Get configuration for specified profile"""
    if profile_name and profile_name in AUTOMATION_PROFILES:
        # Merge profile with base config
        config = AGENT_CONFIG.copy()
        profile = AUTOMATION_PROFILES[profile_name]
        config.update(profile)
        return config
    return AGENT_CONFIG


def print_config() -> None:
    """Print current configuration"""
    print("\n" + "="*60)
    print("  AGENT CONFIGURATION")
    print("="*60)
    
    for key, value in AGENT_CONFIG.items():
        if not key.startswith("_"):
            print(f"  {key}: {value}")
    
    print("="*60 + "\n")


# ══════════════════════════════════════════════════════════
# CONFIGURATION VALIDATION
# ══════════════════════════════════════════════════════════
def validate_config() -> bool:
    """Validate configuration integrity"""
    errors = []
    
    # Check required keys
    required_keys = ["model", "language", "listen_timeout", "tts_rate"]
    for key in required_keys:
        if key not in AGENT_CONFIG:
            errors.append(f"Missing required key: {key}")
    
    # Check value ranges
    if not (150 <= AGENT_CONFIG["tts_rate"] <= 250):
        errors.append("TTS rate must be between 150 and 250 wpm")
    
    if not (0 <= AGENT_CONFIG["tts_volume"] <= 1):
        errors.append("TTS volume must be between 0 and 1")
    
    if AGENT_CONFIG["listen_timeout"] < 5:
        errors.append("Listen timeout should be at least 5 seconds")
    
    if errors:
        print("❌ Configuration errors:")
        for error in errors:
            print(f"   • {error}")
        return False
    
    return True


# Run validation on import
if __name__ == "__main__":
    validate_config()
    print_config()
else:
    # Silently validate
    validate_config()