# Add your Groq API key here
GROQ_API_KEY = "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"
AGENT_CONFIG = {
    # Groq model — llama-3.3-70b-versatile is best free option
    "model":           "llama-3.3-70b-versatile",

    # Voice recognition
    "language":        "en-US",        # change to "ur-PK" for Urdu, etc.
    "listen_timeout":  15,             # seconds to wait for speech
    "phrase_limit":    20,             # max seconds per phrase
    "pause_threshold": 0.8,            # silence that ends a phrase

    # TTS
    "tts_rate":        185,            # words per minute (150–220)

    # Behaviour
    "speak_on_success": False,         # speak confirmation for every action
    "speak_on_error":   True,          # speak on failures
}