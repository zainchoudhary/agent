
import logging
from .voice_agent import ProVoiceAgent

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[
        logging.FileHandler("agent.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

if __name__ == "__main__":
    try:
        agent = ProVoiceAgent()
        agent.start()
    except KeyboardInterrupt:
        logging.info("Agent shutting down...")
    except Exception as e:
        logging.error(f"Agent error: {e}", exc_info=True)
