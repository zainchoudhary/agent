
import subprocess
import threading
from fastapi import FastAPI
import uvicorn

app = FastAPI()

@app.on_event("startup")
def start_agent():
    # Start the hotkey listener in a background thread
    threading.Thread(target=lambda: subprocess.Popen(["python", "agent/voice_agent.py"]), daemon=True).start()

@app.get("/")
def root():
    return {"status": "Agent running"}

if __name__ == "__main__":
    uvicorn.run("agent.main:app", host="127.0.0.1", port=8000, reload=True)
