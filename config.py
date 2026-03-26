# config.py
import os
from dotenv import load_dotenv

load_dotenv()

# ── OpenAI (replaces Gemini in both projects) ──────────────────────
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL   = "gpt-4o-mini"

# ── Project 1 settings ─────────────────────────────────────────────
SETTINGS_FILE = "agent_memory.json"

# ── Project 2 settings ─────────────────────────────────────────────
TRIGGER_WORD  = "agent:"
LOG_FILE      = "office_agent.log"
TOAST_ENABLED = True

OFFICE_PATHS = {
    "excel":      r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    "word":       r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "powerpoint": r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
}
