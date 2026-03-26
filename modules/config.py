import json
import os
from dotenv import load_dotenv

load_dotenv()

TRIGGER_WORD = "agent:"
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
LOG_FILE = "office_agent.log"
TOAST_ENABLED = True

PATHS_FILE = "app_paths.json"
MEMORY_FILE = "known_apps.json"


def load_path_settings():
    if os.path.exists(PATHS_FILE):
        try:
            with open(PATHS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {}


PATH_SETTINGS = load_path_settings()
OFFICE_PATHS = PATH_SETTINGS.get("office_paths", {})


def get_basic_apps_for_os(os_key):
    basic = PATH_SETTINGS.get("basic_apps", {})
    if not isinstance(basic, dict):
        return {}
    # Format A: {"windows": {...}, "macos": {...}}
    apps = basic.get(os_key, {})
    if isinstance(apps, dict):
        return apps
    # Format B: {"chrome": {"windows": "...", "macos": "..."}, ...}
    converted = {}
    for app, cfg in basic.items():
        if isinstance(cfg, dict):
            path = cfg.get(os_key)
            if path:
                converted[app] = path
    return converted


def get_open_candidate_paths():
    data = PATH_SETTINGS.get("open_candidate_paths")
    if not isinstance(data, dict):
        data = PATH_SETTINGS.get("path_candidates")
    if not isinstance(data, dict):
        data = PATH_SETTINGS.get("candidate_paths")
    return data if isinstance(data, dict) else {}


def get_uri_apps():
    data = PATH_SETTINGS.get("uri_apps", {})
    return data if isinstance(data, dict) else {}


def _derive_office_paths_if_missing():
    global OFFICE_PATHS
    if isinstance(OFFICE_PATHS, dict) and OFFICE_PATHS.get("excel") and OFFICE_PATHS.get("word") and OFFICE_PATHS.get("powerpoint"):
        return
    basic = PATH_SETTINGS.get("basic_apps", {})
    if not isinstance(basic, dict):
        return
    derived = {}
    for app in ("excel", "word", "powerpoint"):
        entry = basic.get(app, {})
        if isinstance(entry, dict):
            win = entry.get("windows")
            if win:
                derived[app] = win
    if derived:
        OFFICE_PATHS = derived


_derive_office_paths_if_missing()


def load_memory():
    if os.path.exists(MEMORY_FILE):
        try:
            with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_memory(app_name, path, is_store_app=False):
    memory = load_memory()
    memory[app_name] = {"path": path, "type": "store" if is_store_app else "exe"}
    with open(MEMORY_FILE, "w", encoding="utf-8") as f:
        json.dump(memory, f, indent=2)


def delete_memory(app_name):
    memory = load_memory()
    if app_name in memory:
        del memory[app_name]
        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump(memory, f, indent=2)
        return True
    return False
