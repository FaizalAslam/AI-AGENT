# modules/system_core.py
import glob
import json
import os
import platform
import re
import subprocess
from pathlib import Path

from . import config, openai_client, store_apps

system     = platform.system().lower()
IS_WINDOWS = system == "windows"
IS_MACOS   = system == "darwin"

# ── Load app path data once at startup ───────────────────────────
_DATA_FILES = [
    Path(__file__).parent.parent / "apps_data.json",
    Path(__file__).parent.parent / "app_paths.json",
]


def _load_apps_data():
    last_error = None
    for data_file in _DATA_FILES:
        try:
            if not data_file.exists():
                continue
            with open(data_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    print(f"[system_core] Loaded app config: {data_file.name}")
                    return data
        except Exception as e:
            last_error = e
    if last_error:
        print(f"[system_core] Failed to load app config: {last_error}")
    else:
        print("[system_core] No app config file found (apps_data.json/app_paths.json)")
    return {}


_D = _load_apps_data()


# ── Shorthand accessors ───────────────────────────────────────────
def _uri_apps():
    return _D.get("uri_apps", {})

def _app_aliases():
    return _D.get("app_aliases", {})

def _direct_execs():
    return _D.get("direct_executables", {})

def _store_aliases():
    return _D.get("store_aliases", {})

def _basic_path(app_name):
    entry = _D.get("basic_apps", {}).get(app_name)
    if not entry:
        return None
    key = "macos" if IS_MACOS else "windows"
    return entry.get(key)

def _candidate_paths(app_name):
    raw_paths = _D.get("candidate_paths", {}).get(app_name, [])
    resolved  = []
    for p in raw_paths:
        expanded = os.path.expandvars(p)
        if "*" in expanded:
            matches = glob.glob(expanded)
            resolved.extend(matches)
        elif os.path.exists(expanded):
            resolved.append(expanded)
    return resolved

def _close_procs(app_name):
    entry = _D.get("close_processes", {}).get(app_name)
    if not entry:
        return None
    key = "macos" if IS_MACOS else "windows"
    return entry.get(key)


# ════════════════════════════════════════════════════════════════
#  LAUNCH HELPERS
# ════════════════════════════════════════════════════════════════

def _launch_uri(uri):
    """Launch a URI scheme app — shell=True required for URI schemes."""
    print(f"Launching URI: {uri}")
    try:
        if IS_WINDOWS:
            subprocess.run(
                f'start "" "{uri}"',
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            return True
        if IS_MACOS:
            subprocess.run(["open", uri], check=False)
            return True
        subprocess.run(["xdg-open", uri], check=False)
        return True
    except Exception as e:
        print(f"URI launch failed: {e}")
        return False


def open_path(path, is_store_app=False):
    print(f"Launching: {path}")
    try:
        if is_store_app and IS_WINDOWS:
            # Primary — explorer.exe (most reliable for Store apps)
            success = store_apps.launch_store_app(path)
            if success:
                return True
            # Fallback — shell start
            subprocess.run(
                f'start "" "shell:AppsFolder\\{path}"',
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            return True

        if IS_WINDOWS:
            if path.startswith("start "):
                os.system(path)
                return True
            if "\\" not in path and "/" not in path and not path.lower().endswith(".exe"):
                result = subprocess.run(
                    ["where", path],
                    capture_output=True, text=True, timeout=3
                )
                if result.returncode != 0:
                    return False
            try:
                os.startfile(path)
            except OSError:
                proc = subprocess.run(path, shell=True)
                if proc.returncode != 0:
                    return False
            return True

        if IS_MACOS:
            if path.startswith("/Applications"):
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["open", "-a", path], check=False)
            return True

        subprocess.run(["xdg-open", path], check=False)
        return True

    except Exception as e:
        print(f"Launch failed: {e}")
        return False


def _launch_whatsapp():
    """WhatsApp-specific launcher — 4 methods in order."""
    print("Trying WhatsApp-specific launcher...")

    # Method 1 — candidate install paths
    for p in _candidate_paths("whatsapp"):
        if open_path(p):
            config.save_memory("whatsapp", p, is_store_app=False)
            return True, "Opened WhatsApp from install path"

    # Method 2 — Get-StartApps live scan (most reliable for Store version)
    app_id = store_apps.find_app_id(
        "whatsapp",
        aliases=_D.get("store_aliases", {}).get("whatsapp", [])
    )
    if app_id:
        if store_apps.launch_store_app(app_id):
            config.save_memory("whatsapp", app_id, is_store_app=True)
            return True, "Opened WhatsApp via Store AppID"

    # Method 3 — hardcoded Store IDs from apps_data.json
    for wid in _D.get("whatsapp_store_ids", []):
        if store_apps.launch_store_app(wid):
            return True, "Opened WhatsApp via hardcoded Store ID"

    # Method 4 — URI scheme
    if _launch_uri("whatsapp:"):
        return True, "Opened WhatsApp via URI"

    return False, "WhatsApp not found — check if it is installed from Store"


# ════════════════════════════════════════════════════════════════
#  NORMALIZATION
# ════════════════════════════════════════════════════════════════

def _normalize_app_name(raw_command):
    cleaned = re.sub(r"\s+", " ", (raw_command or "").lower()).strip()
    cleaned = cleaned.replace(".", "")
    cleaned = re.sub(
        r"^(open|launch|start|run|boot|close|shut|exit)\s+", "", cleaned
    ).strip()
    cleaned = re.sub(r"\s+(app|application)$", "", cleaned).strip()
    return _app_aliases().get(cleaned, cleaned)


def normalize_app_name(raw_command):
    return _normalize_app_name(raw_command)


# ════════════════════════════════════════════════════════════════
#  LOOKUP HELPERS
# ════════════════════════════════════════════════════════════════

def _try_where_lookup(app_name):
    if not IS_WINDOWS:
        return None
    base       = _direct_execs().get(app_name, app_name)
    candidates = [base]
    if not base.endswith(".exe"):
        candidates.append(base + ".exe")
    for candidate in candidates:
        try:
            result = subprocess.run(
                ["where", candidate],
                capture_output=True, text=True, timeout=3
            )
            if result.returncode == 0:
                first = (result.stdout.splitlines() or [""])[0].strip()
                if first:
                    return first
        except Exception:
            pass
    return None


def _find_store_app_id(app_name):
    if not IS_WINDOWS:
        return None
    aliases = _store_aliases().get(app_name, [])
    return store_apps.find_app_id(app_name, aliases)


def _is_valid_store_id(app_name, app_id):
    if not app_id:
        return False
    current = _find_store_app_id(app_name)
    if not current:
        return False
    return str(app_id).strip().lower() == str(current).strip().lower()


# ════════════════════════════════════════════════════════════════
#  CLOSE HELPERS
# ════════════════════════════════════════════════════════════════

def close_process_by_name(proc_name):
    print(f"Closing process: {proc_name}")
    try:
        if IS_WINDOWS:
            subprocess.run(
                ["taskkill", "/F", "/IM", proc_name],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        else:
            subprocess.run(
                ["pkill", "-f", proc_name],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        return True
    except Exception as e:
        print(f"Close failed: {e}")
        return False


def close_processes(proc_names):
    if isinstance(proc_names, str):
        return close_process_by_name(proc_names)
    if not isinstance(proc_names, (list, tuple)):
        return False
    ok = False
    for name in proc_names:
        if close_process_by_name(name):
            ok = True
    return ok


def close_store_app_windows(app_name):
    print(f"Fuzzy searching process for: {app_name}")
    try:
        cmd = (
            'powershell -Command "Get-Process | '
            f"Where-Object {{$_.Name -like '*{app_name}*'}} | "
            'Stop-Process -Force"'
        )
        subprocess.run(
            cmd, shell=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        return True
    except Exception as e:
        print(f"Fuzzy close failed: {e}")
        return False


# ════════════════════════════════════════════════════════════════
#  MAIN: FIND AND LAUNCH
# ════════════════════════════════════════════════════════════════

def find_and_launch(raw_command):
    app_name = _normalize_app_name(raw_command)
    print(f"Searching for: '{app_name}'")

    # ── STEP 1: URI apps (Camera, Settings, Alarms, etc.) ────────
    uri = _uri_apps().get(app_name)
    if uri:
        success = _launch_uri(uri)
        return success, "Opened via URI" if success else "URI launch failed"

    # ── STEP 2: WhatsApp special handler ─────────────────────────
    if app_name == "whatsapp":
        return _launch_whatsapp()

    # ── STEP 3: Basic apps (known exe paths) ─────────────────────
    basic = _basic_path(app_name)
    if basic is not None:
        if basic:
            if open_path(basic):
                return True, "Opened Basic App"
        else:
            return False, f"'{app_name}' not available on this OS"

    # ── STEP 4: Memory (previously saved paths) ──────────────────
    memory = config.load_memory()
    if app_name in memory:
        saved    = memory[app_name]
        path     = saved.get("path", "")
        is_store = saved.get("type") == "store"
        if is_store and not _is_valid_store_id(app_name, path):
            config.delete_memory(app_name)
        elif is_store or os.path.exists(path):
            return open_path(path, is_store), "Opened from Memory"
        else:
            config.delete_memory(app_name)

    # ── STEP 5: WHERE lookup (apps in system PATH) ───────────────
    where_path = _try_where_lookup(app_name)
    if where_path:
        config.save_memory(app_name, where_path, is_store_app=False)
        return open_path(where_path), "Opened from PATH"

    # ── STEP 6: Direct exe name attempt ──────────────────────────
    direct = _direct_execs().get(app_name)
    if direct and open_path(direct):
        return True, "Opened by command name"

    # ── STEP 7: Candidate install paths ──────────────────────────
    for candidate in _candidate_paths(app_name):
        if open_path(candidate):
            config.save_memory(app_name, candidate, is_store_app=False)
            return True, "Opened from known install path"

    # ── STEP 8: Windows Store scan (Get-StartApps) ───────────────
    if IS_WINDOWS:
        aliases = _store_aliases().get(app_name, [])
        app_id  = store_apps.find_app_id(app_name, aliases)
        if app_id:
            if store_apps.launch_store_app(app_id):
                config.save_memory(app_name, app_id, is_store_app=True)
                return True, "Opened Store App"

    # ── STEP 9: OpenAI fallback ───────────────────────────────────
    guessed = openai_client.guess_path_with_ai(app_name)
    if guessed == "STORE_APP":
        return False, "OpenAI says Store App — try exact name"
    if guessed and os.path.exists(guessed):
        config.save_memory(app_name, guessed, is_store_app=False)
        return open_path(guessed), "Opened via AI"

    return False, "Not Found"


# ════════════════════════════════════════════════════════════════
#  MAIN: CLOSE APP
# ════════════════════════════════════════════════════════════════

def close_app(raw_command):
    app_name = _normalize_app_name(raw_command)
    print(f"Closing: '{app_name}'")

    # ── STEP 1: Known process list ───────────────────────────────
    procs = _close_procs(app_name)
    if procs is not None:
        if procs:
            if close_processes(procs):
                return True, "Closed from Process List"
        else:
            return False, f"'{app_name}' not available on this OS"

    # ── STEP 2: Memory ───────────────────────────────────────────
    memory = config.load_memory()
    if app_name in memory:
        saved = memory[app_name]
        if saved.get("type") == "store":
            if IS_WINDOWS and close_store_app_windows(app_name):
                return True, "Closed Store App"
        else:
            proc_name = os.path.basename(saved.get("path", ""))
            if proc_name and close_processes([proc_name]):
                return True, "Closed from Memory"

    # ── STEP 3: Fuzzy process search ─────────────────────────────
    if IS_WINDOWS:
        if close_store_app_windows(app_name):
            return True, "Closed via Fuzzy Search"
    elif IS_MACOS:
        if close_processes([app_name]):
            return True, "Closed via Name Search"

    return False, f"Could not close '{app_name}'"
