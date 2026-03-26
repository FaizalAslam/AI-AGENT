import json
import threading
from difflib import SequenceMatcher
from pathlib import Path


_LOCK = threading.Lock()
_MAP_FILE = Path("command_map.json")
FUZZY_THRESHOLD = 85


def _read():
    if not _MAP_FILE.exists():
        return {}
    try:
        return json.loads(_MAP_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _write(data):
    _MAP_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def _score(a, b):
    a = (a or "").strip().lower()
    b = (b or "").strip().lower()
    if not a or not b:
        return 0
    seq = SequenceMatcher(None, a, b).ratio() * 100
    sa = set(a.split())
    sb = set(b.split())
    overlap = (len(sa & sb) / max(1, len(sa | sb))) * 100
    return int((seq * 0.7) + (overlap * 0.3))


def get_cached_actions(app_name, command_text):
    app = (app_name or "").strip().lower()
    query = (command_text or "").strip()
    if not app or not query:
        return None, None, 0

    with _LOCK:
        data = _read()
        app_map = data.get(app, {})
        if not isinstance(app_map, dict):
            return None, None, 0

        # Exact (case-insensitive) match.
        for key, actions in app_map.items():
            if key.strip().lower() == query.lower():
                if isinstance(actions, list):
                    return key, actions, 100
                return None, None, 0

        # Fuzzy match fallback.
        best_key = None
        best_actions = None
        best_score = 0
        for key, actions in app_map.items():
            if not isinstance(actions, list):
                continue
            s = _score(query, key)
            if s > best_score:
                best_score = s
                best_key = key
                best_actions = actions

        if best_score > FUZZY_THRESHOLD and best_actions:
            return best_key, best_actions, best_score
        return None, None, best_score


def save_actions(app_name, command_text, actions):
    app = (app_name or "").strip().lower()
    key = (command_text or "").strip()
    if not app or not key or not isinstance(actions, list):
        return False
    with _LOCK:
        data = _read()
        data.setdefault(app, {})
        data[app][key] = actions
        _write(data)
    return True


def remove_action(app_name, command_text):
    app = (app_name or "").strip().lower()
    key = (command_text or "").strip()
    if not app or not key:
        return False
    with _LOCK:
        data = _read()
        app_map = data.get(app, {})
        if not isinstance(app_map, dict):
            return False
        # Case-insensitive key removal.
        hit = None
        for existing in app_map.keys():
            if existing.strip().lower() == key.lower():
                hit = existing
                break
        if not hit:
            return False
        del app_map[hit]
        data[app] = app_map
        _write(data)
        return True
