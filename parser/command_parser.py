# modules/command_parser.py
import json
import re
import logging
from pathlib import Path

logger = logging.getLogger("OfficeAgent")

BASE = Path(__file__).parent.parent

_COMMAND_FILES = {
    "excel":      BASE / "excel_commands.json",
    "word":       BASE / "word_commands.json",
    "powerpoint": BASE / "powerpoint_commands.json",
    "ppt":        BASE / "powerpoint_commands.json",
}

_COMMAND_CACHE = {}
_STOP_WORDS = {"a", "an", "the", "please", "kindly"}


def _load_commands(app):
    if app in _COMMAND_CACHE:
        return _COMMAND_CACHE[app]
    path = _COMMAND_FILES.get(app)
    if not path or not path.exists():
        logger.warning(f"No command file for app: {app}")
        return {}
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    _COMMAND_CACHE[app] = data
    return data


# ════════════════════════════════════════════════════════════════
#  PARAMETER EXTRACTORS
# ════════════════════════════════════════════════════════════════

def _extract_cell(text):
    match = re.search(r'\b([A-Z]{1,3}[0-9]{1,7})\b', text.upper())
    return match.group(1) if match else None


def _extract_range(text):
    match = re.search(r'\b([A-Z]{1,3}[0-9]{1,7})\s*[:\-to]+\s*([A-Z]{1,3}[0-9]{1,7})\b', text.upper())
    if match:
        return f"{match.group(1)}:{match.group(2)}"
    row_match = re.search(r'\brow\s*(\d+)\b', text.lower())
    if row_match:
        return f"{row_match.group(1)}:{row_match.group(1)}"
    col_match = re.search(r'\bcolumn\s*([A-Za-z]+)\b', text.lower())
    if col_match:
        c = col_match.group(1).upper()
        return f"{c}:{c}"
    return _extract_cell(text)


def _extract_number(text):
    match = re.search(r'\b(\d+(?:\.\d+)?)\b', text)
    return match.group(1) if match else None


def _extract_font_size(text):
    match = re.search(r'\b(\d{1,3})\s*(?:pt|px|point|size)?\b', text.lower())
    if match:
        val = int(match.group(1))
        if 6 <= val <= 200:
            return val
    return None


def _extract_color(text):
    color_map = {
        "red":        "FF0000",
        "green":      "00B050",
        "blue":       "0070C0",
        "yellow":     "FFFF00",
        "orange":     "FFA500",
        "purple":     "7030A0",
        "pink":       "FF69B4",
        "black":      "000000",
        "white":      "FFFFFF",
        "gray":       "808080",
        "grey":       "808080",
        "dark red":   "C00000",
        "dark blue":  "00008B",
        "dark green": "006400",
        "light blue": "ADD8E6",
        "light gray": "D3D3D3",
        "teal":       "008080",
        "cyan":       "00FFFF",
        "magenta":    "FF00FF",
        "gold":       "FFD700",
        "brown":      "A52A2A",
        "navy":       "000080",
    }
    text_lower = text.lower()
    for name, hex_val in color_map.items():
        if name in text_lower:
            return hex_val
    hex_match = re.search(r'#?([0-9A-Fa-f]{6})', text)
    if hex_match:
        return hex_match.group(1).upper()
    return None


def _extract_text_value(text):
    quoted = re.search(r'["\u201c\u201d\u2018\u2019](.*?)["\u201c\u201d\u2018\u2019]', text)
    if quoted:
        return quoted.group(1)
    for kw in ["write", "type", "set", "put", "enter", "add", "with text", "text", "saying", "called", "named", "title", "body", "subtitle", "content"]:
        pattern = rf'{kw}\s+(.+?)(?:\s+in\s+|\s+at\s+|\s+to\s+|\s+on\s+|$)'
        match = re.search(pattern, text.lower())
        if match:
            candidate = match.group(1).strip()
            if len(candidate) > 1:
                return candidate
    return None


def _extract_filename(text):
    match = re.search(r'(?:as|named?|called?|to)\s+["\']?([a-zA-Z0-9_\-\s\.]+)["\']?', text.lower())
    if match:
        return match.group(1).strip()
    return None


def _extract_rows_cols(text):
    rows = cols = None
    row_match = re.search(r'(\d+)\s*(?:row|rows)', text.lower())
    col_match = re.search(r'(\d+)\s*(?:col|cols|column|columns)', text.lower())
    if row_match:
        rows = int(row_match.group(1))
    if col_match:
        cols = int(col_match.group(1))
    by_match = re.search(r'(\d+)\s*[xX×]\s*(\d+)', text)
    if by_match:
        rows = int(by_match.group(1))
        cols = int(by_match.group(2))
    return rows, cols


def _normalize_for_match(text):
    text = re.sub(r"[^a-z0-9\s]", " ", (text or "").lower())
    tokens = [t for t in text.split() if t and t not in _STOP_WORDS]
    return " ".join(tokens), set(tokens)


def _split_sub_commands(raw_command):
    text = (raw_command or "").lower().strip()

    # Keep "X columns and Y rows" as one command (same for row/column inverse).
    protected = re.sub(
        r"(\d+\s*(?:col|cols|column|columns)\s+)and(\s*\d+\s*(?:row|rows))",
        r"\1__AND__\2",
        text
    )
    protected = re.sub(
        r"(\d+\s*(?:row|rows)\s+)and(\s*\d+\s*(?:col|cols|column|columns))",
        r"\1__AND__\2",
        protected
    )

    parts = re.split(r"\s+(?:and|then|also|after that|next|,)\s+", protected)
    return [p.replace("__AND__", " and ").strip() for p in parts if p.strip()]


def _heuristic_action(app, command_text):
    text = (command_text or "").lower()

    # Example: "create a table with 4 columns and 5 rows"
    if app == "excel" and "table" in text and any(w in text for w in ("create", "make", "insert", "add", "build")):
        rows, cols = _extract_rows_cols(text)
        start_cell = _extract_cell(text) or "A1"
        return {
            "action": "create_table",
            "rows": rows or 5,
            "cols": cols or 3,
            "start_cell": start_cell,
        }

    if app == "word":
        if "table" in text and any(w in text for w in ("create", "make", "insert", "add", "build")):
            rows, cols = _extract_rows_cols(text)
            return {"action": "add_table", "rows": rows or 5, "cols": cols or 3}
        if any(w in text for w in ("add heading", "create heading", "heading")):
            return {"action": "add_heading", "text": _extract_text_value(command_text) or "Heading", "level": 1}
        if any(w in text for w in ("add paragraph", "write paragraph", "paragraph", "add text", "write text")):
            return {"action": "add_paragraph", "text": _extract_text_value(command_text) or "New paragraph"}

    if app in ("powerpoint", "ppt"):
        if "slide" in text and any(w in text for w in ("add", "create", "insert", "new")):
            return {"action": "add_slide", "layout": "title_content"}
        if "title" in text and any(w in text for w in ("set", "add", "write")):
            return {
                "action": "set_slide_text",
                "target": "title",
                "text": _extract_text_value(command_text) or "Title",
                "slide_index": _extract_slide_index(command_text) or 0,
            }
        if "table" in text and any(w in text for w in ("add", "create", "insert", "make")):
            rows, cols = _extract_rows_cols(text)
            return {
                "action": "insert_table",
                "rows": rows or 3,
                "cols": cols or 3,
                "slide_index": _extract_slide_index(command_text) or 0,
            }

    return None


def _extract_slide_index(text):
    patterns = [
        r'slide\s*(\d+)',
        r'(\d+)(?:st|nd|rd|th)?\s*slide',
        r'page\s*(\d+)',
    ]
    for p in patterns:
        m = re.search(p, text.lower())
        if m:
            return int(m.group(1)) - 1  # zero-indexed
    return None


def _extract_alignment(text):
    text_lower = text.lower()
    if any(w in text_lower for w in ["center", "centre", "middle", "centered"]):
        return "center"
    if any(w in text_lower for w in ["right", "flush right"]):
        return "right"
    if any(w in text_lower for w in ["justify", "justified", "both sides"]):
        return "justify"
    return "left"


def _extract_sort_order(text):
    text_lower = text.lower()
    if any(w in text_lower for w in ["descend", "z to a", "high to low", "largest first"]):
        return "descending"
    return "ascending"


def _extract_orientation(text):
    text_lower = text.lower()
    if "landscape" in text_lower:
        return "landscape"
    return "portrait"


def _extract_spacing(text):
    if "double" in text.lower():
        return 2.0
    if "1.5" in text:
        return 1.5
    if "single" in text.lower():
        return 1.0
    match = re.search(r'(\d+(?:\.\d+)?)\s*(?:line|spacing)', text.lower())
    if match:
        return float(match.group(1))
    return 1.15


def _extract_formula(text):
    match = re.search(r'=\s*([A-Z]+\s*\(.*?\))', text.upper())
    if match:
        return match.group(0).replace(" ", "")
    return None


def _extract_shape_type(text):
    text_lower = text.lower()
    shape_map = {
        "rectangle": "RECTANGLE",
        "rect":      "RECTANGLE",
        "square":    "RECTANGLE",
        "circle":    "OVAL",
        "oval":      "OVAL",
        "ellipse":   "OVAL",
        "arrow":     "RIGHT_ARROW",
        "triangle":  "TRIANGLE",
        "star":      "STAR_5_POINT",
        "pentagon":  "PENTAGON",
        "diamond":   "DIAMOND",
        "line":      "LINE",
    }
    for keyword, shape in shape_map.items():
        if keyword in text_lower:
            return shape
    return "RECTANGLE"


def _extract_chart_type(text):
    text_lower = text.lower()
    if any(w in text_lower for w in ["bar", "column"]):
        return "bar"
    if any(w in text_lower for w in ["line", "trend"]):
        return "line"
    if any(w in text_lower for w in ["pie", "donut", "doughnut"]):
        return "pie"
    if any(w in text_lower for w in ["area"]):
        return "area"
    if any(w in text_lower for w in ["scatter", "plot"]):
        return "scatter"
    return "bar"


def _extract_transition(text):
    text_lower = text.lower()
    transitions = ["fade", "wipe", "dissolve", "push", "cover", "uncover", "zoom", "split", "reveal", "flash"]
    for t in transitions:
        if t in text_lower:
            return t
    return "fade"


def _extract_target(text):
    text_lower = text.lower()
    if "title" in text_lower:
        return "title"
    if "subtitle" in text_lower:
        return "subtitle"
    if any(w in text_lower for w in ["body", "content", "text"]):
        return "body"
    return "body"


def _extract_level(text):
    for i in range(1, 10):
        if str(i) in text:
            return i
    return 1


def _extract_heading_level(text):
    text_lower = text.lower()
    if any(w in text_lower for w in ["heading 1", "h1", "primary", "main heading", "chapter"]):
        return 1
    if any(w in text_lower for w in ["heading 2", "h2", "subheading", "section"]):
        return 2
    if any(w in text_lower for w in ["heading 3", "h3"]):
        return 3
    if any(w in text_lower for w in ["heading 4", "h4"]):
        return 4
    return None


def _extract_delimiter(text):
    text_lower = text.lower()
    if "comma" in text_lower:
        return ","
    if "tab" in text_lower:
        return "\t"
    if "semicolon" in text_lower:
        return ";"
    if "pipe" in text_lower:
        return "|"
    if "space" in text_lower:
        return " "
    return ","


def _extract_url(text):
    match = re.search(r'https?://[^\s]+', text)
    if match:
        return match.group(0)
    return None


def _extract_find_replace(text):
    pattern = r'(?:replace|change)\s+["\']?(.+?)["\']?\s+(?:with|to)\s+["\']?(.+?)["\']?$'
    match = re.search(pattern, text.lower())
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return None, None


def _extract_number_format(text):
    text_lower = text.lower()
    if any(w in text_lower for w in ["currency", "dollar", "money", "$"]):
        return "$#,##0.00"
    if any(w in text_lower for w in ["percent", "%", "percentage"]):
        return "0.00%"
    if any(w in text_lower for w in ["date"]):
        if "mm/dd" in text_lower:
            return "MM/DD/YYYY"
        return "DD/MM/YYYY"
    if any(w in text_lower for w in ["time"]):
        return "hh:mm:ss"
    if any(w in text_lower for w in ["comma", "thousands"]):
        return "#,##0"
    return "General"


# ════════════════════════════════════════════════════════════════
#  KEYWORD MATCHING
# ════════════════════════════════════════════════════════════════

def _score_match(command_text, keywords):
    text_norm, text_tokens = _normalize_for_match(command_text)
    score = 0
    matched = []
    for kw in keywords:
        kw_norm, kw_tokens = _normalize_for_match(kw)
        if not kw_norm:
            continue

        if kw_norm in text_norm:
            score += max(2, len(kw_tokens) * 2)
            matched.append(kw)
            continue

        if kw_tokens and kw_tokens.issubset(text_tokens):
            score += max(1, len(kw_tokens))
            matched.append(kw)
            continue

        overlap = len(kw_tokens & text_tokens)
        if len(kw_tokens) >= 2 and overlap >= 2:
            score += overlap
            matched.append(kw)
    return score, matched


def _find_matching_commands(app, command_text):
    commands = _load_commands(app)
    scored = []
    for cmd_name, cmd_data in commands.items():
        keywords = cmd_data.get("keywords", [])
        score, matched = _score_match(command_text, keywords)
        if score > 0:
            scored.append((score, cmd_name, cmd_data, matched))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored


# ════════════════════════════════════════════════════════════════
#  PARAMETER RESOLVER
# ════════════════════════════════════════════════════════════════

def _resolve_params(params, command_text, app):
    resolved = {}
    for key, val in params.items():
        if not isinstance(val, str) or not val.startswith("{"):
            resolved[key] = val
            continue

        placeholder = val.strip("{}")

        if placeholder in ("range",):
            resolved[key] = _extract_range(command_text) or "A1"

        elif placeholder in ("cell", "result_cell", "target_cell", "dest_cell"):
            cells = re.findall(r'[A-Z]{1,3}[0-9]{1,7}', command_text.upper())
            if placeholder == "result_cell" and len(cells) >= 1:
                resolved[key] = cells[-1]
            elif cells:
                resolved[key] = cells[0]
            else:
                resolved[key] = "A1"

        elif placeholder == "start_cell":
            cells = re.findall(r'[A-Z]{1,3}[0-9]{1,7}', command_text.upper())
            resolved[key] = cells[0] if cells else "A1"

        elif placeholder == "value":
            resolved[key] = _extract_text_value(command_text) or _extract_number(command_text) or ""

        elif placeholder == "text":
            resolved[key] = _extract_text_value(command_text) or ""

        elif placeholder == "size":
            resolved[key] = _extract_font_size(command_text) or 12

        elif placeholder == "color":
            resolved[key] = _extract_color(command_text) or "000000"

        elif placeholder == "font_name":
            font_match = re.search(
                r'(?:font|typeface|use|change to|apply)\s+([A-Za-z\s]+?)(?:\s+font|\s+size|\s*$)',
                command_text.lower()
            )
            resolved[key] = font_match.group(1).strip().title() if font_match else "Calibri"

        elif placeholder == "alignment":
            resolved[key] = _extract_alignment(command_text)

        elif placeholder == "rows":
            rows, _ = _extract_rows_cols(command_text)
            resolved[key] = rows or 5

        elif placeholder == "cols":
            _, cols = _extract_rows_cols(command_text)
            resolved[key] = cols or 3

        elif placeholder == "row_number":
            m = re.search(r'row\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 1

        elif placeholder == "column":
            m = re.search(r'column\s*([A-Za-z]+)', command_text.lower())
            resolved[key] = m.group(1).upper() if m else "A"

        elif placeholder == "slide_index":
            idx = _extract_slide_index(command_text)
            resolved[key] = idx if idx is not None else 0

        elif placeholder == "target":
            if app == "powerpoint":
                resolved[key] = _extract_target(command_text)
            else:
                resolved[key] = _extract_range(command_text) or "selection"

        elif placeholder == "filename":
            resolved[key] = _extract_filename(command_text) or "output"

        elif placeholder == "file_path":
            path_match = re.search(r'["\']([^"\']+\.[a-z]{3,5})["\']', command_text)
            resolved[key] = path_match.group(1) if path_match else ""

        elif placeholder == "formula":
            resolved[key] = _extract_formula(command_text) or ""

        elif placeholder == "order":
            resolved[key] = _extract_sort_order(command_text)

        elif placeholder == "orientation":
            resolved[key] = _extract_orientation(command_text)

        elif placeholder == "spacing":
            resolved[key] = _extract_spacing(command_text)

        elif placeholder == "shape_type":
            resolved[key] = _extract_shape_type(command_text)

        elif placeholder == "chart_type":
            resolved[key] = _extract_chart_type(command_text)

        elif placeholder == "transition_type":
            resolved[key] = _extract_transition(command_text)

        elif placeholder == "animation_type":
            anim_match = re.search(
                r'(?:animation|animate|effect)\s+([a-z\s]+?)(?:\s+on|\s+to|$)',
                command_text.lower()
            )
            resolved[key] = anim_match.group(1).strip() if anim_match else "appear"

        elif placeholder == "format":
            resolved[key] = _extract_number_format(command_text)

        elif placeholder == "date_format":
            resolved[key] = "DD/MM/YYYY"

        elif placeholder == "delimiter":
            resolved[key] = _extract_delimiter(command_text)

        elif placeholder == "url":
            resolved[key] = _extract_url(command_text) or ""

        elif placeholder == "find_text":
            find, _ = _extract_find_replace(command_text)
            resolved[key] = find or ""

        elif placeholder == "replace_text":
            _, replace = _extract_find_replace(command_text)
            resolved[key] = replace or ""

        elif placeholder in ("level",):
            resolved[key] = _extract_heading_level(command_text) or _extract_level(command_text)

        elif placeholder == "zoom_level":
            m = re.search(r'(\d+)\s*%', command_text)
            resolved[key] = int(m.group(1)) if m else 100

        elif placeholder == "password":
            m = re.search(r'password\s+["\']?([^\s"\']+)["\']?', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder == "sheet_name":
            m = re.search(r'(?:sheet|tab)\s+["\']?([^\s"\']+)["\']?', command_text.lower())
            resolved[key] = m.group(1) if m else "Sheet1"

        elif placeholder == "old_name":
            m = re.search(r'rename\s+["\']?([^\s"\']+)["\']?\s+to', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder == "new_name":
            m = re.search(r'to\s+["\']?([^\s"\']+)["\']?$', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder in ("from_index",):
            m = re.search(r'from\s+(?:slide\s*)?(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) - 1 if m else 0

        elif placeholder == "to_index":
            m = re.search(r'to\s+(?:slide\s*)?(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) - 1 if m else 0

        elif placeholder == "count":
            m = re.search(r'(\d+)\s*(?:column|col)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 2

        elif placeholder == "image_path":
            m = re.search(r'["\']([^"\']+\.(?:png|jpg|jpeg|gif|bmp|svg))["\']', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder == "video_path":
            m = re.search(r'["\']([^"\']+\.(?:mp4|avi|mov|mkv|wmv))["\']', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder == "audio_path":
            m = re.search(r'["\']([^"\']+\.(?:mp3|wav|ogg|aac|flac))["\']', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder in ("criteria", "condition"):
            m = re.search(r'(?:where|if|when|equals?|greater|less)\s+(.+?)(?:\s+in\s+|\s*$)', command_text.lower())
            resolved[key] = m.group(1).strip() if m else ""

        elif placeholder == "items":
            items_match = re.findall(r'["\u201c\u201d](.*?)["\u201c\u201d]', command_text)
            if items_match:
                resolved[key] = items_match
            else:
                resolved[key] = [s.strip() for s in re.split(r'[,;]', command_text)
                                 if len(s.strip()) > 1][-5:] or ["Item 1", "Item 2"]

        elif placeholder == "output_path":
            m = re.search(r'["\']([^"\']+\.pdf)["\']', command_text.lower())
            resolved[key] = m.group(1) if m else "output.pdf"

        elif placeholder == "width":
            m = re.search(r'width\s*[=:]\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 10

        elif placeholder == "height":
            m = re.search(r'height\s*[=:]\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 5

        elif placeholder == "angle":
            m = re.search(r'(\d+)\s*(?:degree|deg|°)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 90

        elif placeholder == "direction":
            if "vertical" in command_text.lower():
                resolved[key] = "vertical"
            else:
                resolved[key] = "horizontal"

        elif placeholder == "speed":
            if "slow" in command_text.lower():
                resolved[key] = "slow"
            elif "fast" in command_text.lower():
                resolved[key] = "fast"
            else:
                resolved[key] = "medium"

        elif placeholder == "seconds":
            m = re.search(r'(\d+(?:\.\d+)?)\s*(?:second|sec|s\b)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 3.0

        elif placeholder == "delay":
            m = re.search(r'(\d+(?:\.\d+)?)\s*(?:second|sec|s\b)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 0.0

        elif placeholder in ("before", "after"):
            m = re.search(rf'{placeholder}\s*[=:]\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 6

        elif placeholder == "top":
            m = re.search(r'top\s*[=:]\s*(\d+(?:\.\d+)?)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 1.0

        elif placeholder == "bottom":
            m = re.search(r'bottom\s*[=:]\s*(\d+(?:\.\d+)?)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 1.0

        elif placeholder == "left":
            m = re.search(r'left\s*[=:]\s*(\d+(?:\.\d+)?)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 1.0

        elif placeholder == "right":
            m = re.search(r'right\s*[=:]\s*(\d+(?:\.\d+)?)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 1.0

        elif placeholder == "position":
            m = re.search(r'position\s*[=:]\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 0

        elif placeholder in ("decimals", "num_chars"):
            m = re.search(r'(\d+)\s*(?:decimal|digit|char)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 2

        elif placeholder == "base":
            cells = re.findall(r'[A-Z]{1,3}[0-9]{1,7}', command_text.upper())
            resolved[key] = cells[0] if cells else "A1"

        elif placeholder == "exponent":
            m = re.search(r'(?:power|exponent)\s+(?:of\s+)?(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 2

        elif placeholder in ("start_date", "end_date"):
            m = re.search(r'[A-Z]{1,3}[0-9]{1,7}', command_text.upper())
            resolved[key] = m.group(0) if m else "A1"

        elif placeholder == "unit":
            if "month" in command_text.lower():
                resolved[key] = "M"
            elif "year" in command_text.lower():
                resolved[key] = "Y"
            else:
                resolved[key] = "D"

        elif placeholder in ("year", "month", "day"):
            m = re.search(r'\b(19|20)\d{2}\b', command_text)
            resolved[key] = m.group(0) if m else placeholder.upper()

        elif placeholder == "error_value":
            resolved[key] = '""'

        elif placeholder in ("true_val", "false_val"):
            resolved[key] = '"Yes"' if "true" in placeholder else '"No"'

        elif placeholder in ("cond1", "cond2", "condition1", "condition2"):
            resolved[key] = "TRUE"

        elif placeholder in ("val1", "val2"):
            resolved[key] = '""'

        elif placeholder == "lookup_value":
            cells = re.findall(r'[A-Z]{1,3}[0-9]{1,7}', command_text.upper())
            resolved[key] = cells[0] if cells else "A1"

        elif placeholder == "table_range":
            resolved[key] = _extract_range(command_text) or "A1:Z100"

        elif placeholder == "col_index":
            m = re.search(r'column\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 2

        elif placeholder == "row_index":
            m = re.search(r'row\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 2

        elif placeholder in ("lookup_array", "return_array"):
            resolved[key] = _extract_range(command_text) or "A1:A10"

        elif placeholder in ("lookup_range", "return_range"):
            resolved[key] = _extract_range(command_text) or "A1:A10"

        elif placeholder in ("sum_range", "range1", "range2"):
            resolved[key] = _extract_range(command_text) or "A1:A10"

        elif placeholder in ("criteria1", "criteria2"):
            resolved[key] = '">0"'

        elif placeholder == "data_source":
            resolved[key] = ""

        elif placeholder == "compare_path":
            resolved[key] = ""

        elif placeholder == "bookmark_name":
            m = re.search(r'(?:bookmark|mark)\s+["\']?([^\s"\']+)["\']?', command_text.lower())
            resolved[key] = m.group(1) if m else "bookmark1"

        elif placeholder == "layout":
            text_lower = command_text.lower()
            if "blank" in text_lower:
                return {**resolved, key: "blank"}
            if "title only" in text_lower:
                return {**resolved, key: "title_only"}
            if "two content" in text_lower or "two column" in text_lower:
                return {**resolved, key: "two_content"}
            resolved[key] = "title_content"

        elif placeholder == "style":
            style_match = re.search(
                r'(?:style|design|format)\s+["\']?([A-Za-z\s]+?)["\']?(?:\s|$)',
                command_text.lower()
            )
            resolved[key] = style_match.group(1).strip().title() if style_match else "Normal"

        elif placeholder == "theme":
            m = re.search(r'theme\s+["\']?([A-Za-z\s]+?)["\']?(?:\s|$)', command_text.lower())
            resolved[key] = m.group(1).strip().title() if m else "Office Theme"

        elif placeholder == "case":
            text_lower = command_text.lower()
            if "upper" in text_lower or "caps" in text_lower:
                return {**resolved, key: "upper"}
            if "lower" in text_lower:
                return {**resolved, key: "lower"}
            resolved[key] = "title"

        elif placeholder == "character":
            m = re.search(r'(?:symbol|character)\s+["\']?(.)["\']?', command_text)
            resolved[key] = m.group(1) if m else "©"

        elif placeholder == "type":
            resolved[key] = "list"

        elif placeholder == "values":
            resolved[key] = ""

        elif placeholder == "icon_name":
            m = re.search(r'icon\s+["\']?([A-Za-z\s]+?)["\']?(?:\s|$)', command_text.lower())
            resolved[key] = m.group(1).strip() if m else "star"

        elif placeholder == "nper":
            m = re.search(r'(\d+)\s*(?:year|month|period)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 12

        elif placeholder == "pv":
            m = re.search(r'(\d+(?:,\d{3})*(?:\.\d+)?)', command_text)
            resolved[key] = m.group(1).replace(",", "") if m else "10000"

        elif placeholder == "rate":
            m = re.search(r'(\d+(?:\.\d+)?)\s*%', command_text)
            resolved[key] = f"{float(m.group(1)) / 100}" if m else "0.05"

        elif placeholder == "start":
            m = re.search(r'start\s*(?:at\s*)?(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 1

        elif placeholder in ("start_row", "end_row"):
            rows = re.findall(r'row\s*(\d+)', command_text.lower())
            if placeholder == "start_row":
                resolved[key] = int(rows[0]) if rows else 1
            else:
                resolved[key] = int(rows[1]) if len(rows) > 1 else 5

        elif placeholder == "source_range":
            resolved[key] = _extract_range(command_text) or "A1:Z100"

        elif placeholder in ("color1", "color2"):
            colors = re.findall(
                r'\b(red|blue|green|yellow|orange|purple|white|black|gray|teal)\b',
                command_text.lower()
            )
            color_map_simple = {
                "red": "FF0000", "blue": "0070C0", "green": "00B050",
                "yellow": "FFFF00", "orange": "FFA500", "purple": "7030A0",
                "white": "FFFFFF", "black": "000000", "gray": "808080", "teal": "008080"
            }
            if placeholder == "color1":
                resolved[key] = color_map_simple.get(colors[0], "FFFFFF") if colors else "FFFFFF"
            else:
                resolved[key] = color_map_simple.get(colors[1], "000000") if len(colors) > 1 else "000000"

        elif placeholder == "scheme":
            m = re.search(r'scheme\s+["\']?([A-Za-z\s]+?)["\']?(?:\s|$)', command_text.lower())
            resolved[key] = m.group(1).strip().title() if m else "Default"

        else:
            resolved[key] = val

    return resolved


# ════════════════════════════════════════════════════════════════
#  MAIN PARSE FUNCTION
# ════════════════════════════════════════════════════════════════

def parse_command(app, raw_command):
    """
    Parse a raw office command string for the given app.
    Returns a list of action dicts ready for the executor.

    Example:
        parse_command("excel", "bold A1:E1 and sum B2 to B10 in C10")
        → [
            {"action": "set_bold",      "range": "A1:E1", "bold": True},
            {"action": "write_formula", "cell": "C10", "formula": "=SUM(B2:B10)"}
          ]
    """
    app = app.lower().strip()
    if app not in _COMMAND_FILES:
        logger.warning(f"Unknown app: {app}")
        return []

    # Split compound commands by conjunctions
    sub_commands = _split_sub_commands(raw_command)

    all_actions = []
    seen_actions = set()

    for sub in sub_commands:
        sub = sub.strip()
        if not sub:
            continue

        matches = _find_matching_commands(app, sub)
        if not matches:
            heuristic = _heuristic_action(app, sub)
            if heuristic:
                all_actions.append(heuristic)
                logger.info(f"Heuristic match for '{sub}' → {heuristic}")
                continue
            logger.info(f"No match found for: '{sub}'")
            continue

        # Take top match (highest score), deduplicate by action+key params
        top_score, cmd_name, cmd_data, matched_kws = matches[0]
        action_name = cmd_data.get("action")
        raw_params  = cmd_data.get("params", {})
        resolved    = _resolve_params(raw_params, sub, app)

        # Dedup key to avoid running same action on same target twice
        dedup_key = f"{action_name}_{resolved.get('range') or resolved.get('cell') or resolved.get('slide_index', '')}"
        if dedup_key in seen_actions:
            continue
        seen_actions.add(dedup_key)

        action_dict = {"action": action_name, **resolved}
        all_actions.append(action_dict)
        logger.info(f"Matched '{cmd_name}' (score={top_score}, kws={matched_kws}) → {action_dict}")

    return all_actions
