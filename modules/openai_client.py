# modules/openai_client.py
import re
import platform
from openai import OpenAI
from . import config


def guess_path_with_ai(app_name):
    if not config.OPENAI_API_KEY:
        print("⚠️ OpenAI Key missing.")
        return None

    system = platform.system().lower()
    if system == "windows":
        prompt = (
            f'Find the absolute Windows file path for "{app_name}".\n'
            f'Rules:\n1. Return ONLY the path. No text. No markdown.\n'
            f'2. Use Program Files path if standard.\n'
            f'3. If Windows Store app, return exactly: STORE_APP'
        )
    elif system == "darwin":
        prompt = (
            f'Find the macOS application path for "{app_name}".\n'
            f'Rules:\n1. Return ONLY the path. No text. No markdown.\n'
            f'2. Use format: /Applications/NAME.app/Contents/MacOS/NAME\n'
            f'3. If App Store app, return exactly: APPS_STORE'
        )
    else:
        prompt = f'Return only the absolute path for {app_name} on Linux.'

    try:
        client   = OpenAI(api_key=config.OPENAI_API_KEY)
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        text = response.choices[0].message.content.strip()

        if "STORE_APP"  in text: return "STORE_APP"
        if "APPS_STORE" in text: return "APPS_STORE"

        match = re.search(r'[a-zA-Z]:\\[^<>:"/|?*\n]+|/[^\s<>:"|?*\n]+', text)
        if match:
            clean = match.group(0).strip()
            print(f"🤖 OpenAI Found: {clean}")
            return clean
        return None

    except Exception as e:
        print(f"OpenAI Error: {e}")
        return None


def clean_text_for_reading(raw_text, page_num=None, is_ocr=False):
    if not config.OPENAI_API_KEY:
        return raw_text
    if not raw_text or len(raw_text.strip()) < 10:
        return raw_text

    page_info = f"This is page {page_num} of a PDF document." if page_num else ""

    if is_ocr:
        source_note = (
            "This text was extracted by OCR so ALL formatting is lost.\n"
            "You must INFER structure from context clues:\n"
            "- Short lines in ALL CAPS or Title Case = likely a heading\n"
            "- Lines starting with a dash, asterisk, or number+dot = likely a bullet\n"
            "- Very short standalone lines = likely labels or headings\n"
            "- Long continuous text = body paragraph"
        )
    else:
        source_note = "This text was extracted directly from a PDF and may have some structure preserved."

    prompt = f"""You are a text-to-speech pre-processor. {page_info}

{source_note}

STRICT RULES:
1. REMOVE: page numbers, headers, footers, URLs, emails, copyright lines, watermarks, document IDs.
2. REMOVE: lines that are only symbols, dashes, underscores, or formatting noise.
3. REMOVE: table of contents entries, bibliography/reference entries.
4. FIX broken words from line wrapping — e.g. "impor-\\ntant" → "important".
5. JOIN sentences split across lines naturally.
6. If you detect a TITLE or MAIN HEADING → output as: "Heading: [text]."
7. If you detect a SUB-HEADING → output as: "Section: [text]."
8. If you detect a BULLET POINT → output as: "Point: [text]."
9. If you detect a NUMBERED LIST ITEM → output as: "Point [number]: [text]."
10. If you detect a figure caption → output as: "Figure: [text]."
11. Keep ALL actual content.
12. Output ONLY the final spoken script. No markdown. No extra labels.

RAW TEXT:
---
{raw_text[:4000]}
---
Output the clean spoken script now:"""

    try:
        client   = OpenAI(api_key=config.OPENAI_API_KEY)
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        cleaned = response.choices[0].message.content.strip()
        if not cleaned or len(cleaned) < 20:
            print("⚠️ OpenAI returned empty — using raw text")
            return raw_text
        print(f"🤖 OpenAI cleaned: {len(raw_text)} → {len(cleaned)} chars")
        return cleaned

    except Exception as e:
        print(f"⚠️ OpenAI clean error: {e} — using raw text")
        return raw_text
