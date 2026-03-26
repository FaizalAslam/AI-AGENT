import threading
import subprocess
import platform
import tempfile
import time
import os
import re
import html
from pypdf import PdfReader
from modules.openai_client import clean_text_for_reading

system = platform.system().lower()
IS_WINDOWS = system == "windows"
IS_MACOS   = system == "darwin"

reader_state = {
    "is_reading":   False,
    "is_paused":    False,
    "current_page": 0,
    "total_pages":  0,
    "speed":        150,
    "pages_text":   [],
}

_lock                = threading.Lock()
_current_speech_proc = None


# ─────────────────────────────────────────
# SSML BUILDER
# ─────────────────────────────────────────

def _to_ssml(text, speed=150):
    """Convert plain text to SSML with prosody rate."""
    rate = max(0.5, min(3.0, speed / 150))
    escaped = html.escape(text)
    return (
        '<?xml version="1.0"?>'
        '<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="en-US">'
        f'<prosody rate="{rate}">{escaped}</prosody>'
        '</speak>'
    )


# ─────────────────────────────────────────
# NATIVE OS TTS — speaks ONE sentence
# ─────────────────────────────────────────

def _speak_one(text, speed=150):
    """Speak a single chunk using SSML on Windows, native on others."""
    global _current_speech_proc

    text = text.strip()
    if not text:
        return

    try:
        if IS_MACOS:
            rate  = max(80, min(300, speed))
            clean = text.replace("'", " ").replace('"', " ").replace("\n", " ")
            _current_speech_proc = subprocess.Popen(
                ['say', '-r', str(rate), clean]
            )
            _current_speech_proc.wait()

        elif IS_WINDOWS:
            ssml = _to_ssml(text, speed)
            ps_script = f"""Add-Type -AssemblyName System.Speech
$synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
$synth.SpeakSsml(@"
{ssml}
"@)
"""
            with tempfile.NamedTemporaryFile(
                mode='w', suffix='.ps1',
                delete=False, encoding='utf-8'
            ) as f:
                f.write(ps_script)
                script_path = f.name

            _current_speech_proc = subprocess.Popen(
                ['powershell', '-ExecutionPolicy', 'Bypass', '-File', script_path],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            _current_speech_proc.wait()

            if os.path.exists(script_path):
                os.remove(script_path)

        else:
            clean = text.replace("'", " ").replace('"', " ").replace("\n", " ")
            _current_speech_proc = subprocess.Popen(
                ['espeak', '-s', str(speed), clean],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            _current_speech_proc.wait()

    except Exception as e:
        print(f"TTS Error: {e}")
    finally:
        _current_speech_proc = None


def _kill_speech():
    """Instantly stop current speech."""
    global _current_speech_proc
    if _current_speech_proc:
        try:
            _current_speech_proc.terminate()
            _current_speech_proc = None
        except:
            pass


# ─────────────────────────────────────────
# SPLIT TEXT INTO SENTENCES
# ─────────────────────────────────────────

def _split_sentences(text):
    """Split into sentences so we can check speed/pause/stop between each."""
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences if s.strip()]


# ─────────────────────────────────────────
# PDF TEXT EXTRACTION
# ─────────────────────────────────────────

def extract_pdf_pages(pdf_path):
    print(f"📖 Loading PDF: {pdf_path}")
    reader     = PdfReader(pdf_path)
    pages_text = []

    for i, page in enumerate(reader.pages):
        text = page.extract_text()

        if not text or len(text.strip()) < 20:
            print(f"⚠️ Page {i+1} no text - running OCR...")
            try:
                from pdf2image import convert_from_path
                from modules.ocr_utils import image_to_text
                images = convert_from_path(pdf_path,
                                           first_page=i+1, last_page=i+1)
                if images:
                    temp_path = f"temp_page_{i}.png"
                    images[0].save(temp_path)
                    text = image_to_text(temp_path)
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
            except Exception as e:
                print(f"OCR fallback failed p{i+1}: {e}")
                text = f"Page {i+1} could not be read."

        pages_text.append(text or f"Page {i+1} is empty.")
        print(f"✅ Page {i+1}/{len(reader.pages)} loaded")

    return pages_text


# ─────────────────────────────────────────
# READING LOOP — sentence by sentence
# ─────────────────────────────────────────

def _reading_loop():
    while True:
        with _lock:
            if not reader_state["is_reading"]:
                print("⏹ Reading stopped")
                break

        if reader_state["is_paused"]:
            time.sleep(0.3)
            continue

        with _lock:
            page_index = reader_state["current_page"]
            total      = reader_state["total_pages"]

        if page_index >= total:
            print("📖 Finished reading all pages")
            with _lock:
                reader_state["is_reading"] = False
            break

        # ✅ Clean through OpenAI before speaking
        raw_text = reader_state["pages_text"][page_index]
        print(f"🤖 Sending page {page_index + 1} to OpenAI for cleanup...")
        page_text = clean_text_for_reading(raw_text, page_num=page_index + 1)
        sentences = _split_sentences(page_text)
        print(f"📖 Reading page {page_index + 1}/{total} "
              f"({len(sentences)} sentences)")

        page_changed = False

        for sentence in sentences:
            with _lock:
                if not reader_state["is_reading"]:
                    return

            while reader_state["is_paused"]:
                time.sleep(0.3)
                with _lock:
                    if not reader_state["is_reading"]:
                        return

            with _lock:
                if reader_state["current_page"] != page_index:
                    page_changed = True
                    break

            with _lock:
                current_speed = reader_state["speed"]

            _speak_one(sentence, current_speed)

        if page_changed:
            continue

        with _lock:
            if not reader_state["is_reading"]:
                break
            if reader_state["is_paused"]:
                continue

        with _lock:
            reader_state["current_page"] += 1
            print(f"➡️ Advanced to page {reader_state['current_page'] + 1}")


# ─────────────────────────────────────────
# PUBLIC CONTROLS
# ─────────────────────────────────────────

def start_reading(pdf_path, start_page=0):
    pages = extract_pdf_pages(pdf_path)
    with _lock:
        reader_state["pages_text"]   = pages
        reader_state["total_pages"]  = len(pages)
        reader_state["current_page"] = start_page
        reader_state["is_reading"]   = True
        reader_state["is_paused"]    = False

    thread = threading.Thread(target=_reading_loop, daemon=True)
    thread.start()
    print(f"▶️ Reading started - {len(pages)} pages")


def pause_reading():
    _kill_speech()
    with _lock:
        reader_state["is_paused"] = True
    print("⏸ Paused")


def resume_reading():
    with _lock:
        reader_state["is_paused"] = False
    print("▶️ Resumed")


def stop_reading():
    _kill_speech()
    with _lock:
        reader_state["is_reading"]   = False
        reader_state["is_paused"]    = False
        reader_state["current_page"] = 0
        reader_state["pages_text"]   = []
    print("⏹ Stopped")


def next_page():
    _kill_speech()
    with _lock:
        if reader_state["current_page"] < reader_state["total_pages"] - 1:
            reader_state["current_page"] += 1
            reader_state["is_paused"]    = False
    print(f"⏭ Skipped to page {reader_state['current_page'] + 1}")


def prev_page():
    _kill_speech()
    with _lock:
        if reader_state["current_page"] > 0:
            reader_state["current_page"] -= 1
            reader_state["is_paused"]    = False
    print(f"⏮ Back to page {reader_state['current_page'] + 1}")


def set_speed(wpm):
    with _lock:
        reader_state["speed"] = int(wpm)
    _kill_speech()
    print(f"🔊 Speed changed to {wpm} WPM — applying immediately")


def get_status():
    with _lock:
        return {
            "is_reading":   reader_state["is_reading"],
            "is_paused":    reader_state["is_paused"],
            "current_page": reader_state["current_page"] + 1,
            "total_pages":  reader_state["total_pages"],
            "speed":        reader_state["speed"],
        }
