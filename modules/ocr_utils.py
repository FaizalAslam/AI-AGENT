import os
import threading
import queue as queue_module
import pyautogui
from PIL import ImageGrab

_reader       = None
_snip_queue   = queue_module.Queue()
_result_queue = queue_module.Queue()


def get_reader():
    global _reader
    if _reader is None:
        print("🔍 Loading OCR Engine (first time only)...")
        import easyocr
        _reader = easyocr.Reader(['en'])
    return _reader


def image_to_text(image_path):
    if not os.path.exists(image_path):
        return f"Error: File not found - {image_path}"
    results = get_reader().readtext(image_path, detail=0)
    text    = "\n".join(results)
    print(f"📄 OCR Complete: {len(text)} chars extracted")
    return text


def capture_fullscreen():
    """Take full screenshot using PIL — no pyautogui conflicts."""
    import time
    time.sleep(0.8)
    path = "ocr_screenshot.png"
    img  = ImageGrab.grab()
    img.save(path)
    print(f"📸 Screenshot saved: {path}")
    return path


def run_snip_overlay_main_thread():
    """
    MUST run on the main thread (tkinter requirement on Windows).
    """
    import tkinter as tk
    from tkinter import Canvas

    print("✂️ Snip overlay listener ready on main thread")

    while True:
        try:
            _snip_queue.get(timeout=1)
        except:
            continue

        coords = {"x1": 0, "y1": 0, "x2": 0, "y2": 0, "done": False}

        root = tk.Tk()
        root.attributes("-fullscreen", True)
        root.attributes("-alpha", 0.35)
        root.attributes("-topmost", True)
        root.configure(bg="black")
        root.title("Click and Drag | ESC to cancel")
        root.focus_force()

        canvas = Canvas(root, cursor="cross", bg="black")
        canvas.pack(fill=tk.BOTH, expand=True)
        rect = [None]

        def on_press(e):
            coords["x1"] = e.x_root
            coords["y1"] = e.y_root

        def on_drag(e):
            if rect[0]:
                canvas.delete(rect[0])
            rect[0] = canvas.create_rectangle(
                coords["x1"] - root.winfo_rootx(),
                coords["y1"] - root.winfo_rooty(),
                canvas.canvasx(e.x),
                canvas.canvasy(e.y),
                outline="#00aaff", width=2, dash=(4, 2)
            )

        def on_release(e):
            coords["x2"]   = e.x_root
            coords["y2"]   = e.y_root
            coords["done"] = True
            root.quit()
            root.destroy()

        def on_escape(e):
            root.quit()
            root.destroy()

        canvas.bind("<ButtonPress-1>",   on_press)
        canvas.bind("<B1-Motion>",       on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        root.bind("<Escape>",            on_escape)
        root.mainloop()

        if not coords["done"]:
            _result_queue.put(None)
            print("✂️ Snip cancelled via ESC")
            continue

        x1 = min(coords["x1"], coords["x2"])
        y1 = min(coords["y1"], coords["y2"])
        x2 = max(coords["x1"], coords["x2"])
        y2 = max(coords["y1"], coords["y2"])

        w = x2 - x1
        h = y2 - y1

        if w < 2 or h < 2:
            print(f"✂️ Selection too small ({w}x{h}px) — retrying")
            _result_queue.put(None)
            continue

        if w < 20:
            x1 = max(0, x1 - 10)
            x2 = x2 + 10
        if h < 20:
            y1 = max(0, y1 - 10)
            y2 = y2 + 10

        import time
        time.sleep(0.15)
        img  = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        path = "ocr_snip.png"
        img.save(path)
        print(f"✂️ Snip saved: {path} ({x2-x1}x{y2-y1}px)")
        _result_queue.put(path)


def _minimize_browser():
    """
    Minimize browser by sending Alt+Space then N (minimize).
    Does NOT use Win+D which opens Task Manager as a side effect.
    """
    import platform, subprocess, time
    sys = platform.system().lower()
    try:
        if sys == "windows":
            subprocess.run(
                ['powershell', '-Command',
                 'Add-Type -AssemblyName System.Windows.Forms;'
                 '[System.Windows.Forms.SendKeys]::SendWait("%{SPACE}n")'],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
            time.sleep(0.4)
        elif sys == "darwin":
            subprocess.run(
                ['osascript', '-e',
                 'tell application "Google Chrome" to set miniaturized of window 1 to true'],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
            time.sleep(0.3)
    except Exception as e:
        print(f"Minimize error: {e}")


def trigger_snip_and_ocr(last_ocr_ref):
    """Called by Ctrl+Shift+S hotkey — background thread."""
    import time
    print("✂️ Snip hotkey triggered...")
    _minimize_browser()
    time.sleep(0.8)

    _snip_queue.put("snip")

    try:
        path = _result_queue.get(timeout=60)
    except:
        print("✂️ Snip timed out")
        return

    if path:
        text = image_to_text(path)
        last_ocr_ref["text"]    = text
        last_ocr_ref["pending"] = True
        print(f"✅ Snip OCR done: {len(text)} chars")
    else:
        print("✂️ Snip cancelled")


def trigger_screenshot_and_ocr(last_ocr_ref):
    """Called by Ctrl+Shift+F hotkey — background thread."""
    import time
    print("📸 Screenshot hotkey triggered...")
    _minimize_browser()
    time.sleep(0.8)
    path = capture_fullscreen()
    text = image_to_text(path)
    last_ocr_ref["text"]    = text
    last_ocr_ref["pending"] = True
    print(f"✅ Screenshot OCR done: {len(text)} chars")


# Global to track the running TTS process
_speak_process = None


def speak_text(text):
    global _speak_process

    stop_speaking()

    # ✅ Clean through OpenAI
    try:
        from modules.openai_client import clean_text_for_reading
        print("🤖 Sending OCR text to OpenAI for cleanup...")
        text = clean_text_for_reading(text, is_ocr=True)
    except Exception as e:
        print(f"⚠️ OpenAI cleanup skipped: {e} — using raw text")

    import platform, subprocess, tempfile, os
    sys = platform.system().lower()

    # Sanitize
    text = text[:3000].replace('"', " ").replace("'", " ").replace("\n", " ").strip()

    try:
        if sys == "darwin":
            _speak_process = subprocess.Popen(['say', text])

        elif sys == "windows":
            from modules.pdf_reader import _to_ssml
            ssml = _to_ssml(text, speed=150)

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

            _speak_process = subprocess.Popen(
                ['powershell', '-ExecutionPolicy', 'Bypass', '-File', script_path],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

        else:
            _speak_process = subprocess.Popen(
                ['espeak', text],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )

        print(f"🔊 Speaking started (PID {_speak_process.pid})")

    except Exception as e:
        print(f"Speak error: {e}")
        _speak_process = None


def stop_speaking():
    global _speak_process
    if _speak_process and _speak_process.poll() is None:
        try:
            import platform, subprocess
            if platform.system().lower() == "windows":
                subprocess.run(
                    ['taskkill', '/F', '/T', '/PID', str(_speak_process.pid)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
            else:
                _speak_process.terminate()
                _speak_process.wait(timeout=2)
            print("⏹ Speech stopped")
        except Exception as e:
            print(f"Stop speak error: {e}")
    _speak_process = None


def save_as_txt(text):
    """
    Sends a Save As request to the main thread via pdf_utils dialog queue.
    Returns the saved path, or None if cancelled.
    """
    from modules.pdf_utils import _ask
    path = _ask(
        kind         = "save_file",
        default_name = "ocr_result.txt",
        title        = "Save Text File As",
        filetypes    = [("Text Files", "*.txt"), ("All Files", "*.*")]
    )
    if not path:
        print("💾 TXT save cancelled")
        return None
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"💾 Saved: {path}")
    return path


def copy_to_clipboard(text):
    try:
        import pyperclip
        pyperclip.copy(text)
        print("📋 Copied")
    except ImportError:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        root.clipboard_clear()
        root.clipboard_append(text)
        root.update()
        root.after(500, root.destroy)
        root.mainloop()
        print("📋 Copied via tkinter")
