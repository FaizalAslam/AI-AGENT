# modules/ui.py
import threading
import queue
import tkinter as tk
from tkinter import filedialog
import platform
import subprocess


def speak_text(text):
    text = str(text).replace(",", "").replace(".", "").strip()
    sys = platform.system().lower()
    try:
        if sys == "darwin":
            subprocess.Popen(["say", text]).wait()
        elif sys == "windows":
            ps = f'Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak("{text}")'
            subprocess.Popen(
                ["powershell", "-Command", ps],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            ).wait()
        else:
            subprocess.Popen(["espeak", text],
                             stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL).wait()
    except Exception as e:
        print(f"Speak error: {e}")


def speak(text):
    """Backward-compatible alias used by server routes."""
    speak_text(text)


def _run_file_dialog(result_queue, title, filetypes):
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        root.destroy()
        result_queue.put(path or "")
    except Exception as e:
        print(f"Dialog error: {e}")
        result_queue.put("")


def file_selector(title="Select a File", filetypes=None):
    if filetypes is None:
        filetypes = [("All Files", "*.*")]
    q = queue.Queue()
    t = threading.Thread(target=_run_file_dialog, args=(q, title, filetypes))
    t.start()
    t.join(timeout=60)
    try:
        return q.get_nowait() or ""
    except:
        return ""


def manual_selector():
    return file_selector(
        title="Select the Application Executable",
        filetypes=[("Executables", "*.exe *.lnk *.app"), ("All Files", "*.*")]
    )
