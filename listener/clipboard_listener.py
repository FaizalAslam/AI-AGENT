# listener/clipboard_listener.py
import time
import logging
import pyperclip

logger = logging.getLogger("OfficeAgent")

TRIGGER = "agent:"

class ClipboardListener:
    def __init__(self, cmd_buf):
        self.cmd_buf  = cmd_buf
        self._last    = ""
        self._last_err = ""
        self._last_err_at = 0.0

    def start(self):
        while True:
            try:
                current = pyperclip.paste()
                if current != self._last:
                    self._last = current
                    if isinstance(current, str) and current.strip().lower().startswith(TRIGGER):
                        logger.info(f"Clipboard candidate: {current.strip()}")
                        self.cmd_buf.set_candidate(current.strip())
            except Exception as e:
                msg = str(e)
                now = time.time()
                # WinError 0 OpenClipboard is noisy/transient; throttle logs.
                if msg != self._last_err or (now - self._last_err_at) > 10:
                    logger.warning(f"Clipboard error: {msg}")
                    self._last_err = msg
                    self._last_err_at = now
            time.sleep(0.5)
