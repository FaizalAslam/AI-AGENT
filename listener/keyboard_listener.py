# listener/keyboard_listener.py
import threading
import logging
from pynput import keyboard

logger = logging.getLogger("OfficeAgent")

class KeyboardListener:
    def __init__(self, on_command_callback, cmd_buf):
        self.on_command  = on_command_callback
        self.cmd_buf     = cmd_buf
        self._ctrl_v     = False

    def start(self):
        with keyboard.Listener(
            on_press=self._on_press
        ) as listener:
            listener.join()

    def _on_press(self, key):
        try:
            # Detect Ctrl+V (paste)
            if key == keyboard.Key.ctrl_l or key == keyboard.Key.ctrl_r:
                self._ctrl_v = True

            # On Enter — check if there's a buffered candidate
            if key == keyboard.Key.enter:
                candidate = self.cmd_buf.get_candidate()
                if candidate:
                    self.cmd_buf.clear()
                    logger.info(f"Keyboard trigger: {candidate}")
                    threading.Thread(
                        target=self.on_command,
                        args=(candidate,),
                        daemon=True
                    ).start()
        except Exception as e:
            logger.error(f"Keyboard listener error: {e}")
