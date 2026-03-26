# utils/command_buffer.py
import time
import threading

class CommandBuffer:
    CANDIDATE_TTL = 30  # seconds

    def __init__(self):
        self._lock      = threading.Lock()
        self._candidate = None
        self._set_at    = 0

    def set_candidate(self, text):
        with self._lock:
            self._candidate = text
            self._set_at    = time.time()

    def get_candidate(self):
        with self._lock:
            if self._candidate and (time.time() - self._set_at) <= self.CANDIDATE_TTL:
                return self._candidate
            return None

    def clear(self):
        with self._lock:
            self._candidate = None
            self._set_at    = 0
