import logging
import re
import threading
import time

try:
    import speech_recognition as sr
except Exception:  # pragma: no cover - optional dependency at runtime
    sr = None

logger = logging.getLogger("OfficeAgent")

SUPPORTED_APPS = ("excel", "word", "powerpoint", "ppt")
SYSTEM_PREFIXES = ("open ", "launch ", "start ", "run ", "boot ", "close ", "shut ", "exit ")


class VoiceListener:
    """
    Wake-word driven microphone listener.

    Flow:
    1) Hear wake word "agent".
    2) Open a 5-second command window.
    3) Accept one follow-up command, then sleep again.
    4) If no command within 5s, sleep and require wake word again.
    """

    def __init__(self, on_command_callback, wake_word="agent"):
        self.on_command = on_command_callback
        self.wake_word = (wake_word or "agent").lower().strip(": ")
        self.command_window_seconds = 5
        self._stop_event = threading.Event()
        self._thread = None
        self._is_running = False
        self._last_heard = ""
        self._last_heard_at = 0.0
        self._armed_until = 0.0
        self._last_error = ""

    @property
    def available(self):
        return sr is not None

    @property
    def is_running(self):
        return self._is_running

    @property
    def last_heard(self):
        return self._last_heard

    @property
    def last_heard_at(self):
        return self._last_heard_at

    @property
    def last_error(self):
        return self._last_error

    @property
    def armed(self):
        return time.time() < self._armed_until

    @property
    def armed_seconds_left(self):
        return max(0.0, self._armed_until - time.time())

    def _arm(self):
        self._armed_until = time.time() + float(self.command_window_seconds)

    def _disarm(self):
        self._armed_until = 0.0

    def start(self):
        if not self.available:
            self._last_error = "SpeechRecognition is not installed."
            return False
        if self._is_running:
            return True
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True, name="VoiceListener")
        self._thread.start()
        self._is_running = True
        logger.info("Voice listener started")
        return True

    def stop(self):
        self._stop_event.set()
        self._is_running = False
        self._disarm()
        logger.info("Voice listener stopped")
        return True

    def _normalize_wake_or_direct(self, spoken_text):
        text = (spoken_text or "").strip()
        if not text:
            return None
        lower = text.lower().strip()
        wake = self.wake_word

        # Already structured command.
        if lower.startswith("agent:"):
            return text

        # Wake only: "agent"
        if re.fullmatch(rf"{re.escape(wake)}[:\s,.-]*", lower):
            return "__WAKE__"

        # Wake + command in one sentence: "agent open chrome" or "agent excel ..."
        if lower.startswith(wake):
            rest = lower[len(wake):].strip(" :,-")
            if not rest:
                return "__WAKE__"

            office = re.match(r"^(excel|word|powerpoint|ppt)\s*:?\s+(.+)$", rest, re.DOTALL)
            if office:
                app = office.group(1).strip()
                cmd = office.group(2).strip()
                if app in SUPPORTED_APPS and cmd:
                    return f"agent: {app}: {cmd}"

            if rest.startswith(SYSTEM_PREFIXES):
                return f"agent {rest}"

        return None

    def _normalize_followup(self, spoken_text):
        """Follow-up command spoken after wake word, within command window."""
        text = (spoken_text or "").strip()
        if not text:
            return None
        lower = text.lower().strip()

        if lower.startswith(SYSTEM_PREFIXES):
            return f"agent {lower}"

        office = re.match(r"^(excel|word|powerpoint|ppt)\s*:?\s+(.+)$", lower, re.DOTALL)
        if office:
            app = office.group(1).strip()
            cmd = office.group(2).strip()
            if app in SUPPORTED_APPS and cmd:
                return f"agent: {app}: {cmd}"

        return None

    def _run(self):
        recognizer = sr.Recognizer()
        while not self._stop_event.is_set():
            try:
                with sr.Microphone() as source:
                    recognizer.adjust_for_ambient_noise(source, duration=0.4)
                    audio = recognizer.listen(source, timeout=3, phrase_time_limit=8)

                heard = recognizer.recognize_google(audio)
                self._last_heard = heard
                self._last_heard_at = time.time()
                self._last_error = ""

                now = time.time()
                if self._armed_until and now >= self._armed_until:
                    self._disarm()

                candidate = self._normalize_wake_or_direct(heard)
                if candidate == "__WAKE__":
                    self._arm()
                    logger.info("Wake detected. Listening for command (5s window).")
                    continue

                if candidate:
                    logger.info(f"Voice trigger: {candidate}")
                    threading.Thread(target=self.on_command, args=(candidate,), daemon=True).start()
                    self._disarm()
                    continue

                if self.armed:
                    follow = self._normalize_followup(heard)
                    if follow:
                        logger.info(f"Voice follow-up trigger: {follow}")
                        threading.Thread(target=self.on_command, args=(follow,), daemon=True).start()
                        self._disarm()
            except sr.WaitTimeoutError:
                if self._armed_until and time.time() >= self._armed_until:
                    self._disarm()
                continue
            except sr.UnknownValueError:
                if self._armed_until and time.time() >= self._armed_until:
                    self._disarm()
                continue
            except Exception as e:
                self._last_error = str(e)
                logger.error(f"Voice listener error: {e}")
                # Recover from intermittent microphone backend errors.
                time.sleep(1.5)
