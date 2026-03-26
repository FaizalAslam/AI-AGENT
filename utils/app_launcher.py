# utils/app_launcher.py
import logging
import os
import subprocess
import time
import pythoncom
import win32com.client
from modules.config import OFFICE_PATHS

logger = logging.getLogger("OfficeAgent")

COM_IDS = {
    "excel":      "Excel.Application",
    "word":       "Word.Application",
    "powerpoint": "PowerPoint.Application",
    "ppt":        "PowerPoint.Application",
}


class AppLauncher:

    def __init__(self):
        self._instances = {}

    def _candidate_paths(self, app_name):
        configured = OFFICE_PATHS.get(app_name, "")
        candidates = [configured]
        if configured:
            candidates.append(configured.replace("Office16", "Office15"))
            candidates.append(configured.replace("Office16", "Office17"))
            if "Program Files\\" in configured:
                candidates.append(configured.replace("Program Files", "Program Files (x86)"))
        seen = set()
        for path in candidates:
            if path and path not in seen:
                seen.add(path)
                yield path

    def _connect_active(self, com_id):
        try:
            return win32com.client.GetActiveObject(com_id)
        except Exception:
            return None

    def _is_alive(self, app):
        try:
            _ = app.Visible
            return True
        except Exception:
            return False

    def get(self, app_name):
        pythoncom.CoInitialize()
        com_id = COM_IDS.get(app_name)
        if not com_id:
            logger.error(f"No COM ID for: {app_name}")
            return None
        cached = self._instances.get(app_name)
        if cached and self._is_alive(cached):
            return cached
        app = self._connect_active(com_id)
        if app:
            self._instances[app_name] = app
            return app
        logger.info(f"🚀 Launching {app_name}...")
        launch_path = next(
            (p for p in self._candidate_paths(app_name) if os.path.exists(p)), None
        )
        if launch_path:
            try:
                subprocess.Popen(launch_path)
            except Exception as e:
                logger.error(f"Failed to launch {app_name}: {e}")
        for _ in range(10):
            time.sleep(1)
            app = self._connect_active(com_id)
            if app:
                self._instances[app_name] = app
                return app
        try:
            app = win32com.client.Dispatch(com_id)
            self._instances[app_name] = app
            return app
        except Exception as e:
            logger.error(f"All methods failed for {app_name}: {e}")
            return None
