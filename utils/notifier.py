# utils/notifier.py
import logging
from modules.config import TOAST_ENABLED

logger = logging.getLogger("OfficeAgent")


def notify(app, message):
    logger.info(f"🔔 [{app}] {message}")
    if not TOAST_ENABLED:
        return
    try:
        from plyer import notification
        notification.notify(
            title=f"Office Agent — {app}",
            message=message,
            timeout=3
        )
    except Exception:
        pass
