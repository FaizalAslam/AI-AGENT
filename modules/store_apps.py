# modules/store_apps.py
import subprocess
import json
import logging

logger = logging.getLogger("OfficeAgent")
_cache = None


def get_store_apps():
    global _cache
    if _cache is not None:
        return _cache
    try:
        ps_script = (
            "Get-StartApps | "
            "Where-Object { $_.AppId -notmatch '^{' } | "
            "Select-Object Name, AppId | "
            "ConvertTo-Json -Compress"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True, text=True, timeout=15
        )
        if result.returncode != 0 or not result.stdout.strip():
            logger.warning(f"Get-StartApps failed: {result.stderr}")
            _cache = {}
            return _cache

        raw = json.loads(result.stdout.strip())
        if isinstance(raw, dict):
            raw = [raw]

        _cache = {
            item["Name"].lower().strip(): item["AppId"]
            for item in raw
            if item.get("Name") and item.get("AppId")
        }
        logger.info(f"Store apps loaded: {len(_cache)} entries")
        return _cache

    except Exception as e:
        logger.error(f"get_store_apps error: {e}")
        _cache = {}
        return _cache


def find_app_id(app_name, aliases=None):
    store_map = get_store_apps()
    if not store_map:
        return None

    name_clean = app_name.lower().strip()

    if name_clean in store_map:
        return store_map[name_clean]

    probes = [name_clean, name_clean.replace(" ", "")]
    if aliases:
        probes.extend([a.lower() for a in aliases])

    for key, app_id in store_map.items():
        key_clean = key.replace(" ", "").lower()
        for probe in probes:
            p = probe.replace(" ", "").lower()
            if p and len(p) >= 3 and (p in key_clean or key_clean in p):
                return app_id

    return _find_via_appx_package(app_name)


def _find_via_appx_package(app_name):
    try:
        ps_script = (
            f"$pkg = Get-AppxPackage -Name '*{app_name}*' | "
            "Select-Object -First 1; "
            "if ($pkg) { $pkg.PackageFamilyName }"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True, text=True, timeout=10
        )
        pfn = result.stdout.strip()
        if pfn and "!" not in pfn:
            return f"{pfn}!App"
        return pfn or None
    except Exception as e:
        logger.error(f"_find_via_appx_package error: {e}")
        return None


def launch_store_app(app_id):
    try:
        target = f"shell:AppsFolder\\{app_id}"
        subprocess.run(
            ["explorer.exe", target],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        return True
    except Exception as e:
        logger.error(f"launch_store_app failed: {e}")
        return False


def refresh_cache():
    global _cache
    _cache = None
    return get_store_apps()
