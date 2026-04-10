"""
Razer Battery Monitor
=====================
System tray app that displays your Razer wireless device's battery percentage
and fires a Windows toast notification when it drops below a configurable threshold.

Reads battery state from Razer Synapse log files — requires Synapse 3 or 4 running.
One tray icon is created per device that reports battery information.

Usage:
    pythonw battery_monitor.pyw           # run silently (no console window)
    python  battery_monitor.pyw           # run with console (for debugging)
    python  battery_monitor.pyw --help    # show config options
"""

import os
import sys
import time
import threading
import re
import ctypes
import argparse
import json
import logging
import signal
import subprocess
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
import pystray

# ============================================================================
# CONFIGURATION DEFAULTS
# ============================================================================

DEFAULT_ALERT_THRESHOLD = 30        # percent — fires toast at or below this
DEFAULT_POLL_INTERVAL   = 2         # seconds between log checks
DEFAULT_ALERT_COOLDOWN  = 300       # seconds before re-alerting (5 min)
DEFAULT_SYNAPSE_VERSION = "auto"    # "3", "4", or "auto"
APP_NAME                = "Razer Battery Monitor"
APP_ID                  = "RazerBatteryMonitor"

# Log file paths per Synapse version
LOG_PATHS = {
    "3": Path.home() / "AppData/Local/Razer/Synapse3/Log/Razer Synapse 3.log",
    "4": Path.home() / "AppData/Local/Razer/RazerAppEngine/User Data/Logs",
}

# Glob pattern for Synapse 4 rotated logs (systray_systrayv2.log, systray_systrayv21.log, etc.)
LOG4_GLOB = "systray_systrayv2*.log"

# ============================================================================
# LOGGING SETUP
# ============================================================================

log = logging.getLogger(APP_NAME)

# ============================================================================
# DPI AWARENESS
# ============================================================================

def get_dpi_scale() -> float:
    """Get the system DPI scale factor (1.0 = 96 DPI, 1.5 = 144 DPI, etc.)."""
    try:
        ctypes.windll.user32.SetProcessDPIAware()
        dc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(dc, 88)  # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, dc)
        return dpi / 96.0
    except Exception:
        return 1.0

# ============================================================================
# LOG FILE DETECTION
# ============================================================================

def find_synapse4_log() -> Path | None:
    """
    Find the most recently modified Synapse 4 log file.
    Synapse rotates logs as systray_systrayv2.log, systray_systrayv21.log, etc.
    """
    log_dir = LOG_PATHS["4"]
    if not log_dir.is_dir():
        return None
    candidates = sorted(log_dir.glob(LOG4_GLOB), key=lambda p: p.stat().st_mtime)
    return candidates[-1] if candidates else None


def resolve_log_path(version: str) -> tuple[Path, str]:
    """
    Find the Synapse log file. If version is "auto", try 4 first (newer),
    then fall back to 3.  Returns (path, detected_version).
    """
    if version == "4":
        path = find_synapse4_log()
        if path:
            return path, "4"
        log.error(f"Synapse 4 log not found in: {LOG_PATHS['4']}/{LOG4_GLOB}")
        sys.exit(1)

    if version == "3":
        path = LOG_PATHS["3"]
        if path.exists():
            return path, "3"
        log.error(f"Synapse 3 log not found at: {path}")
        sys.exit(1)

    # Auto-detect: prefer 4, fall back to 3
    path = find_synapse4_log()
    if path:
        log.info(f"Auto-detected Synapse 4 log: {path}")
        return path, "4"

    path = LOG_PATHS["3"]
    if path.exists():
        log.info(f"Auto-detected Synapse 3 log: {path}")
        return path, "3"

    log.error(
        "Could not find Synapse log file in either location.\n"
        f"  Synapse 3: {LOG_PATHS['3']}\n"
        f"  Synapse 4: {LOG_PATHS['4']}/{LOG4_GLOB}\n"
        "Is Razer Synapse installed and running?"
    )
    sys.exit(1)

# ============================================================================
# LOG PARSING
# ============================================================================

# Legacy format (older Synapse 3): "... level 80  state 0 ..."
BATTERY_RE_LEGACY = re.compile(r"level\s+(\d+)\s+state\s+(\d+)")

# JSON format (newer Synapse 3 / Synapse 4) — spans multiple lines:
#   "powerStatus": {
#       "chargingStatus": "Charging",
#       "level": 34
#   }
# re.DOTALL so .+ crosses newlines; non-greedy so it grabs the first closing brace
POWER_STATUS_RE = re.compile(
    r'"powerStatus"\s*:\s*\{.+?\}', re.DOTALL
)

# Synapse 4: each line of the form "[timestamp] info: Device  [JSON array of devices]"
# We capture everything from the opening [ to end-of-line.
DEVICE_LOG_RE = re.compile(r'info: Device\s+(\[.+)')

# Charging status values seen in the wild (lowercase for comparison)
CHARGING_VALUES = {"charging", "connected"}

# How many bytes from the end of the log file to read when looking for the
# latest battery entry.  64 KB is plenty — battery entries are small and
# Synapse writes them frequently.
TAIL_CHUNK_SIZE = 65536


def _device_name_from_obj(dev: dict) -> str:
    """Extract the English device name from a Synapse 4 device object."""
    name_obj = dev.get("name", {})
    name = name_obj.get("en") or next(iter(name_obj.values()), "Unknown Device")
    return name.title()


def parse_devices_from_text(text: str) -> list[dict]:
    """
    Search a block of text for battery status across all devices.
    Returns the state from the LAST matching log entry (most recent).
    Each entry is {"device": str, "percent": int, "charging": bool}.

    Handles Synapse 4 device JSON arrays first; falls back to legacy formats
    for Synapse 3 logs (device name will be "Unknown Device" in those cases).
    """
    # --- Synapse 4: parse device JSON arrays ---
    # Each matching line is a full device-list snapshot; later lines override
    # earlier ones, so we process all matches and keep the last value per device.
    result: dict[str, dict] = {}  # device_name -> status (last write wins)

    for m in DEVICE_LOG_RE.finditer(text):
        try:
            devices = json.loads(m.group(1))
            for dev in devices:
                ps = dev.get("powerStatus")
                if ps is None:
                    continue
                level = ps.get("level")
                if level is None:
                    continue
                level = int(level)
                if not (0 <= level <= 100):
                    continue
                name = _device_name_from_obj(dev)
                status_str = ps.get("chargingStatus", "")
                result[name] = {
                    "device":   name,
                    "percent":  level,
                    "charging": status_str.lower() in CHARGING_VALUES,
                }
        except (json.JSONDecodeError, KeyError, ValueError, TypeError):
            continue

    if result:
        return list(result.values())

    # --- Newer Synapse 3: JSON powerStatus (no device name available) ---
    last = None
    for m in POWER_STATUS_RE.finditer(text):
        try:
            fragment = "{" + m.group(0) + "}"
            data = json.loads(fragment)
            ps = data["powerStatus"]
            level = ps.get("level")
            status = ps.get("chargingStatus", "")
            if level is not None and 0 <= int(level) <= 100:
                last = {
                    "device":   "Unknown Device",
                    "percent":  int(level),
                    "charging": status.lower() in CHARGING_VALUES,
                }
        except (json.JSONDecodeError, KeyError, ValueError, AttributeError):
            continue

    if last:
        return [last]

    # --- Legacy Synapse 3: "level N state N" ---
    for m in BATTERY_RE_LEGACY.finditer(text):
        level = int(m.group(1))
        if 0 <= level <= 100:
            last = {
                "device":   "Unknown Device",
                "percent":  level,
                "charging": int(m.group(2)) != 0,
            }

    return [last] if last else []


def read_all_statuses(log_path: Path) -> list[dict]:
    """Read the tail of the log file and return battery status for all devices."""
    if not log_path.exists():
        return []
    try:
        file_size = log_path.stat().st_size
        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            if file_size > TAIL_CHUNK_SIZE:
                f.seek(file_size - TAIL_CHUNK_SIZE)
            text = f.read()
        return parse_devices_from_text(text)
    except Exception as e:
        log.warning(f"Error reading log: {e}")
    return []

# ============================================================================
# ICON RENDERING
# ============================================================================

class IconRenderer:
    """
    Dynamically renders a tray icon showing the battery percentage as a number
    with a color-coded background. No pre-made icon images needed.
    """

    # Color thresholds: (max_percent, bg_color, text_color)
    COLORS = [
        (15,  "#CC3333", "#FFFFFF"),  # red — critical
        (30,  "#DD8800", "#FFFFFF"),  # orange — low
        (60,  "#4488CC", "#FFFFFF"),  # blue — moderate
        (100, "#33AA55", "#FFFFFF"),  # green — healthy
    ]

    CHARGING_BG   = "#2266AA"
    CHARGING_TEXT  = "#FFFFFF"

    def __init__(self, base_size: int = 32):
        scale = get_dpi_scale()
        self.size = int(base_size * scale)
        # Use a font size that fits 2-3 chars in the icon
        self.font_size = max(10, int(self.size * 0.55))
        self.font_size_small = max(8, int(self.size * 0.45))
        self._font = None
        self._font_small = None

    def _get_font(self, small: bool = False):
        """Load a font. Falls back through several options."""
        attr = "_font_small" if small else "_font"
        cached = getattr(self, attr)
        if cached:
            return cached

        size = self.font_size_small if small else self.font_size
        # Try common Windows fonts that render well at small sizes
        for font_name in ("segoeuib.ttf", "arialbd.ttf", "arial.ttf"):
            try:
                font = ImageFont.truetype(font_name, size)
                setattr(self, attr, font)
                return font
            except OSError:
                continue
        font = ImageFont.load_default()
        setattr(self, attr, font)
        return font

    def _pick_colors(self, percent: int, charging: bool) -> tuple[str, str]:
        if charging:
            return self.CHARGING_BG, self.CHARGING_TEXT
        for threshold, bg, fg in self.COLORS:
            if percent <= threshold:
                return bg, fg
        return self.COLORS[-1][1], self.COLORS[-1][2]

    def render(self, percent: int, charging: bool) -> Image.Image:
        """Create a tray icon image showing the battery percentage."""
        bg_color, text_color = self._pick_colors(percent, charging)

        img = Image.new("RGBA", (self.size, self.size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Rounded rectangle background
        margin = max(1, self.size // 16)
        radius = max(2, self.size // 6)
        draw.rounded_rectangle(
            [margin, margin, self.size - margin, self.size - margin],
            radius=radius,
            fill=bg_color,
        )

        # Battery percentage text
        text = str(percent)
        use_small = len(text) == 3  # "100" needs a smaller font
        font = self._get_font(small=use_small)

        bbox = draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        tx = (self.size - tw) // 2
        ty = (self.size - th) // 2 - bbox[1]  # compensate for font ascent offset

        draw.text((tx, ty), text, fill=text_color, font=font)

        # Charging indicator: small lightning-ish dot in top-right
        if charging:
            dot_r = max(2, self.size // 10)
            cx = self.size - margin - dot_r - 1
            cy = margin + dot_r + 1
            draw.ellipse(
                [cx - dot_r, cy - dot_r, cx + dot_r, cy + dot_r],
                fill="#FFDD00",
            )

        return img

    def render_unknown(self) -> Image.Image:
        """Render a '?' icon for when battery status is unknown."""
        img = Image.new("RGBA", (self.size, self.size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        margin = max(1, self.size // 16)
        radius = max(2, self.size // 6)
        draw.rounded_rectangle(
            [margin, margin, self.size - margin, self.size - margin],
            radius=radius,
            fill="#666666",
        )
        font = self._get_font()
        bbox = draw.textbbox((0, 0), "?", font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        tx = (self.size - tw) // 2
        ty = (self.size - th) // 2 - bbox[1]
        draw.text((tx, ty), "?", fill="#CCCCCC", font=font)
        return img

# ============================================================================
# TOAST NOTIFICATIONS
# ============================================================================

class AlertManager:
    """
    Fires a Windows toast notification when battery drops to/below threshold.
    Debounces so you only get alerted once per discharge cycle — it resets
    when the battery goes back above the threshold (i.e., you charged it).
    """

    def __init__(self, threshold: int, cooldown: int, device_name: str = ""):
        self.threshold = threshold
        self.cooldown = cooldown
        self.device_name = device_name
        self._alerted = False      # have we fired for this discharge cycle?
        self._last_alert = 0.0     # timestamp of last alert

    def check(self, percent: int, charging: bool):
        # Reset alert state when charging or back above threshold
        if charging or percent > self.threshold:
            self._alerted = False
            return

        # Fire if at/below threshold and haven't alerted this cycle
        if percent <= self.threshold and not self._alerted:
            now = time.time()
            if now - self._last_alert >= self.cooldown:
                self._fire(percent)
                self._alerted = True
                self._last_alert = now

    def _fire(self, percent: int):
        title = f"{self.device_name} Battery Low" if self.device_name else "Wireless Device Battery Low"
        msg = f"Battery is at {percent}%. Time to charge it!"
        # Use PowerShell with the registered PowerShell AUMID — works on
        # Windows 10/11 without requiring this app to be registered.
        ps_script = (
            "[Windows.UI.Notifications.ToastNotificationManager,"
            " Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null; "
            "$template = [Windows.UI.Notifications.ToastNotificationManager]::"
            "GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02); "
            f'$template.SelectSingleNode("//text[@id=1]").InnerText = "{title}"; '
            f'$template.SelectSingleNode("//text[@id=2]").InnerText = "{msg}"; '
            "$notif = [Windows.UI.Notifications.ToastNotification]::new($template); "
            "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("
            "'{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
            "\\WindowsPowerShell\\v1.0\\powershell.exe').Show($notif)"
        )
        try:
            subprocess.run(
                ["powershell", "-NonInteractive", "-NoProfile", "-Command", ps_script],
                capture_output=True,
                timeout=10,
            )
            log.info(f"Alert fired: {self.device_name} {percent}%")
        except Exception as e:
            log.warning(f"Failed to show notification: {e}")

# ============================================================================
# LOG WATCHER (poll-based)
# ============================================================================

def watch_log(log_path: Path, poll_interval: float, callback):
    """
    Periodically read the tail of the log file and fire the callback when
    the battery status changes for any device.

    Poll-based instead of line-tailing because the JSON powerStatus blocks
    span multiple lines, making readline()-based tailing unreliable.
    Battery percentage doesn't change fast enough for this to matter.
    """
    last_statuses: dict[str, dict] = {}  # device_name -> status

    while True:
        try:
            statuses = read_all_statuses(log_path)
            statuses_by_name = {s["device"]: s for s in statuses}
            if statuses_by_name != last_statuses:
                last_statuses = statuses_by_name
                callback(statuses)
                for s in statuses:
                    log.debug(f"{s['device']}: {s['percent']}% "
                              f"({'charging' if s['charging'] else 'discharging'})")
        except Exception as e:
            log.warning(f"Watch error: {e}")

        time.sleep(poll_interval)

# ============================================================================
# PER-DEVICE TRAY ICON
# ============================================================================

class DeviceIcon:
    """Manages a single system tray icon for one Razer device."""

    def __init__(self, device_name: str, threshold: int, cooldown: int,
                 renderer: IconRenderer, on_quit):
        self.device_name = device_name
        self.status = {"device": device_name, "percent": None, "charging": False}
        self.alert_mgr = AlertManager(threshold, cooldown, device_name)
        self._renderer = renderer
        self._on_quit = on_quit

        safe_id = re.sub(r"\W+", "_", device_name)
        self.icon = pystray.Icon(
            f"{APP_ID}_{safe_id}",
            renderer.render_unknown(),
            device_name,
            menu=self._make_menu(),
        )

    def _make_menu(self):
        return pystray.Menu(
            pystray.MenuItem(
                lambda item: (
                    f"{self.device_name}: {self.status['percent']}%"
                    if self.status["percent"] is not None
                    else f"{self.device_name}: Waiting..."
                ),
                action=None,
                enabled=False,
            ),
            pystray.MenuItem(
                lambda item: "Charging" if self.status["charging"] else "Discharging",
                action=None,
                enabled=False,
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", lambda *_: self._on_quit()),
        )

    def update(self, status: dict):
        self.status = status
        p, c = status["percent"], status["charging"]
        self.icon.icon = self._renderer.render(p, c)
        state = "Charging" if c else "Discharging"
        self.icon.title = f"{self.device_name}: {p}% ({state})"
        self.alert_mgr.check(p, c)

    def start(self):
        self.icon.run_detached()

    def stop(self):
        try:
            self.icon.stop()
        except Exception:
            pass

# ============================================================================
# TRAY APPLICATION
# ============================================================================

class BatteryTrayApp:
    def __init__(self, args):
        self.log_path, self.synapse_version = resolve_log_path(args.synapse)
        self.threshold      = args.threshold
        self.poll_interval  = args.poll_interval
        self.alert_cooldown = args.alert_cooldown
        self.renderer       = IconRenderer()
        self._device_icons: dict[str, DeviceIcon] = {}
        self._stop_event    = threading.Event()
        self._lock          = threading.Lock()

    def _quit_all(self):
        with self._lock:
            icons = list(self._device_icons.values())
        for di in icons:
            di.stop()
        self._stop_event.set()

    def _update(self, statuses: list[dict]):
        with self._lock:
            # Remove waiting placeholder on first real data
            if "__waiting__" in self._device_icons and statuses:
                self._device_icons.pop("__waiting__").stop()

            for status in statuses:
                name = status["device"]
                if name not in self._device_icons:
                    di = DeviceIcon(
                        name, self.threshold, self.alert_cooldown,
                        self.renderer, self._quit_all,
                    )
                    self._device_icons[name] = di
                    di.start()
                    log.info(f"New device icon: {name}")
                self._device_icons[name].update(status)

    def run(self):
        initial = read_all_statuses(self.log_path)
        if initial:
            log.info(f"Found {len(initial)} device(s) with battery info")
            self._update(initial)
        else:
            log.info("No initial battery data found, waiting for log updates...")
            # Show a placeholder so there's something visible in the tray
            waiting = DeviceIcon(
                "Razer Monitor", self.threshold, self.alert_cooldown,
                self.renderer, self._quit_all,
            )
            self._device_icons["__waiting__"] = waiting
            waiting.start()

        # Start log watcher in background thread
        watcher = threading.Thread(
            target=watch_log,
            args=(self.log_path, self.poll_interval, self._update),
            daemon=True,
        )
        watcher.start()

        # Allow Ctrl+C / SIGTERM to shut down all tray icons cleanly
        def _stop(sig, frame):
            self._quit_all()
        signal.signal(signal.SIGINT, _stop)
        signal.signal(signal.SIGTERM, _stop)

        # Block until quit is requested
        self._stop_event.wait()

# ============================================================================
# ENTRY POINT
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description=APP_NAME,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  pythonw battery_monitor.pyw                   # run with defaults\n"
            "  pythonw battery_monitor.pyw --threshold 20    # alert at 20%%\n"
            "  pythonw battery_monitor.pyw --synapse 4       # force Synapse 4 log\n"
        ),
    )
    parser.add_argument(
        "--threshold", type=int, default=DEFAULT_ALERT_THRESHOLD,
        help=f"Battery percent at or below which to fire an alert (default: {DEFAULT_ALERT_THRESHOLD})",
    )
    parser.add_argument(
        "--poll-interval", type=float, default=DEFAULT_POLL_INTERVAL,
        help=f"Seconds between log file checks (default: {DEFAULT_POLL_INTERVAL})",
    )
    parser.add_argument(
        "--alert-cooldown", type=int, default=DEFAULT_ALERT_COOLDOWN,
        help=f"Minimum seconds between repeated alerts (default: {DEFAULT_ALERT_COOLDOWN})",
    )
    parser.add_argument(
        "--synapse", choices=["3", "4", "auto"], default=DEFAULT_SYNAPSE_VERSION,
        help=f"Which Synapse log to read (default: {DEFAULT_SYNAPSE_VERSION})",
    )
    parser.add_argument(
        "--debug", action="store_true",
        help="Enable debug logging to console",
    )

    args = parser.parse_args()

    if not 0 <= args.threshold <= 100:
        parser.error("--threshold must be between 0 and 100")
    if args.poll_interval <= 0:
        parser.error("--poll-interval must be greater than 0")
    if args.alert_cooldown < 0:
        parser.error("--alert-cooldown must be >= 0")

    # Configure logging
    level = logging.DEBUG if args.debug else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    log.info(f"Starting {APP_NAME}")
    log.info(f"  Threshold:    {args.threshold}%")
    log.info(f"  Poll interval: {args.poll_interval}s")
    log.info(f"  Synapse:       {args.synapse}")

    app = BatteryTrayApp(args)
    app.run()


if __name__ == "__main__":
    main()
