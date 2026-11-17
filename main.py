import os
import sys
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
import shutil
import tempfile
import subprocess
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from openpyxl import load_workbook
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QPixmap, QPainter, QColor, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QFrame,
    QVBoxLayout, QHBoxLayout, QGridLayout, QMessageBox, QLineEdit, QProgressBar, QScrollArea, QSizePolicy
)

APP_NAME = "MY TIME TRAVEL"
BASE_DIR = Path(__file__).resolve().parent
SOURCE_DB = BASE_DIR / "World_TimeZone_Database.xlsx"
BACKGROUND = BASE_DIR / "assets" / "background_original.jpg"
REGIONAL_DEFAULTS = BASE_DIR / "regional_defaults.json"


# Hide CMD / PowerShell windows when subprocess commands run on Windows.
def hidden_subprocess_options():
    if os.name != "nt":
        return {}

    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE

    return {
        "startupinfo": startupinfo,
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
    }


# Common fallback map. tzlocal supplies a much larger IANA -> Windows map
# when available; this map keeps common zones working even if that lookup fails.
FALLBACK_WINDOWS = {
    "Etc/UTC": "UTC", "UTC": "UTC",
    "Europe/London": "GMT Standard Time",
    "Europe/Berlin": "W. Europe Standard Time",
    "Europe/Paris": "Romance Standard Time",
    "Europe/Madrid": "Romance Standard Time",
    "Europe/Rome": "W. Europe Standard Time",
    "Europe/Amsterdam": "W. Europe Standard Time",
    "Europe/Zurich": "W. Europe Standard Time",
    "Europe/Vienna": "W. Europe Standard Time",
    "Europe/Stockholm": "W. Europe Standard Time",
    "Europe/Moscow": "Russian Standard Time",
    "Europe/Istanbul": "Turkey Standard Time",
    "America/New_York": "Eastern Standard Time",
    "America/Toronto": "Eastern Standard Time",
    "America/Chicago": "Central Standard Time",
    "America/Denver": "Mountain Standard Time",
    "America/Los_Angeles": "Pacific Standard Time",
    "America/Phoenix": "US Mountain Standard Time",
    "America/Anchorage": "Alaskan Standard Time",
    "Pacific/Honolulu": "Hawaiian Standard Time",
    "Asia/Kolkata": "India Standard Time",
    "Asia/Dubai": "Arabian Standard Time",
    "Asia/Karachi": "Pakistan Standard Time",
    "Asia/Dhaka": "Bangladesh Standard Time",
    "Asia/Kathmandu": "Nepal Standard Time",
    "Asia/Bangkok": "SE Asia Standard Time",
    "Asia/Singapore": "Singapore Standard Time",
    "Asia/Hong_Kong": "China Standard Time",
    "Asia/Shanghai": "China Standard Time",
    "Asia/Tokyo": "Tokyo Standard Time",
    "Asia/Seoul": "Korea Standard Time",
    "Asia/Jerusalem": "Israel Standard Time",
    "Asia/Riyadh": "Arab Standard Time",
    "Australia/Sydney": "AUS Eastern Standard Time",
    "Australia/Perth": "W. Australia Standard Time",
    "Australia/Adelaide": "Cen. Australia Standard Time",
    "Pacific/Auckland": "New Zealand Standard Time",
    "Africa/Cairo": "Egypt Standard Time",
    "Africa/Johannesburg": "South Africa Standard Time",
    "America/Sao_Paulo": "E. South America Standard Time",
    "America/Argentina/Buenos_Aires": "Argentina Standard Time",
}


def app_data_dir():
    root = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    path = Path(root) / "MY_TIME_TRAVEL"
    path.mkdir(parents=True, exist_ok=True)
    return path


def database_path():
    """Use a writable per-user copy; never modify the bundled original."""
    user_db = app_data_dir() / "World_TimeZone_Database.xlsx"
    if not user_db.exists():
        if not SOURCE_DB.exists():
            raise FileNotFoundError("Bundled World_TimeZone_Database.xlsx is missing.")
        shutil.copy2(SOURCE_DB, user_db)
    return user_db


def windows_zone_for(iana_name):
    # Prefer tzlocal's comprehensive mapping, then the built-in safe fallback.
    try:
        from tzlocal.windows_tz import tz_names
        # tzlocal's table is commonly Windows-ID -> IANA-zone; invert it safely.
        direct = tz_names.get(iana_name)
        if direct and isinstance(direct, str) and "/" not in direct:
            return direct
        for windows_id, mapped in tz_names.items():
            if isinstance(mapped, str) and mapped == iana_name:
                return windows_id
            if isinstance(mapped, (list, tuple, set)) and iana_name in mapped:
                return windows_id
    except Exception:
        pass
    return FALLBACK_WINDOWS.get(iana_name)


def current_offset(iana_name):
    now = datetime.now(ZoneInfo(iana_name))
    offset = now.utcoffset()
    if offset is None:
        return ""
    total = int(offset.total_seconds() // 60)
    sign = "+" if total >= 0 else "-"
    total = abs(total)
    return f"UTC{sign}{total // 60:02d}:{total % 60:02d}"


def dst_difference_text(iana_name):
    """Return the zone's seasonal DST adjustment, without guessing a street-level location."""
    try:
        zone = ZoneInfo(iana_name)
        now = datetime.now(zone)
        current = now.utcoffset()
        jan = datetime(now.year, 1, 15, 12, 0, tzinfo=zone).utcoffset()
        jul = datetime(now.year, 7, 15, 12, 0, tzinfo=zone).utcoffset()
        if current is None or jan is None or jul is None:
            return "DST Difference: Not available"
        minutes = int(abs((jul - jan).total_seconds()) // 60)
        sign = "+" if (now.dst() and now.dst().total_seconds() > 0) else ""
        return f"DST Difference: {minutes} minutes seasonal change; current {now.strftime('%z')}"
    except Exception:
        return "DST Difference: Not available"


def current_windows_timezone():
    try:
        p = subprocess.run(
            ["tzutil", "/g"],
            capture_output=True,
            text=True,
            timeout=5,
            **hidden_subprocess_options()
        )
        if p.returncode == 0 and p.stdout.strip():
            return p.stdout.strip()
    except Exception:
        pass
    return "Unknown"


def get_existing_headers(ws):
    return [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]


def ensure_database_sheets(wb):
    if "Locations" not in wb.sheetnames:
        ws = wb.create_sheet("Locations")
        ws.append(["Country", "ISO2", "Location / Zone label", "IANA Time Zone",
                   "Current UTC Offset*", "Windows Time Zone (common mapping)", "Source / Status"])
    if "Windows Zones" not in wb.sheetnames:
        ws = wb.create_sheet("Windows Zones")
        ws.append(["Windows Time Zone ID", "Example / Notes"])
    return wb["Locations"], wb["Windows Zones"]


def find_timezone_record(path, iana):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if "Locations" not in wb.sheetnames:
            return None
        ws = wb["Locations"]
        headers = get_existing_headers(ws)
        try:
            col = headers.index("IANA Time Zone") + 1
        except ValueError:
            col = 4
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) >= col and str(row[col - 1] or "").strip().casefold() == iana.casefold():
                return {
                    "country": row[0] if len(row) > 0 else "",
                    "city": row[2] if len(row) > 2 else "",
                    "iana": row[col - 1],
                    "windows": row[5] if len(row) > 5 else "",
                    "status": row[6] if len(row) > 6 else "Existing database record",
                }
        return None
    finally:
        wb.close()


def learn_timezone(path, location):
    """Validate, back up, and append a new zone without duplicate IANA IDs."""
    iana = (location.get("timezone") or "").strip()
    if not iana:
        return "not_added", "IP lookup did not return a time zone."
    try:
        ZoneInfo(iana)
    except (ZoneInfoNotFoundError, ValueError):
        return "not_added", f"'{iana}' is not a valid IANA time zone; it was not saved."
    existing = find_timezone_record(path, iana)
    if existing:
        return "exists", "This IANA time zone is already in the Excel database."

    wb = load_workbook(path)
    try:
        locations_ws, windows_ws = ensure_database_sheets(wb)
        # Check again inside the workbook before appending.
        headers = get_existing_headers(locations_ws)
        try:
            tz_col = headers.index("IANA Time Zone") + 1
        except ValueError:
            tz_col = 4
        for row in locations_ws.iter_rows(min_row=2, values_only=True):
            if len(row) >= tz_col and str(row[tz_col - 1] or "").strip().casefold() == iana.casefold():
                return "exists", "This IANA time zone is already in the Excel database."

        # Preserve existing column order and append a provenance column if absent.
        if "Source / Status" not in headers:
            locations_ws.cell(1, locations_ws.max_column + 1, "Source / Status")
            headers = get_existing_headers(locations_ws)

        values = {
            "Country": location.get("country", ""),
            "ISO2": location.get("country_code", ""),
            "Location / Zone label": location.get("city", "") or location.get("region", "") or iana,
            "IANA Time Zone": iana,
            "Current UTC Offset*": current_offset(iana),
            "Windows Time Zone (common mapping)": windows_zone_for(iana) or "",
            "Source / Status": "IP Lookup - Auto Learned (unverified city)",
        }
        next_row = locations_ws.max_row + 1
        for col, header in enumerate(headers, start=1):
            if header in values:
                locations_ws.cell(next_row, col, values[header])

        win_id = values["Windows Time Zone (common mapping)"]
        if win_id:
            known = {str(r[0]).strip() for r in windows_ws.iter_rows(min_row=2, values_only=True) if r and r[0]}
            if win_id not in known:
                windows_ws.append([win_id, "Auto-added from validated IANA mapping"])
        else:
            # Keep the zone record, but never invent a Windows ID.
            pass

        # Atomic-ish save: backup original first, save temp in same directory, replace.
        backup = path.with_suffix(path.suffix + ".bak")
        if not backup.exists():
            shutil.copy2(path, backup)
        fd, temp_name = tempfile.mkstemp(prefix="mtt_db_", suffix=".xlsx", dir=str(path.parent))
        os.close(fd)
        try:
            wb.save(temp_name)
            os.replace(temp_name, path)
        finally:
            if os.path.exists(temp_name):
                os.remove(temp_name)
        return "added", ("New time zone saved to Excel." if win_id else
                         "New IANA zone saved. Windows mapping is missing; marked for review.")
    finally:
        wb.close()


def _lookup_provider(url, provider):
    req = urllib.request.Request(url, headers={"User-Agent": "MY-TIME-TRAVEL/2.1"})
    with urllib.request.urlopen(req, timeout=8) as response:
        data = json.loads(response.read().decode("utf-8"))
    if provider == "ipapi":
        if data.get("error"):
            raise RuntimeError(data.get("reason", "IP lookup rejected request"))
        result = {
            "ip": data.get("ip", ""), "city": data.get("city", ""),
            "region": data.get("region", ""), "country": data.get("country_name", ""),
            "country_code": data.get("country_code", ""), "timezone": data.get("timezone", ""),
            "org": data.get("org", ""), "latitude": data.get("latitude"),
            "longitude": data.get("longitude"), "accuracy": data.get("accuracy"),
        }
    else:
        if not data.get("success", False):
            raise RuntimeError(data.get("message", "IP lookup rejected request"))
        tz = data.get("timezone") or {}
        result = {
            "ip": data.get("ip", ""), "city": data.get("city", ""),
            "region": data.get("region", ""), "country": data.get("country", ""),
            "country_code": data.get("country_code", ""), "timezone": tz.get("id", ""),
            "org": (data.get("connection") or {}).get("isp", ""),
            "latitude": data.get("latitude"), "longitude": data.get("longitude"),
            "accuracy": data.get("accuracy"),
        }
    if not result["ip"] or not result["timezone"]:
        raise RuntimeError("Provider response did not include public IP and time zone.")
    result["provider"] = provider
    return result


def fetch_ip_location():
    # Query both independent providers at once; use the first complete valid result.
    endpoints = [
        ("https://ipapi.co/json/", "ipapi"),
        ("https://ipwho.is/", "ipwho"),
    ]
    errors = []
    pool = ThreadPoolExecutor(max_workers=2)
    futures = [pool.submit(_lookup_provider, url, provider) for url, provider in endpoints]
    try:
        for future in as_completed(futures):
            try:
                result = future.result()
                for pending in futures:
                    if pending is not future:
                        pending.cancel()
                return result
            except Exception as exc:
                errors.append(str(exc))
    finally:
        # Do not wait for a slower fallback after a complete result has arrived.
        pool.shutdown(wait=False, cancel_futures=True)
    raise RuntimeError("Both IP lookup services failed.\n" + "\n".join(errors))


def apply_windows_timezone(windows_id):
    if not windows_id or any(ch in windows_id for ch in "\r\n"):
        raise ValueError("Windows time zone ID is empty or invalid.")
    subprocess.run(
        ["tzutil", "/s", windows_id],
        check=True,
        capture_output=True,
        text=True,
        timeout=10,
        **hidden_subprocess_options()
    )


def run_powershell(script, timeout=15):
    """Run a fixed PowerShell command and return stdout; never use shell=True."""
    exe = shutil.which("powershell.exe") or shutil.which("powershell")
    if not exe:
        raise RuntimeError("Windows PowerShell was not found on this computer.")
    result = subprocess.run(
        [exe, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        timeout=timeout,
        **hidden_subprocess_options()
    )
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "PowerShell command failed.").strip()
        raise RuntimeError(detail)
    return result.stdout.strip()


def load_regional_defaults():
    try:
        with REGIONAL_DEFAULTS.open("r", encoding="utf-8") as handle:
            return json.load(handle).get("countries", {})
    except Exception:
        return {}


def windows_geoid_for_country(iso2):
    """Resolve ISO2 to a unique Windows GeoID using Windows .NET RegionInfo."""
    code = (iso2 or "").strip().upper()
    if len(code) != 2 or not code.isalpha():
        return None
    script = ("$ErrorActionPreference='Stop'; "
              "$codes=@(); "
              "foreach($c in [Globalization.CultureInfo]::GetCultures([Globalization.CultureTypes]::SpecificCultures)){ "
              "try{$r=[Globalization.RegionInfo]::new($c.Name); "
              f"if($r.TwoLetterISORegionName -eq '{code}' -and $r.GeoId -gt 0){{$codes += [int]$r.GeoId}} "
              "}catch{} }; "
              "$u=@($codes | Sort-Object -Unique); "
              "if($u.Count -eq 1){[string]$u[0]} elseif($u.Count -gt 1){'AMBIGUOUS'} else {'MISSING'}")
    value = run_powershell(script).strip()
    if value.isdigit():
        return int(value)
    return None


def validate_windows_culture(iso2, culture):
    """Validate that .NET accepts the culture and it belongs to the detected ISO2."""
    code = (iso2 or "").strip().upper()
    # Culture is read only from our bundled, generated mapping; reject unsafe text.
    import re
    if not re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z]{4})?-[A-Z]{2}", culture or ""):
        return False
    safe_culture = culture
    safe_code = code
    script = ("$ErrorActionPreference='Stop'; "
              f"$c=[Globalization.CultureInfo]::GetCultureInfo('{safe_culture}'); "
              "$r=[Globalization.RegionInfo]::new($c.Name); "
              f"if($r.TwoLetterISORegionName -ne '{safe_code}'){{throw 'Culture territory does not match detected country.'}}; "
              "'OK'")
    try:
        return run_powershell(script) == "OK"
    except Exception:
        return False


def apply_windows_region(geoid):
    if not isinstance(geoid, int) or geoid <= 0:
        raise ValueError("Windows GeoID is missing or invalid.")
    run_powershell(f"$ErrorActionPreference='Stop'; Set-WinHomeLocation -GeoId {geoid} -ErrorAction Stop")


def apply_regional_format(culture):
    import re
    if not re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z]{4})?-[A-Z]{2}", culture or ""):
        raise ValueError("Regional format culture is invalid.")
    run_powershell(f"$ErrorActionPreference='Stop'; Set-Culture -CultureInfo '{culture}' -ErrorAction Stop")


def current_windows_region():
    try:
        return run_powershell("$x=Get-WinHomeLocation; [string]$x.GeoId + ' (' + [string]$x.HomeLocation + ')'")
    except Exception:
        return "Unknown"


def current_regional_format():
    try:
        return run_powershell("(Get-Culture).Name")
    except Exception:
        return "Unknown"


class LookupWorker(QThread):
    succeeded = Signal(dict)
    failed = Signal(str)

    def run(self):
        try:
            self.succeeded.emit(fetch_ip_location())
        except Exception as exc:
            self.failed.emit(str(exc))


class BackgroundWidget(QWidget):
    def __init__(self, image_path):
        super().__init__()
        self.pixmap = QPixmap(str(image_path)) if image_path.exists() else QPixmap()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#020806"))
        if not self.pixmap.isNull():
            # Fit the complete 9:16 portrait image into the available background area.
            # KeepAspectRatio avoids the heavy crop/zoom used by the previous build.
            scaled = self.pixmap.scaled(self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            painter.drawPixmap((self.width()-scaled.width())//2,
                               (self.height()-scaled.height())//2, scaled)
        painter.fillRect(self.rect(), QColor(0, 8, 5, 28))
        painter.end()


class GlassPanel(QFrame):
    def __init__(self, title):
        super().__init__()
        self.setObjectName("glassPanel")
        self.setStyleSheet("""
            QFrame#glassPanel { background: rgba(0, 12, 10, 112);
                border: 1px solid #00e878; border-radius: 16px; }
            QLabel { color: #f3fff9; background: transparent; }
        """)
        self.box = QVBoxLayout(self)
        self.box.setContentsMargins(7, 5, 7, 5)
        self.box.setSpacing(2)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        heading = QLabel(title)
        heading.setWordWrap(True)
        heading.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        heading.setFont(QFont("Segoe UI", 10, QFont.Bold))
        heading.setStyleSheet("color:#42ff8b;")
        self.box.addWidget(heading)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        # Compact dashboard: fit the cards into a maximized desktop window.
        # Scrolling remains available only when the actual display is too small.
        self.resize(1180, 760)
        self.setMinimumSize(820, 560)
        self.detected = None
        self.suggested_windows_id = None
        self.suggested_geoid = None
        self.suggested_culture = None
        self.regional_defaults = load_regional_defaults()
        self.worker = None
        try:
            self.db_path = database_path()
            self.db_ready = True
        except Exception as exc:
            self.db_path = None
            self.db_ready = False
            self.db_error = str(exc)

        # Do not force a tall content canvas: let the dashboard fit naturally.
        self.content_root = BackgroundWidget(BACKGROUND)
        self.content_root.setMinimumSize(0, 0)
        self.content_root.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; } QScrollBar:vertical { width: 12px; background: rgba(0,0,0,100); } QScrollBar::handle:vertical { background: #00b96b; border-radius: 5px; min-height: 30px; }")
        scroll.setWidget(self.content_root)
        self.setCentralWidget(scroll)
        root = self.content_root
        outer = QVBoxLayout(root)
        outer.setContentsMargins(8, 5, 8, 5)
        outer.setSpacing(3)

        top = QHBoxLayout()
        brand = QLabel("🌐  MY TIME TRAVEL")
        brand.setFont(QFont("Segoe UI", 17, QFont.Bold))
        brand.setStyleSheet("color:white; background:transparent;")
        self.online = QLabel("●  READY")
        self.online.setStyleSheet("color:#45ff87; background:rgba(0,30,12,200); padding:6px 12px; border:1px solid #00e878; border-radius:14px; font-weight:bold;")
        top.addWidget(brand)
        top.addStretch()
        top.addWidget(self.online)
        outer.addLayout(top)

        tagline = QLabel("WORLD LOCATION  •  VPN / PROXY  •  TIME ZONE MANAGER  •  AUTO-LEARN DATABASE")
        tagline.setStyleSheet("color:#9dffbf; background:transparent; font-weight:bold;")
        outer.addWidget(tagline)

        columns = QHBoxLayout()
        columns.setSpacing(9)
        self.location_panel = GlassPanel("📍  PUBLIC IP LOCATION")
        self.location_panel.setMinimumWidth(0)
        self.location_panel.setMinimumHeight(390)
        self.location_panel.box.setContentsMargins(10, 9, 10, 9)
        self.location_panel.box.setSpacing(7)
        self.location_grid = QGridLayout()
        self.location_grid.setContentsMargins(0, 0, 0, 0)
        self.location_grid.setHorizontalSpacing(7)
        self.location_grid.setVerticalSpacing(7)
        self.location_panel.box.addLayout(self.location_grid)
        self.ip_label = self.location_line("Public IP: Not checked")
        self.country_label = self.location_line("Country: —")
        self.city_label = self.location_line("City: —")
        self.network_label = self.location_line("Network / ISP: —")
        self.provider_label = self.location_line("Lookup provider: —")
        self.detected_zone_label = self.location_line("Detected IANA time zone: —")
        self.coordinates_label = self.location_line("Latitude / Longitude: —")
        self.accuracy_label = self.location_line("Accuracy: IP-based estimate")
        self.address_label = self.location_line("Location: —")
        self.dst_label = self.location_line("DST Difference: —")
        for row, pair in enumerate([
            (self.ip_label, None), (self.country_label, self.city_label),
            (self.network_label, self.provider_label),
            (self.detected_zone_label, self.coordinates_label),
            (self.accuracy_label, self.dst_label), (self.address_label, None)
        ]):
            left, right_label = pair
            self.location_grid.addWidget(left, row, 0, 1, 2 if right_label is None else 1)
            if right_label is not None:
                self.location_grid.addWidget(right_label, row, 1)
        self.detect_button = QPushButton("⌕   DETECT LOCATION")
        self.detect_button.clicked.connect(self.detect_location)
        self.button_style(self.detect_button, "#006fc9")
        self.location_panel.box.addWidget(self.detect_button)

        right_widget = QWidget()
        right_widget.setMinimumWidth(0)
        right_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        right = QVBoxLayout(right_widget)
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(4)
        self.time_panel = GlassPanel("◷  TIME ZONE MATCH")
        self.current_zone_label = self.line(self.time_panel, "Current Windows zone: " + current_windows_timezone())
        self.suggested_zone_label = self.line(self.time_panel, "Suggested Windows zone: —")
        self.db_label = self.line(self.time_panel, "Database: Ready (user copy)" if self.db_ready else "Database: ERROR: " + self.db_error)
        self.db_status_label = self.line(self.time_panel, "Database match: Waiting for detection")
        self.action_panel = GlassPanel("⚙  TIME ZONE CONFIRMATION")
        self.status_label = self.line(self.action_panel, "Status: Waiting for location detection")
        self.apply_button = QPushButton("✓   YES – APPLY TIME ZONE")
        self.apply_button.setEnabled(False)
        self.apply_button.clicked.connect(self.confirm_apply)
        self.button_style(self.apply_button, "#008b38")
        self.cancel_button = QPushButton("✕   NO – CANCEL")
        self.cancel_button.clicked.connect(self.cancel_apply)
        self.button_style(self.cancel_button, "#9d1726")
        row = QHBoxLayout()
        row.addWidget(self.apply_button)
        row.addWidget(self.cancel_button)
        self.action_panel.box.addLayout(row)
        note = QLabel("Windows time zone changes only after your YES confirmation. IP location may be approximate.")
        note.setWordWrap(True)
        note.setStyleSheet("color:#b9ffd2; background:transparent;")
        self.action_panel.box.addWidget(note)

        self.region_panel = GlassPanel("🌍  WINDOWS REGION")
        self.current_region_label = self.line(self.region_panel, "Current Windows Region: " + current_windows_region())
        self.suggested_region_label = self.line(self.region_panel, "Suggested Windows Region: —")
        self.region_button = QPushButton("✓   YES – APPLY WINDOWS REGION")
        self.region_button.setEnabled(False)
        self.region_button.clicked.connect(self.confirm_region_apply)
        self.button_style(self.region_button, "#006b72")
        self.region_panel.box.addWidget(self.region_button)

        self.format_panel = GlassPanel("🔤  REGIONAL FORMAT")
        self.current_format_label = self.line(self.format_panel, "Current Regional Format: " + current_regional_format())
        self.suggested_format_label = self.line(self.format_panel, "Suggested Regional Format: —")
        self.format_button = QPushButton("✓   YES – APPLY REGIONAL FORMAT")
        self.format_button.setEnabled(False)
        self.format_button.clicked.connect(self.confirm_format_apply)
        self.button_style(self.format_button, "#6652a5")
        self.format_panel.box.addWidget(self.format_button)
        self.region_note = QLabel("Region = Windows home location. Regional Format = date/number formatting. Neither changes Windows display language, keyboard, GPS or IP address.")
        self.region_note.setWordWrap(True)
        self.region_note.setStyleSheet("color:#b9ffd2; background:transparent;")
        self.format_panel.box.addWidget(self.region_note)

        right.addWidget(self.time_panel)
        right.addWidget(self.action_panel)
        right.addWidget(self.region_panel)
        right.addWidget(self.format_panel)
        # Use the full available width: wider cards reduce wrapped labels and wasted margins.
        self.location_panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        columns.addWidget(self.location_panel, 5, Qt.AlignTop)
        columns.addWidget(right_widget, 6, Qt.AlignTop)
        outer.addLayout(columns, 1)

        footer = QLabel("AUTO-LEARN: Missing validated IANA zones are added to your writable Excel copy • Backup is created before first update")
        footer.setWordWrap(True)
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("color:#4dff91; background:rgba(0,0,0,175); padding:3px; border:1px solid #087c43; border-radius:10px;")
        clock_row = QHBoxLayout()
        clock_row.setSpacing(8)
        self.current_clock = QLabel("")
        self.current_clock.setAlignment(Qt.AlignCenter)
        self.current_clock.setWordWrap(True)
        self.current_clock.setMinimumHeight(44)
        self.current_clock.setStyleSheet("color:white; background:rgba(0,0,0,185); padding:4px; border:1px solid #087c43; border-radius:8px;")
        self.target_clock = QLabel("DETECTED / NEW TIME ZONE\nDetect a location to preview its live time")
        self.target_clock.setAlignment(Qt.AlignCenter)
        self.target_clock.setWordWrap(True)
        self.target_clock.setMinimumHeight(44)
        self.target_clock.setStyleSheet("color:#9dffbf; background:rgba(0,0,0,185); padding:4px; border:1px solid #087c43; border-radius:8px;")
        clock_row.addWidget(self.current_clock, 1)
        clock_row.addWidget(self.target_clock, 1)
        outer.addLayout(clock_row)
        outer.addWidget(footer)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_clock)
        self.timer.start(1000)
        self.update_clock()
        self.setStyleSheet("""
            QWidget { font-family: 'Segoe UI'; }
            QPushButton { color:white; font-size:9px; font-weight:bold; padding:4px 5px;
                border-radius:10px; border:1px solid #51e8ff; }
            QPushButton:disabled { color:#82918a; background:#25312d; border-color:#56645d; }
        """)

    def line(self, panel, text):
        label = QLabel(text)
        label.setWordWrap(True)
        label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        label.setFont(QFont("Segoe UI", 8))
        label.setStyleSheet("color:#f2fff8; padding:0px 1px; background:transparent;")
        panel.box.addWidget(label)
        return label

    def location_line(self, text):
        label = QLabel(text)
        label.setWordWrap(True)
        label.setMinimumHeight(36)
        label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        label.setFont(QFont("Segoe UI", 8, QFont.DemiBold))
        label.setStyleSheet("color:#f2fff8; padding:6px 8px; background:rgba(0,25,20,175); border:1px solid rgba(0,232,120,150); border-radius:7px;")
        return label

    def button_style(self, button, color):
        button.setStyleSheet(f"QPushButton {{ background:{color}; color:white; padding:4px 6px; border:1px solid #51e8ff; border-radius:7px; font-size:9px; font-weight:bold; }} QPushButton:hover {{ border:2px solid white; }}")

    def update_clock(self):
        # Left: actual Windows/system time. Right: live preview for the detected IANA zone.
        local_now = datetime.now().astimezone()
        local_zone = current_windows_timezone()
        self.current_clock.setText(
            "CURRENT WINDOWS TIME\n" + local_zone + "\n" +
            local_now.strftime("%A, %d %B %Y   •   %I:%M:%S %p")
        )
        iana = (self.detected or {}).get("timezone") if self.detected else None
        if iana:
            try:
                target_now = datetime.now(ZoneInfo(iana))
                self.target_clock.setText(
                    "DETECTED / NEW TIME ZONE\n" + iana + "\n" +
                    target_now.strftime("%A, %d %B %Y   •   %I:%M:%S %p")
                )
            except (ZoneInfoNotFoundError, ValueError):
                self.target_clock.setText("DETECTED / NEW TIME ZONE\n" + iana + "\nLive preview unavailable")
        else:
            self.target_clock.setText("DETECTED / NEW TIME ZONE\nDetect a location to preview its live time")

    def detect_location(self):
        if not self.db_ready:
            QMessageBox.critical(self, APP_NAME, "Database could not be prepared:\n" + self.db_error)
            return
        self.detect_button.setEnabled(False)
        self.online.setText("●  LOOKING UP…")
        self.status_label.setText("Status: Contacting IP lookup service…")
        self.worker = LookupWorker()
        self.worker.succeeded.connect(self.on_lookup_success)
        self.worker.failed.connect(self.on_lookup_error)
        self.worker.finished.connect(lambda: self.detect_button.setEnabled(True))
        self.worker.start()

    def on_lookup_success(self, data):
        self.detected = data
        self.ip_label.setText("Public IP: " + (data.get("ip") or "Unknown"))
        self.country_label.setText("Country: " + (data.get("country") or "Unknown") + " (" + (data.get("country_code") or "—") + ")")
        city = data.get("city") or "Not supplied"
        region = data.get("region") or ""
        self.city_label.setText("City / region: " + city + (", " + region if region else ""))
        self.network_label.setText("Network / ISP: " + (data.get("org") or "Not supplied"))
        self.provider_label.setText("Lookup provider: " + data.get("provider", "Unknown"))
        iana = data.get("timezone") or ""
        self.detected_zone_label.setText("Detected IANA time zone: " + (iana or "Not supplied"))
        lat, lon = data.get("latitude"), data.get("longitude")
        if lat is not None and lon is not None:
            self.coordinates_label.setText(f"Latitude / Longitude: {float(lat):.5f}, {float(lon):.5f}")
        else:
            self.coordinates_label.setText("Latitude / Longitude: Not supplied by provider")
        accuracy = data.get("accuracy")
        self.accuracy_label.setText(
            f"Accuracy: about {accuracy} m (provider estimate)" if accuracy else
            "Accuracy: Approximate IP location; radius not supplied"
        )
        address_parts = [x for x in [city if city != "Not supplied" else "", region, data.get("country", "")] if x]
        self.address_label.setText("Location: " + (", ".join(address_parts) if address_parts else "Not supplied") + " (IP estimate)")
        self.dst_label.setText(dst_difference_text(iana) if iana else "DST Difference: Not available")
        self.update_clock()
        self.update_regional_suggestions(data)
        self.suggested_windows_id = None
        try:
            record = find_timezone_record(self.db_path, iana) if iana else None
            if record:
                self.db_status_label.setText("Database match: FOUND in Excel")
                windows_id = record.get("windows")
                if not windows_id:
                    windows_id = windows_zone_for(iana)
                self.suggested_windows_id = windows_id or None
            else:
                status, message = learn_timezone(self.db_path, data)
                self.db_status_label.setText("Database match: " + ("AUTO-ADDED" if status == "added" else "ALREADY PRESENT" if status == "exists" else "NOT ADDED"))
                self.status_label.setText("Status: " + message)
                # Re-read after append to ensure the stored record is what we use.
                record = find_timezone_record(self.db_path, iana) if iana else None
                self.suggested_windows_id = (record or {}).get("windows") or windows_zone_for(iana)
        except Exception as exc:
            self.db_status_label.setText("Database match: Excel update/read failed")
            self.status_label.setText("Status: Could not update database: " + str(exc))
            QMessageBox.warning(self, APP_NAME, "IP location was detected, but the Excel database could not be updated.\n\n" + str(exc))
        if self.suggested_windows_id:
            self.suggested_zone_label.setText("Suggested Windows zone: " + self.suggested_windows_id)
            self.apply_button.setEnabled(True)
            if self.status_label.text().startswith("Status: Waiting") or self.status_label.text().startswith("Status: Contacting"):
                self.status_label.setText("Status: Review the detected location, then choose YES or NO.")
        else:
            self.suggested_zone_label.setText("Suggested Windows zone: NOT MAPPED — no automatic change")
            self.apply_button.setEnabled(False)
            self.status_label.setText("Status: IANA zone saved/known, but Windows mapping needs review.")
        self.online.setText("●  ONLINE")

    def update_regional_suggestions(self, data):
        code = (data.get("country_code") or "").strip().upper()
        if not self.db_path or not code:
            self.suggested_geoid = None
            self.suggested_culture = None
            self.suggested_region_label.setText("Suggested Windows Region: NOT MAPPED — review required")
            self.suggested_format_label.setText("Suggested Regional Format: NOT MAPPED — review required")
            self.region_button.setEnabled(False)
            self.format_button.setEnabled(False)
            return
        try:
            wb = load_workbook(self.db_path, read_only=True, data_only=True)
            try:
                ws = wb["Countries"] if "Countries" in wb.sheetnames else None
                supported = set()
                if ws:
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if len(row) > 1 and row[1]:
                            supported.add(str(row[1]).strip().upper())
            finally:
                wb.close()
            if code not in supported:
                raise RuntimeError("Country code is not present in your Excel Countries sheet.")
            geoid = windows_geoid_for_country(code)
            self.suggested_geoid = geoid
            self.suggested_region_label.setText(
                f"Suggested Windows Region: {data.get('country') or code} (ISO {code}) — GeoID {geoid}" if geoid
                else f"Suggested Windows Region: {data.get('country') or code} — NOT MAPPED / AMBIGUOUS; review required")
            self.region_button.setEnabled(bool(geoid))

            entry = self.regional_defaults.get(code) or {}
            culture = entry.get("culture")
            ambiguous = bool(entry.get("ambiguous"))
            valid = bool(culture and not ambiguous and validate_windows_culture(code, culture))
            self.suggested_culture = culture if valid else None
            if valid:
                self.suggested_format_label.setText(f"Suggested Regional Format: {culture} (CLDR default-language suggestion)")
            elif ambiguous:
                self.suggested_format_label.setText(f"Suggested Regional Format: {culture or 'Multiple options'} — AMBIGUOUS; manual review required")
            else:
                self.suggested_format_label.setText("Suggested Regional Format: NOT MAPPED / unavailable on this Windows system; review required")
            self.format_button.setEnabled(valid)
        except Exception as exc:
            self.suggested_geoid = None
            self.suggested_culture = None
            self.suggested_region_label.setText("Suggested Windows Region: REVIEW REQUIRED — " + str(exc))
            self.suggested_format_label.setText("Suggested Regional Format: REVIEW REQUIRED")
            self.region_button.setEnabled(False)
            self.format_button.setEnabled(False)

    def confirm_region_apply(self):
        if self.suggested_geoid is None:
            return
        country = (self.detected or {}).get("country") or (self.detected or {}).get("country_code") or "detected country"
        answer = QMessageBox.question(self, APP_NAME,
            f"Change Windows Home Location / Region to {country}?\n\nGeoID: {self.suggested_geoid}\n\nThis is separate from time zone and regional date/number formats.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            self.status_label.setText("Status: Windows Region change cancelled.")
            return
        try:
            apply_windows_region(self.suggested_geoid)
            self.current_region_label.setText("Current Windows Region: " + current_windows_region())
            self.status_label.setText("Status: Windows Region change command completed.")
            QMessageBox.information(self, APP_NAME, "Windows Region change command completed.")
        except Exception as exc:
            QMessageBox.critical(self, APP_NAME, "Could not change Windows Region.\n\n" + str(exc))

    def confirm_format_apply(self):
        if not self.suggested_culture:
            return
        answer = QMessageBox.question(self, APP_NAME,
            f"Change Regional Format to {self.suggested_culture}?\n\nThis changes date, time and number formatting for your user account. It does not change Windows display language or keyboard.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            self.status_label.setText("Status: Regional Format change cancelled.")
            return
        try:
            apply_regional_format(self.suggested_culture)
            self.current_format_label.setText("Current Regional Format: " + current_regional_format())
            self.status_label.setText("Status: Regional Format command completed. It may take effect in new apps/sessions.")
            QMessageBox.information(self, APP_NAME, "Regional Format command completed. Some apps may need to be reopened.")
        except Exception as exc:
            QMessageBox.critical(self, APP_NAME, "Could not change Regional Format.\n\n" + str(exc))

    def on_lookup_error(self, message):
        self.region_button.setEnabled(False)
        self.format_button.setEnabled(False)
        self.online.setText("●  LOOKUP ERROR")
        self.status_label.setText("Status: Location lookup failed.")
        QMessageBox.warning(self, APP_NAME, "Could not detect public IP location.\nCheck internet/VPN and try again.\n\n" + message)

    def confirm_apply(self):
        if not self.suggested_windows_id:
            return
        answer = QMessageBox.question(
            self, APP_NAME,
            "Apply this Windows time zone?\n\n" + self.suggested_windows_id +
            "\n\nThis changes the Windows time zone setting, not GPS or physical location.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if answer != QMessageBox.Yes:
            self.cancel_apply()
            return
        try:
            apply_windows_timezone(self.suggested_windows_id)
            self.current_zone_label.setText("Current Windows zone: " + current_windows_timezone())
            self.update_clock()
            self.status_label.setText("Status: Windows time zone command completed.")
            QMessageBox.information(self, APP_NAME, "Windows time zone change command completed.")
        except Exception as exc:
            self.status_label.setText("Status: Windows did not accept the time zone change.")
            QMessageBox.critical(self, APP_NAME, "Could not change the Windows time zone.\n\n" + str(exc) +
                                 "\n\nCheck the Windows account permissions and try again.")

    def cancel_apply(self):
        self.status_label.setText("Status: Change cancelled. Windows time zone was not changed.")
        self.apply_button.setEnabled(bool(self.suggested_windows_id))


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

