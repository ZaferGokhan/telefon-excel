# -*- coding: utf-8 -*-
"""USB ile takılan telefonun bilgilerini okuyup Excel'e yazar."""

from __future__ import annotations

import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import traceback
import winsound
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
SETTINGS_PATH = SCRIPT_DIR / "ayar.json"
LOG_PATH = SCRIPT_DIR / "kayit.log"
DEFAULT_XLSX = SCRIPT_DIR / "telefonlar.xlsx"

COLUMNS = [
    ("sira", "Sıra"),
    ("zaman", "Tarih Saat"),
    ("serial", "Seri No"),
    ("model", "Model"),
    ("imei1", "IMEI 1"),
    ("imei2", "IMEI 2"),
    ("csc", "CSC"),
    ("csc_ver", "CSC Sürüm"),
    ("build", "Build Number"),
    ("baseband", "Baseband"),
    ("android", "Android"),
    ("oneui", "One UI"),
    ("durum", "Durum"),
]
HEADERS = [title for _, title in COLUMNS]
TEXT_KEYS = {"serial", "imei1", "imei2", "build", "baseband", "csc_ver"}
TREE_KEYS = (
    "sira",
    "model",
    "imei1",
    "imei2",
    "csc",
    "build",
    "baseband",
    "android",
    "oneui",
    "durum",
)
COLUMN_WIDTHS = {
    "A": 8,
    "B": 20,
    "C": 20,
    "D": 16,
    "E": 22,
    "F": 22,
    "G": 12,
    "H": 46,
    "I": 32,
    "J": 34,
    "K": 12,
    "L": 12,
    "M": 34,
}
MULTI_CSC = {
    "OXM",
    "OYM",
    "OWO",
    "OLE",
    "OLM",
    "OJV",
    "OXA",
    "ODD",
    "ODM",
    "OZS",
    "ROM",
    "SAM",
}
IMEI_PROP_KEYS = (
    "persist.radio.imei1",
    "ril.imei1",
    "ril.IMEI1",
    "persist.radio.imei",
    "ril.gsm.imei",
    "gsm.imei",
    "ro.ril.oem.imei",
    "ro.ril.oem.imei1",
    "ro.boot.imei",
    "persist.radio.imei0",
    "persist.radio.imei2",
    "ril.imei2",
    "ril.IMEI2",
    "ro.ril.oem.imei2",
)
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
CELL_FONT = Font(name="Calibri", size=11)


class AdbError(Exception):
    pass


class ExcelLocked(Exception):
    def __init__(self, backup: Path):
        super().__init__(str(backup))
        self.backup = backup


def find_adb() -> str | None:
    candidates = []
    if os.environ.get("ADB"):
        candidates.append(Path(os.environ["ADB"]))
    which = None
    for folder in os.environ.get("PATH", "").split(os.pathsep):
        exe = Path(folder) / "adb.exe"
        if exe.is_file():
            which = exe
            break
    if which:
        candidates.append(which)
    candidates.append(SCRIPT_DIR / "adb.exe")
    local = os.environ.get("LOCALAPPDATA", "")
    candidates.append(Path(local) / "Android" / "Sdk" / "platform-tools" / "adb.exe")
    candidates.append(Path(r"C:\Users\Zafer\Documents\Samsung Collect 1.0\adb.exe"))
    for path in candidates:
        if path and path.is_file():
            return str(path)
    return None


def luhn_ok(number: str) -> bool:
    if not number.isdigit():
        return False
    total = 0
    double = False
    for char in reversed(number):
        digit = int(char)
        if double:
            digit *= 2
            if digit > 9:
                digit -= 9
        total += digit
        double = not double
    return total % 10 == 0


def add_imei(found: list[str], value: str, trusted: bool = False) -> None:
    if not value:
        return
    chunks = re.split(r"[,;|/]+", value)
    if len(chunks) == 1:
        chunks = [value]
    for chunk in chunks:
        digits = re.sub(r"\D", "", chunk)
        if len(digits) == 16:
            digits = digits[:15]
        if len(digits) != 15:
            continue
        if not trusted and not luhn_ok(digits):
            continue
        if digits not in found:
            found.append(digits)


def imeis_from_props(props: dict[str, str]) -> list[str]:
    found: list[str] = []
    used = set()
    for key in IMEI_PROP_KEYS:
        used.add(key.lower())
        add_imei(found, props.get(key, ""), trusted=True)
    for key, value in props.items():
        if key.lower() in used:
            continue
        if "imei" in key.lower():
            add_imei(found, value, trusted=True)
    return found


def imeis_from_labeled_text(text: str) -> list[str]:
    found: list[str] = []
    for line in text.splitlines():
        if not re.search(r"imei|device\s*id", line, re.I):
            continue
        flat = re.sub(r"[^0-9]", "", line)
        index = 0
        while index + 15 <= len(flat):
            candidate = flat[index : index + 15]
            if luhn_ok(candidate):
                if candidate not in found:
                    found.append(candidate)
                index += 15
            else:
                index += 1
    return found


def imeis_from_parcel(text: str) -> list[str]:
    if "Parcel" not in text:
        return []
    digits = re.sub(r"\D", "", "".join(re.findall(r"'([^']*)'", text)))
    found: list[str] = []
    index = 0
    while index + 15 <= len(digits):
        candidate = digits[index : index + 15]
        if luhn_ok(candidate):
            if candidate not in found:
                found.append(candidate)
            index += 15
        else:
            index += 1
    if not found and len(digits) in (15, 16) and luhn_ok(digits[:15]):
        found.append(digits[:15])
    return found


def imeis_from_ui_xml(xml: str) -> list[str]:
    found: list[str] = []
    texts = re.findall(r'(?:text|content-desc)="([^"]*)"', xml)
    for text in texts:
        flat = re.sub(r"\D", "", text)
        if len(flat) in (15, 16):
            candidate = flat[:15]
            if luhn_ok(candidate) and candidate not in found:
                found.append(candidate)
            continue
        index = 0
        while index + 15 <= len(flat):
            candidate = flat[index : index + 15]
            if luhn_ok(candidate):
                if candidate not in found:
                    found.append(candidate)
                index += 15
            else:
                index += 1
    return found


def format_one_ui(raw: str) -> str:
    """Samsung kodu: 60100 = 6.1, 60101 = 6.1.1, 20500 = 2.5, 70000 = 7.0."""
    raw = (raw or "").strip()
    if not raw or raw.lower() in {"0", "null", "unknown"}:
        return ""
    if not raw.isdigit():
        return raw
    number = int(raw)
    if number < 10000:
        return raw
    major = number // 10000
    minor = (number % 10000) // 100
    patch = number % 100
    if major <= 0 or major > 30:
        return raw
    if patch:
        return f"{major}.{minor}.{patch}"
    return f"{major}.{minor}"


def tidy_baseband(value: str) -> str:
    parts = []
    for piece in value.split(","):
        piece = piece.strip()
        if piece and piece not in parts and piece.lower() not in {"unknown", "none"}:
            parts.append(piece)
    return ", ".join(parts)


def first_prop(props: dict[str, str], keys: tuple[str, ...]) -> str:
    for key in keys:
        value = props.get(key, "").strip()
        if value and value.lower() not in {"unknown", "none", "null"}:
            return value
    return ""


def extract_csc(props: dict[str, str]) -> tuple[str, str]:
    code = ""
    for key in (
        "ro.csc.sales_code",
        "persist.sys.sales_code",
        "ro.boot.sales_code",
        "ro.csc.omcnw_code",
    ):
        value = props.get(key, "").strip().upper()
        if re.fullmatch(r"[A-Z0-9]{3}", value):
            code = value
            break
    if not code:
        for key in ("persist.sys.omcnw_path", "persist.sys.omc_path", "persist.sys.omc_etcpath"):
            match = re.search(r"/([A-Z]{3})(?:/conf)?/?$", props.get(key, "").strip())
            if match:
                code = match.group(1)
                break
    version = first_prop(
        props,
        ("mdc.omc.config_version", "ril.official_cscver", "ro.omc.build.version"),
    )
    if not code and version:
        tokens = [part for part in version.split("_") if re.fullmatch(r"[A-Z]{3}", part)]
        country = [part for part in tokens if part not in MULTI_CSC]
        if country:
            code = country[-1]
        elif tokens:
            code = tokens[-1]
    return code, version


def extract_oneui(props: dict[str, str]) -> str:
    for key in (
        "ro.build.version.oneui",
        "ro.build.version.oneui_level",
        "ro.product.oneui.version",
    ):
        if props.get(key, "").strip():
            return format_one_ui(props[key])
    for key, value in props.items():
        if "oneui" in key.lower() and value.strip():
            return format_one_ui(value)
    return ""


def parse_props(text: str) -> dict[str, str]:
    props = {}
    for line in text.splitlines():
        if line.startswith("[") and "]: [" in line and line.endswith("]"):
            key, value = line[1:].split("]: [", 1)
            props[key] = value[:-1]
    return props


def excel_safe(value: str) -> str:
    value = "" if value is None else str(value).strip()
    if value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def append_log(line: str) -> None:
    try:
        with LOG_PATH.open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except OSError:
        pass


class Adb:
    def __init__(self, path: str):
        self.path = path

    def run(self, args: list[str], timeout: int = 20) -> str:
        try:
            completed = subprocess.run(
                [self.path, *args],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                creationflags=NO_WINDOW,
            )
        except subprocess.TimeoutExpired as exc:
            raise AdbError("ADB zaman aşımına uğradı") from exc
        text = ((completed.stdout or "") + (completed.stderr or ""))
        text = text.replace("\r\n", "\n").replace("\r", "")
        lowered = text.lower()
        if "not found" in lowered and "device" in lowered:
            raise AdbError(text.strip() or "telefon bulunamadı")
        if "unauthorized" in lowered:
            raise AdbError("unauthorized")
        if "error: closed" in lowered or "device offline" in lowered:
            raise AdbError(text.strip() or "telefon koptu")
        return text

    def start(self) -> None:
        self.run(["start-server"], timeout=25)

    def devices(self) -> list[tuple[str, str]]:
        text = self.run(["devices"], timeout=15)
        found = []
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith("List of"):
                continue
            parts = line.split()
            if len(parts) >= 2:
                found.append((parts[0], parts[1]))
        return found

    def shell(self, serial: str, command: str, timeout: int = 20) -> str:
        return self.run(["-s", serial, "shell", command], timeout=timeout)


class ExcelStore:
    def __init__(self, path: Path):
        self.path = path
        self.rows: list[dict[str, str]] = []
        self.dirty = False
        self.lock = threading.Lock()

    def load(self) -> None:
        self.rows = []
        if not self.path.exists():
            return
        workbook = load_workbook(self.path, read_only=True, data_only=True)
        try:
            sheet = workbook.active
            iterator = sheet.iter_rows(values_only=True)
            header = next(iterator, None)
            if not header:
                return
            titles = ["" if cell is None else str(cell).strip() for cell in header]
            if titles[: len(HEADERS)] != HEADERS:
                raise ValueError("Excel başlıkları bu programın dosyası değil")
            keys = [key for key, _ in COLUMNS]
            for raw in iterator:
                if raw is None or all(cell is None or str(cell).strip() == "" for cell in raw):
                    continue
                item = {}
                for index, key in enumerate(keys):
                    value = raw[index] if index < len(raw) else ""
                    item[key] = "" if value is None else str(value).strip()
                self.rows.append(item)
        finally:
            workbook.close()

    def snapshot(self) -> list[dict[str, str]]:
        with self.lock:
            return [dict(row) for row in self.rows]

    def upsert(self, record: dict[str, str]) -> tuple[str, Path | None, str]:
        with self.lock:
            action = self._apply(record)
            self.dirty = True
            try:
                saved = self._save_unlocked()
                self.dirty = False
                return action, saved, ""
            except ExcelLocked as exc:
                return action, exc.backup, (
                    "Excel dosyası açık. Kapatın; kayıt yedeğe yazıldı ve dosya boşalınca "
                    f"asıl dosyaya işlenecek. Yedek: {exc.backup}"
                )

    def flush(self) -> str:
        with self.lock:
            if not self.dirty and self.path.exists():
                return ""
            if not self.rows and not self.dirty:
                return ""
            try:
                self._save_unlocked()
                self.dirty = False
                return ""
            except ExcelLocked as exc:
                return str(exc.backup)

    def _apply(self, record: dict[str, str]) -> str:
        clean = {key: str(record.get(key, "") or "").strip() for key, _ in COLUMNS}
        index = self._find(clean)
        if index is None:
            numbers = []
            for row in self.rows:
                try:
                    numbers.append(int(float(row.get("sira", "") or 0)))
                except ValueError:
                    continue
            clean["sira"] = str((max(numbers) if numbers else 0) + 1)
            if not clean["zaman"]:
                clean["zaman"] = time.strftime("%Y-%m-%d %H:%M:%S")
            self.rows.append(clean)
            return "new"
        previous = self.rows[index]
        clean["sira"] = previous.get("sira") or clean["sira"]
        if not clean["zaman"]:
            clean["zaman"] = time.strftime("%Y-%m-%d %H:%M:%S")
        self.rows[index] = clean
        return "update"

    def _find(self, record: dict[str, str]) -> int | None:
        serial = record.get("serial", "")
        imei = record.get("imei1", "")
        if serial:
            for index, row in enumerate(self.rows):
                if row.get("serial") == serial:
                    return index
        if imei:
            for index, row in enumerate(self.rows):
                if row.get("imei1") == imei:
                    return index
        return None

    def _save_unlocked(self) -> Path:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(".tmp.xlsx")
        self._write(temporary)
        try:
            temporary.replace(self.path)
            return self.path
        except OSError:
            backup = self.path.with_name(self.path.stem + "_yedek.xlsx")
            try:
                if backup.exists():
                    backup.unlink()
                temporary.replace(backup)
            except OSError:
                backup = temporary
            raise ExcelLocked(backup)

    def _write(self, path: Path) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Telefonlar"
        sheet.sheet_view.zoomScale = 120
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{max(1, len(self.rows) + 1)}"
        sheet.row_dimensions[1].height = 22
        for index, title in enumerate(HEADERS, start=1):
            cell = sheet.cell(1, index, title)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row_index, row in enumerate(self.rows, start=2):
            for col_index, (key, _) in enumerate(COLUMNS, start=1):
                cell = sheet.cell(row_index, col_index, excel_safe(row.get(key, "")))
                cell.font = CELL_FONT
                cell.alignment = Alignment(vertical="center")
                if key in TEXT_KEYS:
                    cell.number_format = "@"
            sheet.row_dimensions[row_index].height = 18
        for letter, width in COLUMN_WIDTHS.items():
            sheet.column_dimensions[letter].width = width
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        workbook.properties.title = "Telefon kayıtları"
        workbook.save(path)


def merge_imeis(*groups: list[str]) -> list[str]:
    found: list[str] = []
    for group in groups:
        for imei in group:
            if imei not in found:
                found.append(imei)
    return found[:4]


def section_text(text: str, name: str) -> str:
    # Marker'da # yok. Cihaz kabuğu # işaretini yorum satırı sanıyor.
    match = re.search(rf"__{name}__\n(.*?)(?:\n__[A-Z0-9]+__|\Z)", text, re.S)
    return match.group(1) if match else ""


def read_phone(adb: Adb, serial: str, allow_screen: bool, progress) -> dict[str, str]:
    progress(f"Özellikler okunuyor ({serial})")
    props = parse_props(adb.shell(serial, "getprop", timeout=20))
    model = first_prop(
        props,
        ("ro.product.model", "ro.product.odm.model", "ro.product.system.model"),
    )
    build = first_prop(props, ("ro.build.display.id", "ro.build.PDA", "ro.build.id"))
    baseband = tidy_baseband(
        first_prop(
            props,
            ("gsm.version.baseband", "vendor.gsm.version.baseband", "ril.sw_ver"),
        )
    )
    android = first_prop(
        props,
        ("ro.build.version.release", "ro.build.version.release_or_codename"),
    )
    oneui = extract_oneui(props)
    csc, csc_ver = extract_csc(props)
    sheet_serial = first_prop(props, ("ro.serialno", "ro.boot.serialno")) or serial
    imeis = imeis_from_props(props)
    notes = []
    if imeis:
        notes.append("IMEI sistem kaydından geldi")

    if len(imeis) < 2:
        progress("IMEI servisi sorgulanıyor")
        try:
            # Yeni Android: slot 0 / slot 1. Paket adı olmadan çoğu telefonda reddediliyor.
            probed = adb.shell(
                serial,
                "echo __A__; service call iphonesubinfo 4 i32 0 s16 com.android.shell 2>&1; "
                "echo __B__; service call iphonesubinfo 4 i32 1 s16 com.android.shell 2>&1; "
                "echo __C__; service call iphonesubinfo 1 s16 com.android.shell 2>&1; "
                "echo __END__",
                timeout=15,
            )
        except AdbError as exc:
            notes.append(f"IMEI servisi: {exc}")
            probed = ""
        extra = merge_imeis(
            imeis_from_parcel(section_text(probed, "A")),
            imeis_from_parcel(section_text(probed, "B")),
            imeis_from_parcel(section_text(probed, "C")),
        )
        before = len(imeis)
        imeis = merge_imeis(imeis, extra)
        if len(imeis) > before:
            notes.append("IMEI servis sorgusundan geldi")

    if len(imeis) < 2:
        progress("İkinci IMEI aranıyor")
        try:
            legacy = adb.shell(
                serial,
                "echo __D__; dumpsys iphonesubinfo 2>&1; "
                "echo __E__; service call iphonesubinfo 1 2>&1; "
                "echo __F__; service call iphonesubinfo 1 i32 0 2>&1; "
                "echo __G__; service call iphonesubinfo 1 i32 1 2>&1; "
                "echo __H__; service call iphonesubinfo 3 i32 0 2>&1; "
                "echo __I__; service call iphonesubinfo 3 i32 1 2>&1; "
                "echo __END__",
                timeout=15,
            )
        except AdbError as exc:
            notes.append(f"Eski IMEI yöntemi: {exc}")
            legacy = ""
        extra = merge_imeis(
            imeis_from_labeled_text(section_text(legacy, "D")),
            imeis_from_parcel(section_text(legacy, "E")),
            imeis_from_parcel(section_text(legacy, "F")),
            imeis_from_parcel(section_text(legacy, "G")),
            imeis_from_parcel(section_text(legacy, "H")),
            imeis_from_parcel(section_text(legacy, "I")),
        )
        before = len(imeis)
        imeis = merge_imeis(imeis, extra)
        if len(imeis) > before:
            notes.append("IMEI eski servis sorgusundan geldi")

    locked = False
    if not imeis and allow_screen:
        progress("Ekran kilidi kontrol ediliyor")
        try:
            window = adb.shell(
                serial,
                "dumpsys window | grep -E 'mShowingLockscreen|isStatusBarKeyguard|mDreamingLockscreen'",
                timeout=12,
            )
        except AdbError:
            window = ""
        locked = any(
            flag in window
            for flag in (
                "mShowingLockscreen=true",
                "isStatusBarKeyguard=true",
                "mDreamingLockscreen=true",
            )
        )
        if locked:
            notes.append("ekran kilitli, *#06# açılamadı")
        else:
            progress("IMEI ekrandan okunuyor, telefona dokunmayın")
            screen_imeis = imeis_from_screen(adb, serial)
            if screen_imeis:
                imeis = merge_imeis(imeis, screen_imeis)
                notes.append("IMEI *#06# ekranından geldi")
            else:
                notes.append("*#06# ekranında IMEI bulunamadı")

    if not imeis:
        durum = "IMEI okunamadı, ekran kilitli" if locked else "IMEI okunamadı"
    elif not csc:
        durum = "Tamam, CSC boş"
    else:
        durum = "Tamam"

    record = {
        "sira": "",
        "zaman": time.strftime("%Y-%m-%d %H:%M:%S"),
        "serial": sheet_serial,
        "model": model,
        "imei1": imeis[0] if imeis else "",
        "imei2": imeis[1] if len(imeis) > 1 else "",
        "csc": csc,
        "csc_ver": csc_ver,
        "build": build,
        "baseband": baseband,
        "android": android,
        "oneui": oneui,
        "durum": durum,
        "_log": " | ".join(notes) if notes else "okundu",
    }
    return record


def imeis_from_screen(adb: Adb, serial: str) -> list[str]:
    try:
        adb.shell(serial, "input keyevent KEYCODE_WAKEUP", timeout=8)
    except AdbError:
        pass
    commands = (
        "am start -a android.provider.Telephony.SECRET_CODE -d android_secret_code://06",
        "am start -a android.intent.action.DIAL -d 'tel:*%2306%23'",
    )
    for command in commands:
        try:
            adb.shell(serial, command, timeout=10)
        except AdbError:
            continue
        time.sleep(1.4)
        xml = dump_ui(adb, serial)
        imeis = imeis_from_ui_xml(xml)
        if not imeis:
            time.sleep(1.0)
            imeis = imeis_from_ui_xml(dump_ui(adb, serial))
        if imeis:
            close_screen(adb, serial)
            return imeis
    close_screen(adb, serial)
    return []


def dump_ui(adb: Adb, serial: str) -> str:
    for target in ("/data/local/tmp/imei_ui.xml", "/sdcard/imei_ui.xml"):
        try:
            output = adb.shell(serial, f"uiautomator dump {target} && cat {target}", timeout=15)
        except AdbError:
            continue
        if "<hierarchy" in output or "<node" in output:
            try:
                adb.shell(serial, f"rm -f {target}", timeout=6)
            except AdbError:
                pass
            return output
    try:
        return adb.shell(serial, "uiautomator dump /dev/stdout", timeout=15)
    except AdbError as exc:
        return str(exc)


def close_screen(adb: Adb, serial: str) -> None:
    try:
        adb.shell(serial, "input keyevent KEYCODE_HOME", timeout=6)
    except AdbError:
        pass


class Watcher(threading.Thread):
    def __init__(self, adb: Adb, store: ExcelStore, events: queue.Queue, allow_screen):
        super().__init__(daemon=True)
        self.adb = adb
        self.store = store
        self.events = events
        self.allow_screen = allow_screen
        self.stop = threading.Event()
        self.paused = False
        self.reread = False
        self.settled: set[str] = set()
        self.misses: dict[str, int] = {}
        self.fails: dict[str, int] = {}
        self.seen_unauth: set[str] = set()
        self.last_status: tuple[str, str] | None = None

    def run(self) -> None:
        try:
            self.adb.start()
        except AdbError as exc:
            self.status("error", f"ADB başlamadı: {exc}")
            return
        self.status("wait", "Telefon bekleniyor")
        while not self.stop.is_set():
            if self.store.dirty:
                error = self.store.flush()
                if not error:
                    self.status("ok", "Excel kapanmıştı, bekleyen kayıt dosyaya yazıldı")
            try:
                devices = self.adb.devices()
            except AdbError as exc:
                self.status("error", f"ADB hatası: {exc}")
                self.stop.wait(1.5)
                continue
            self.events.put(("devices", devices))
            current = {serial for serial, _state in devices}
            for serial in list(self.settled):
                if serial in current:
                    self.misses.pop(serial, None)
                    continue
                self.misses[serial] = self.misses.get(serial, 0) + 1
                if self.misses[serial] >= 2:
                    self.settled.discard(serial)
                    self.misses.pop(serial, None)
                    self.fails.pop(serial, None)
            if self.paused:
                self.status("wait", "Duraklatıldı. Devam deyince sıradaki telefon okunur.")
                self.stop.wait(0.6)
                continue
            ready = [serial for serial, state in devices if state == "device"]
            unauthorized = [serial for serial, state in devices if state == "unauthorized"]
            if self.reread:
                self.reread = False
                for serial in ready:
                    self.settled.discard(serial)
                    self.fails.pop(serial, None)
            if unauthorized and not any(serial not in self.settled for serial in ready):
                self.status(
                    "warn",
                    "Telefonda “Bu bilgisayara izin ver” deyin. Kutudaki onay işaretini de seçin.",
                )
            target = next((serial for serial in ready if serial not in self.settled), None)
            fresh_unauth = [serial for serial in unauthorized if serial not in self.seen_unauth]
            self.seen_unauth = set(unauthorized)
            if fresh_unauth and not target:
                self.events.put(("reading", ""))
            if target:
                self.events.put(("reading", target))
                self.read_one(target)
            elif not devices:
                self.status("wait", "Telefon bekleniyor. Takın, yeşili görünce çıkarın.")
            self.stop.wait(0.7)

    def read_one(self, serial: str) -> None:
        try:
            record = read_phone(
                self.adb,
                serial,
                bool(self.allow_screen[0]),
                lambda text: self.status("busy", text),
            )
        except Exception as exc:
            self.fails[serial] = self.fails.get(serial, 0) + 1
            append_log(f"{time.strftime('%H:%M:%S')} HATA {serial}: {exc}")
            self.events.put(("log", f"HATA {serial}: {exc}"))
            if self.fails[serial] >= 2:
                self.settled.add(serial)
                self.status("error", f"Bu telefon okunamadı ({serial}). Çıkarıp tekrar takın.")
            else:
                self.status("warn", f"Okuma yarıda kaldı, bir kez daha denenecek ({serial})")
            return
        action, _path, error = self.store.upsert(record)
        self.settled.add(serial)
        self.fails.pop(serial, None)
        self.events.put(("saved", {"action": action, "record": record, "error": error}))
        stamp = time.strftime("%H:%M:%S")
        summary = (
            f"{stamp} {record.get('model') or serial}  "
            f"IMEI {record.get('imei1') or '-'}  CSC {record.get('csc') or '-'}  "
            f"{record.get('durum')}  ({record.get('_log')})"
        )
        append_log(summary)
        self.events.put(("log", summary))
        if error:
            self.status("warn", error)
        elif record.get("durum") == "Tamam":
            verb = "güncellendi" if action == "update" else "kaydedildi"
            self.status(
                "ok",
                f"{record.get('model') or 'Telefon'} {verb}. Çıkarıp yenisini takın.",
            )
        else:
            self.status(
                "warn",
                f"{record.get('model') or 'Telefon'} yazıldı ama {record.get('durum')}. "
                "Kilidi açıp Yeniden oku'ya basın veya çıkarıp tekrar takın.",
            )

    def status(self, kind: str, text: str) -> None:
        if self.last_status == (kind, text):
            return
        self.last_status = (kind, text)
        self.events.put(("status", (kind, text)))


HEAD_DOTS = 96
DEFAULT_LABEL = (14, 40, 5)
FONT_CANDIDATES = (
    r"C:\Windows\Fonts\arialbd.ttf",
    r"C:\Windows\Fonts\arial.ttf",
    r"C:\Windows\Fonts\segoeuib.ttf",
)


def find_nelko_port(preferred: str = "") -> str:
    from serial.tools import list_ports

    spp_uuid = "00001101-0000-1000-8000-00805F9B34FB"
    ports = list(list_ports.comports())

    def outgoing(port) -> bool:
        hwid = (port.hwid or "").upper()
        if spp_uuid not in hwid and "BTHENUM" not in hwid:
            return False
        # Gelen kutu (MAC sıfır) yazıcıya bağlanmaz, hemen düşer.
        return "000000000000" not in hwid

    outgoing_ports = [port for port in ports if outgoing(port)]
    if preferred:
        for port in outgoing_ports:
            if port.device.upper() == preferred.upper():
                return port.device
    named = [
        port.device
        for port in outgoing_ports
        if "nelko" in f"{port.description} {port.hwid}".lower() or "p21" in f"{port.description} {port.hwid}".lower()
    ]
    if named:
        return named[0]
    if outgoing_ports:
        return outgoing_ports[0].device
    return ""


def parse_label_roll(block: bytes) -> tuple[int, int, int] | None:
    if len(block) != 16:
        return None
    tag_format = block[3] >> 4
    width = block[13]
    length = block[11]
    gap = block[8] if tag_format >= 2 and block[8] else 5
    if not (8 <= width <= 20 and 15 <= length <= 80):
        return None
    return int(width), int(length), int(gap)


def label_dots(width_mm: int, length_mm: int, dpi: int = 203) -> tuple[int, int]:
    width = min(HEAD_DOTS, int(round(width_mm / 25.4 * dpi)) // 8 * 8)
    height = max(80, int(round(length_mm / 25.4 * dpi)))
    return width, height


def _font(size: int):
    from PIL import ImageFont

    for path in FONT_CANDIDATES:
        if Path(path).is_file():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def _text_size(draw, text: str, font) -> tuple[int, int]:
    box = draw.textbbox((0, 0), text, font=font)
    return box[2] - box[0], box[3] - box[1]


def _fit_font(draw, text: str, max_width: int, max_height: int):
    size = max(8, max_height)
    while size > 8:
        font = _font(size)
        width, height = _text_size(draw, text, font)
        if width <= max_width and height <= max_height:
            return font
        size -= 1
    return _font(8)


CODE128 = (
    "11011001100", "11001101100", "11001100110", "10010011000", "10010001100",
    "10001001100", "10011001000", "10011000100", "10001100100", "11001001000",
    "11001000100", "11000100100", "10110011100", "10011011100", "10011001110",
    "10111001100", "10011101100", "10011100110", "11001110010", "11001011100",
    "11001001110", "11011100100", "11001110100", "11101101110", "11101001100",
    "11100101100", "11100100110", "11101100100", "11100110100", "11100110010",
    "11011011000", "11011000110", "11000110110", "10100011000", "10001011000",
    "10001000110", "10110001000", "10001101000", "10001100010", "11010001000",
    "11000101000", "11000100010", "10110111000", "10110001110", "10001101110",
    "10111011000", "10111000110", "10001110110", "11101110110", "11010001110",
    "11000101110", "11011101000", "11011100010", "11011101110", "11101011000",
    "11101000110", "11100010110", "11101101000", "11101100010", "11100011010",
    "11101111010", "11001000010", "11110001010", "10100110000", "10100001100",
    "10010110000", "10010000110", "10000101100", "10000100110", "10110010000",
    "10110000100", "10011010000", "10011000010", "10000110100", "10000110010",
    "11000010010", "11001010000", "11110111010", "11000010100", "10001111010",
    "10100111100", "10010111100", "10010011110", "10111100100", "10011110100",
    "10011110010", "11110100100", "11110010100", "11110010010", "11011011110",
    "11011110110", "11110110110", "10101111000", "10100011110", "10001011110",
    "10111101000", "10111100010", "11110101000", "11110100010", "10111011110",
    "10111101110", "11101011110", "11110101110", "11010000100", "11010010000",
    "11010011100",
)
CODE128_STOP = "1100011101011"


def _code128_bits(values: list[int]) -> str:
    checksum = values[0]
    for position, value in enumerate(values[1:], start=1):
        checksum += position * value
    values.append(checksum % 103)
    return "".join(CODE128[value] for value in values) + CODE128_STOP


def code128_modules(text: str) -> str:
    """Rakamsa sıkışık Code C, diğer yazılarda Code B."""
    text = str(text).strip()
    if text.isdigit():
        values = [105]
        index = 0
        while index + 1 < len(text):
            values.append(int(text[index : index + 2]))
            index += 2
        if index < len(text):
            values.append(100)
            values.append(ord(text[index]) - 32)
        return _code128_bits(values)
    values = [104]
    for char in text:
        code = ord(char)
        if code < 32 or code > 126:
            code = ord("?")
        values.append(code - 32)
    return _code128_bits(values)


def render_barcode_bitmap(text: str, width: int, height: int, caption: str | None = None) -> bytes:
    """Uzun kenarda Code 128, yanında yazı. 0 biti siyah basar."""
    from PIL import Image, ImageDraw

    shown = (caption if caption is not None else str(text)).strip()
    page = Image.new("1", (width, height), 1)
    expect = (width // 8) * height
    if not str(text).strip():
        return page.tobytes().ljust(expect, b"\xff")[:expect]
    bits = code128_modules(str(text).strip())
    text_band = 30 if width >= 90 else max(18, width // 3)
    bar_height = max(20, width - text_band - 8)
    quiet_modules = 8
    usable = max(1, height - 6)
    module = max(1, usable // (len(bits) + quiet_modules * 2))
    while module > 1 and module * (len(bits) + quiet_modules * 2) > height - 4:
        module -= 1
    barcode_span = module * (len(bits) + quiet_modules * 2)
    y = max(0, (height - barcode_span) // 2) + module * quiet_modules
    x = max(1, (width - text_band - bar_height) // 2)
    draw = ImageDraw.Draw(page)
    for bit in bits:
        if bit == "1":
            draw.rectangle((x, y, x + bar_height - 1, y + module - 1), fill=0)
        y += module
    text_img = Image.new("1", (height - 4, text_band), 1)
    text_draw = ImageDraw.Draw(text_img)
    font = _fit_font(text_draw, shown, text_img.width - 4, text_band - 2)
    text_width, text_height = _text_size(text_draw, shown, font)
    text_draw.text(
        ((text_img.width - text_width) // 2, max(0, (text_img.height - text_height) // 2)),
        shown,
        fill=0,
        font=font,
    )
    turned = text_img.rotate(270, expand=True, fillcolor=1).convert("1", dither=Image.Dither.NONE)
    page.paste(turned, (width - turned.width, max(0, (height - turned.height) // 2)))
    expect = (width // 8) * height
    data = page.convert("1", dither=Image.Dither.NONE).tobytes()
    if len(data) < expect:
        data = data.ljust(expect, b"\xff")
    return data[:expect]


def render_imei_bitmap(imeis: list[str], width: int, height: int) -> bytes:
    imei = re.sub(r"\D", "", imeis[0] if imeis else "")[:15]
    if len(imei) == 15:
        caption = f"{imei[:6]} {imei[6:12]} {imei[12:]}"
    else:
        caption = imei
    return render_barcode_bitmap(imei, width, height, caption)


def label_job(bitmap: bytes, width: int, height: int, width_mm: int, length_mm: int, gap_mm: int) -> bytes:
    return (
        f"\r\nSIZE {width_mm}.0 mm,{length_mm}.0 mm\r\n"
        f"GAP {gap_mm}.0 mm,0 mm\r\n"
        "DIRECTION 1,1\r\n"
        "DENSITY 15\r\n"
        "CLS\r\n"
        f"BITMAP 0,0,{width // 8},{height},1,"
    ).encode() + bitmap + b"\r\nPRINT 1\r\n"


def _write_chunks(port, payload: bytes) -> None:
    for offset in range(0, len(payload), 512):
        port.write(payload[offset : offset + 512])
        port.flush()
        time.sleep(0.02)


def print_imei_label(port_name: str, imeis: list[str]) -> str:
    import serial

    lines = [line for line in imeis if line]
    if not lines:
        raise RuntimeError("Basılacak IMEI yok")
    try:
        port = serial.Serial(port_name, 115200, timeout=2, write_timeout=8)
    except serial.SerialException as exc:
        raise RuntimeError(f"Yazıcı portu açılmadı ({port_name}): {exc}") from exc
    try:
        time.sleep(0.4)
        port.reset_input_buffer()
        port.write(b"\x1b!o\r\n")
        port.flush()
        roll = parse_label_roll(port.read(16)) or DEFAULT_LABEL
        width_mm, length_mm, gap_mm = roll
        width, height = label_dots(width_mm, length_mm)
        bitmap = render_imei_bitmap(lines, width, height)
        _write_chunks(port, label_job(bitmap, width, height, width_mm, length_mm, gap_mm))
    finally:
        port.close()
    return f"{width_mm}x{length_mm} mm etikete {len(lines)} IMEI basıldı ({port_name})"


def barcode_too_long(text: str, height: int) -> bool:
    return len(code128_modules(text)) + 8 > max(1, height - 4)


def wait_for_label(port, length_mm: int, gap_mm: int) -> None:
    """Sonraki işi, şerit fiziksel olarak çıkmadan gönderme. Erken CLS baskıyı yarıda keser."""
    seconds = max(2.8, (length_mm + gap_mm) / 18.0)
    time.sleep(seconds)
    try:
        port.reset_input_buffer()
        port.write(b"\x1b!?\r\n")
        port.flush()
        reply = port.read(1)
    except Exception:
        return
    if reply and reply[0] in (0x02, 0x10, 0x20):
        time.sleep(1.4)


def print_text_labels(port_name: str, texts: list[str], progress=None, stop=None) -> str:
    import serial

    cleaned = [str(item).strip() for item in texts if str(item).strip()]
    if not cleaned:
        raise RuntimeError("Basılacak satır yok")
    try:
        port = serial.Serial(port_name, 115200, timeout=2, write_timeout=8)
    except serial.SerialException as exc:
        raise RuntimeError(f"Yazıcı portu açılmadı ({port_name}): {exc}") from exc
    try:
        time.sleep(0.4)
        port.reset_input_buffer()
        port.write(b"\x1b!o\r\n")
        port.flush()
        width_mm, length_mm, gap_mm = parse_label_roll(port.read(16)) or DEFAULT_LABEL
        width, height = label_dots(width_mm, length_mm)
        done = 0
        skipped = 0
        for index, text in enumerate(cleaned, start=1):
            if stop is not None and stop.is_set():
                break
            if barcode_too_long(text, height):
                skipped += 1
                if progress:
                    progress(index, len(cleaned), f"çok uzun, atlandı: {text}")
                continue
            if progress:
                progress(index, len(cleaned), text)
            bitmap = render_barcode_bitmap(text, width, height)
            _write_chunks(port, label_job(bitmap, width, height, width_mm, length_mm, gap_mm))
            done += 1
            if index < len(cleaned) and not (stop is not None and stop.is_set()):
                if progress:
                    progress(index, len(cleaned), f"{text}  — sonraki etiket bekleniyor")
                wait_for_label(port, length_mm, gap_mm)
    finally:
        port.close()
    note = f"{done} etiket basıldı"
    if skipped:
        note += f", {skipped} satır etikete sığmadı"
    return note


def load_settings() -> dict:
    if not SETTINGS_PATH.exists():
        return {}
    try:
        return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(data: dict) -> None:
    try:
        SETTINGS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def beep(ok: bool) -> None:
    try:
        if ok:
            winsound.Beep(880, 120)
            winsound.Beep(1175, 140)
        else:
            winsound.Beep(280, 320)
    except RuntimeError:
        pass


def run_app() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    settings = load_settings()
    excel_path = Path(settings.get("excel") or DEFAULT_XLSX)
    adb_path = find_adb()

    root = tk.Tk()
    root.title("Telefon Excel")
    root.configure(bg="#eef2f6")
    root.attributes("-topmost", True)
    width, height = 440, 700
    primary_w = ctypes.windll.user32.GetSystemMetrics(0)
    x = max(0, primary_w - width - 16)
    y = 12
    root.geometry(f"{width}x{height}+{x}+{y}")
    root.minsize(400, 560)

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(".", font=("Segoe UI", 10))
    style.configure("Treeview", rowheight=26, font=("Segoe UI", 10))
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    style.configure("TButton", padding=(10, 6))

    events: queue.Queue = queue.Queue()
    store = ExcelStore(excel_path)
    load_error = ""
    try:
        store.load()
    except PermissionError:
        load_error = "Excel dosyası açık. Kapatıp programı yeniden başlatın veya başka dosya seçin."
    except Exception as exc:
        load_error = f"Seçili Excel okunamadı, yeni dosyaya geçilecek: {exc}"
        excel_path = SCRIPT_DIR / f"telefonlar_{time.strftime('%H%M%S')}.xlsx"
        store = ExcelStore(excel_path)

    allow_screen = [bool(settings.get("ekran_imei", True))]
    printer_port = [str(settings.get("yazici") or "")]
    watcher = None
    session_count = {"n": 0}
    refreshing = {"on": False}
    shown_serial = {"value": ""}

    banner = tk.Frame(root, bg="#1F4E79")
    banner.pack(fill="x")
    status_var = tk.StringVar(value="Başlatılıyor")
    status_label = tk.Label(
        banner,
        textvariable=status_var,
        bg="#1F4E79",
        fg="white",
        font=("Segoe UI", 16, "bold"),
        anchor="w",
        justify="left",
        padx=14,
        pady=12,
        wraplength=400,
    )
    status_label.pack(fill="x")

    body = ttk.Frame(root, padding=(10, 8))
    body.pack(fill="both", expand=True)

    file_row = ttk.Frame(body)
    file_row.pack(fill="x")
    path_var = tk.StringVar(value=store.path.name)

    def persist_settings() -> None:
        save_settings(
            {
                "excel": str(store.path),
                "ekran_imei": bool(allow_screen[0]),
                "yazici": printer_port[0],
            }
        )

    def refresh_tree(select_serial: str = "") -> None:
        rows = store.snapshot()
        refreshing["on"] = True
        chosen = None
        try:
            tree.delete(*tree.get_children())
            for index, row in enumerate(rows):
                tag = "ok" if str(row.get("durum", "")).startswith("Tamam") else "warn"
                iid = str(index)
                tree.insert("", "end", iid=iid, values=[row.get(key, "") for key in ("sira", "model", "imei1", "csc")], tags=(tag,))
                if select_serial and row.get("serial") == select_serial:
                    chosen = iid
            if chosen:
                tree.selection_set(chosen)
                tree.see(chosen)
        finally:
            refreshing["on"] = False
        count_var.set(f"Dosyada {len(rows)} telefon    Bu oturumda {session_count['n']} okuma")

    def show_record(record: dict) -> None:
        shown_serial["value"] = str(record.get("serial") or "")
        for key, var in detail_vars.items():
            value = record.get(key, "") or "—"
            var.set(value)
        detail_value_labels["imei1"].configure(
            foreground="#9b1c1c" if record.get("imei1", "") == "" else "#16325c"
        )
        detail_value_labels["csc"].configure(
            foreground="#9b1c1c" if record.get("csc", "") == "" else "#16325c"
        )

    def set_status(kind: str, text: str) -> None:
        colors = {
            "wait": ("#1F4E79", "white"),
            "busy": ("#8A4B08", "white"),
            "ok": ("#0E7A32", "white"),
            "warn": ("#8A5A00", "white"),
            "error": ("#9B1C1C", "white"),
        }
        bg, fg = colors.get(kind, ("#1F4E79", "white"))
        banner.configure(bg=bg)
        status_label.configure(bg=bg, fg=fg)
        status_var.set(text)

    def choose_file() -> None:
        selected = filedialog.asksaveasfilename(
            title="Excel dosyası",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=store.path.name,
            initialdir=str(store.path.parent),
        )
        if not selected:
            return
        path = Path(selected)
        if path.suffix.lower() != ".xlsx":
            path = path.with_suffix(".xlsx")
        error = store.flush()
        if error:
            messagebox.showwarning("Telefon Excel", "Eski dosya hâlâ açık. Önce Excel'i kapatın.")
            return
        try:
            new_store_rows = ExcelStore(path)
            new_store_rows.load()
        except Exception as exc:
            messagebox.showerror("Telefon Excel", str(exc))
            return
        store.path = path
        store.rows = new_store_rows.rows
        store.dirty = False
        path_var.set(path.name)
        persist_settings()
        refresh_tree()

    def open_excel() -> None:
        if not store.path.exists():
            messagebox.showinfo("Telefon Excel", "Henüz satır yok. İlk telefonu takınca dosya oluşur.")
            return
        set_status(
            "warn",
            "Excel açıldı. Dosya açıkken yeni telefon yedeğe yazılır. Bitince Excel'i kapatın.",
        )
        os.startfile(store.path)

    def open_folder() -> None:
        os.startfile(store.path.parent)

    ttk.Label(file_row, textvariable=path_var).pack(side="left")
    ttk.Button(file_row, text="Excel", command=open_excel).pack(side="right")
    ttk.Button(file_row, text="Gözat", command=choose_file).pack(side="right", padx=(0, 4))

    details = ttk.Frame(body)
    details.pack(fill="x", pady=(8, 4))
    detail_vars = {}
    detail_value_labels = {}
    detail_fields = (
        ("model", "Model"),
        ("imei1", "IMEI 1"),
        ("imei2", "IMEI 2"),
        ("csc", "CSC"),
        ("build", "Build"),
        ("baseband", "Baseband"),
        ("android", "Android"),
        ("oneui", "One UI"),
    )
    for row_index, (key, title) in enumerate(detail_fields):
        ttk.Label(details, text=title, width=10).grid(row=row_index, column=0, sticky="w", pady=1)
        var = tk.StringVar(value="—")
        detail_vars[key] = var
        label = ttk.Label(details, textvariable=var, font=("Segoe UI", 12, "bold"))
        label.grid(row=row_index, column=1, sticky="w", padx=(8, 0), pady=1)
        detail_value_labels[key] = label
    details.columnconfigure(1, weight=1)

    def print_label() -> None:
        imei1 = detail_vars["imei1"].get().strip()
        imeis = [imei1] if imei1 and imei1 != "—" else []
        if not imeis:
            messagebox.showinfo("Telefon Excel", "Önce bir telefon okuyun. Etikette basılacak IMEI yok.")
            return
        port_name = find_nelko_port(printer_port[0])
        if not port_name:
            os.startfile("ms-settings:bluetooth")
            messagebox.showinfo(
                "Telefon Excel",
                "P21 açık değil ya da eşleşme kopmuş.\n\nYazıcıyı açın. Bluetooth listesinde bağlı kalması gerekmez; "
                "ayar orada hemen düşer. Eşleşmiş görünüyorsa Etiket'e tekrar basın.",
            )
            return
        printer_port[0] = port_name
        persist_settings()
        set_status("busy", f"Etiket basılıyor ({port_name})")

        def worker() -> None:
            try:
                note = print_imei_label(port_name, imeis)
            except Exception as exc:
                events.put(("log", f"Etiket hatası: {exc}"))
                events.put(("status", ("error", f"Etiket basılamadı: {exc}")))
                return
            events.put(("log", note))
            events.put(("status", ("ok", f"Etiket çıktı. {note}")))

        threading.Thread(target=worker, daemon=True).start()

    ttk.Button(body, text="IMEI 1 etiketi bas", command=print_label).pack(fill="x", pady=(6, 2))

    button_row = ttk.Frame(body)
    button_row.pack(fill="x", pady=(6, 4))
    screen_var = tk.BooleanVar(value=allow_screen[0])

    def toggle_screen() -> None:
        allow_screen[0] = bool(screen_var.get())
        persist_settings()

    def toggle_pause() -> None:
        if watcher is None:
            return
        watcher.paused = not watcher.paused
        pause_button.configure(text="Devam" if watcher.paused else "Durdur")
        watcher.last_status = None

    def force_reread() -> None:
        if watcher is None:
            return
        watcher.reread = True
        watcher.last_status = None
        set_status("busy", "Bağlı telefon yeniden okunuyor")

    pause_button = ttk.Button(button_row, text="Durdur", command=toggle_pause)
    pause_button.pack(side="left")
    ttk.Button(button_row, text="Yeniden oku", command=force_reread).pack(side="left", padx=6)
    top_var = tk.BooleanVar(value=True)

    def toggle_top() -> None:
        root.attributes("-topmost", bool(top_var.get()))

    ttk.Checkbutton(button_row, text="Üstte", variable=top_var, command=toggle_top).pack(side="left", padx=4)
    count_var = tk.StringVar(value="")
    ttk.Label(body, textvariable=count_var).pack(fill="x", pady=(0, 4))
    ttk.Checkbutton(
        body,
        text="IMEI gelmezse *#06# aç",
        variable=screen_var,
        command=toggle_screen,
    ).pack(anchor="w")

    table_wrap = ttk.Frame(body)
    table_wrap.pack(fill="both", expand=True)
    columns = ("sira", "model", "imei1", "csc")
    tree = ttk.Treeview(table_wrap, columns=columns, show="headings", selectmode="browse")
    headings = {
        "sira": ("#", 36),
        "model": ("Model", 96),
        "imei1": ("IMEI", 150),
        "csc": ("CSC", 56),
    }
    tree.configure(height=6)
    for key in columns:
        title, width = headings[key]
        tree.heading(key, text=title)
        tree.column(key, width=width, anchor="w")
    tree.tag_configure("ok", background="#f3fbf6")
    tree.tag_configure("warn", background="#fff6e8")
    yscroll = ttk.Scrollbar(table_wrap, orient="vertical", command=tree.yview)
    xscroll = ttk.Scrollbar(table_wrap, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
    tree.grid(row=0, column=0, sticky="nsew")
    yscroll.grid(row=0, column=1, sticky="ns")
    xscroll.grid(row=1, column=0, sticky="ew")
    table_wrap.rowconfigure(0, weight=1)
    table_wrap.columnconfigure(0, weight=1)

    def on_select(_event=None) -> None:
        if refreshing["on"]:
            return
        selected = tree.selection()
        if not selected:
            return
        rows = store.snapshot()
        index = int(selected[0])
        if 0 <= index < len(rows):
            show_record(rows[index])

    tree.bind("<<TreeviewSelect>>", on_select)

    log = tk.Text(body, height=3, font=("Consolas", 8), bg="#0f1720", fg="#d6deeb", relief="flat")
    log.pack(fill="x", pady=(8, 0))
    log.configure(state="disabled")

    def add_log(line: str) -> None:
        log.configure(state="normal")
        log.insert("end", line + "\n")
        log.see("end")
        line_count = int(log.index("end-1c").split(".")[0])
        if line_count > 400:
            log.delete("1.0", "200.0")
        log.configure(state="disabled")

    adb_label = ttk.Label(body, text="", foreground="#5c6b7a")
    adb_label.pack(fill="x", pady=(6, 0))

    def poll() -> None:
        try:
            while True:
                kind, payload = events.get_nowait()
                if kind == "status":
                    set_status(payload[0], payload[1])
                elif kind == "log":
                    add_log(payload)
                elif kind == "reading":
                    shown_serial["value"] = ""
                    for var in detail_vars.values():
                        var.set("…")
                elif kind == "saved":
                    session_count["n"] += 1
                    record = payload["record"]
                    refresh_tree(record.get("serial") or "")
                    show_record(record)
                    ok = record.get("durum") == "Tamam" and not payload["error"]
                    beep(ok)
                elif kind == "devices":
                    serials = [serial for serial, state in payload if state in ("device", "unauthorized")]
                    if shown_serial["value"] and shown_serial["value"] not in serials and serials:
                        shown_serial["value"] = ""
                        for var in detail_vars.values():
                            var.set("…")
                elif kind == "alert":
                    messagebox.showinfo("Telefon Excel", payload)
        except queue.Empty:
            pass
        root.after(150, poll)

    def on_close() -> None:
        if watcher is not None:
            watcher.stop.set()
        error = store.flush()
        if error:
            messagebox.showwarning(
                "Telefon Excel",
                "Excel hâlâ açık. Son kayıtlar program kapanınca da bellekte kalmaz; "
                f"yedek dosyaya bakın:\n{error}",
            )
        persist_settings()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    refresh_tree()
    existing = store.snapshot()
    if existing:
        show_record(existing[-1])
    root.lift()
    root.focus_force()
    if load_error:
        set_status("warn", load_error)
        add_log(load_error)
    if not adb_path:
        set_status("error", "adb.exe bulunamadı. telefon-excel klasöründe adb.exe durmalı.")
        adb_label.configure(text="ADB yok")
    else:
        adb_label.configure(text=f"ADB: {adb_path}")
        watcher = Watcher(Adb(adb_path), store, events, allow_screen)
        watcher.start()
    poll()
    root.mainloop()


def run_self_test() -> None:
    import tempfile

    assert format_one_ui("60100") == "6.1"
    assert format_one_ui("60101") == "6.1.1"
    assert format_one_ui("70000") == "7.0"
    assert format_one_ui("20500") == "2.5"
    assert format_one_ui("50100") == "5.1"
    assert format_one_ui("80000") == "8.0"
    assert format_one_ui("80500") == "8.5"
    assert format_one_ui("6.1") == "6.1"
    assert format_one_ui("") == ""
    assert luhn_ok("350046302364350")
    assert tidy_baseband("S948BXXS1AZC7,S948BXXS1AZC7") == "S948BXXS1AZC7"
    props = parse_props(
        "[ro.product.model]: [SM-S948B]\n"
        "[ro.build.display.id]: [S948BXXS1AZC7]\n"
        "[gsm.version.baseband]: [S948BXXS1AZC7,S948BXXS1AZC7]\n"
        "[ro.build.version.release]: [16]\n"
        "[ro.build.version.oneui]: [80000]\n"
        "[ro.csc.sales_code]: [TUR]\n"
        "[ril.official_cscver]: [SAOMC_SM-S948B_OXM_TUR_16_0001]\n"
        "[persist.radio.imei1]: [350046302364350]\n"
        "[persist.radio.imei2]: []\n"
    )
    assert props["ro.product.model"] == "SM-S948B"
    assert extract_oneui(props) == "8.0"
    assert extract_csc(props) == ("TUR", "SAOMC_SM-S948B_OXM_TUR_16_0001")
    assert imeis_from_props(props) == ["350046302364350"]
    hidden = {
        "ril.official_cscver": "SAOMC_SM-S918B_OXM_EUX_14_0002",
    }
    assert extract_csc(hidden)[0] == "EUX"
    parcel = "Result: Parcel(\n  0x00000000: 00000000 '" + ".".join("350046302364350") + "'\n)"
    assert imeis_from_parcel(parcel) == ["350046302364350"]
    live_parcel = (
        "Result: Parcel(\n"
        "0x00000000: 00000000 0000000f 00350033 00300037 '........3.5.7.0.'\n"
        "0x00000010: 00310038 00340031 00340032 00380033 '8.1.1.4.2.4.3.8.'\n"
        "0x00000020: 00330033 00000031                   '3.3.1...        ')\n"
    )
    assert imeis_from_parcel(live_parcel) == ["357081142438331"]
    assert extract_csc({"persist.sys.omcnw_path": "/optics/configs/carriers/single/XSA/conf"})[0] == "XSA"
    xml = '<node text="IMEI1" /><node text="350046302364350" /><node text="IMEI2" /><node text="350046300767737" />'
    assert imeis_from_ui_xml(xml) == ["350046302364350", "350046300767737"]
    assert code128_modules("ABC").startswith(CODE128[104])
    assert code128_modules("356039915961831").startswith(CODE128[105])
    assert code128_modules("356039915961831").endswith(CODE128_STOP)
    bitmap = render_imei_bitmap(["356039915961831"], 96, 320)
    assert len(bitmap) == (96 // 8) * 320
    assert any(byte != 0xFF for byte in bitmap)
    folder = Path(tempfile.mkdtemp())
    path = folder / "t.xlsx"
    store = ExcelStore(path)
    store.load()
    record = {key: "" for key, _ in COLUMNS}
    record.update(
        {
            "zaman": "2026-10-29 09:00:00",
            "serial": "R5CWTEST",
            "model": "SM-S948B",
            "imei1": "350046302364350",
            "csc": "TUR",
            "build": "S948BXXS1AZC7",
            "baseband": "S948BXXS1AZC7",
            "android": "16",
            "oneui": "8.0",
            "durum": "Tamam",
        }
    )
    action, saved, error = store.upsert(record)
    assert action == "new" and error == "" and saved == path
    record["oneui"] = "8.5"
    record["zaman"] = "2026-10-29 09:05:00"
    action, _saved, error = store.upsert(record)
    assert action == "update" and error == ""
    again = ExcelStore(path)
    again.load()
    assert len(again.rows) == 1
    assert again.rows[0]["imei1"] == "350046302364350"
    assert again.rows[0]["oneui"] == "8.5"
    assert again.rows[0]["sira"] == "1"
    adb_path = find_adb()
    assert adb_path, "adb.exe yok"
    devices = Adb(adb_path).devices()
    assert isinstance(devices, list)
    print("self-test ok")
    print("adb:", adb_path)
    print("cihaz:", devices or "takili telefon yok")


def main() -> None:
    if "--self-test" in sys.argv:
        run_self_test()
        return
    try:
        run_app()
    except Exception:
        text = traceback.format_exc()
        append_log(text)
        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Telefon Excel", text)
            root.destroy()
        except Exception:
            print(text, file=sys.stderr)


if __name__ == "__main__":
    main()
