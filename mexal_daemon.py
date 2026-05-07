import tkinter as tk
from tkinter import messagebox
import os
import time
import json
import getpass
import shutil
import re
import subprocess
import ctypes
import threading
import urllib.parse
import urllib.request
import webbrowser
import smtplib
import ssl
import sys
from email.message import EmailMessage
from dataclasses import dataclass
from typing import Optional
from datetime import datetime

import ttkbootstrap as ttk
from ttkbootstrap.constants import *

from PyPDF2 import PdfReader, PdfWriter

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    _GDRIVE_AVAILABLE = True
except ImportError:
    _GDRIVE_AVAILABLE = False

# Colore badge per tipo documento
_DOC_BOOTSTYLE: dict[str, str] = {
    "DDT": "success",
    "FC":  "primary",
    "PC":  "warning",
    "OC":  "info",
    "OF":  "secondary",
    "?":   "danger",
}
_DOC_TAG_COLOR: dict[str, str] = {
    "DDT": "#198754",
    "FC":  "#0d6efd",
    "PC":  "#e65100",
    "OC":  "#0dcaf0",
    "OF":  "#6c757d",
    "?":   "#dc3545",
}


def _smtp_config() -> dict[str, object]:
    host = os.environ.get("SMTP_HOST", "").strip()
    port_raw = os.environ.get("SMTP_PORT", "").strip()
    user = os.environ.get("SMTP_USER", "").strip()
    password = os.environ.get("SMTP_PASS", "").strip()
    from_addr = os.environ.get("SMTP_FROM", "").strip() or user

    port = 0
    if port_raw:
        try:
            port = int(port_raw)
        except Exception:
            port = 0

    return {
        "host": host,
        "port": port,
        "user": user,
        "password": password,
        "from_addr": from_addr,
    }


def _app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _load_dotenv(env_path: str) -> None:
    try:
        if not os.path.isfile(env_path):
            return
        with open(env_path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                if not k:
                    continue
                os.environ.setdefault(k, v)
    except Exception as e:
        _log(f"dotenv load failed: {e}")


DOTENV_PATH = os.path.join(_app_dir(), ".env")
_load_dotenv(DOTENV_PATH)
_load_dotenv(os.path.join(_app_dir(), "local.env"))

GDRIVE_CREDENTIALS        = os.environ.get("GDRIVE_CREDENTIALS", "").strip()
GDRIVE_INBOX_PD_FOLDER_ID = os.environ.get("GDRIVE_INBOX_PD_FOLDER_ID", "").strip()
TELEGRAM_BOT_USERNAME     = os.environ.get("TELEGRAM_BOT_USERNAME", "FdA_AutoBOT_bot").strip()
TELEGRAM_BOT_TOKEN        = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
TELEGRAM_NOTIFY_CHAT_ID   = os.environ.get("TELEGRAM_NOTIFY_CHAT_ID", "").strip()
_GDRIVE_SCOPES            = ["https://www.googleapis.com/auth/drive"]

# Codici articolo Mexal degli espositori refrigerati — propone Procedura Documentale come default
_ESPOSITORE_CODES: set[str] = {"FDA-002", "FDA-003", "FDA-004", "FDA-045", "FDA-014"}


def _is_espositore_ddt(pdf_path: str) -> bool:
    """Ritorna True se il DDT contiene almeno un espositore (per codice o parola chiave)."""
    try:
        reader = PdfReader(pdf_path)
        text = " ".join(page.extract_text() or "" for page in reader.pages).upper()
        if "ESPOSITORE" in text:
            return True
        for code in _ESPOSITORE_CODES:
            if code.upper() in text:
                return True
    except Exception:
        pass
    return False


def _is_trasporto_vettore(pdf_path: str) -> bool:
    """Ritorna True se il DDT riporta 'VETTORE' come mezzo di trasporto."""
    try:
        reader = PdfReader(pdf_path)
        text = " ".join(page.extract_text() or "" for page in reader.pages).upper()
        return "VETTORE" in text
    except Exception:
        return False


_TG_DOC_EMOJI: dict[str, str] = {
    "DDT": "📦", "FC": "🧾", "PC": "📋", "OC": "📥", "OF": "📤",
}


def _tg_notify(doc: "ParsedDoc", doc_id: str, is_espositore: bool) -> None:
    """Invia notifica Telegram con bottoni inline — HTTP in thread background, non blocca tkinter."""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_NOTIFY_CHAT_ID:
        return

    emoji = _TG_DOC_EMOJI.get(doc.doc_code, "📄")
    num_str  = f" n°{doc.doc_number}" if doc.doc_number else ""
    date_str = f" — {doc.doc_date}"   if doc.doc_date   else ""
    esp_str  = " (Espositore)"        if is_espositore   else ""

    text = (
        f"{emoji} {doc.doc_type}{num_str}{date_str}\n"
        f"{doc.recipient}{esp_str}\n"
        f"Salvato locale ✅"
    )

    if doc.doc_code == "DDT" and is_espositore:
        buttons = [[
            {"text": "\U0001f5a8️ Stampa DDT",     "callback_data": f"stampa_ddt:{doc_id}"},
            {"text": "\U0001f4cb Avvia Procedura", "callback_data": f"avvia_procedura:{doc_id}"},
        ]]
    elif doc.doc_code == "DDT":
        buttons = [[
            {"text": "\U0001f5a8️ Stampa DDT",      "callback_data": f"stampa_ddt:{doc_id}"},
            {"text": "\U0001f4e6 Nuova Spedizione", "callback_data": f"nuova_spedizione:{doc_id}"},
        ]]
    else:
        buttons = None

    payload: dict = {"chat_id": TELEGRAM_NOTIFY_CHAT_ID, "text": text}
    if buttons:
        payload["reply_markup"] = {"inline_keyboard": buttons}

    def _send() -> None:
        try:
            data = json.dumps(payload).encode()
            req = urllib.request.Request(
                f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
                data=data,
                headers={"Content-Type": "application/json"},
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                result = json.loads(resp.read().decode())
            msg_id = result.get("result", {}).get("message_id", "?")
            _log(f"Telegram: msg_id={msg_id} {doc.doc_type} {doc.doc_number} {doc.recipient[:40]}")
        except Exception as e:
            _log(f"Telegram FAIL: {e}")

    threading.Thread(target=_send, daemon=True).start()


def _tg_send_simple(text: str) -> None:
    """Invia un messaggio Telegram semplice (no bottoni) in thread background."""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_NOTIFY_CHAT_ID:
        return
    payload = {"chat_id": TELEGRAM_NOTIFY_CHAT_ID, "text": text}
    def _send() -> None:
        try:
            data = json.dumps(payload).encode()
            req = urllib.request.Request(
                f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
                data=data,
                headers={"Content-Type": "application/json"},
            )
            urllib.request.urlopen(req, timeout=10)
        except Exception as e:
            _log(f"Telegram FAIL (simple): {e}")
    threading.Thread(target=_send, daemon=True).start()


def send_email_smtp(
    *,
    host: str,
    port: int,
    user: str,
    password: str,
    from_addr: str,
    to_addrs: list[str],
    subject: str,
    body: str,
    attachment_path: str,
) -> None:
    msg = EmailMessage()
    msg["From"] = from_addr
    msg["To"] = ", ".join([a for a in to_addrs if a])
    msg["Subject"] = subject
    msg.set_content(body or "")

    with open(attachment_path, "rb") as f:
        data = f.read()

    filename = os.path.basename(attachment_path)
    msg.add_attachment(data, maintype="application", subtype="pdf", filename=filename)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(host, port, context=context) as server:
        server.login(user, password)
        server.send_message(msg)


def _powershell_escape(s: str) -> str:
    return s.replace("'", "''")


def _create_shortcut_ps(
    *,
    link_path: str,
    target_path: str,
    arguments: str,
    working_dir: str,
    description: str,
) -> None:
    link_path = os.path.abspath(link_path)
    target_path = os.path.abspath(target_path)
    working_dir = os.path.abspath(working_dir)
    ps = (
        "$WshShell = New-Object -ComObject WScript.Shell;"
        f"$Shortcut = $WshShell.CreateShortcut('{_powershell_escape(link_path)}');"
        f"$Shortcut.TargetPath = '{_powershell_escape(target_path)}';"
        f"$Shortcut.Arguments = '{_powershell_escape(arguments)}';"
        f"$Shortcut.WorkingDirectory = '{_powershell_escape(working_dir)}';"
        f"$Shortcut.Description = '{_powershell_escape(description)}';"
        "$Shortcut.Save();"
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        check=True,
        creationflags=0x08000000,
    )


def _remove_file_silent(path: str) -> None:
    try:
        if os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass


def _ps_str(s: str) -> str:
    return s.replace("'", "''")


def _register_task_scheduler(task_name: str, exe: str, arguments: str, workdir: str) -> None:
    """Registra un Task Scheduler con restart-on-failure e delay 60s al logon."""
    xml = f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <Delay>PT60S</Delay>
    </LogonTrigger>
  </Triggers>
  <Actions Context="Author">
    <Exec>
      <Command>{exe}</Command>
      <Arguments>{arguments}</Arguments>
      <WorkingDirectory>{workdir}</WorkingDirectory>
    </Exec>
  </Actions>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <RestartOnFailure>
      <Interval>PT1M</Interval>
      <Count>5</Count>
    </RestartOnFailure>
    <IdleSettings><StopOnIdleEnd>false</StopOnIdleEnd><RestartOnIdle>false</RestartOnIdle></IdleSettings>
  </Settings>
  <Principals>
    <Principal id="Author">
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
</Task>"""
    xml_path = os.path.join(os.environ.get("TEMP", "C:\\Temp"), f"{task_name}.xml")
    with open(xml_path, "w", encoding="utf-16") as f:
        f.write(xml)
    result = subprocess.run(
        ["schtasks", "/Create", "/TN", task_name, "/XML", xml_path, "/F"],
        creationflags=0x08000000, capture_output=True,
    )
    if result.returncode != 0:
        err = result.stderr.decode(errors="replace") or result.stdout.decode(errors="replace")
        raise RuntimeError(f"schtasks fallito (rc={result.returncode}): {err.strip()}")
    _remove_file_silent(xml_path)


def install_windows_shortcuts() -> None:
    if getattr(sys, "frozen", False):
        exe = os.path.abspath(sys.executable)
        workdir = os.path.dirname(exe)
        pyw_exe = exe
        args_task = ""
    else:
        script_path = os.path.abspath(__file__)
        workdir = os.path.dirname(script_path)
        py_exe = sys.executable
        pyw_exe = py_exe
        if py_exe.lower().endswith("python.exe"):
            cand = py_exe[:-10] + "pythonw.exe"
            if os.path.isfile(cand):
                pyw_exe = cand
        exe = py_exe
        args_task = f'"{script_path}"'

    # Task Scheduler (avvio affidabile + restart automatico)
    _register_task_scheduler("MexalDaemon", pyw_exe, args_task, workdir)

    # Watchdog separato ogni 5 min
    watchdog_path = os.path.join(workdir, "mexal_watchdog.pyw")
    if os.path.isfile(watchdog_path):
        _register_task_scheduler("MexalWatchdog", pyw_exe, f'"{watchdog_path}"', workdir)

    # Collegamento Desktop (per avvio manuale)
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    os.makedirs(desktop_dir, exist_ok=True)
    _create_shortcut_ps(
        link_path=os.path.join(desktop_dir, "Mexal Automation Daemon.lnk"),
        target_path=exe,
        arguments=args_task,
        working_dir=workdir,
        description="Mexal Automation Daemon",
    )


def uninstall_windows_shortcuts() -> None:
    for task in ("MexalDaemon", "MexalWatchdog"):
        try:
            subprocess.run(
                ["schtasks", "/Delete", "/TN", task, "/F"],
                check=False, creationflags=0x08000000, capture_output=True,
            )
        except Exception:
            pass
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    _remove_file_silent(os.path.join(desktop_dir, "Mexal Automation Daemon.lnk"))


_SINGLE_INSTANCE_MUTEX = None


def _ensure_single_instance(mutex_name: str) -> None:
    global _SINGLE_INSTANCE_MUTEX
    try:
        ERROR_ALREADY_EXISTS = 183
        handle = ctypes.windll.kernel32.CreateMutexW(None, True, mutex_name)
        last_error = ctypes.windll.kernel32.GetLastError()
        if last_error == ERROR_ALREADY_EXISTS:
            raise SystemExit(0)
        _SINGLE_INSTANCE_MUTEX = handle
    except SystemExit:
        raise
    except Exception as e:
        _log(f"Single-instance lock failed: {e}")


USER = getpass.getuser()
_BASE_PATH_DEFAULT = os.path.join("C:/Users", USER, "Desktop", "AMMINISTRAZIONE_2025")
BOLLE_DIR            = os.environ.get("BOLLE_DIR")            or os.path.join(_BASE_PATH_DEFAULT, "BOLLE_2025")
FATTURE_DIR          = os.environ.get("FATTURE_DIR")          or os.path.join(_BASE_PATH_DEFAULT, "FATTURE_2025")
PREVENTIVI_DIR       = os.environ.get("PREVENTIVI_DIR")       or os.path.join(_BASE_PATH_DEFAULT, "PREVENTIVI_2025")
ORDINI_DIR           = os.environ.get("ORDINI_DIR")           or os.path.join(_BASE_PATH_DEFAULT, "ORDINI_2025")
ORDINI_FORNITORI_DIR = os.environ.get("ORDINI_FORNITORI_DIR") or os.path.join(_BASE_PATH_DEFAULT, "ORDINI FORNITORI_2025")

PATHS = {
    # descrizioni
    "Bolla": BOLLE_DIR,
    "DDT": BOLLE_DIR,
    "Fattura": FATTURE_DIR,
    "Preventivo": PREVENTIVI_DIR,
    "Ordine cliente": ORDINI_DIR,
    "Ordine fornitore": ORDINI_FORNITORI_DIR,
    # codici
    "DDT": BOLLE_DIR,
    "FC": FATTURE_DIR,
    "PC": PREVENTIVI_DIR,
    "OC": ORDINI_DIR,
    "OF": ORDINI_FORNITORI_DIR,
}

LOG_FILE = "mexal_daemon.log"


def _log(msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def _safe_get_mtime(path: str) -> Optional[float]:
    try:
        return os.path.getmtime(path)
    except Exception:
        return None


def _safe_get_size(path: str) -> Optional[int]:
    try:
        return os.path.getsize(path)
    except Exception:
        return None

def _detect_mexal_temp_dir() -> str:
    override = os.environ.get("MEXAL_TEMP")
    if override and os.path.isdir(override):
        _log(f"MEXAL_TEMP override: {override}")
        return override

    candidates = [
        r"C:\Passepartout\PassClient\mxdesk1205143000\temp",
        r"C:\Passepartout\PassClient1\mxdesk1205143000\temp",
    ]

    def newest_pdf_mtime(base_dir: str) -> float:
        newest = 0.0
        try:
            for root_dir, _, files in os.walk(base_dir):
                for fn in files:
                    if fn.lower().endswith(".pdf"):
                        full_path = os.path.join(root_dir, fn)
                        # NOTE: this runs at import time; keep it independent of later definitions.
                        try:
                            mt = os.path.getmtime(full_path)
                        except Exception:
                            mt = None
                        if mt and mt > newest:
                            newest = mt
        except Exception:
            return 0.0
        return newest

    existing = [c for c in candidates if os.path.isdir(c)]
    if not existing:
        return candidates[0]

    existing.sort(key=newest_pdf_mtime, reverse=True)
    chosen = existing[0]
    _log(f"MEXAL_TEMP autodetect candidates={existing} chosen={chosen}")
    return chosen


MEXAL_TEMP = _detect_mexal_temp_dir()
STATE_FILE = "documenti_state.json"
SEEN_FILE = "watcher_seen.json"

_log(f"Startup. MEXAL_TEMP={MEXAL_TEMP}")


@dataclass(frozen=True)
class ParsedDoc:
    source_path: str
    created_at: float
    doc_code: str
    doc_type: str
    doc_number: str
    doc_date: str
    recipient: str
    dest_cap: str = ""
    dest_citta: str = ""
    dest_provincia: str = ""
    dest_tel: str = ""
    dest_email: str = ""


def _load_json(path: str, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path: str, data) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _safe_get_mtime(path: str) -> Optional[float]:
    try:
        return os.path.getmtime(path)
    except Exception:
        return None


def _safe_get_size(path: str) -> Optional[int]:
    try:
        return os.path.getsize(path)
    except Exception:
        return None


def extract_first_page_text(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    if not reader.pages:
        return ""
    text = reader.pages[0].extract_text() or ""
    return text


_HEADER_RE = re.compile(
    r"^(?P<tipo>[A-Za-zÀ-ÿ]+)\b.*?\b(n\.?|nr\.?|n°)\s*(?P<num>[0-9/]+)\b.*?\bdel\s+(?P<data>\d{2}/\d{2}/\d{4})\b",
    re.IGNORECASE,
)

_DOCNUM_RE = re.compile(
    r"\b(?:n\s*[\.:°º]?|nr\s*[\.:°º]?)\s*(?P<num>[0-9]+(?:\s*/\s*[0-9]+)?)\b",
    re.IGNORECASE,
)
_DOCDATE_RE = re.compile(
    r"\b(?:del|data)(?=[\s\d])\.?\s*(?P<data>\d{2}[\./-]\d{2}[\./-]\d{4})\b",
    re.IGNORECASE,
)
_ADDR_RE = re.compile(
    r"(?P<cap>\d{5})\s+(?P<citta>[A-ZÀ-Ü][A-ZÀ-Ü\s'\.]+?)\s+(?P<prov>[A-Z]{2,3})\b"
)
_ANYDATE_RE = re.compile(r"\b(?P<data>\d{2}[\./-]\d{2}[\./-]\d{4})\b")
_NUM_AFTER_N_RE = re.compile(r"\bn\b[^0-9]*(?P<num>[0-9/]+)", re.IGNORECASE)
_SERIES_PROG_RE = re.compile(r"\b(?P<serie>\d+)\s*/\s*(?P<prog>\d+)\b")
# Telefono + email sulla stessa riga (formato Mexal: "Tel.0541 123 Mail info@azienda.it")
_TEL_MAIL_RE = re.compile(
    r"Tel\.?\s*(?P<tel>[\d][\d\s\+\-\/\.]{3,}?)\s+Mail\s*(?P<email>[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
    re.IGNORECASE,
)
_TEL_ONLY_RE = re.compile(
    r"Tel\.?\s*(?P<tel>[\d][\d\s\+\-\/\.]{4,}?)(?:\s|$)",
    re.IGNORECASE,
)


def _doc_code_from_lines(lines: list[str]) -> tuple[str, str]:
    first_lines_raw = [ln.lower() for ln in lines[:10]]
    first_compact = [re.sub(r"[^a-z0-9]+", "", ln) for ln in first_lines_raw]
    first_norm = [
        re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", ln)).strip()
        for ln in first_lines_raw
    ]
    first5_joined_norm = " ".join(first_norm[:5]).strip()
    first5_joined_compact = re.sub(r"[^a-z0-9]+", "", first5_joined_norm)

    header_raw = " ".join(lines[:20]).lower()
    header_norm = re.sub(r"[^a-z0-9]+", " ", header_raw)
    header_norm = re.sub(r"\s+", " ", header_norm).strip()
    header_compact = re.sub(r"[^a-z0-9]+", "", header_raw)

    # Priorità assoluta: se nelle prime righe c'è "fattura" allora è una fattura.
    # Serve a evitare falsi positivi DDT quando l'estrazione testo contiene "ddt" in altre zone.
    if any("fattura" in c for c in first_compact) or any("fattura" in n for n in first_norm):
        return "FC", "Fattura"

    # Caso noto: "D.D.T. consegna" (deve comparire come etichetta in alto, non altrove)
    # Gestisce anche il caso in cui "D.D.T." e "consegna" siano su righe diverse.
    if (
        "ddtconsegna" in first5_joined_compact
        or "ddt consegna" in first5_joined_norm
        or "d d t consegna" in first5_joined_norm
        or any("ddt consegna" in n or "d d t consegna" in n for n in first_norm[:5])
    ):
        return "DDT", "DDT"

    if "preventivo" in header_compact or "preventivo" in header_norm:
        return "PC", "Preventivo"

    # Ordini: evitiamo falsi positivi (es. DDT che contiene la parola "cliente" o riferimenti ad "ordine").
    # Richiediamo la dicitura in alto (prime righe) e in forma coerente.
    if (
        any("ordine cliente" in n for n in first_norm)
        or any("ordinecliente" in c for c in first_compact)
        or "ordinecliente" in first5_joined_compact
    ):
        return "OC", "Ordine cliente"
    if (
        any("ordine fornitore" in n for n in first_norm)
        or any("ordinefornitore" in c for c in first_compact)
        or any("ordine forn" in n for n in first_norm)
        or any("ordineforn" in c for c in first_compact)
        or "ordinefornitore" in first5_joined_compact
    ):
        return "OF", "Ordine fornitore"

    # Fattura: fallback (in teoria già coperta sopra dalle prime righe)
    if ("fattura" in header_compact or "fattura" in header_norm) and "ordine" not in header_norm:
        return "FC", "Fattura"

    # DDT: spesso appare come D.D.T. nel PDF.
    # Evitiamo falsi positivi: se troviamo "fattura" nel testo compatto non deve diventare DDT.
    if "ddt" in header_compact and "fattura" not in header_compact:
        return "DDT", "DDT"
    if "documento di trasporto" in header_raw:
        return "DDT", "DDT"
    if "bolla" in header_norm:
        return "DDT", "DDT"
    return "?", "Sconosciuto"


def parse_mexal_pdf(pdf_path: str) -> Optional[ParsedDoc]:
    mtime = _safe_get_mtime(pdf_path)
    if mtime is None:
        return None

    try:
        text = extract_first_page_text(pdf_path)
    except Exception:
        return None

    raw_lines = [ln.strip() for ln in text.splitlines()]
    lines = [re.sub(r"\s+", " ", ln) for ln in raw_lines if ln.strip()]

    doc_code, doc_type = _doc_code_from_lines(lines)
    doc_number = ""
    doc_date = ""

    if doc_code == "?":
        preview = " | ".join(lines[:20])
        _log(f"Doc type UNKNOWN. file={pdf_path} preview={preview}")

    # Numero/data possono essere sulla stessa riga (es. Preventivo) oppure su righe separate (es. DDT).
    for ln in lines[:20]:
        m = _HEADER_RE.match(ln)
        if m:
            doc_number = doc_number or m.group("num")
            doc_date = doc_date or m.group("data")
        else:
            if not doc_number:
                mnum = _DOCNUM_RE.search(ln)
                if mnum:
                    doc_number = mnum.group("num")
                else:
                    # fallback per casi dove l'estrazione separa "n" dalla punteggiatura
                    mnum2 = _NUM_AFTER_N_RE.search(ln)
                    if mnum2:
                        doc_number = mnum2.group("num")
            if not doc_date:
                mdat = _DOCDATE_RE.search(ln)
                if mdat:
                    doc_date = mdat.group("data").replace(".", "/").replace("-", "/")

        if doc_number and doc_date:
            break

    # Fallback: alcuni DDT potrebbero non contenere chiaramente "del" nell'estrazione; prendiamo la prima data trovata.
    if doc_code == "DDT" and not doc_date:
        for ln in lines[:25]:
            m_any = _ANYDATE_RE.search(ln)
            if m_any:
                doc_date = m_any.group("data").replace(".", "/").replace("-", "/")
                break

    # Serie/progressivo: spesso il numero è nel formato "3/ 1234" dove 3 è la serie e 1234 è il progressivo.
    # Per OC/OF/FC lo normalizziamo come "3-1234".
    if doc_code in {"OC", "OF", "FC"}:
        for ln in lines[:20]:
            m_sp = _SERIES_PROG_RE.search(ln)
            if m_sp:
                doc_number = f"{m_sp.group('serie')}-{m_sp.group('prog')}"
                break

    recipient = ""
    dest_idx = None
    for i, ln in enumerate(lines):
        if "destinatario" in ln.lower():
            dest_idx = i
            break

    if dest_idx is not None:
        for ln in lines[dest_idx + 1 : dest_idx + 8]:
            if ln.strip():
                recipient = ln.strip()
                break

    if not recipient:
        recipient = "(Destinatario non trovato)"
    else:
        # Mexal spesso rende "Destinatario" e "Destinazione" come due colonne.
        # Nell'estrazione testo possono finire sulla stessa riga in duplicato.
        # In tal caso prendiamo la prima colonna.
        parts = [p.strip() for p in re.split(r"\t+|\s{2,}", recipient) if p.strip()]
        if len(parts) >= 2:
            recipient = parts[0]
        else:
            # Caso: la stessa stringa ripetuta due volte nella stessa riga
            m_dup = re.match(r"^(?P<a>.+?)\s+(?P=a)\s*$", recipient)
            if m_dup:
                recipient = m_dup.group("a").strip()

    # Estrai CAP / Città / Provincia dalle righe dopo il destinatario
    dest_cap = dest_citta = dest_provincia = ""
    if dest_idx is not None:
        for ln in lines[dest_idx + 1 : dest_idx + 14]:
            m_addr = _ADDR_RE.search(ln)
            if m_addr:
                dest_cap = m_addr.group("cap")
                dest_citta = m_addr.group("citta").strip()
                dest_provincia = m_addr.group("prov").upper()
                break

    # Normalizza numero: rimuove spazi interni (es. "3/ 61" → "3/61")
    doc_number = re.sub(r"\s+", "", doc_number)

    # Estrai telefono e email — salta DDT Grenke (numero di contatto Grenke, non del cliente)
    dest_tel = dest_email = ""
    is_grenke = "grenke" in recipient.lower()
    if not is_grenke:
        for ln in lines:
            m_tm = _TEL_MAIL_RE.search(ln)
            if m_tm:
                dest_tel   = re.sub(r"[\s\.]", "", m_tm.group("tel")).strip()
                dest_email = m_tm.group("email").strip()
                break
            m_t = _TEL_ONLY_RE.search(ln)
            if m_t and not dest_tel:
                dest_tel = re.sub(r"[\s\.]", "", m_t.group("tel")).strip()

    return ParsedDoc(
        source_path=pdf_path,
        created_at=mtime,
        doc_code=doc_code,
        doc_type=doc_type,
        doc_number=doc_number,
        doc_date=doc_date,
        recipient=recipient,
        dest_cap=dest_cap,
        dest_citta=dest_citta,
        dest_provincia=dest_provincia,
        dest_tel=dest_tel,
        dest_email=dest_email,
    )


def save_first_page_only(input_path: str, output_path: str) -> None:
    reader = PdfReader(input_path)
    writer = PdfWriter()
    writer.add_page(reader.pages[0])
    with open(output_path, "wb") as f_out:
        writer.write(f_out)


def _gdrive_get_credentials():
    """Credenziali Google OAuth — stesso pattern dell'orchestratore FdA."""
    if not _GDRIVE_AVAILABLE:
        raise RuntimeError("Librerie Google non installate (pip install google-api-python-client google-auth).")
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    client_id     = os.environ.get("GOOGLE_CLIENT_ID", "").strip()
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET", "").strip()
    refresh_token = os.environ.get("GOOGLE_REFRESH_TOKEN", "").strip()
    if not client_id or not client_secret or not refresh_token:
        raise RuntimeError(
            "Credenziali Google mancanti nel .env.\n"
            "Copia GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET e GOOGLE_REFRESH_TOKEN\n"
            "dal file .env dell'orchestratore FdA."
        )
    creds = Credentials(
        token=None, refresh_token=refresh_token,
        client_id=client_id, client_secret=client_secret,
        token_uri="https://oauth2.googleapis.com/token",
        scopes=_GDRIVE_SCOPES,
    )
    if not creds.valid:
        creds.refresh(Request())
    return creds


def _gdrive_upload_to_inbox_pd(local_path: str, filename: str) -> str:
    """Carica un PDF su GDrive nella cartella Inbox_PD. Restituisce l'ID file."""
    if not GDRIVE_INBOX_PD_FOLDER_ID:
        raise RuntimeError("GDRIVE_INBOX_PD_FOLDER_ID non configurato nel .env")
    creds = _gdrive_get_credentials()
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    meta = {"name": filename, "parents": [GDRIVE_INBOX_PD_FOLDER_ID]}
    media = MediaFileUpload(local_path, mimetype="application/pdf", resumable=False)
    f = service.files().create(body=meta, media_body=media, fields="id").execute()
    return f.get("id", "")


def print_pdf(path: str, copies: int = 1) -> None:
    copies = max(1, int(copies or 1))

    # Prefer SumatraPDF if installed: supports silent printing and copies.
    sumatra_candidates = [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    ]
    sumatra = next((p for p in sumatra_candidates if os.path.isfile(p)), None)
    if sumatra:
        # -print-settings supports copies; use default printer.
        args = [
            sumatra,
            "-silent",
            "-print-to-default",
            "-print-settings",
            f"{copies}x",
            path,
        ]
        subprocess.run(args, check=True)
        return

    # Fallback: uses the default PDF handler. Copies may be ignored by some viewers.
    for _ in range(copies):
        try:
            os.startfile(path, "print")
        except OSError as e:
            # WinError 1155: no application associated with the specified file for this operation
            if getattr(e, "winerror", None) == 1155:
                raise RuntimeError(
                    "Nessuna applicazione associata alla stampa PDF (WinError 1155). "
                    "Installa un lettore PDF che supporti la stampa da shell (consigliato: SumatraPDF) "
                    "oppure imposta un'app predefinita per i PDF con la funzione di stampa."
                )
            raise
        time.sleep(1.0)


class MexalDaemonApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.withdraw()

        self.state = _load_json(STATE_FILE, {"docs": {}})
        # seen[path] = last_mtime_processed
        self.seen = _load_json(SEEN_FILE, {"seen": {}})
        self._size_history: dict[str, list[int]] = {}

        self.overlay: Optional[tk.Toplevel] = None
        self.list_window: Optional[tk.Toplevel] = None
        self.list_tree: Optional[ttk.Treeview] = None

        self._last_detected: list[ParsedDoc] = []
        self._current_doc: Optional[ParsedDoc] = None
        self._tick_count = 0

        self.root.after(1000, self._tick)

    def _refresh_list_tree(self) -> None:
        tree = self.list_tree
        if not tree or not tree.winfo_exists():
            return
        try:
            for iid in list(tree.get_children()):
                tree.delete(iid)
        except Exception:
            return
        docs = self._collect_last_docs(limit=5)
        for doc in docs:
            created_str = time.strftime("%d/%m %H:%M", time.localtime(doc.created_at))
            tag = doc.doc_code if doc.doc_code in _DOC_TAG_COLOR else "?"
            tree.insert(
                "", "end",
                iid=self._doc_id(doc),
                values=(doc.doc_code, doc.doc_type, doc.recipient, doc.doc_date, created_str),
                tags=(tag,),
            )
        for code, color in _DOC_TAG_COLOR.items():
            tree.tag_configure(code, foreground=color)
        if tree.get_children():
            tree.selection_set(tree.get_children()[0])
            tree.event_generate("<<TreeviewSelect>>")

    def _tick(self):
        try:
            self._tick_count += 1
            if self._tick_count == 1:
                _log("Tick loop started")
            new_docs = self._scan_for_new_docs()
            if new_docs:
                self._last_detected = new_docs
                self._current_doc = new_docs[0]
                for doc in new_docs:
                    doc_id = self._doc_id(doc)
                    dest = self._do_save(doc, doc_id)
                    if dest is not None:
                        is_esp = doc.doc_code == "DDT" and _is_espositore_ddt(doc.source_path)
                        _tg_notify(doc, doc_id, is_esp)
                        if is_esp and _GDRIVE_AVAILABLE and GDRIVE_INBOX_PD_FOLDER_ID \
                                and os.environ.get("GOOGLE_CLIENT_ID") \
                                and os.environ.get("GOOGLE_REFRESH_TOKEN"):
                            threading.Thread(
                                target=self._do_gdrive_upload_bg,
                                args=(doc, doc_id),
                                daemon=True,
                            ).start()
                        if doc.doc_code == "DDT":
                            threading.Thread(
                                target=self._upsert_destinatario_bg,
                                args=(doc,),
                                daemon=True,
                            ).start()
                self._show_overlay(new_docs[0])
        except Exception as e:
            _log(f"Tick error (#{self._tick_count}): {e}")
        finally:
            try:
                self.root.after(1000, self._tick)
            except Exception as e:
                _log(f"After error: {e}")

    def _scan_for_new_docs(self) -> list[ParsedDoc]:
        if not os.path.isdir(MEXAL_TEMP):
            if self._tick_count == 1 or self._tick_count % 10 == 0:
                _log(f"Scan: MEXAL_TEMP non esiste: {MEXAL_TEMP}")
            return []

        pdfs = []
        for root_dir, _, files in os.walk(MEXAL_TEMP):
            for fn in files:
                if fn.lower().endswith(".pdf"):
                    full_path = os.path.join(root_dir, fn)
                    mtime = _safe_get_mtime(full_path)
                    if mtime is None:
                        continue
                    pdfs.append((full_path, mtime))

        pdfs.sort(key=lambda x: x[1], reverse=True)

        if self._tick_count == 1 or self._tick_count % 10 == 0:
            _log(f"Scan: found_pdfs={len(pdfs)} (showing up to 20)")

        parsed: list[ParsedDoc] = []
        for path, mtime in pdfs[:20]:
            seen_map = self.seen.setdefault("seen", {})
            last_mtime = seen_map.get(path)
            if last_mtime is not None and mtime <= float(last_mtime):
                continue

            size = _safe_get_size(path)
            if size is None:
                continue

            hist = self._size_history.setdefault(path, [])
            hist.append(size)
            if len(hist) > 3:
                del hist[0]

            if len(hist) < 2 or hist[-1] != hist[-2]:
                if self._tick_count % 10 == 0:
                    _log(f"Scan: not_stable_yet: {os.path.basename(path)} size_hist={hist}")
                continue

            doc = parse_mexal_pdf(path)
            if not doc:
                _log(f"Scan: parse_failed: {path}")
                continue

            parsed.append(doc)
            seen_map[path] = mtime

        if parsed:
            _log(f"Scan: new_docs={len(parsed)} first={os.path.basename(parsed[0].source_path)} type={parsed[0].doc_type}")
            _save_json(SEEN_FILE, self.seen)

        return parsed

    def _show_overlay(self, doc: ParsedDoc) -> None:
        if self.overlay and self.overlay.winfo_exists():
            self.overlay.lift()
            return

        win = tk.Toplevel(self.root)
        self.overlay = win
        win.title("Mexal — Nuovo documento")
        win.attributes("-topmost", True)
        win.resizable(False, False)

        # Header colorato
        hdr = ttk.Frame(win, bootstyle="primary", padding=(16, 10))
        hdr.grid(row=0, column=0, sticky="ew")
        ttk.Label(
            hdr,
            text="📄  Nuovo documento Mexal",
            font=("Segoe UI", 11, "bold"),
            bootstyle="inverse-primary",
        ).grid(row=0, column=0, sticky="w")

        # Body
        body = ttk.Frame(win, padding=(20, 14, 20, 6))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)

        color = _DOC_BOOTSTYLE.get(doc.doc_code, "secondary")
        ttk.Label(
            body,
            text=f"  {doc.doc_code}  ",
            bootstyle=f"inverse-{color}",
            font=("Segoe UI", 9, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Label(
            body,
            text=doc.recipient,
            font=("Segoe UI", 11, "bold"),
        ).grid(row=0, column=1, sticky="w")

        if doc.doc_date:
            ttk.Label(
                body,
                text=f"{doc.doc_type}  •  {doc.doc_date}",
                font=("Segoe UI", 9),
                bootstyle="secondary",
            ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        ttk.Separator(body).grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 8))

        ttk.Label(
            body,
            text="Vuoi processarlo adesso?",
            font=("Segoe UI", 10),
        ).grid(row=3, column=0, columnspan=2, sticky="w")

        # Bottoni
        btn_row = ttk.Frame(win, padding=(20, 8, 20, 16))
        btn_row.grid(row=2, column=0, sticky="ew")
        btn_row.columnconfigure(0, weight=1)
        btn_row.columnconfigure(1, weight=1)
        ttk.Button(
            btn_row, text="Ignora", bootstyle="secondary-outline",
            command=self._overlay_no, width=14,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(
            btn_row, text="✅  Processa", bootstyle="success",
            command=self._overlay_yes, width=14,
        ).grid(row=0, column=1, sticky="ew")

        win.update_idletasks()
        w = max(win.winfo_reqwidth(), 360)
        h = win.winfo_reqheight()
        x = int((win.winfo_screenwidth() - w) / 2)
        y = int((win.winfo_screenheight() - h) / 2)
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.protocol("WM_DELETE_WINDOW", self._overlay_no)

    def _overlay_no(self):
        if self.overlay and self.overlay.winfo_exists():
            self.overlay.destroy()
        self.overlay = None

    def _overlay_yes(self):
        self._overlay_no()
        doc = self._current_doc
        if not doc:
            return
        doc_id = self._doc_id(doc)
        dest_path = self._do_save(doc, doc_id)
        if dest_path is None:
            return
        if doc.doc_code == "DDT":
            is_esp = _is_espositore_ddt(doc.source_path)
            is_vettore = _is_trasporto_vettore(doc.source_path)
            self._dialog_ddt_azione(doc, doc_id, is_esp, is_vettore)
        else:
            messagebox.showinfo("Completato", f"Documento salvato in:\n{dest_path}")

    def _doc_id(self, doc: ParsedDoc) -> str:
        base = os.path.basename(doc.source_path)
        return f"{base}|{int(doc.created_at)}"

    def _preferred_save_dir(self, doc: ParsedDoc) -> str:
        if doc.doc_code in PATHS:
            chosen = PATHS[doc.doc_code]
            _log(f"SaveDir: doc_code={doc.doc_code} doc_type={doc.doc_type} chosen={chosen}")
            return chosen
        if doc.doc_type in PATHS:
            chosen = PATHS[doc.doc_type]
            _log(f"SaveDir: doc_code={doc.doc_code} doc_type={doc.doc_type} chosen={chosen}")
            return chosen
        chosen = os.path.join(BASE_PATH, "DOCUMENTI_2025")
        _log(f"SaveDir: doc_code={doc.doc_code} doc_type={doc.doc_type} chosen={chosen}")
        return chosen

    def _get_doc_state(self, doc_id: str) -> dict:
        st = self.state.setdefault("docs", {}).setdefault(
            doc_id,
            {
                "saved": False,
                "printed": False,
                "emailed": False,
                "gdrive_uploaded": False,
                "dest_path": "",
                "meta": {},
            },
        )
        st.setdefault("gdrive_uploaded", False)
        return st

    def _show_list_window(self):
        if self.list_window and self.list_window.winfo_exists():
            self._refresh_list_tree()
            try:
                self.list_window.deiconify()
            except Exception:
                pass
            self.list_window.lift()
            try:
                self.list_window.focus_force()
            except Exception:
                pass
            return

        win = tk.Toplevel(self.root)
        self.list_window = win
        win.title("Mexal — Documenti recenti")
        win.attributes("-topmost", True)
        win.minsize(720, 320)
        win.columnconfigure(0, weight=1)
        win.rowconfigure(1, weight=1)
        try:
            win.deiconify()
            win.focus_force()
        except Exception:
            pass

        # Header
        hdr = ttk.Frame(win, bootstyle="primary", padding=(16, 10))
        hdr.grid(row=0, column=0, sticky="ew")
        ttk.Label(
            hdr,
            text="📋  Documenti Mexal — ultimi 5",
            font=("Segoe UI", 11, "bold"),
            bootstyle="inverse-primary",
        ).grid(row=0, column=0, sticky="w")

        # Main frame
        main = ttk.Frame(win, padding=(12, 10, 12, 12))
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)

        # Treeview con scrollbar
        tree_wrap = ttk.Frame(main)
        tree_wrap.grid(row=0, column=0, sticky="nsew")
        tree_wrap.columnconfigure(0, weight=1)

        cols = ("tipo", "descr", "destinatario", "data_doc", "creato")
        tree = ttk.Treeview(
            tree_wrap, columns=cols, show="headings", height=6, bootstyle="primary",
        )
        self.list_tree = tree

        tree.heading("tipo",         text="Tipo",        anchor="center")
        tree.heading("descr",        text="Descrizione")
        tree.heading("destinatario", text="Destinatario")
        tree.heading("data_doc",     text="Data",        anchor="center")
        tree.heading("creato",       text="Rilevato",    anchor="center")

        tree.column("tipo",         width=62,  stretch=False, anchor="center")
        tree.column("descr",        width=120, stretch=False)
        tree.column("destinatario", width=300, stretch=True)
        tree.column("data_doc",     width=90,  stretch=False, anchor="center")
        tree.column("creato",       width=120, stretch=False, anchor="center")

        sb = ttk.Scrollbar(tree_wrap, orient="vertical", command=tree.yview, bootstyle="primary-round")
        tree.configure(yscrollcommand=sb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        # Popola e applica colori per tipo
        docs = self._collect_last_docs(limit=5)
        for doc in docs:
            created_str = time.strftime("%d/%m %H:%M", time.localtime(doc.created_at))
            tag = doc.doc_code if doc.doc_code in _DOC_TAG_COLOR else "?"
            tree.insert(
                "", "end",
                iid=self._doc_id(doc),
                values=(doc.doc_code, doc.doc_type, doc.recipient, doc.doc_date, created_str),
                tags=(tag,),
            )
        for code, color in _DOC_TAG_COLOR.items():
            tree.tag_configure(code, foreground=color)

        if tree.get_children():
            tree.selection_set(tree.get_children()[0])

        ttk.Separator(main).grid(row=1, column=0, sticky="ew", pady=(10, 8))

        # Riga pulsanti principali
        btn_row1 = ttk.Frame(main)
        btn_row1.grid(row=2, column=0, sticky="ew")
        for i in range(4):
            btn_row1.columnconfigure(i, weight=1)

        btn_save  = ttk.Button(btn_row1, text="💾  Salva",   bootstyle="success",   command=lambda: self._action_save(tree))
        btn_print = ttk.Button(btn_row1, text="🖨️  Stampa",  bootstyle="secondary", command=lambda: self._action_print(tree))
        btn_email = ttk.Button(btn_row1, text="✉️  Email",   bootstyle="info",      command=lambda: self._action_email(tree))
        btn_view  = ttk.Button(btn_row1, text="👁️  Vedi",    bootstyle="light",     command=lambda: self._action_view(tree))

        btn_save.grid( row=0, column=0, sticky="ew", padx=(0, 4))
        btn_print.grid(row=0, column=1, sticky="ew", padx=(0, 4))
        btn_email.grid(row=0, column=2, sticky="ew", padx=(0, 4))
        btn_view.grid( row=0, column=3, sticky="ew")

        # Pulsante GDrive (seconda riga, full width)
        gdrive_ok = (
            _GDRIVE_AVAILABLE
            and bool(GDRIVE_INBOX_PD_FOLDER_ID)
            and bool(os.environ.get("GOOGLE_CLIENT_ID"))
            and bool(os.environ.get("GOOGLE_REFRESH_TOKEN"))
        )
        btn_gdrive = ttk.Button(
            main,
            text="☁️  Invia a GDrive Inbox_PD (Procedura Documentale)",
            bootstyle="primary-outline",
            command=lambda: self._action_gdrive_pd(tree),
        )
        btn_gdrive.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        if not gdrive_ok:
            btn_gdrive.state(["disabled"])

        # Pulsante Spedizione (terza riga, full width)
        spedizioni_ok = True  # app locale sempre disponibile
        btn_spedizione = ttk.Button(
            main,
            text="📦  Nuova Spedizione",
            bootstyle="success-outline",
            command=lambda: self._action_nuova_spedizione(tree),
        )
        btn_spedizione.grid(row=4, column=0, sticky="ew", pady=(4, 0))

        def refresh_buttons(*_):
            sel = tree.selection()
            if not sel:
                for b in (btn_save, btn_print, btn_email, btn_view, btn_gdrive, btn_spedizione):
                    b.state(["disabled"])
                return

            doc_id = sel[0]
            st = self._get_doc_state(doc_id)
            doc = self._find_doc_by_id(doc_id)

            saved = st.get("saved", False)
            btn_save.state(["disabled"] if saved else ["!disabled"])
            btn_print.state(["!disabled"] if saved else ["disabled"])
            btn_email.state(["!disabled"] if saved else ["disabled"])
            btn_view.state(["!disabled"] if saved else ["disabled"])

            # GDrive: abilita solo se DDT + credenziali ok + non già caricato
            is_ddt = doc and doc.doc_code == "DDT"
            already_uploaded = st.get("gdrive_uploaded", False)
            if gdrive_ok and is_ddt and not already_uploaded:
                btn_gdrive.state(["!disabled"])
                btn_gdrive.configure(text="☁️  Invia a GDrive Inbox_PD (Procedura Documentale)")
            elif already_uploaded:
                btn_gdrive.state(["disabled"])
                btn_gdrive.configure(text="✅  Già caricato su GDrive Inbox_PD")
            else:
                btn_gdrive.state(["disabled"])

            # Spedizione: abilita solo se DDT + non già inviato
            already_shipped = st.get("spedizione_creata", False)
            if is_ddt and not already_shipped:
                btn_spedizione.state(["!disabled"])
                btn_spedizione.configure(text="📦  Nuova Spedizione")
            elif already_shipped:
                btn_spedizione.state(["disabled"])
                btn_spedizione.configure(text="✅  Spedizione già creata")
            else:
                btn_spedizione.state(["disabled"])

        tree.bind("<<TreeviewSelect>>", refresh_buttons)
        refresh_buttons()

        win.protocol("WM_DELETE_WINDOW", win.withdraw)

    def _collect_last_docs(self, limit: int = 5) -> list[ParsedDoc]:
        pdfs = []
        if not os.path.isdir(MEXAL_TEMP):
            return []

        for root_dir, _, files in os.walk(MEXAL_TEMP):
            for fn in files:
                if fn.lower().endswith(".pdf"):
                    full_path = os.path.join(root_dir, fn)
                    mtime = _safe_get_mtime(full_path)
                    if mtime is None:
                        continue
                    pdfs.append((full_path, mtime))

        pdfs.sort(key=lambda x: x[1], reverse=True)

        docs: list[ParsedDoc] = []
        for path, _ in pdfs:
            doc = parse_mexal_pdf(path)
            if doc:
                docs.append(doc)
            if len(docs) >= limit:
                break

        # Forza l'ultimo documento rilevato in cima (deve sempre apparire nel modale)
        current = self._current_doc or (self._last_detected[0] if self._last_detected else None)
        if current:
            current_id = self._doc_id(current)
            docs = [d for d in docs if self._doc_id(d) != current_id]
            docs.insert(0, current)
            docs = docs[:limit]

        return docs

    def _find_doc_by_id(self, doc_id: str) -> Optional[ParsedDoc]:
        docs = self._collect_last_docs(limit=10)
        for d in docs:
            if self._doc_id(d) == doc_id:
                return d
        return None

    def _do_save(self, doc: "ParsedDoc", doc_id: str) -> Optional[str]:
        """Salva il documento in locale. Ritorna il percorso di destinazione o None in caso di errore."""
        st = self._get_doc_state(doc_id)
        if st.get("saved"):
            return st.get("dest_path", "")

        save_dir = self._preferred_save_dir(doc)
        os.makedirs(save_dir, exist_ok=True)

        numero = (doc.doc_number or "").strip()
        intestatario = (doc.recipient or "").strip()
        data_doc = (doc.doc_date or "").strip()

        if doc.doc_code == "DDT":
            parts = [numero, intestatario, data_doc]
        elif doc.doc_code in {"PC", "FC"}:
            parts = [numero, intestatario]
        else:
            tipo = (doc.doc_code or "?").strip()
            parts = [tipo, numero, intestatario, data_doc]

        parts = [p for p in parts if p] or ["Documento"]
        filename = re.sub(r"[\\/:*?\"<>|]", "-", " ".join(parts)) + ".pdf"
        filename = re.sub(r"\s+", " ", filename)
        dest_path = os.path.join(save_dir, filename)

        try:
            if doc.doc_type.lower() == "fattura":
                save_first_page_only(doc.source_path, dest_path)
            else:
                shutil.copy2(doc.source_path, dest_path)
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante il salvataggio:\n{e}")
            return None

        st["saved"] = True
        st["dest_path"] = dest_path
        st["meta"] = {
            "doc_code": doc.doc_code,
            "doc_type": doc.doc_type,
            "doc_number": doc.doc_number,
            "recipient": doc.recipient,
            "doc_date": doc.doc_date,
            "source_path": doc.source_path,
            "created_at": doc.created_at,
        }
        _save_json(STATE_FILE, self.state)
        return dest_path

    def _dialog_ddt_azione(self, doc: "ParsedDoc", doc_id: str, is_espositore: bool, is_vettore: bool) -> None:
        """Dopo il salvataggio di un DDT, gestisce i 3 casi:
        - non-espositore → Spedizione default
        - espositore + vettore → Procedura default
        - espositore + non vettore → chiede stampa → Procedura
        """
        # Caso 3: espositore consegnato direttamente (non vettore) → stampa + procedura
        if is_espositore and not is_vettore:
            self._dialog_espositore_diretto(doc, doc_id)
            return

        gdrive_ok = (
            _GDRIVE_AVAILABLE
            and bool(GDRIVE_INBOX_PD_FOLDER_ID)
            and bool(os.environ.get("GOOGLE_CLIENT_ID"))
            and bool(os.environ.get("GOOGLE_REFRESH_TOKEN"))
        )
        spedizioni_url = os.environ.get("SPEDIZIONI_API_URL", "http://localhost:8000")

        dlg = tk.Toplevel(self.root)
        dlg.title("DDT salvato — Prossimo passo")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)

        hdr = ttk.Frame(dlg, bootstyle="success", padding=(16, 10))
        hdr.grid(row=0, column=0, sticky="ew")
        ttk.Label(
            hdr,
            text="✅  DDT salvato in locale",
            font=("Segoe UI", 11, "bold"),
            bootstyle="inverse-success",
        ).grid(row=0, column=0, sticky="w")

        body = ttk.Frame(dlg, padding=(20, 14, 20, 6))
        body.grid(row=1, column=0, sticky="nsew")

        # Caso 1: non-espositore → default Spedizione
        # Caso 2: espositore + vettore → default Procedura
        if is_espositore:
            hint = "Rilevato espositore — consigliata Procedura Documentale"
        else:
            hint = "DDT standard — consigliata Nuova Spedizione"
        ttk.Label(body, text=hint, font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", pady=(0, 4))

        addr_parts = [p for p in [doc.dest_cap, doc.dest_citta, doc.dest_provincia] if p]
        if addr_parts:
            addr_str = f"{doc.dest_cap} {doc.dest_citta} ({doc.dest_provincia})"
            ttk.Label(body, text=addr_str, font=("Segoe UI", 9), bootstyle="secondary").grid(
                row=1, column=0, sticky="w", pady=(0, 8)
            )

        btn_row = ttk.Frame(dlg, padding=(20, 8, 20, 4))
        btn_row.grid(row=2, column=0, sticky="ew")
        btn_row.columnconfigure(0, weight=1)
        btn_row.columnconfigure(1, weight=1)

        def do_procedura():
            dlg.destroy()
            self._do_save(doc, doc_id)
            if not gdrive_ok:
                messagebox.showerror("GDrive", "Credenziali Google non configurate.")
                return
            self._do_gdrive_upload(doc, doc_id)

        def do_spedizione():
            dlg.destroy()
            self._do_save(doc, doc_id)
            self._do_spedizione(doc, doc_id, spedizioni_url)

        proc_style = "primary" if is_espositore else "primary-outline"
        sped_style = "success-outline" if is_espositore else "success"

        ttk.Button(
            btn_row, text="☁️  Procedura Documentale",
            bootstyle=proc_style, command=do_procedura, width=22,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(
            btn_row, text="📦  Nuova Spedizione",
            bootstyle=sped_style, command=do_spedizione, width=22,
        ).grid(row=0, column=1, sticky="ew")

        btn_row2 = ttk.Frame(dlg, padding=(20, 0, 20, 12))
        btn_row2.grid(row=3, column=0, sticky="ew")
        btn_row2.columnconfigure(0, weight=1)
        btn_row2.columnconfigure(1, weight=1)

        def do_stampa():
            copies = self._ask_copies()
            if copies is None:
                return
            st_doc = self._get_doc_state(doc_id)
            dest = st_doc.get("dest_path", "")
            if not dest:
                messagebox.showerror("Errore", "Documento non ancora salvato.")
                return
            try:
                print_pdf(dest, copies=copies)
                st_doc["printed"] = True
                _save_json(STATE_FILE, self.state)
            except Exception as exc:
                messagebox.showerror("Errore stampa", str(exc))

        def do_email():
            st_doc = self._get_doc_state(doc_id)
            dest = st_doc.get("dest_path", "")
            if not dest:
                messagebox.showerror("Errore", "Documento non ancora salvato.")
                return
            fields = self._ask_email(doc)
            if not fields:
                return
            to_addr = fields.get("to", "").strip()
            subject = fields.get("subject", "").strip()
            body_text = fields.get("body", "").strip()
            to_addrs = [a.strip() for a in re.split(r"[;,\s]+", to_addr) if a.strip()]
            if not to_addrs:
                messagebox.showwarning("Attenzione", "Inserisci un destinatario valido.")
                return
            cfg = _smtp_config()
            host = str(cfg.get("host") or "").strip()
            port = int(cfg.get("port") or 0)
            user = str(cfg.get("user") or "").strip()
            password = str(cfg.get("password") or "").strip()
            from_addr = str(cfg.get("from_addr") or "").strip()
            if not host or not port or not user or not password:
                ok = self._smtp_settings_wizard()
                if not ok:
                    return
                cfg = _smtp_config()
                host = str(cfg.get("host") or "").strip()
                port = int(cfg.get("port") or 0)
                user = str(cfg.get("user") or "").strip()
                password = str(cfg.get("password") or "").strip()
                from_addr = str(cfg.get("from_addr") or "").strip()
            try:
                send_email_smtp(
                    host=host, port=port, user=user, password=password,
                    from_addr=from_addr, to_addrs=to_addrs,
                    subject=subject, body=body_text,
                    attachment_path=os.path.abspath(dest),
                )
                st_doc["emailed"] = True
                _save_json(STATE_FILE, self.state)
                messagebox.showinfo("Email", "Email inviata.")
            except Exception as exc:
                messagebox.showerror("Errore email", str(exc))

        ttk.Button(
            btn_row2, text="🖨️  Stampa",
            bootstyle="warning-outline", command=do_stampa, width=22,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(
            btn_row2, text="📧  Invia Email",
            bootstyle="info-outline", command=do_email, width=22,
        ).grid(row=0, column=1, sticky="ew")

        ttk.Button(
            dlg, text="Chiudi",
            bootstyle="secondary-link",
            command=dlg.destroy,
        ).grid(row=4, column=0, pady=(0, 8))

        dlg.update_idletasks()
        w = max(dlg.winfo_reqwidth(), 420)
        h = dlg.winfo_reqheight()
        x = int((dlg.winfo_screenwidth() - w) / 2)
        y = int((dlg.winfo_screenheight() - h) / 2)
        dlg.geometry(f"{w}x{h}+{x}+{y}")

    def _dialog_espositore_diretto(self, doc: "ParsedDoc", doc_id: str) -> None:
        """Espositore consegnato direttamente (non vettore): chiede stampa DDT poi avvia Procedura."""
        gdrive_ok = (
            _GDRIVE_AVAILABLE
            and bool(GDRIVE_INBOX_PD_FOLDER_ID)
            and bool(os.environ.get("GOOGLE_CLIENT_ID"))
            and bool(os.environ.get("GOOGLE_REFRESH_TOKEN"))
        )

        dlg = tk.Toplevel(self.root)
        dlg.title("Espositore — consegna diretta")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)

        hdr = ttk.Frame(dlg, bootstyle="warning", padding=(16, 10))
        hdr.grid(row=0, column=0, sticky="ew")
        ttk.Label(
            hdr,
            text="🏪  Espositore — consegna diretta",
            font=("Segoe UI", 11, "bold"),
            bootstyle="inverse-warning",
        ).grid(row=0, column=0, sticky="w")

        body = ttk.Frame(dlg, padding=(20, 14, 20, 6))
        body.grid(row=1, column=0, sticky="nsew")
        ttk.Label(
            body,
            text="Trasporto non a vettore.\nVuoi stampare il DDT prima di avviare la Procedura Documentale?",
            font=("Segoe UI", 10),
            justify="left",
        ).grid(row=0, column=0, sticky="w", pady=(0, 12))

        btn_row = ttk.Frame(dlg, padding=(20, 8, 20, 16))
        btn_row.grid(row=2, column=0, sticky="ew")
        btn_row.columnconfigure(0, weight=1)
        btn_row.columnconfigure(1, weight=1)

        def do_stampa_poi_procedura():
            dlg.destroy()
            copies = self._ask_copies()
            if copies is not None:
                st = self._get_doc_state(doc_id)
                dest = st.get("dest_path") or doc.source_path
                try:
                    print_pdf(dest, copies=copies)
                    st["printed"] = True
                    _save_json(STATE_FILE, self.state)
                except Exception as e:
                    messagebox.showerror("Errore stampa", str(e))
            if gdrive_ok:
                self._do_gdrive_upload(doc, doc_id)
            else:
                messagebox.showerror("GDrive", "Credenziali Google non configurate.")

        def do_solo_procedura():
            dlg.destroy()
            if gdrive_ok:
                self._do_gdrive_upload(doc, doc_id)
            else:
                messagebox.showerror("GDrive", "Credenziali Google non configurate.")

        ttk.Button(
            btn_row, text="🖨️  Stampa + Procedura",
            bootstyle="warning", command=do_stampa_poi_procedura, width=22,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(
            btn_row, text="☁️  Solo Procedura",
            bootstyle="primary-outline", command=do_solo_procedura, width=22,
        ).grid(row=0, column=1, sticky="ew")

        ttk.Button(
            dlg, text="Solo salvataggio — chiudi",
            bootstyle="secondary-link",
            command=dlg.destroy,
        ).grid(row=3, column=0, pady=(0, 8))

        dlg.update_idletasks()
        w = max(dlg.winfo_reqwidth(), 420)
        h = dlg.winfo_reqheight()
        x = int((dlg.winfo_screenwidth() - w) / 2)
        y = int((dlg.winfo_screenheight() - h) / 2)
        dlg.geometry(f"{w}x{h}+{x}+{y}")

    def _do_gdrive_upload(self, doc: "ParsedDoc", doc_id: str) -> None:
        numero = (doc.doc_number or "").strip()
        dest = (doc.recipient or "").strip()
        data = (doc.doc_date or "").strip()
        parts = [p for p in [numero, dest, data] if p] or ["DDT"]
        filename = re.sub(r"[\\/:*?\"<>|]", "-", " ".join(parts)) + ".pdf"
        filename = re.sub(r"\s+", " ", filename)
        try:
            file_id = _gdrive_upload_to_inbox_pd(doc.source_path, filename)
            _log(f"GDrive upload OK: {filename} → id={file_id}")
        except Exception as e:
            _log(f"GDrive upload FAIL: {e}")
            messagebox.showerror("Errore GDrive", str(e))
            return
        st = self._get_doc_state(doc_id)
        st["gdrive_uploaded"] = True
        _save_json(STATE_FILE, self.state)
        self._dialog_dopo_gdrive(filename)

    def _do_gdrive_upload_bg(self, doc: "ParsedDoc", doc_id: str) -> None:
        """Upload GDrive in background thread — nessuna messagebox, safe da thread."""
        st = self._get_doc_state(doc_id)
        if st.get("gdrive_uploaded"):
            return
        dest_path = st.get("dest_path") or doc.source_path
        if not dest_path or not os.path.isfile(dest_path):
            _log(f"GDrive auto-upload: file non trovato: {dest_path}")
            return
        numero   = (doc.doc_number or "").strip()
        dest     = (doc.recipient  or "").strip()
        data_doc = (doc.doc_date   or "").strip()
        parts    = [p for p in [numero, dest, data_doc] if p] or ["DDT"]
        filename = re.sub(r"[\\/:*?\"<>|]", "-", " ".join(parts)) + ".pdf"
        filename = re.sub(r"\s+", " ", filename)
        try:
            file_id = _gdrive_upload_to_inbox_pd(dest_path, filename)
            _log(f"GDrive auto-upload OK: {filename} → id={file_id}")
            st["gdrive_uploaded"] = True
            _save_json(STATE_FILE, self.state)
            _tg_send_simple(
                f"☁️ GDrive Inbox\\_PD ✅\n"
                f"{doc.doc_type} {doc.doc_number} — {dest[:40]}"
            )
        except Exception as e:
            _log(f"GDrive auto-upload FAIL: {e}")
            _tg_send_simple(
                f"⚠️ GDrive upload fallito\n"
                f"{doc.doc_type} {doc.doc_number}: {str(e)[:80]}"
            )

    def _upsert_destinatario_bg(self, doc: "ParsedDoc") -> None:
        """UPSERT anagrafica destinatario nel DB Spedizioni — background thread."""
        if not doc.recipient or doc.recipient == "(Destinatario non trovato)":
            return
        spedizioni_url = os.environ.get("SPEDIZIONI_API_URL", "http://localhost:8000")
        payload: dict = {"nome": doc.recipient}
        if doc.dest_cap:       payload["cap"]       = doc.dest_cap
        if doc.dest_citta:     payload["citta"]     = doc.dest_citta
        if doc.dest_provincia: payload["provincia"] = doc.dest_provincia
        if doc.dest_email:     payload["email"]     = doc.dest_email
        if doc.dest_tel:       payload["telefono"]  = doc.dest_tel
        try:
            data = json.dumps(payload).encode()
            req = urllib.request.Request(
                f"{spedizioni_url}/api/destinatari/upsert",
                data=data,
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                result = json.loads(resp.read().decode())
            action  = result.get("action", "?")
            dest_id = result.get("destinatario_id", "?")
            _log(f"Spedizioni UPSERT: {action} id={dest_id} nome={doc.recipient[:40]}")
        except Exception as e:
            _log(f"Spedizioni UPSERT FAIL: {e}")

    def _do_spedizione(self, doc: "ParsedDoc", doc_id: str, spedizioni_url: str) -> None:
        try:
            boundary = "----FormBoundary7MA4YWxkTrZu0gW"
            with open(doc.source_path, "rb") as f:
                pdf_data = f.read()
            filename = os.path.basename(doc.source_path)
            body = (
                f"--{boundary}\r\n"
                f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
                f"Content-Type: application/pdf\r\n\r\n"
            ).encode() + pdf_data + f"\r\n--{boundary}--\r\n".encode()
            req = urllib.request.Request(
                f"{spedizioni_url}/api/spedizioni/da-ddt?draft=true",
                data=body,
                headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=15) as resp:
                result = json.loads(resp.read().decode())
            spedizione_id = result.get("id") or result.get("spedizione_id") or result.get("data", {}).get("id")
            _log(f"Spedizione bozza creata: id={spedizione_id}")
        except Exception as e:
            _log(f"Spedizione FAIL: {e}")
            messagebox.showerror("Errore Spedizione", str(e))
            return
        st = self._get_doc_state(doc_id)
        st["spedizione_creata"] = True
        _save_json(STATE_FILE, self.state)
        if spedizione_id:
            webbrowser.open(f"{spedizioni_url}/?spedizione={spedizione_id}")
        else:
            webbrowser.open(spedizioni_url)
        messagebox.showinfo("Spedizione creata", "✅ Bozza spedizione creata.\n\nIl browser si è aperto per completare i dettagli.")

    def _action_save(self, tree: ttk.Treeview):
        sel = tree.selection()
        if not sel:
            return
        doc_id = sel[0]
        doc = self._find_doc_by_id(doc_id)
        if not doc:
            messagebox.showerror("Errore", "Documento non trovato.")
            return
        dest_path = self._do_save(doc, doc_id)
        if dest_path:
            messagebox.showinfo("Completato", f"Documento salvato in:\n{dest_path}")
            tree.event_generate("<<TreeviewSelect>>")

    def _action_print(self, tree: ttk.Treeview):
        sel = tree.selection()
        if not sel:
            return
        doc_id = sel[0]
        st = self._get_doc_state(doc_id)
        if not st.get("saved") or not st.get("dest_path"):
            return

        copies = self._ask_copies()
        if copies is None:
            return

        try:
            print_pdf(st["dest_path"], copies=copies)
        except Exception as e:
            messagebox.showerror("Errore", f"Errore stampa:\n{e}")
            return

        st["printed"] = True
        _save_json(STATE_FILE, self.state)

    def _ask_copies(self) -> Optional[int]:
        dlg = tk.Toplevel(self.root)
        dlg.title("Stampa")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Numero copie:").grid(row=0, column=0, sticky="w")
        var = tk.StringVar(value="1")
        entry = ttk.Entry(frm, textvariable=var, width=6)
        entry.grid(row=0, column=1, sticky="w", padx=(8, 0))
        entry.focus_set()

        result: dict[str, Optional[int]] = {"value": None}

        def ok():
            try:
                n = int(var.get().strip())
                if n < 1:
                    raise ValueError
            except Exception:
                messagebox.showwarning("Attenzione", "Inserisci un numero valido (>= 1).")
                return
            result["value"] = n
            dlg.destroy()

        def cancel():
            dlg.destroy()

        ttk.Button(frm, text="Annulla", command=cancel).grid(row=1, column=0, pady=(10, 0), sticky="ew", padx=(0, 8))
        ttk.Button(frm, text="OK", command=ok).grid(row=1, column=1, pady=(10, 0), sticky="ew")

        dlg.grab_set()
        self.root.wait_window(dlg)
        return result["value"]

    def _smtp_settings_wizard(self) -> bool:
        cfg = _smtp_config()

        dlg = tk.Toplevel(self.root)
        dlg.title("Impostazioni SMTP")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")

        host_var = tk.StringVar(value=str(cfg.get("host") or "smtp.gmail.com"))
        port_var = tk.StringVar(value=str(cfg.get("port") or 465))
        user_var = tk.StringVar(value=str(cfg.get("user") or ""))
        pass_var = tk.StringVar(value=str(cfg.get("password") or ""))
        from_var = tk.StringVar(value=str(cfg.get("from_addr") or ""))

        ttk.Label(frm, text="Host:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=host_var, width=38).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ttk.Label(frm, text="Porta:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frm, textvariable=port_var, width=10).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(6, 0))

        ttk.Label(frm, text="Utente (SMTP_USER):").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frm, textvariable=user_var, width=38).grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=(6, 0))

        ttk.Label(frm, text="Password App (SMTP_PASS):").grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frm, textvariable=pass_var, width=38, show="*").grid(row=3, column=1, sticky="ew", padx=(8, 0), pady=(6, 0))

        ttk.Label(frm, text="Mittente (opzionale SMTP_FROM):").grid(row=4, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frm, textvariable=from_var, width=38).grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=(6, 0))

        note = f"Salvataggio in: {DOTENV_PATH}"
        ttk.Label(frm, text=note).grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 0))

        result: dict[str, bool] = {"ok": False}

        def save():
            host = host_var.get().strip()
            port_s = port_var.get().strip()
            user = user_var.get().strip()
            password = pass_var.get().strip()
            from_addr = from_var.get().strip()

            try:
                port = int(port_s)
            except Exception:
                messagebox.showwarning("Attenzione", "Porta non valida.")
                return

            if not host or not port or not user or not password:
                messagebox.showwarning("Attenzione", "Compila host/porta/utente/password.")
                return

            try:
                with open(DOTENV_PATH, "w", encoding="utf-8") as f:
                    f.write(f"SMTP_HOST={host}\n")
                    f.write(f"SMTP_PORT={port}\n")
                    f.write(f"SMTP_USER={user}\n")
                    f.write(f"SMTP_PASS={password}\n")
                    if from_addr:
                        f.write(f"SMTP_FROM={from_addr}\n")
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile salvare .env:\n{e}")
                return

            os.environ["SMTP_HOST"] = host
            os.environ["SMTP_PORT"] = str(port)
            os.environ["SMTP_USER"] = user
            os.environ["SMTP_PASS"] = password
            if from_addr:
                os.environ["SMTP_FROM"] = from_addr

            result["ok"] = True
            dlg.destroy()

        def cancel():
            dlg.destroy()

        ttk.Button(frm, text="Annulla", command=cancel).grid(row=6, column=0, pady=(10, 0), sticky="ew", padx=(0, 8))
        ttk.Button(frm, text="Salva", command=save).grid(row=6, column=1, pady=(10, 0), sticky="ew")

        dlg.grab_set()
        dlg.focus_force()
        self.root.wait_window(dlg)
        return result["ok"]

    def _ask_email(self, doc: ParsedDoc) -> Optional[dict[str, str]]:
        dlg = tk.Toplevel(self.root)
        dlg.title("Email")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="A:").grid(row=0, column=0, sticky="w")
        to_var = tk.StringVar(value="")
        to_entry = ttk.Entry(frm, textvariable=to_var, width=42)
        to_entry.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        subj_default = f"{(doc.doc_code or '').strip()} {(doc.doc_number or '').strip()} {(doc.recipient or '').strip()}".strip()
        ttk.Label(frm, text="Oggetto:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        subj_var = tk.StringVar(value=subj_default)
        subj_entry = ttk.Entry(frm, textvariable=subj_var, width=42)
        subj_entry.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        ttk.Label(frm, text="Testo:").grid(row=2, column=0, sticky="nw", pady=(8, 0))
        body = tk.Text(frm, width=42, height=6)
        body.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        recip = (doc.recipient or "").strip()
        default_body = (
            f"Buongiorno {recip},\n"
            "in allegato troverà il documento in oggetto.\n\n"
            "Cordiali saluti\n"
            "Fior d'Acqua Team"
        ).strip()
        body.insert("1.0", default_body)

        result: dict[str, Optional[dict[str, str]]] = {"value": None}

        def ok():
            result["value"] = {
                "to": to_var.get().strip(),
                "subject": subj_var.get().strip(),
                "body": body.get("1.0", "end").strip(),
            }
            dlg.destroy()

        def cancel():
            dlg.destroy()

        ttk.Button(frm, text="Annulla", command=cancel).grid(row=3, column=0, pady=(10, 0), sticky="ew", padx=(0, 8))
        ttk.Button(frm, text="OK", command=ok).grid(row=3, column=1, pady=(10, 0), sticky="ew")

        dlg.grab_set()
        to_entry.focus_set()
        self.root.wait_window(dlg)
        return result["value"]

    def _action_email(self, tree: ttk.Treeview):
        sel = tree.selection()
        if not sel:
            return
        doc_id = sel[0]
        st = self._get_doc_state(doc_id)
        if not st.get("saved") or not st.get("dest_path"):
            return

        doc = self._find_doc_by_id(doc_id)
        if not doc:
            messagebox.showerror("Errore", "Documento non trovato.")
            return

        fields = self._ask_email(doc)
        if not fields:
            return

        to_addr = fields.get("to", "").strip()
        subject = fields.get("subject", "").strip()
        body = fields.get("body", "").strip()
        attachment_path = st["dest_path"]

        to_addrs = [a.strip() for a in re.split(r"[;,\s]+", to_addr) if a.strip()]
        if not to_addrs:
            messagebox.showwarning("Attenzione", "Inserisci almeno un destinatario valido (campo A:).")
            return

        cfg = _smtp_config()
        host = str(cfg.get("host") or "").strip()
        port = int(cfg.get("port") or 0)
        user = str(cfg.get("user") or "").strip()
        password = str(cfg.get("password") or "").strip()
        from_addr = str(cfg.get("from_addr") or "").strip()

        if not host or not port or not user or not password or not from_addr:
            ok = self._smtp_settings_wizard()
            if not ok:
                return

            cfg = _smtp_config()
            host = str(cfg.get("host") or "").strip()
            port = int(cfg.get("port") or 0)
            user = str(cfg.get("user") or "").strip()
            password = str(cfg.get("password") or "").strip()
            from_addr = str(cfg.get("from_addr") or "").strip()

            if not host or not port or not user or not password or not from_addr:
                messagebox.showerror("Errore", "Configurazione SMTP non valida.")
                return

        try:
            send_email_smtp(
                host=host,
                port=port,
                user=user,
                password=password,
                from_addr=from_addr,
                to_addrs=to_addrs,
                subject=subject,
                body=body,
                attachment_path=os.path.abspath(attachment_path),
            )
        except Exception as e:
            _log(f"SMTP send failed: {e}")
            messagebox.showerror("Errore", f"Invio email fallito:\n{e}")
            return

        st["emailed"] = True
        _save_json(STATE_FILE, self.state)
        tree.event_generate("<<TreeviewSelect>>")
        messagebox.showinfo("Email", "Email inviata.")

    def _dialog_dopo_gdrive(self, filename: str) -> None:
        dlg = tk.Toplevel(self.root)
        dlg.title("GDrive — Caricamento completato")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)

        # Header
        hdr = ttk.Frame(dlg, bootstyle="success", padding=(16, 10))
        hdr.grid(row=0, column=0, sticky="ew")
        ttk.Label(
            hdr,
            text="☁️  DDT caricato su GDrive Inbox_PD",
            font=("Segoe UI", 11, "bold"),
            bootstyle="inverse-success",
        ).grid(row=0, column=0, sticky="w")

        # Body
        body = ttk.Frame(dlg, padding=(20, 14, 20, 6))
        body.grid(row=1, column=0, sticky="nsew")

        ttk.Label(
            body,
            text=filename,
            font=("Segoe UI", 9),
            bootstyle="secondary",
        ).grid(row=0, column=0, sticky="w")

        ttk.Separator(body).grid(row=1, column=0, sticky="ew", pady=(10, 8))

        ttk.Label(
            body,
            text="Quando hai contattato il cliente e sei pronto,\navvia la Procedura Documentale dal bot Telegram.",
            font=("Segoe UI", 10),
            justify="left",
        ).grid(row=2, column=0, sticky="w")

        # Bottoni
        btn_row = ttk.Frame(dlg, padding=(20, 10, 20, 16))
        btn_row.grid(row=2, column=0, sticky="ew")
        btn_row.columnconfigure(0, weight=1)
        btn_row.columnconfigure(1, weight=1)

        def _apri_telegram():
            import urllib.parse
            import webbrowser
            testo = urllib.parse.quote("nuova procedura")
            webbrowser.open(f"tg://resolve?domain={TELEGRAM_BOT_USERNAME}&text={testo}")
            dlg.destroy()

        ttk.Button(
            btn_row, text="Lo faccio dopo", bootstyle="secondary-outline",
            command=dlg.destroy, width=16,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(
            btn_row, text="📱  Avvia ora su Telegram", bootstyle="success",
            command=_apri_telegram, width=24,
        ).grid(row=0, column=1, sticky="ew")

        dlg.update_idletasks()
        w = max(dlg.winfo_reqwidth(), 400)
        h = dlg.winfo_reqheight()
        x = int((dlg.winfo_screenwidth() - w) / 2)
        y = int((dlg.winfo_screenheight() - h) / 2)
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.grab_set()

    def _action_view(self, tree: ttk.Treeview):
        sel = tree.selection()
        if not sel:
            return
        doc_id = sel[0]
        st = self._get_doc_state(doc_id)
        if not st.get("saved") or not st.get("dest_path"):
            return
        try:
            os.startfile(st["dest_path"])
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire il file:\n{e}")

    def _action_gdrive_pd(self, tree: ttk.Treeview):
        sel = tree.selection()
        if not sel:
            return
        doc_id = sel[0]
        doc = self._find_doc_by_id(doc_id)
        if not doc:
            messagebox.showerror("Errore", "Documento non trovato.")
            return
        self._do_save(doc, doc_id)
        self._do_gdrive_upload(doc, doc_id)
        tree.event_generate("<<TreeviewSelect>>")

    def _action_nuova_spedizione(self, tree: ttk.Treeview):
        sel = tree.selection()
        if not sel:
            return
        doc_id = sel[0]
        doc = self._find_doc_by_id(doc_id)
        if not doc:
            messagebox.showerror("Errore", "Documento non trovato.")
            return
        self._do_save(doc, doc_id)
        spedizioni_url = os.environ.get("SPEDIZIONI_API_URL", "http://localhost:8000")
        self._do_spedizione(doc, doc_id, spedizioni_url)
        tree.event_generate("<<TreeviewSelect>>")


def _is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _relaunch_as_admin() -> None:
    """Rilancia lo script corrente con UAC (ShellExecute runas)."""
    script = os.path.abspath(__file__)
    args   = " ".join(f'"{a}"' for a in sys.argv[1:])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {args}', None, 1)


if __name__ == "__main__":
    if "--install-startup" in sys.argv:
        if not _is_admin():
            _relaunch_as_admin()
            raise SystemExit(0)
        install_windows_shortcuts()
        print("OK: task schedulati (MexalDaemon + MexalWatchdog) e collegamento Desktop creati.")
        raise SystemExit(0)

    if "--uninstall-startup" in sys.argv:
        if not _is_admin():
            _relaunch_as_admin()
            raise SystemExit(0)
        uninstall_windows_shortcuts()
        print("OK: task rimossi e collegamento Desktop eliminato.")
        raise SystemExit(0)

    _ensure_single_instance("MexalAutomationDaemon")

    root = ttk.Window(themename="flatly")
    MexalDaemonApp(root)
    root.mainloop()
