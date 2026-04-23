"""
Watchdog per MexalAutomationDaemon.
Viene avviato al logon da Task Scheduler (FdA\MexalWatchdog).
Controlla ogni 5 minuti se il daemon è in esecuzione e lo rilancia se necessario.
"""
import ctypes
import os
import subprocess
import sys
import time

DAEMON_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mexal_daemon.py")
MUTEX_NAME    = "MexalAutomationDaemon"
INTERVAL_SEC  = 300  # 5 minuti


def _daemon_is_running() -> bool:
    SYNCHRONIZE = 0x00100000
    h = ctypes.windll.kernel32.OpenMutexW(SYNCHRONIZE, False, MUTEX_NAME)
    if h:
        ctypes.windll.kernel32.CloseHandle(h)
        return True
    return False


def _launch_daemon() -> None:
    py = sys.executable
    pyw = py[:-10] + "pythonw.exe" if py.lower().endswith("python.exe") else py
    workdir = os.path.dirname(DAEMON_SCRIPT)
    subprocess.Popen([pyw, DAEMON_SCRIPT], cwd=workdir)


# Attendi 90 secondi al primo avvio (Windows si sta ancora inizializzando)
time.sleep(90)

while True:
    try:
        if not _daemon_is_running():
            _launch_daemon()
    except Exception:
        pass
    time.sleep(INTERVAL_SEC)
