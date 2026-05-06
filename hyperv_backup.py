import sys
import os
import json
import re
import base64
import subprocess
import threading
import ctypes
import winreg
import shutil
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

import schedule

# qtawesome импортируется ПОСЛЕ создания QApplication (см. точку входа)
# Здесь только объявляем заглушки
QTA_AVAILABLE = False
qta = None

def _try_load_qta():
    """Вызвать после QApplication.__init__()"""
    global QTA_AVAILABLE, qta
    if QTA_AVAILABLE:
        return True
    try:
        import qtawesome as _qta
        qta = _qta
        QTA_AVAILABLE = True
        return True
    except ImportError:
        return False

def qicon(name: str, color: str = "#f0f2f5", size: int = 16) -> "QIcon":
    """Возвращает QIcon из qtawesome. Вызывать только после QApplication!"""
    if QTA_AVAILABLE and qta is not None:
        try:
            return qta.icon(name, color=color, scale_factor=1.0)
        except Exception:
            pass
    return QIcon()

def qpixmap(name: str, color: str = "#f0f2f5", size: int = 18) -> "QPixmap":
    """Возвращает QPixmap из qtawesome. Вызывать только после QApplication!"""
    if QTA_AVAILABLE and qta is not None:
        try:
            return qta.icon(name, color=color).pixmap(size, size)
        except Exception:
            pass
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    return pm
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QLineEdit, QSpinBox, QComboBox,
    QCheckBox, QTableWidget, QTableWidgetItem, QHeaderView, QPlainTextEdit,
    QProgressBar, QFileDialog, QMessageBox, QDialog, QDialogButtonBox,
    QSystemTrayIcon, QMenu, QStackedWidget, QButtonGroup, QStyle,
    QFrame, QScrollArea, QSizePolicy, QTabWidget
)
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal, QProcess, QObject, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QColor, QFont, QIcon, QPixmap, QPainter, QLinearGradient, QBrush, QPalette

# ─── Константы ────────────────────────────────────────────────────────────────

APP_VERSION   = "5.0"
CONFIG_FILE   = "hyperv_backup_config.json"
SCHEDULE_FILE = "hyperv_backup_schedule.json"
LAST_BACKUP_FILE = "last_backup.json"
AUTOSTART_KEY  = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTOSTART_NAME = "HyperVBackupManager"

_VM_STATE_INT: dict[int, str] = {
    0: "Other", 1: "Running", 2: "Off", 3: "Stopping",
    4: "Saved",  6: "Paused",  9: "Starting", 10: "Reset",
    11: "Saving", 32768: "FastSaved", 32769: "FastSaving",
}
_VM_STATE_STR: dict[str, str] = {
    "Running": "Running", "Off": "Off", "Paused": "Paused",
    "Saved": "Saved", "Starting": "Starting", "Stopping": "Stopping",
    "Reset": "Stopping", "Saving": "Saving", "FastSaved": "Saved",
    "FastSaving": "Saving", "Other": "Other",
    "Работает": "Running", "Выключена": "Off", "Выключен": "Off",
    "Приостановлена": "Paused", "Сохранена": "Saved",
    "Запуск": "Starting", "Запускается": "Starting",
    "Остановка": "Stopping", "Останавливается": "Stopping",
    "Сохранение": "Saving", "Приостановка": "Paused",
    "Возобновление": "Starting",
}

def resolve_vm_state(raw) -> str:
    if isinstance(raw, int): return _VM_STATE_INT.get(raw, "Other")
    if isinstance(raw, str): return _VM_STATE_STR.get(raw.strip(), raw.strip())
    return "Other"

# ─── Тёмная тема ──────────────────────────────────────────────────────────────

DARK_BG       = "#4a4d5a"   # основной фон — средне-серый
DARK_SURFACE  = "#525664"   # поверхности панелей
DARK_CARD     = "#5a5e6e"   # карточки
DARK_BORDER   = "#6e7282"   # рамки
DARK_SIDEBAR  = "#666870"   # сайдбар — средне-серый
DARK_TEXT     = "#f0f2f5"   # основной текст
DARK_MUTED    = "#c8ccd8"   # второстепенный (светлее для читаемости на сером)
DARK_ACCENT   = "#6366f1"        # indigo
DARK_ACCENT2  = "#8b5cf6"        # violet
DARK_SUCCESS  = "#22c55e"
DARK_ERROR    = "#ef4444"
DARK_WARN     = "#f59e0b"
DARK_INFO     = "#38bdf8"

VM_STATES = {
    "Running":  {"text": "● Работает",   "bg": QColor("#14532d"), "fg": QColor("#86efac")},
    "Off":      {"text": "● Выключена",  "bg": QColor("#450a0a"), "fg": QColor("#fca5a5")},
    "Paused":   {"text": "⏸ Приостан.", "bg": QColor("#451a03"), "fg": QColor("#fcd34d")},
    "Saved":    {"text": "● Сохранена",  "bg": QColor("#1e3a5f"), "fg": QColor("#93c5fd")},
    "Starting": {"text": "● Запуск...",  "bg": QColor("#3b2900"), "fg": QColor("#fde68a")},
    "Stopping": {"text": "● Стоп...",    "bg": QColor("#3b1515"), "fg": QColor("#f87171")},
    "Saving":   {"text": "● Сохран...",  "bg": QColor("#0c3045"), "fg": QColor("#7dd3fc")},
    "Reset":    {"text": "↺ Сброс",      "bg": QColor("#3e1d7a"), "fg": QColor("#c4b5fd")},
    "Other":    {"text": "? Другое",     "bg": QColor("#1e293b"), "fg": QColor("#94a3b8")},
}
DEFAULT_STATE = {"text": "? Неизв.", "bg": QColor("#1e293b"), "fg": QColor("#94a3b8")}

VM_JOB_STATUS = {
    "pending":  {"text": "⏳ Ожидание",  "color": DARK_MUTED},
    "running":  {"text": "⚙️ Идёт...",   "color": DARK_ACCENT},
    "done":     {"text": "✅ Готово",     "color": DARK_SUCCESS},
    "error":    {"text": "❌ Ошибка",     "color": DARK_ERROR},
    "skipped":  {"text": "⏭ Пропущена", "color": DARK_MUTED},
}

LOG_COLORS = {
    "SUCCESS": DARK_SUCCESS,
    "ERROR":   DARK_ERROR,
    "WARNING": DARK_WARN,
    "INFO":    DARK_INFO,
    "SYSTEM":  DARK_MUTED,
}

STYLESHEET = f"""
QMainWindow, QDialog {{
    background: {DARK_BG};
}}
QWidget {{
    color: {DARK_TEXT};
    font-family: 'Segoe UI Variable', 'Segoe UI', sans-serif;
    font-size: 13px;
    background: transparent;
}}
QScrollArea, QScrollArea > QWidget > QWidget {{
    background: transparent;
    border: none;
}}

/* ── Таблицы ── */
QTableWidget {{
    background: {DARK_CARD};
    gridline-color: {DARK_BORDER};
    border: 1px solid {DARK_BORDER};
    border-radius: 8px;
    selection-background-color: #3730a3;
}}
QTableWidget::item {{
    padding: 6px 8px;
    border: none;
    color: {DARK_TEXT};
}}
QTableWidget::item:selected {{
    background: #3730a3;
    color: #c7d2fe;
}}
QHeaderView::section {{
    background: {DARK_SURFACE};
    color: {DARK_MUTED};
    padding: 8px;
    border: none;
    border-bottom: 1px solid {DARK_BORDER};
    font-weight: 600;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.8px;
}}
QTableWidget::corner-button {{ background: {DARK_SURFACE}; }}

/* ── Лог ── */
QPlainTextEdit {{
    background: {DARK_SURFACE};
    color: {DARK_TEXT};
    border: 1px solid {DARK_BORDER};
    border-radius: 6px;
    font-family: 'Cascadia Code', 'Consolas', 'Courier New', monospace;
    font-size: 12px;
    padding: 6px;
    selection-background-color: {DARK_ACCENT};
}}

/* ── Кнопки ── */
QPushButton {{
    background: {DARK_SURFACE};
    color: {DARK_TEXT};
    border: 1px solid {DARK_BORDER};
    padding: 8px 16px;
    border-radius: 7px;
    font-weight: 500;
    font-size: 13px;
}}
QPushButton:hover {{
    background: {DARK_CARD};
    border-color: {DARK_ACCENT};
    color: #a5b4fc;
}}
QPushButton:pressed {{
    background: #3730a3;
    border-color: {DARK_ACCENT};
}}
QPushButton:disabled {{
    background: {DARK_SURFACE};
    color: #374151;
    border-color: {DARK_BORDER};
}}
QPushButton#primary {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {DARK_ACCENT}, stop:1 {DARK_ACCENT2});
    color: #fff;
    border: none;
    font-weight: 600;
}}
QPushButton#primary:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4f46e5, stop:1 #7c3aed);
}}
QPushButton#primary:disabled {{
    background: #2e2b6b;
    color: #4c4799;
    border: none;
}}
QPushButton#danger {{
    background: #450a0a;
    color: #fca5a5;
    border: 1px solid #7f1d1d;
    font-weight: 600;
}}
QPushButton#danger:hover {{
    background: #7f1d1d;
    color: #fecaca;
}}
QPushButton#success {{
    background: #14532d;
    color: #86efac;
    border: 1px solid #166534;
    font-weight: 600;
}}
QPushButton#success:hover {{
    background: #166534;
}}

/* ── Поля ввода ── */
QLineEdit, QSpinBox {{
    background: {DARK_SURFACE};
    color: {DARK_TEXT};
    border: 1px solid {DARK_BORDER};
    border-radius: 6px;
    padding: 7px 10px;
    selection-background-color: {DARK_ACCENT};
}}
QLineEdit:focus, QSpinBox:focus {{
    border-color: {DARK_ACCENT};
    background: {DARK_CARD};
}}
QComboBox {{
    background: {DARK_SURFACE};
    color: {DARK_TEXT};
    border: 1px solid {DARK_BORDER};
    border-radius: 6px;
    padding: 7px 10px;
    min-width: 80px;
}}
QComboBox:focus {{ border-color: {DARK_ACCENT}; }}
QComboBox::drop-down {{ border: none; width: 28px; }}
QComboBox::down-arrow {{ color: {DARK_MUTED}; }}
QComboBox QAbstractItemView {{
    background: {DARK_CARD};
    color: {DARK_TEXT};
    border: 1px solid {DARK_BORDER};
    selection-background-color: #3730a3;
    selection-color: #c7d2fe;
    outline: none;
}}
QSpinBox::up-button, QSpinBox::down-button {{
    background: {DARK_BORDER};
    border: none;
    width: 20px;
    border-radius: 3px;
}}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {{
    background: {DARK_ACCENT};
}}

/* ── Чекбокс ── */
QCheckBox {{ spacing: 10px; color: {DARK_TEXT}; }}
QCheckBox::indicator {{
    width: 18px; height: 18px;
    border: 2px solid {DARK_BORDER};
    border-radius: 4px;
    background: {DARK_SURFACE};
}}
QCheckBox::indicator:checked {{
    background: {DARK_ACCENT};
    border-color: {DARK_ACCENT};
    image: none;
}}
QCheckBox::indicator:hover {{ border-color: {DARK_ACCENT}; }}

/* ── Прогресс-бар ── */
QProgressBar {{
    height: 8px;
    border: none;
    border-radius: 4px;
    background: {DARK_BORDER};
    text-align: center;
}}
QProgressBar::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {DARK_ACCENT}, stop:1 {DARK_ACCENT2});
    border-radius: 4px;
}}

/* ── Скроллбар ── */
QScrollBar:vertical {{
    background: {DARK_BG};
    width: 8px;
    border: none;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {DARK_BORDER};
    border-radius: 4px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {DARK_MUTED}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{
    background: {DARK_BG};
    height: 8px;
    border: none;
}}
QScrollBar::handle:horizontal {{
    background: {DARK_BORDER};
    border-radius: 4px;
    min-width: 24px;
}}
QScrollBar::handle:horizontal:hover {{ background: {DARK_MUTED}; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}

/* ── Сайдбар ── */
#sidebar {{
    background: {DARK_SIDEBAR};
    border-right: 1px solid {DARK_BORDER};
}}
#nav_btn {{
    background: transparent;
    border: none;
    padding: 11px 16px;
    margin: 2px 8px;
    border-radius: 8px;
    text-align: left;
    font-size: 13px;
    color: {DARK_MUTED};
    font-weight: 500;
}}
#nav_btn:hover {{
    background: #787a88;
    color: {DARK_TEXT};
}}
#nav_btn:checked {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4e4b8a, stop:1 #5a3f95);
    color: #e0e7ff;
    font-weight: 600;
    border-left: 3px solid {DARK_ACCENT};
}}

/* ── Карточки дашборда ── */
#dash_card {{
    background: {DARK_CARD};
    border: 1px solid {DARK_BORDER};
    border-radius: 12px;
}}
#dash_card:hover {{ border-color: #3730a3; }}
#card {{
    background: {DARK_CARD};
    border: 1px solid {DARK_BORDER};
    border-radius: 10px;
}}
#stat_value {{
    color: {DARK_TEXT};
    font-size: 26px;
    font-weight: 700;
    font-family: 'Segoe UI Variable Display', 'Segoe UI', sans-serif;
}}
#quick_btn {{
    background: {DARK_SURFACE};
    border: 1px solid {DARK_BORDER};
    border-radius: 8px;
    padding: 9px 16px;
    font-weight: 500;
    color: {DARK_TEXT};
}}
#quick_btn:hover {{
    background: {DARK_CARD};
    border-color: {DARK_ACCENT};
    color: #a5b4fc;
}}

/* ── Статус-бар ── */
QStatusBar {{
    background: {DARK_SURFACE};
    color: {DARK_MUTED};
    border-top: 1px solid {DARK_BORDER};
    font-size: 12px;
}}
QStatusBar::item {{ border: none; }}

/* ── Диалоги ── */
QMessageBox {{ background: {DARK_CARD}; }}
QMessageBox QLabel {{ color: {DARK_TEXT}; }}
QDialogButtonBox QPushButton {{ min-width: 80px; }}

/* ── Tabs ── */
QTabWidget::pane {{
    border: 1px solid {DARK_BORDER};
    border-radius: 8px;
    background: {DARK_CARD};
    top: -1px;
}}
QTabBar::tab {{
    background: {DARK_SIDEBAR};
    color: {DARK_TEXT};
    padding: 9px 22px;
    border: 1px solid {DARK_BORDER};
    border-bottom: none;
    border-radius: 6px 6px 0 0;
    font-weight: 500;
    margin-right: 3px;
}}
QTabBar::tab:selected {{
    background: {DARK_CARD};
    color: #c7d2fe;
    border-color: {DARK_BORDER};
    border-bottom: 2px solid {DARK_ACCENT};
    font-weight: 600;
}}
QTabBar::tab:hover:!selected {{ background: #787a88; color: {DARK_TEXT}; }}

/* ── Разделитель ── */
QFrame[frameShape="4"], QFrame[frameShape="5"] {{
    color: {DARK_BORDER};
    background: {DARK_BORDER};
}}
"""

# ─── Вспомогательные ──────────────────────────────────────────────────────────

def _safe_decode(data: bytes) -> str:
    if isinstance(data, str): return data
    for enc in ("utf-8", "cp1251", "cp866", "latin-1"):
        try: return data.decode(enc)
        except (UnicodeDecodeError, AttributeError): pass
    return data.decode("utf-8", errors="replace")


class LogEntry:
    __slots__ = ("ts", "msg", "level")
    def __init__(self, msg: str, level: str = "INFO"):
        self.ts = datetime.now().strftime("%H:%M:%S")
        self.msg = msg
        self.level = level


class LogWorker(QObject):
    log_signal = pyqtSignal(LogEntry)
    def __init__(self):
        super().__init__()
        self._queue: list[LogEntry] = []
        self._lock = threading.Lock()
        self._timer = QTimer()
        self._timer.timeout.connect(self._flush)
        self._timer.start(80)
    def add(self, msg: str, level: str = "INFO"):
        with self._lock: self._queue.append(LogEntry(msg, level))
    def _flush(self):
        with self._lock:
            items, self._queue = self._queue, []
        for e in items:
            self.log_signal.emit(e)


class VMFetcher(QThread):
    finished = pyqtSignal(list)
    error    = pyqtSignal(str)
    _PS = r"""
[Console]::OutputEncoding=[System.Text.Encoding]::UTF8
$ErrorActionPreference='Stop'
try {
    $vms=Get-VM|ForEach-Object{[PSCustomObject]@{
        Name=[string]$_.Name; State=[string]$_.State
        Id=[string]$_.Id; MemoryAssigned=[long]$_.MemoryAssigned
        ProcessorCount=[int]$_.ProcessorCount}}
    if($vms){$vms|ConvertTo-Json -Compress -Depth 2}else{Write-Output '[]'}
}catch{Write-Error "Ошибка: $($_.Exception.Message)";exit 1}
"""
    def run(self):
        try:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            r = subprocess.run(
                ["powershell","-NoProfile","-ExecutionPolicy","Bypass","-Command",self._PS],
                capture_output=True, timeout=20, startupinfo=si)
            out = _safe_decode(r.stdout)
            err = _safe_decode(r.stderr)
            if r.returncode != 0:
                msg = err.strip() or out.strip() or "Неизвестная ошибка"
                self.error.emit("Hyper-V недоступен" if "не установлен" in msg else f"PowerShell: {msg}")
                return
            raw = out.strip()
            if not raw or raw == "[]":
                self.finished.emit([]); return
            data = json.loads(raw)
            self.finished.emit([data] if isinstance(data, dict) else data)
        except subprocess.TimeoutExpired: self.error.emit("Таймаут >20 сек")
        except FileNotFoundError:         self.error.emit("PowerShell не найден")
        except json.JSONDecodeError as e: self.error.emit(f"JSON: {e}")
        except Exception as e:            self.error.emit(str(e))


class SizeCalcWorker(QThread):
    finished = pyqtSignal(int, str, int)   # row, label, bytes
    def __init__(self, path: str, row: int):
        super().__init__(); self.path = path; self.row = row
    def run(self):
        total = 0
        try:
            for root, _, files in os.walk(self.path):
                for f in files:
                    try: total += os.path.getsize(os.path.join(root, f))
                    except OSError: pass
        except Exception: pass
        if   total >= 1024**3: s = f"{total/1024**3:.2f} ГБ"
        elif total >= 1024**2: s = f"{total/1024**2:.0f} МБ"
        elif total > 0:        s = f"{total/1024:.0f} КБ"
        else:                  s = "0 Б"
        self.finished.emit(self.row, s, total)


# ─── SMTP Email Worker ────────────────────────────────────────────────────────

class EmailWorker(QThread):
    """Отправляет email в фоновом потоке, чтобы не блокировать UI."""
    result = pyqtSignal(bool, str)   # success, message

    def __init__(self, cfg: dict, subject: str, html_body: str,
                 attach_log: str | None = None):
        super().__init__()
        self.cfg = cfg
        self.subject = subject
        self.html_body = html_body
        self.attach_log = attach_log   # путь к лог-файлу или None

    def run(self):
        try:
            msg = MIMEMultipart("alternative")
            msg["Subject"] = self.subject
            msg["From"]    = self.cfg.get("smtp_from", self.cfg.get("smtp_user", ""))
            to_list = [x.strip() for x in self.cfg.get("smtp_to", "").split(",") if x.strip()]
            if not to_list:
                self.result.emit(False, "Не указан адрес получателя"); return
            msg["To"] = ", ".join(to_list)

            # Текстовая версия (fallback)
            plain = re.sub(r"<[^>]+>", "", self.html_body)
            msg.attach(MIMEText(plain, "plain", "utf-8"))
            msg.attach(MIMEText(self.html_body, "html", "utf-8"))

            # Вложение лога
            if self.attach_log and os.path.exists(self.attach_log):
                with open(self.attach_log, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                fn = os.path.basename(self.attach_log)
                part.add_header("Content-Disposition", "attachment", filename=fn)
                msg.attach(part)

            port = int(self.cfg.get("smtp_port", 587))
            host = self.cfg.get("smtp_host", "")
            user = self.cfg.get("smtp_user", "")
            pwd  = self.cfg.get("smtp_password", "")
            use_tls = self.cfg.get("smtp_tls", True)
            use_ssl = self.cfg.get("smtp_ssl", False)

            if use_ssl:
                ctx = ssl.create_default_context()
                with smtplib.SMTP_SSL(host, port, context=ctx, timeout=15) as s:
                    if user: s.login(user, pwd)
                    s.sendmail(msg["From"], to_list, msg.as_string())
            else:
                with smtplib.SMTP(host, port, timeout=15) as s:
                    s.ehlo()
                    if use_tls: s.starttls(context=ssl.create_default_context())
                    if user: s.login(user, pwd)
                    s.sendmail(msg["From"], to_list, msg.as_string())

            self.result.emit(True, f"Отправлено: {', '.join(to_list)}")
        except Exception as e:
            self.result.emit(False, str(e))


# ─── Настройки Email (диалог теста) ──────────────────────────────────────────

class EmailTestDialog(QDialog):
    def __init__(self, cfg: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Тест отправки email")
        self.setMinimumWidth(360)
        self.setStyleSheet(STYLESHEET)
        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.addWidget(QLabel("Нажмите кнопку — будет отправлено тестовое письмо на указанный адрес."))
        self._status = QLabel("Статус: ожидание")
        self._status.setWordWrap(True)
        self._status.setStyleSheet(f"color:{DARK_MUTED};font-size:12px;")
        lay.addWidget(self._status)
        btn = QPushButton("📨  Отправить тест"); btn.setObjectName("primary")
        btn.clicked.connect(self._send)
        lay.addWidget(btn)
        close = QPushButton("Закрыть"); close.clicked.connect(self.accept)
        lay.addWidget(close)
        self.cfg = cfg
        self._worker = None

    def _send(self):
        self._status.setText("Отправка...")
        self._status.setStyleSheet(f"color:{DARK_INFO};font-size:12px;")
        body = self._build_test_html()
        self._worker = EmailWorker(self.cfg, "🧪 Тест: HyperV Backup Manager", body)
        self._worker.result.connect(self._on_result)
        self._worker.start()

    def _on_result(self, ok: bool, msg: str):
        if ok:
            self._status.setText(f"✅ {msg}")
            self._status.setStyleSheet(f"color:{DARK_SUCCESS};font-size:12px;")
        else:
            self._status.setText(f"❌ {msg}")
            self._status.setStyleSheet(f"color:{DARK_ERROR};font-size:12px;")

    @staticmethod
    def _build_test_html() -> str:
        return f"""
        <div style="font-family:Segoe UI,sans-serif;max-width:560px;background:#0f1117;color:#e2e8f0;border-radius:12px;overflow:hidden;">
          <div style="background:linear-gradient(135deg,#4f46e5,#7c3aed);padding:24px 28px;">
            <h2 style="margin:0;color:#fff;font-size:20px;">🧪 Тестовое письмо</h2>
            <p style="margin:4px 0 0;color:#c7d2fe;font-size:13px;">HyperV Backup Manager Pro v{APP_VERSION}</p>
          </div>
          <div style="padding:24px 28px;">
            <p>SMTP-подключение работает корректно. Уведомления будут приходить на этот адрес.</p>
            <p style="color:#64748b;font-size:12px;margin-top:16px;">
              Отправлено: {datetime.now():%d.%m.%Y %H:%M:%S}
            </p>
          </div>
        </div>"""


# ─── Страница настроек Email ──────────────────────────────────────────────────

class EmailSettingsWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setSpacing(16)
        lay.setContentsMargins(0, 0, 0, 0)

        # ── SMTP ──────────────────────────────────────────────────────────────
        smtp_card = QWidget(); smtp_card.setObjectName("card")
        sg = QGridLayout(smtp_card)
        sg.setContentsMargins(20, 16, 20, 16)
        sg.setSpacing(10)
        sg.addWidget(self._section("📧  SMTP — СЕРВЕР ОТПРАВКИ"), 0, 0, 1, 3)

        fields = [
            ("SMTP сервер:", "smtp_host", "smtp.gmail.com"),
            ("Порт:",        "smtp_port", "587"),
            ("Логин:",       "smtp_user", "user@gmail.com"),
            ("Пароль:",      "smtp_password", ""),
            ("Отправитель:", "smtp_from", ""),
            ("Получатель(и):","smtp_to",  "admin@company.ru"),
        ]
        self._fields: dict[str, QLineEdit] = {}
        for i, (lbl, key, ph) in enumerate(fields, 1):
            sg.addWidget(QLabel(lbl), i, 0)
            le = QLineEdit(); le.setPlaceholderText(ph)
            if key == "smtp_password":
                le.setEchoMode(QLineEdit.EchoMode.Password)
            self._fields[key] = le
            sg.addWidget(le, i, 1, 1, 2)

        # TLS / SSL
        row = len(fields) + 1
        self._tls_check = QCheckBox("STARTTLS (порт 587)")
        self._tls_check.setChecked(True)
        self._ssl_check = QCheckBox("SSL/TLS (порт 465)")
        sg.addWidget(self._tls_check, row, 1)
        sg.addWidget(self._ssl_check, row, 2)
        lay.addWidget(smtp_card)

        # ── Триггеры ──────────────────────────────────────────────────────────
        trig_card = QWidget(); trig_card.setObjectName("card")
        tg = QVBoxLayout(trig_card)
        tg.setContentsMargins(20, 16, 20, 16)
        tg.setSpacing(8)
        tg.addWidget(self._section("🔔  КОГДА ОТПРАВЛЯТЬ"))
        self._trigger_checks: dict[str, QCheckBox] = {}
        triggers = [
            ("on_success",  "✅  Успешный бэкап"),
            ("on_error",    "❌  Ошибка при бэкапе"),
            ("on_warning",  "⚠️  Бэкап с предупреждениями"),
            ("on_delete",   "🗑  Удаление старых бэкапов"),
        ]
        for key, label in triggers:
            cb = QCheckBox(label)
            cb.setChecked(True)
            self._trigger_checks[key] = cb
            tg.addWidget(cb)
        lay.addWidget(trig_card)

        # ── Содержимое письма ─────────────────────────────────────────────────
        content_card = QWidget(); content_card.setObjectName("card")
        cg = QVBoxLayout(content_card)
        cg.setContentsMargins(20, 16, 20, 16)
        cg.setSpacing(8)
        cg.addWidget(self._section("📄  СОДЕРЖИМОЕ ПИСЬМА"))
        self._content_checks: dict[str, QCheckBox] = {}
        contents = [
            ("include_vm_list",  "💻  Список ВМ и статусы"),
            ("include_size",     "📦  Размер бэкапа"),
            ("include_path",     "📁  Путь к папке"),
            ("attach_log",       "📎  Вложить лог-файл (.txt)"),
        ]
        for key, label in contents:
            cb = QCheckBox(label)
            cb.setChecked(True)
            self._content_checks[key] = cb
            cg.addWidget(cb)
        lay.addWidget(content_card)

        bb = QHBoxLayout(); bb.setSpacing(8)
        test_btn = QPushButton("  Тест отправки")
        test_btn.setIcon(qicon("fa5s.paper-plane", "#c8ccd8", 13))
        test_btn.clicked.connect(self._test_email)
        save_btn = QPushButton("  Сохранить")
        save_btn.setIcon(qicon("fa5s.save", "#ffffff", 13))
        save_btn.setObjectName("primary"); save_btn.clicked.connect(self._save)
        bb.addStretch(); bb.addWidget(test_btn); bb.addWidget(save_btn)
        lay.addLayout(bb)
        lay.addStretch()

        self._load()

    @staticmethod
    def _section(text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet(f"color:{DARK_ACCENT};font-size:11px;font-weight:bold;"
                          "text-transform:uppercase;letter-spacing:1px;padding:2px 0;")
        return lbl

    def _get_cfg(self) -> dict:
        cfg = {k: le.text().strip() for k, le in self._fields.items()}
        cfg["smtp_tls"] = self._tls_check.isChecked()
        cfg["smtp_ssl"] = self._ssl_check.isChecked()
        for k, cb in self._trigger_checks.items():
            cfg[k] = cb.isChecked()
        for k, cb in self._content_checks.items():
            cfg[k] = cb.isChecked()
        return cfg

    def _save(self):
        if not os.path.exists(CONFIG_FILE):
            full = {}
        else:
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    full = json.load(f)
            except Exception:
                full = {}
        full["email"] = self._get_cfg()
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(full, f, indent=2, ensure_ascii=False)
            QMessageBox.information(self, "Успех", "Настройки email сохранены")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _load(self):
        if not os.path.exists(CONFIG_FILE): return
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                full = json.load(f)
            cfg = full.get("email", {})
            for k, le in self._fields.items():
                if k in cfg: le.setText(str(cfg[k]))
            self._tls_check.setChecked(bool(cfg.get("smtp_tls", True)))
            self._ssl_check.setChecked(bool(cfg.get("smtp_ssl", False)))
            for k, cb in self._trigger_checks.items():
                cb.setChecked(bool(cfg.get(k, True)))
            for k, cb in self._content_checks.items():
                cb.setChecked(bool(cfg.get(k, True)))
        except Exception:
            pass

    def _test_email(self):
        cfg = self._get_cfg()
        if not cfg.get("smtp_host"):
            QMessageBox.warning(self, "Предупреждение", "Укажите SMTP сервер")
            return
        dlg = EmailTestDialog(cfg, self)
        dlg.exec()

    def get_email_config(self) -> dict:
        return self._get_cfg()


# ─── Диалог редактирования задания ───────────────────────────────────────────

class EditJobDialog(QDialog):
    def __init__(self, job: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Редактировать задание")
        self.setMinimumWidth(380)
        self.setStyleSheet(STYLESHEET)
        self.job = dict(job)
        lay = QGridLayout(self)
        lay.setSpacing(10)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.addWidget(QLabel("Время:"), 0, 0)
        tr = QHBoxLayout()
        self.h = QComboBox(); self.h.addItems([f"{i:02d}" for i in range(24)])
        self.m = QComboBox(); self.m.addItems([f"{i:02d}" for i in range(60)])
        hh, mm = job["time"].split(":")
        self.h.setCurrentText(hh); self.m.setCurrentText(mm)
        tr.addWidget(self.h); tr.addWidget(QLabel(":")); tr.addWidget(self.m); tr.addStretch()
        lay.addLayout(tr, 0, 1)
        lay.addWidget(QLabel("Частота:"), 1, 0)
        self.freq = QComboBox()
        self.freq.addItems(["Ежедневно", "Еженедельно", "Ежемесячно"])
        self.freq.setCurrentText(job["frequency"])
        lay.addWidget(self.freq, 1, 1)
        lay.addWidget(QLabel("Режим:"), 2, 0)
        self.mode = QComboBox()
        self.mode.addItems(["Вместе (одна папка)", "По одной (отдельные папки)"])
        self.mode.setCurrentIndex(job.get("separate", 0))
        lay.addWidget(self.mode, 2, 1)
        self.en = QCheckBox("Активно"); self.en.setChecked(job.get("enabled", True))
        lay.addWidget(self.en, 3, 0, 1, 2)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self.accept); bb.rejected.connect(self.reject)
        lay.addWidget(bb, 4, 0, 1, 2)

    def result_job(self) -> dict:
        self.job["time"]      = f"{self.h.currentText()}:{self.m.currentText()}"
        self.job["frequency"] = self.freq.currentText()
        self.job["separate"]  = self.mode.currentIndex()
        self.job["enabled"]   = self.en.isChecked()
        return self.job


# ─── Сайдбар ─────────────────────────────────────────────────────────────────

class Sidebar(QWidget):
    pageChanged = pyqtSignal(int)

    # (текст, fa-иконка, индекс)
    NAV_ITEMS = [
        ("Дашборд",    "fa5s.chart-bar",      0),
        ("ВМ",         "fa5s.desktop",        1),
        ("Запуск",     "fa5s.rocket",         2),
        ("Расписание", "fa5s.clock",          3),
        ("История",    "fa5s.archive",        4),
        ("Логи",       "fa5s.scroll",         5),
        ("Настройки",  "fa5s.cog",            6),
    ]

    def __init__(self):
        super().__init__()
        self.setObjectName("sidebar")
        self.setFixedWidth(220)
        self.setStyleSheet(f"""
            Sidebar {{
                background-color: #666870;
            }}
            Sidebar QLabel {{
                background: transparent;
            }}
            Sidebar QFrame {{
                background: {DARK_BORDER};
            }}
            QPushButton#nav_btn {{
                background: transparent;
                border: none;
                padding: 10px 16px 10px 12px;
                margin: 2px 8px;
                border-radius: 8px;
                text-align: left;
                font-size: 13px;
                color: #f0f2f5;
                font-weight: 500;
            }}
            QPushButton#nav_btn:hover {{
                background: #787a88;
                color: #ffffff;
            }}
            QPushButton#nav_btn:checked {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #4e4b8a, stop:1 #5a3f95);
                color: #e0e7ff;
                font-weight: 600;
                border-left: 3px solid #6366f1;
            }}
        """)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 12, 0, 12)
        root.setSpacing(0)

        # ── Логотип ──────────────────────────────────────────────────────────
        logo_row = QHBoxLayout(); logo_row.setContentsMargins(16, 4, 16, 12)
        bolt_lbl = QLabel()
        bolt_lbl.setPixmap(qpixmap("fa5s.bolt", "#a5b4fc", 18))
        logo_row.addWidget(bolt_lbl)
        logo_txt = QLabel("HV Backup")
        logo_txt.setStyleSheet(
            "color:#a5b4fc;font-size:15px;font-weight:700;letter-spacing:0.5px;")
        logo_row.addWidget(logo_txt); logo_row.addStretch()
        root.addLayout(logo_row)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"background:{DARK_BORDER};margin:0 10px 6px 10px;max-height:1px;")
        root.addWidget(sep)

        # ── Кнопки навигации ─────────────────────────────────────────────────
        self.group = QButtonGroup(self); self.group.setExclusive(True)
        for label, icon_name, idx in self.NAV_ITEMS:
            btn = QPushButton(f"  {label}")
            btn.setObjectName("nav_btn")
            btn.setCheckable(True)
            btn.setIcon(qicon(icon_name, "#c8ccd8", 16))
            btn.setIconSize(__import__('PyQt6.QtCore', fromlist=['QSize']).QSize(18, 18))
            if idx == 0: btn.setChecked(True)
            self.group.addButton(btn, idx)
            root.addWidget(btn)

        self.group.buttonClicked.connect(lambda b: self.pageChanged.emit(self.group.id(b)))
        root.addStretch()

        # ── Версия ────────────────────────────────────────────────────────────
        ver = QLabel(f"v{APP_VERSION}")
        ver.setStyleSheet(f"color:{DARK_MUTED};font-size:11px;padding:8px 20px;")
        root.addWidget(ver)


# ─── Дашборд ─────────────────────────────────────────────────────────────────

class Dashboard(QWidget):
    def __init__(self, app_ref):
        super().__init__()
        self.app = app_ref
        self.root = QVBoxLayout(self)
        self.root.setSpacing(14)
        self.root.setContentsMargins(0, 0, 0, 0)

        # Карточки метрик
        cards_lay = QHBoxLayout(); cards_lay.setSpacing(10)
        self.cards = {}
        self._card_icon_lbls = {}
        card_defs = [
            ("vms",  "fa5s.desktop",    "Виртуальных машин", DARK_ACCENT),
            ("jobs", "fa5s.clock",      "Активных заданий",  DARK_ACCENT2),
            ("disk", "fa5s.hdd",        "Свободно на диске", DARK_SUCCESS),
            ("last", "fa5s.history",    "Последний бэкап",   DARK_INFO),
        ]
        for key, icon_name, label, color in card_defs:
            cw = QWidget(); cw.setObjectName("dash_card")
            cl = QVBoxLayout(cw); cl.setContentsMargins(18, 16, 18, 16); cl.setSpacing(4)
            icon_lbl = QLabel()
            icon_lbl.setPixmap(qpixmap(icon_name, color, 24))
            self._card_icon_lbls[key] = icon_lbl
            cl.addWidget(icon_lbl)
            vl = QLabel("—"); vl.setObjectName("stat_value")
            vl.setStyleSheet(f"color:{color};font-size:24px;font-weight:700;")
            cl.addWidget(vl)
            cl.addWidget(QLabel(label, styleSheet=f"color:{DARK_MUTED};font-size:11px;"))
            cards_lay.addWidget(cw)
            self.cards[key] = vl
        self.root.addLayout(cards_lay)

        # Быстрые действия + Алерты
        mid = QHBoxLayout(); mid.setSpacing(10)
        qa = QWidget(); qa.setObjectName("dash_card")
        ql = QVBoxLayout(qa); ql.setContentsMargins(16, 14, 16, 14); ql.setSpacing(8)
        hdr = QHBoxLayout(); hdr.setSpacing(6)
        hdr_icon = QLabel(); hdr_icon.setPixmap(qpixmap("fa5s.bolt", DARK_ACCENT, 13))
        hdr.addWidget(hdr_icon)
        hdr.addWidget(QLabel("Быстрые действия",
                             styleSheet=f"font-weight:600;color:{DARK_ACCENT};font-size:12px;"))
        hdr.addStretch()
        ql.addLayout(hdr)
        qr = QHBoxLayout(); qr.setSpacing(6)
        quick_btns = [
            ("Старт",       "fa5s.play",        "#22c55e", self.app._start_backup),
            ("Обновить ВМ", "fa5s.sync-alt",    "#38bdf8", self.app.load_vms),
            ("Открыть",     "fa5s.folder-open", "#f59e0b", self.app._browse_dir),
            ("Стоп",        "fa5s.stop",        "#ef4444", self.app._stop_backup),
        ]
        for txt, ico, col, fn in quick_btns:
            b = QPushButton(f" {txt}")
            b.setObjectName("quick_btn")
            b.setIcon(qicon(ico, col, 14))
            b.clicked.connect(fn)
            qr.addWidget(b)
        ql.addLayout(qr)
        mid.addWidget(qa, 3)

        al = QWidget(); al.setObjectName("dash_card")
        al_l = QVBoxLayout(al); al_l.setContentsMargins(16, 14, 16, 14); al_l.setSpacing(6)
        ah = QHBoxLayout(); ah.setSpacing(6)
        ah_icon = QLabel(); ah_icon.setPixmap(qpixmap("fa5s.exclamation-triangle", DARK_ERROR, 13))
        ah.addWidget(ah_icon)
        ah.addWidget(QLabel("Системные алерты",
                            styleSheet=f"font-weight:600;color:{DARK_ERROR};font-size:12px;"))
        ah.addStretch()
        al_l.addLayout(ah)
        self.alerts_box = QVBoxLayout(); al_l.addLayout(self.alerts_box)
        mid.addWidget(al, 2)
        self.root.addLayout(mid)

        # Последние бэкапы + Живой лог
        bot = QHBoxLayout(); bot.setSpacing(10)
        rec = QWidget(); rec.setObjectName("dash_card")
        rec_l = QVBoxLayout(rec); rec_l.setContentsMargins(16, 12, 16, 12)
        rh = QHBoxLayout(); rh.setSpacing(6)
        rh_icon = QLabel(); rh_icon.setPixmap(qpixmap("fa5s.archive", DARK_ACCENT, 13))
        rh.addWidget(rh_icon)
        rh.addWidget(QLabel("Последние бэкапы",
                            styleSheet=f"font-weight:600;color:{DARK_ACCENT};font-size:12px;"))
        rh.addStretch()
        rec_l.addLayout(rh)
        self.rec_list = QPlainTextEdit(); self.rec_list.setReadOnly(True)
        rec_l.addWidget(self.rec_list)
        bot.addWidget(rec, 1)

        log = QWidget(); log.setObjectName("dash_card")
        log_l = QVBoxLayout(log); log_l.setContentsMargins(16, 12, 16, 12)
        lh = QHBoxLayout(); lh.setSpacing(6)
        lh_icon = QLabel(); lh_icon.setPixmap(qpixmap("fa5s.terminal", DARK_ACCENT, 13))
        lh.addWidget(lh_icon)
        lh.addWidget(QLabel("Живой лог (последние 12 событий)",
                            styleSheet=f"font-weight:600;color:{DARK_ACCENT};font-size:12px;"))
        lh.addStretch()
        log_l.addLayout(lh)
        self.live_log = QPlainTextEdit(); self.live_log.setReadOnly(True)
        log_l.addWidget(self.live_log)
        bot.addWidget(log, 1)
        self.root.addLayout(bot)

        self.timer = QTimer(); self.timer.timeout.connect(self._refresh); self.timer.start(5000)
        self._refresh()

    def _refresh(self):
        try:
            path_edit = getattr(self.app, '_path_edit', None)
            path = path_edit.text().strip() if path_edit else ""
            vm_table = getattr(self.app, '_vm_table', None)
            vm_count = vm_table.rowCount() if vm_table else 0

            self.cards["vms"].setText(str(vm_count))
            job_count = sum(1 for j in self.app._scheduled_jobs if j.get("enabled", True))
            self.cards["jobs"].setText(str(job_count))

            try:
                drv = os.path.splitdrive(path)[0] + "\\" if path else "C:\\"
                u = shutil.disk_usage(drv)
                fg = u.free / 1024**3
                pct = u.free / u.total * 100
                col = DARK_SUCCESS if pct > 20 else DARK_WARN if pct > 10 else DARK_ERROR
                self.cards["disk"].setText(f"{fg:.1f} ГБ")
                self.cards["disk"].setStyleSheet(f"color:{col};font-size:24px;font-weight:700;")
            except:
                self.cards["disk"].setText("—")

            last = "—"
            if os.path.exists(LAST_BACKUP_FILE):
                try:
                    with open(LAST_BACKUP_FILE) as f:
                        lb = json.load(f).get("last_backup", "—")
                    if lb and lb != "—":
                        last = datetime.strptime(lb, "%Y-%m-%d %H:%M:%S").strftime("%d.%m %H:%M")
                except: pass
            self.cards["last"].setText(last)

            self.live_log.clear()
            for e in self.app._all_logs[-12:]:
                col = LOG_COLORS.get(e.level, DARK_TEXT)
                self.live_log.appendHtml(
                    f'<span style="color:{col}">[{e.ts}] {e.msg}</span>')
            self.live_log.verticalScrollBar().setValue(
                self.live_log.verticalScrollBar().maximum())

            if path and os.path.isdir(path):
                pat = re.compile(r"^\d{4}-\d{2}-\d{2}_\d{6}$")
                try:
                    bs = sorted(
                        [e for e in os.scandir(path) if e.is_dir() and pat.match(e.name)],
                        key=lambda x: x.name, reverse=True)[:5]
                    txt = "\n".join(f"📁 {b.name}" for b in bs) if bs else "Нет сохранённых бэкапов"
                    self.rec_list.setPlainText(txt)
                except:
                    self.rec_list.setPlainText("Ошибка чтения папки")
            else:
                self.rec_list.setPlainText("Путь не указан")

            self._clear_alerts()
            is_adm = False
            try: is_adm = bool(ctypes.windll.shell32.IsUserAnAdmin())
            except: pass
            if not is_adm:
                self._add_alert("⚠️ Запущено без прав Администратора.")
            try:
                if path:
                    drv = os.path.splitdrive(path)[0] + "\\"
                    u = shutil.disk_usage(drv)
                    if u.free / u.total < 0.1:
                        self._add_alert(f"🔴 Менее 10% свободно на {drv}")
            except: pass
            if vm_table:
                paused = sum(1 for r in range(vm_count)
                             if "Приостан" in (vm_table.item(r, 1).text() or ""))
                if paused > 0:
                    self._add_alert(f"⏸ {paused} ВМ приостановлены")
        except Exception as ex:
            self.app._log(f"Dashboard refresh: {ex}", "ERROR")

    def _add_alert(self, msg):
        lbl = QLabel(msg); lbl.setWordWrap(True)
        lbl.setStyleSheet(f"color:{DARK_WARN};font-size:12px;")
        self.alerts_box.addWidget(lbl)

    def _clear_alerts(self):
        while self.alerts_box.count():
            w = self.alerts_box.itemAt(0).widget()
            self.alerts_box.removeWidget(w); w.deleteLater()


# ─── Главное окно ─────────────────────────────────────────────────────────────

class HyperVBackupApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"HyperV Backup Manager Pro  v{APP_VERSION}")
        self.resize(1280, 860)
        self.setMinimumSize(1050, 720)

        self._log_worker = LogWorker()
        self._log_worker.log_signal.connect(self._append_log_entry)
        self._all_logs: list[LogEntry] = []

        self._scheduled_jobs:  list[dict] = []
        self._backup_process:  QProcess | None = None
        self._restore_process: QProcess | None = None
        self._vms_data:        list[dict] = []
        self._size_workers:    list[SizeCalcWorker] = []
        self._email_workers:   list[EmailWorker] = []

        self._backup_queue:      list[str] = []
        self._backup_queue_idx   = 0
        self._backup_separate    = False
        self._backup_export_root = ""
        self._session_sizes:     dict[str, int] = {}  # vm_name -> bytes

        self._progress = QProgressBar()
        self._progress.setFixedWidth(240)
        self._progress.hide()

        self._email_widget: EmailSettingsWidget | None = None

        self._setup_ui()
        self._setup_tray()
        self._load_config()
        self._load_schedule()
        self._check_admin()
        QTimer.singleShot(200, self.load_vms)
        self._start_scheduler()

    # ── Трей ──────────────────────────────────────────────────────────────────
    def _setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon))
        menu = QMenu()
        menu.addAction("🖥  Показать", self.showNormal)
        menu.addSeparator()
        menu.addAction("▶  Запустить бэкап", self._start_backup)
        menu.addSeparator()
        menu.addAction("❌  Выход", QApplication.quit)
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(
            lambda r: self.showNormal() if r == QSystemTrayIcon.ActivationReason.DoubleClick else None)
        self.tray_icon.show()

    def closeEvent(self, event):
        if self.tray_icon.isVisible():
            event.ignore(); self.hide()
            self.tray_icon.showMessage(
                "Свёрнуто", "Приложение работает в трее",
                QSystemTrayIcon.MessageIcon.Information, 2000)
        else:
            event.accept()

    # ── UI ────────────────────────────────────────────────────────────────────
    def _setup_ui(self):
        self.setStyleSheet(STYLESHEET)
        central = QWidget(); self.setCentralWidget(central)
        root = QHBoxLayout(central); root.setSpacing(0); root.setContentsMargins(0, 0, 0, 0)

        self.sidebar = Sidebar()
        root.addWidget(self.sidebar)

        # Контентная область
        content_wrap = QWidget()
        content_wrap.setStyleSheet(f"background:{DARK_BG};")
        cwl = QVBoxLayout(content_wrap); cwl.setContentsMargins(0, 0, 0, 0); cwl.setSpacing(0)

        self.stack = QStackedWidget()
        cwl.addWidget(self.stack)

        self._email_widget = EmailSettingsWidget()
        self.pages = [
            Dashboard(self),           # 0
            self._create_vms_page(),   # 1
            self._create_run_page(),   # 2
            self._create_schedule_page(), # 3
            self._create_backups_page(),  # 4
            self._create_log_page(),   # 5
            self._create_settings_page(), # 6
        ]
        for p in self.pages: self.stack.addWidget(p)
        root.addWidget(content_wrap, 1)

        self.sidebar.pageChanged.connect(self.stack.setCurrentIndex)
        self.sidebar.pageChanged.connect(self._save_active_page)
        QTimer.singleShot(100, self._restore_active_page)

        self.statusBar().addPermanentWidget(self._progress)
        self._set_status("Готов к работе")

    def _save_active_page(self, idx: int):
        try:
            cfg = {}
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f: cfg = json.load(f)
            cfg["last_page"] = idx
            with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(cfg, f, indent=2, ensure_ascii=False)
        except: pass

    def _restore_active_page(self):
        try:
            cfg = {}
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f: cfg = json.load(f)
            idx = cfg.get("last_page", 0)
            if 0 <= idx < len(self.pages):
                btn = self.sidebar.group.button(idx)
                if btn: btn.setChecked(True)
                self.stack.setCurrentIndex(idx)
        except: pass

    # ── Утилиты UI ────────────────────────────────────────────────────────────
    def _section(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet(f"color:{DARK_ACCENT};font-size:11px;font-weight:bold;"
                          "text-transform:uppercase;letter-spacing:1px;padding:2px 0;")
        return lbl

    def _page_wrap(self, widget: QWidget) -> QScrollArea:
        sa = QScrollArea(); sa.setWidgetResizable(True)
        sa.setWidget(widget)
        return sa

    # ── ВМ ────────────────────────────────────────────────────────────────────
    def _create_vms_page(self):
        tab = QWidget(); tab.setStyleSheet(f"background:{DARK_BG};")
        lay = QVBoxLayout(tab); lay.setContentsMargins(16, 12, 16, 12); lay.setSpacing(8)

        ctrl = QHBoxLayout(); ctrl.setSpacing(6)
        for lbl, ico, fn in [
            ("Обновить",     "fa5s.sync-alt",        self.load_vms),
            ("Все",          "fa5s.check-square",    self._select_all_vms),
            ("Снять",        "fa5s.square",          self._deselect_all_vms),
            ("Запущенные",   "fa5s.play-circle",     self._select_running_vms),
            ("Выключенные",  "fa5s.stop-circle",     self._select_stopped_vms),
        ]:
            b = QPushButton(f"  {lbl}")
            b.setIcon(qicon(ico, "#c8ccd8", 13))
            b.clicked.connect(fn); ctrl.addWidget(b)
        ctrl.addStretch()
        lay.addLayout(ctrl)

        sr = QHBoxLayout()
        sr.addWidget(QLabel("🔍"))
        self._search_edit = QLineEdit()
        self._search_edit.setPlaceholderText("Поиск по имени...")
        self._search_edit.textChanged.connect(self._filter_vms)
        sr.addWidget(self._search_edit)
        lay.addLayout(sr)

        self._vm_info_lbl = QLabel("Всего: 0  |  Выбрано: 0  |  Запущено: 0")
        self._vm_info_lbl.setStyleSheet(f"color:{DARK_MUTED};font-size:12px;padding:2px 0;")
        lay.addWidget(self._vm_info_lbl)

        self._vm_table = QTableWidget()
        self._vm_table.setColumnCount(4)
        self._vm_table.setHorizontalHeaderLabels(["✓", "Состояние", "Имя ВМ", "Память"])
        for i, m in [
            (0, QHeaderView.ResizeMode.ResizeToContents),
            (1, QHeaderView.ResizeMode.ResizeToContents),
            (2, QHeaderView.ResizeMode.Stretch),
            (3, QHeaderView.ResizeMode.ResizeToContents),
        ]:
            self._vm_table.horizontalHeader().setSectionResizeMode(i, m)
        self._vm_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._vm_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._vm_table.setAlternatingRowColors(True)
        self._vm_table.setStyleSheet(f"QTableWidget{{alternate-background-color:{DARK_SURFACE};}}")
        self._vm_table.itemChanged.connect(lambda _: self._update_vm_info())
        lay.addWidget(self._vm_table)
        return tab

    def load_vms(self):
        self._set_status("Загрузка ВМ...", busy=True)
        self._log("🔍 Получение списка виртуальных машин...", "SYSTEM")
        self._fetcher = VMFetcher()
        self._fetcher.finished.connect(self._on_vms_loaded)
        self._fetcher.error.connect(
            lambda e: (self._log(f"❌ {e}", "ERROR"), self._set_status("Ошибка загрузки")))
        self._fetcher.finished.connect(lambda _: self._set_status("Готово"))
        self._fetcher.start()

    def _on_vms_loaded(self, data: list):
        self._vms_data = data
        self._vm_table.blockSignals(True); self._vm_table.setRowCount(0)
        for vm in data:
            name = vm.get("Name", "Unknown")
            sk = resolve_vm_state(vm.get("State", 0))
            mem = vm.get("MemoryAssigned", 0) or 0
            mem_s = f"{mem/1024**3:.1f} ГБ" if mem else "—"
            si = VM_STATES.get(sk, DEFAULT_STATE)
            row = self._vm_table.rowCount(); self._vm_table.insertRow(row)
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            chk.setCheckState(Qt.CheckState.Unchecked)
            self._vm_table.setItem(row, 0, chk)
            st = QTableWidgetItem(si["text"])
            st.setBackground(si["bg"]); st.setForeground(si["fg"])
            st.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._vm_table.setItem(row, 1, st)
            self._vm_table.setItem(row, 2, QTableWidgetItem(name))
            mi = QTableWidgetItem(mem_s)
            mi.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            mi.setForeground(QColor(DARK_MUTED))
            self._vm_table.setItem(row, 3, mi)
        self._vm_table.blockSignals(False)
        self._log(f"✅ Загружено ВМ: {len(data)}", "SUCCESS")
        self._update_vm_info()

    def _get_selected_vms(self) -> list[str]:
        r = []
        for row in range(self._vm_table.rowCount()):
            if not self._vm_table.isRowHidden(row):
                c = self._vm_table.item(row, 0)
                if c and c.checkState() == Qt.CheckState.Checked:
                    r.append(self._vm_table.item(row, 2).text())
        return r

    def _select_all_vms(self):   self._set_checkboxes(Qt.CheckState.Checked)
    def _deselect_all_vms(self): self._set_checkboxes(Qt.CheckState.Unchecked)

    def _select_running_vms(self):
        self._vm_table.blockSignals(True)
        for row in range(self._vm_table.rowCount()):
            if self._vm_table.isRowHidden(row): continue
            si = self._vm_table.item(row, 1)
            if si and "Работает" in si.text():
                self._vm_table.item(row, 0).setCheckState(Qt.CheckState.Checked)
        self._vm_table.blockSignals(False); self._update_vm_info()

    def _select_stopped_vms(self):
        self._vm_table.blockSignals(True)
        for row in range(self._vm_table.rowCount()):
            if self._vm_table.isRowHidden(row): continue
            si = self._vm_table.item(row, 1)
            if si and "Выключена" in si.text():
                self._vm_table.item(row, 0).setCheckState(Qt.CheckState.Checked)
        self._vm_table.blockSignals(False); self._update_vm_info()

    def _set_checkboxes(self, state):
        self._vm_table.blockSignals(True)
        for row in range(self._vm_table.rowCount()):
            if not self._vm_table.isRowHidden(row):
                self._vm_table.item(row, 0).setCheckState(state)
        self._vm_table.blockSignals(False); self._update_vm_info()

    def _filter_vms(self, text: str):
        text = text.lower()
        for row in range(self._vm_table.rowCount()):
            name = self._vm_table.item(row, 2).text().lower()
            self._vm_table.setRowHidden(row, bool(text) and text not in name)
        self._update_vm_info()

    def _update_vm_info(self):
        vis = sel = run = 0
        for row in range(self._vm_table.rowCount()):
            if not self._vm_table.isRowHidden(row):
                vis += 1
                if self._vm_table.item(row, 0).checkState() == Qt.CheckState.Checked:
                    sel += 1
                if "Работает" in (self._vm_table.item(row, 1).text() or ""):
                    run += 1
        self._vm_info_lbl.setText(f"Всего: {vis}  |  ✅ Выбрано: {sel}  |  ● Запущено: {run}")

    # ── Расписание ────────────────────────────────────────────────────────────
    def _create_schedule_page(self):
        tab = QWidget(); tab.setStyleSheet(f"background:{DARK_BG};")
        lay = QVBoxLayout(tab); lay.setContentsMargins(16, 12, 16, 12); lay.setSpacing(10)

        ac = QWidget(); ac.setObjectName("card")
        ag = QGridLayout(ac); ag.setContentsMargins(16, 14, 16, 14); ag.setSpacing(10)
        ag.addWidget(self._section("➕  НОВОЕ ЗАДАНИЕ"), 0, 0, 1, 3)
        ag.addWidget(QLabel("ВМ:"), 1, 0)
        self._sched_vms_combo = QComboBox()
        self._sched_vms_combo.addItems(["Все запущенные", "Выбранные вручную"])
        ag.addWidget(self._sched_vms_combo, 1, 1)
        ag.addWidget(QLabel("Время:"), 2, 0)
        tr = QHBoxLayout()
        self._sched_hour = QComboBox(); self._sched_hour.addItems([f"{i:02d}" for i in range(24)])
        self._sched_hour.setCurrentText("23")
        self._sched_min = QComboBox(); self._sched_min.addItems([f"{i:02d}" for i in range(60)])
        tr.addWidget(self._sched_hour); tr.addWidget(QLabel(":")); tr.addWidget(self._sched_min); tr.addStretch()
        ag.addLayout(tr, 2, 1)
        ag.addWidget(QLabel("Частота:"), 3, 0)
        self._sched_freq = QComboBox()
        self._sched_freq.addItems(["Ежедневно", "Еженедельно", "Ежемесячно"])
        ag.addWidget(self._sched_freq, 3, 1)
        ag.addWidget(QLabel("Режим:"), 4, 0)
        self._sched_mode = QComboBox()
        self._sched_mode.addItems(["Вместе", "По одной ВМ"])
        ag.addWidget(self._sched_mode, 4, 1)
        ab = QPushButton("➕  Добавить задание"); ab.setObjectName("primary")
        ab.clicked.connect(self._add_schedule_job)
        ag.addWidget(ab, 5, 0, 1, 2)
        lay.addWidget(ac)

        self._jobs_table = QTableWidget(); self._jobs_table.setColumnCount(8)
        self._jobs_table.setHorizontalHeaderLabels(
            ["ID", "Время", "Частота", "Режим", "ВМ", "Последний запуск", "Следующий", "Статус"])
        for i in range(8):
            self._jobs_table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
        self._jobs_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._jobs_table.setColumnHidden(0, True)
        self._jobs_table.setAlternatingRowColors(True)
        self._jobs_table.setStyleSheet(f"QTableWidget{{alternate-background-color:{DARK_SURFACE};}}")
        lay.addWidget(self._jobs_table)

        bb = QHBoxLayout(); bb.setSpacing(6)
        eb = QPushButton("✏️  Изменить"); eb.clicked.connect(self._edit_job)
        db = QPushButton("🗑️  Удалить"); db.setObjectName("danger"); db.clicked.connect(self._delete_job)
        rb = QPushButton("▶  Запустить сейчас"); rb.setObjectName("primary"); rb.clicked.connect(self._run_job_now)
        bb.addWidget(eb); bb.addWidget(db); bb.addStretch(); bb.addWidget(rb)
        lay.addLayout(bb)
        return tab

    def _add_schedule_job(self):
        ts = f"{self._sched_hour.currentText()}:{self._sched_min.currentText()}"
        freq = self._sched_freq.currentText()
        vtype = self._sched_vms_combo.currentIndex()
        vms = self._get_selected_vms() if vtype == 1 else ["all"]
        sep = self._sched_mode.currentIndex()
        if vtype == 1 and not vms:
            QMessageBox.warning(self, "Внимание", "Не выбраны ВМ!"); return
        nid = max((j.get("id", 0) for j in self._scheduled_jobs), default=0) + 1
        job = {"id": nid, "time": ts, "frequency": freq, "vms": vms, "vms_type": vtype,
               "separate": sep, "enabled": True, "last_run": None,
               "next_run": self._calc_next_run(ts, freq)}
        self._scheduled_jobs.append(job)
        self._save_schedule(); self._refresh_jobs_list(); self._register_schedule(job)
        self._log(f"✅ Задание #{nid}: {freq} {ts}", "SUCCESS")

    @staticmethod
    def _calc_next_run(ts: str, freq: str) -> str:
        h, m = map(int, ts.split(":"))
        nxt = datetime.now().replace(hour=h, minute=m, second=0, microsecond=0)
        if nxt <= datetime.now():
            if freq == "Ежедневно":    nxt += timedelta(days=1)
            elif freq == "Еженедельно": nxt += timedelta(weeks=1)
            else:
                mo = nxt.month % 12 + 1
                yr = nxt.year + (1 if nxt.month == 12 else 0)
                nxt = nxt.replace(year=yr, month=mo)
        return nxt.strftime("%Y-%m-%d %H:%M")

    def _register_schedule(self, job: dict):
        if not job.get("enabled", True): return
        t = job["time"]
        if   job["frequency"] == "Ежедневно":    schedule.every().day.at(t).do(self._run_scheduled_backup, job)
        elif job["frequency"] == "Еженедельно":  schedule.every().week.at(t).do(self._run_scheduled_backup, job)
        else:                                     schedule.every().day.at(t).do(self._check_monthly_run, job)

    def _check_monthly_run(self, job):
        last = job.get("last_run")
        if last:
            try:
                ld = datetime.strptime(last, "%Y-%m-%d %H:%M:%S")
                if ld.month == datetime.now().month and ld.year == datetime.now().year: return
            except ValueError: pass
        self._run_scheduled_backup(job)

    def _run_scheduled_backup(self, job: dict):
        self._log("═" * 60, "SYSTEM")
        self._log(f"⏰ Расписание: {job['time']}  {job['frequency']}", "INFO")
        vms = job["vms"] if job.get("vms_type") == 1 else []
        self._run_backup_process(vms, separate=bool(job.get("separate", 0)), is_scheduled=True)
        job["last_run"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        job["next_run"] = self._calc_next_run(job["time"], job["frequency"])
        self._save_schedule(); self._refresh_jobs_list()

    def _refresh_jobs_list(self):
        self._jobs_table.blockSignals(True); self._jobs_table.setRowCount(0)
        for j in self._scheduled_jobs:
            row = self._jobs_table.rowCount(); self._jobs_table.insertRow(row)
            vt = "Все запущенные" if j.get("vms_type") == 0 else f"{len(j['vms'])} ВМ"
            md = "По одной" if j.get("separate", 0) else "Вместе"
            st = "✅ Активно" if j.get("enabled", True) else "⏸ Приостановлено"
            sc = QColor(DARK_SUCCESS) if j.get("enabled", True) else QColor(DARK_WARN)
            for col, val in enumerate([str(j["id"]), j["time"], j["frequency"], md, vt,
                                        j.get("last_run") or "—", j.get("next_run") or "—", st]):
                it = QTableWidgetItem(val)
                it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if col == 7: it.setForeground(sc)
                self._jobs_table.setItem(row, col, it)
        self._jobs_table.blockSignals(False)

    def _edit_job(self):
        row = self._jobs_table.currentRow()
        if row < 0: QMessageBox.warning(self, "Внимание", "Выберите задание"); return
        jid = int(self._jobs_table.item(row, 0).text())
        job = next((j for j in self._scheduled_jobs if j["id"] == jid), None)
        if not job: return
        dlg = EditJobDialog(job, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            upd = dlg.result_job()
            idx = next(i for i, j in enumerate(self._scheduled_jobs) if j["id"] == jid)
            self._scheduled_jobs[idx] = upd
            schedule.clear()
            for j in self._scheduled_jobs: self._register_schedule(j)
            self._save_schedule(); self._refresh_jobs_list()
            self._log(f"✏️ Задание #{jid} изменено", "INFO")

    def _delete_job(self):
        row = self._jobs_table.currentRow()
        if row < 0: QMessageBox.warning(self, "Внимание", "Выберите задание"); return
        jid = int(self._jobs_table.item(row, 0).text())
        if QMessageBox.question(self, "Подтверждение", f"Удалить задание #{jid}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self._scheduled_jobs = [j for j in self._scheduled_jobs if j["id"] != jid]
            schedule.clear()
            for j in self._scheduled_jobs: self._register_schedule(j)
            self._save_schedule(); self._refresh_jobs_list()
            self._log(f"🗑️ Задание #{jid} удалено", "WARNING")

    def _run_job_now(self):
        row = self._jobs_table.currentRow()
        if row < 0: QMessageBox.warning(self, "Внимание", "Выберите задание"); return
        jid = int(self._jobs_table.item(row, 0).text())
        job = next((j for j in self._scheduled_jobs if j["id"] == jid), None)
        if job: self._run_scheduled_backup(job)

    # ── Запуск ────────────────────────────────────────────────────────────────
    def _create_run_page(self):
        tab = QWidget(); tab.setStyleSheet(f"background:{DARK_BG};")
        root = QVBoxLayout(tab); root.setContentsMargins(16, 12, 16, 12); root.setSpacing(10)

        mode_card = QWidget(); mode_card.setObjectName("card")
        ml = QVBoxLayout(mode_card); ml.setContentsMargins(18, 14, 18, 14); ml.setSpacing(8)
        ml.addWidget(self._section("🚀  ЗАПУСК БЭКАПА"))
        mode_row = QHBoxLayout(); mode_row.setSpacing(12)
        mode_row.addWidget(QLabel("Режим бэкапа:"))
        self._run_mode_combo = QComboBox()
        self._run_mode_combo.addItems([
            "Вместе  (одна общая папка)",
            "По одной  (отдельная папка на каждую ВМ)"
        ])
        self._run_mode_combo.currentIndexChanged.connect(self._on_run_mode_changed)
        mode_row.addWidget(self._run_mode_combo, 1)
        ml.addLayout(mode_row)
        self._mode_hint = QLabel()
        self._mode_hint.setStyleSheet(f"color:{DARK_MUTED};font-size:12px;padding:2px 0 4px 0;")
        ml.addWidget(self._mode_hint)

        btn_row = QHBoxLayout(); btn_row.setSpacing(8)
        self._start_btn = QPushButton("  Запустить бэкап")
        self._start_btn.setIcon(qicon("fa5s.play", "#ffffff", 14))
        self._start_btn.setObjectName("primary"); self._start_btn.clicked.connect(self._start_backup)
        self._stop_btn  = QPushButton("  Остановить")
        self._stop_btn.setIcon(qicon("fa5s.stop", "#fca5a5", 14))
        self._stop_btn.setObjectName("danger"); self._stop_btn.setEnabled(False)
        self._stop_btn.clicked.connect(self._stop_backup)
        self._skip_btn  = QPushButton("  Следующая ВМ")
        self._skip_btn.setIcon(qicon("fa5s.step-forward", "#c8ccd8", 14))
        self._skip_btn.setEnabled(False); self._skip_btn.clicked.connect(self._skip_current_vm)
        btn_row.addWidget(self._start_btn, 2); btn_row.addWidget(self._skip_btn, 1); btn_row.addWidget(self._stop_btn, 1)
        ml.addLayout(btn_row)
        root.addWidget(mode_card)

        tbl_card = QWidget(); tbl_card.setObjectName("card")
        tl = QVBoxLayout(tbl_card); tl.setContentsMargins(18, 12, 18, 12); tl.setSpacing(6)
        tl.addWidget(self._section("📊  ПРОГРЕСС ВМ"))
        self._vm_progress_table = QTableWidget(); self._vm_progress_table.setColumnCount(4)
        self._vm_progress_table.setHorizontalHeaderLabels(["Виртуальная машина", "Прогресс", "Статус", "Папка бэкапа"])
        for i, m in [
            (0, QHeaderView.ResizeMode.Stretch),
            (1, QHeaderView.ResizeMode.Fixed),
            (2, QHeaderView.ResizeMode.ResizeToContents),
            (3, QHeaderView.ResizeMode.Stretch),
        ]:
            self._vm_progress_table.horizontalHeader().setSectionResizeMode(i, m)
        self._vm_progress_table.horizontalHeader().setDefaultSectionSize(140)
        self._vm_progress_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._vm_progress_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._vm_progress_table.setAlternatingRowColors(True)
        self._vm_progress_table.setStyleSheet(f"QTableWidget{{alternate-background-color:{DARK_SURFACE};}}")
        self._vm_progress_table.verticalHeader().setDefaultSectionSize(34)
        tl.addWidget(self._vm_progress_table)
        root.addWidget(tbl_card, 1)

        self._run_log = QPlainTextEdit(); self._run_log.setReadOnly(True)
        self._run_log.setMaximumBlockCount(500); self._run_log.setFixedHeight(110)
        root.addWidget(self._run_log)
        self._on_run_mode_changed(0)
        return tab

    def _on_run_mode_changed(self, idx: int):
        hints = [
            "Все выбранные ВМ экспортируются в общую папку одной командой.",
            "Каждая ВМ получает собственную подпапку. Можно пропустить отдельную ВМ.",
        ]
        self._mode_hint.setText(hints[idx])
        self._skip_btn.setEnabled(False)

    def _start_backup(self):
        path = self._path_edit.text().strip()
        if not path:
            QMessageBox.critical(self, "Ошибка", "Укажите папку"); return
        try: os.makedirs(path, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Нет доступа: {e}"); return
        separate = (self._run_mode_combo.currentIndex() == 1)
        self._run_backup_process(self._get_selected_vms(), separate=separate, is_scheduled=False)

    def _run_backup_process(self, selected_vms: list[str], *, separate: bool, is_scheduled: bool):
        if self._backup_process and self._backup_process.state() == QProcess.ProcessState.Running:
            QMessageBox.warning(self, "Внимание", "Бэкап уже выполняется!"); return
        path = self._path_edit.text().strip()
        ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self._backup_export_root = os.path.join(path, ts)
        self._backup_separate = separate
        self._session_sizes = {}
        self._log("═" * 70, "SYSTEM")
        self._log("🚀 НАЧАЛО БЭКАПА" + (" (расписание)" if is_scheduled else "") +
                  ("  [по одной]" if separate else "  [вместе]"), "INFO")
        self._log(f"📅 {datetime.now():%Y-%m-%d %H:%M:%S}", "INFO")
        self._log(f"📁 {self._backup_export_root}", "INFO")
        self._start_btn.setEnabled(False); self._stop_btn.setEnabled(True)
        self._run_log.clear()

        if not selected_vms:
            self._backup_queue = [
                self._vm_table.item(r, 2).text() for r in range(self._vm_table.rowCount())
                if "Работает" in (self._vm_table.item(r, 1).text() or "")
            ]
        else:
            self._backup_queue = list(selected_vms)

        if not self._backup_queue:
            self._log("⚠️ Нет ВМ для бэкапа", "WARNING")
            self._start_btn.setEnabled(True); self._stop_btn.setEnabled(False); return

        self._build_progress_table(); self._backup_queue_idx = 0
        self._set_status(f"Бэкап: 0/{len(self._backup_queue)} ВМ", busy=True)
        self._progress.setRange(0, len(self._backup_queue))
        self._progress.setValue(0); self._progress.show()
        self._progress.setFormat("%v/%m ВМ")

        if separate:
            self._skip_btn.setEnabled(True); self._backup_next_vm()
        else:
            self._skip_btn.setEnabled(False); self._backup_all_together()

    def _backup_all_together(self):
        export_dir = self._backup_export_root
        os.makedirs(export_dir, exist_ok=True)
        vms_ps = "@(" + ",".join(f"'{v}'" for v in self._backup_queue) + ")"
        cap = "-Capture" if "конфигурация" in self._backup_type_combo.currentText() else ""
        ret = self._retention_spin.value()
        ps = f"""$ErrorActionPreference='Stop';$exportDir="{export_dir.replace(chr(92),'/')}";$vmNames={vms_ps}
Write-Host "INFO: Всего ВМ: $($vmNames.Count)";$ok=0;$fail=0
foreach($vm in $vmNames){{Write-Host "PROC: $vm";try{{Export-VM -Name $vm -Path $exportDir {cap} -ErrorAction Stop;Write-Host "OK: $vm";$ok++}}catch{{Write-Host "ERR: $vm — $($_.Exception.Message)";$fail++}}}}
Write-Host "STAT: ok=$ok fail=$fail"
if({ret}-gt 0){{$p=Split-Path $exportDir -Parent;$cut=(Get-Date).AddDays(-{ret})
Get-ChildItem $p -Directory|Where-Object{{$_.Name-match '^\\d{{4}}-\\d{{2}}-\\d{{2}}' -and $_.CreationTime-lt $cut}}|ForEach-Object{{Remove-Item $_.FullName -Recurse -Force -EA SilentlyContinue;Write-Host "DEL: $($_.Name)"}}}}
Write-Host "DONE:"
"""
        self._launch_ps(ps, self._read_together_output, self._together_finished)

    def _read_together_output(self):
        proc = self._backup_process
        if not proc: return
        while proc.canReadLine():
            raw  = proc.readLine().data()
            line = re.sub(r"<[^>]+>|_x[0-9A-F]{4}_", "", _safe_decode(raw)).strip()
            if not line: continue
            if line.startswith("PROC:"):
                vm = line[5:].strip(); self._set_vm_status(vm, "running"); self._run_log_append(f"⏳ {vm}")
            elif line.startswith("OK:"):
                vm = line[3:].strip(); self._set_vm_status(vm, "done")
                self._backup_queue_idx += 1; self._progress.setValue(self._backup_queue_idx)
                self._set_status(f"Бэкап: {self._backup_queue_idx}/{len(self._backup_queue)} ВМ", busy=True)
                self._run_log_append(f"✅ {vm}"); self._log(f"✅ {vm}", "SUCCESS")
            elif line.startswith("ERR:"):
                parts = line[4:].split("—", 1)
                vm = parts[0].strip() if len(parts) > 1 else "Unknown"
                self._set_vm_status(vm, "error"); self._run_log_append(f"❌ {vm}")
                self._log(f"❌ {line[4:].strip()}", "ERROR")
                self._backup_queue_idx += 1; self._progress.setValue(self._backup_queue_idx)
            elif line.startswith("STAT:"):
                self._log(f"📊 {line[5:].strip()}", "INFO")
            elif line.startswith("DEL:"):
                deleted = line[4:].strip()
                self._log(f"🗑 Удалён: {deleted}", "WARNING")
                self._send_email_notification("delete", deleted_name=deleted)
            elif line.startswith("INFO:"):
                self._log(line[5:].strip(), "INFO")

    def _together_finished(self, code: int, _): self._finish_backup_session(code)

    def _backup_next_vm(self):
        idx = self._backup_queue_idx
        if idx >= len(self._backup_queue):
            self._finish_backup_session(0); return
        vm = self._backup_queue[idx]
        self._set_vm_status(vm, "running"); self._run_log_append(f"⏳ Экспорт: {vm}")
        self._set_status(f"Бэкап {idx+1}/{len(self._backup_queue)}: {vm}", busy=True)
        export_dir = os.path.join(self._backup_export_root, vm)
        os.makedirs(export_dir, exist_ok=True); self._set_vm_folder(vm, export_dir)
        cap = "-Capture" if "конфигурация" in self._backup_type_combo.currentText() else ""
        ps = f"""$ErrorActionPreference='Stop';$exportDir="{export_dir.replace(chr(92),'/')}"
try{{Export-VM -Name '{vm}' -Path $exportDir {cap} -ErrorAction Stop;Write-Host "OK: {vm}"}}catch{{Write-Host "ERR: {vm} — $($_.Exception.Message)"}}
Write-Host "NEXT:"
"""
        self._launch_ps(ps, self._read_single_output, self._single_finished)

    def _read_single_output(self):
        proc = self._backup_process
        if not proc: return
        while proc.canReadLine():
            line = re.sub(r"<[^>]+>|_x[0-9A-F]{4}_", "", _safe_decode(proc.readLine().data())).strip()
            if not line or line.isdigit(): continue
            if line.startswith("OK:"):
                vm = line[3:].strip(); self._set_vm_status(vm, "done")
                self._log(f"✅ {vm}", "SUCCESS"); self._run_log_append(f"✅ {vm}")
            elif line.startswith("ERR:"):
                self._set_vm_status(self._backup_queue[self._backup_queue_idx], "error")
                self._log(f"❌ {line[4:].strip()}", "ERROR")
                self._run_log_append(f"❌ {line[4:].strip()}")
            else:
                self._log(line, "INFO")

    def _single_finished(self, code: int, _):
        # Считаем размер папки этой ВМ
        vm = self._backup_queue[self._backup_queue_idx] if self._backup_queue_idx < len(self._backup_queue) else ""
        if vm:
            vm_path = os.path.join(self._backup_export_root, vm)
            w = SizeCalcWorker(vm_path, self._backup_queue_idx)
            w.finished.connect(self._on_vm_size_calculated)
            w.start(); self._size_workers.append(w)
        self._backup_queue_idx += 1; self._progress.setValue(self._backup_queue_idx)
        if self._backup_queue_idx < len(self._backup_queue):
            self._backup_next_vm()
        else:
            self._retention_cleanup(); self._finish_backup_session(code)

    def _on_vm_size_calculated(self, row: int, label: str, total_bytes: int):
        vm = self._backup_queue[row] if row < len(self._backup_queue) else ""
        if vm: self._session_sizes[vm] = total_bytes

    def _skip_current_vm(self):
        if not self._backup_separate: return
        idx = self._backup_queue_idx
        if idx >= len(self._backup_queue): return
        vm = self._backup_queue[idx]
        proc = self._backup_process
        if proc and proc.state() == QProcess.ProcessState.Running:
            try: proc.finished.disconnect()
            except: pass
            proc.terminate()
            if not proc.waitForFinished(2000): proc.kill()
        self._set_vm_status(vm, "skipped")
        self._log(f"⏭ Пропущена ВМ: {vm}", "WARNING")
        self._run_log_append(f"⏭ Пропущена: {vm}")
        self._backup_queue_idx += 1; self._progress.setValue(self._backup_queue_idx)
        if self._backup_queue_idx < len(self._backup_queue): self._backup_next_vm()
        else: self._retention_cleanup(); self._finish_backup_session(0)

    def _launch_ps(self, script: str, out_slot, fin_slot):
        encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        self._backup_process = QProcess()
        self._backup_process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        self._backup_process.readyReadStandardOutput.connect(out_slot)
        self._backup_process.finished.connect(fin_slot)
        self._backup_process.start("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded])

    def _finish_backup_session(self, code: int):
        self._start_btn.setEnabled(True); self._stop_btn.setEnabled(False); self._skip_btn.setEnabled(False)
        self._progress.hide(); self._set_status("Готов к работе")
        done = sum(1 for r in range(self._vm_progress_table.rowCount())
                   if self._vm_progress_table.item(r, 2) and "Готово" in self._vm_progress_table.item(r, 2).text())
        err  = sum(1 for r in range(self._vm_progress_table.rowCount())
                   if self._vm_progress_table.item(r, 2) and "Ошибка" in self._vm_progress_table.item(r, 2).text())
        skip = len(self._backup_queue) - done - err
        self._run_log_append("─" * 40)
        self._run_log_append(f"✅ Готово: {done}   ❌ Ошибок: {err}   ⏭ Пропущено: {skip}")
        self._log(f"📊 Итог — Готово: {done}  Ошибок: {err}  Пропущено: {skip}", "INFO")

        if err == 0 and done > 0:
            self._log("✅ Бэкап успешно завершён!", "SUCCESS")
            self._save_last_backup()
            self._notify("Бэкап завершён", f"Готово: {done} ВМ")
            self._send_email_notification("success", done=done, err=err, skip=skip)
        elif err > 0 and done > 0:
            self._log(f"⚠️ Бэкап завершён с ошибками ({err} ВМ)", "WARNING")
            self._send_email_notification("warning", done=done, err=err, skip=skip)
        elif err > 0 and done == 0:
            self._log("❌ Бэкап завершился с ошибками", "ERROR")
            self._send_email_notification("error", done=done, err=err, skip=skip)

        QTimer.singleShot(400, self._refresh_backups_table)

    def _stop_backup(self):
        proc = self._backup_process
        if proc and proc.state() == QProcess.ProcessState.Running:
            try: proc.finished.disconnect()
            except: pass
            proc.terminate()
            if not proc.waitForFinished(3000): proc.kill()
            self._log("🛑 Бэкап остановлен", "WARNING")
        self._finish_backup_session(-1)

    def _retention_cleanup(self):
        ret = self._retention_spin.value()
        if ret <= 0: return
        parent = os.path.dirname(self._backup_export_root)
        cut = datetime.now() - timedelta(days=ret)
        pat = re.compile(r"^\d{4}-\d{2}-\d{2}_\d{6}$")
        deleted_names = []
        try:
            for e in os.scandir(parent):
                if e.is_dir() and pat.match(e.name):
                    try:
                        dt = datetime.strptime(e.name, "%Y-%m-%d_%H%M%S")
                        if dt < cut:
                            shutil.rmtree(e.path, ignore_errors=True)
                            self._log(f"🗑 Удалён старый бэкап: {e.name}", "WARNING")
                            deleted_names.append(e.name)
                    except Exception: pass
        except Exception: pass
        if deleted_names:
            self._send_email_notification("delete", deleted_name=", ".join(deleted_names))

    def _build_progress_table(self):
        self._vm_progress_table.setRowCount(0)
        for vm in self._backup_queue:
            row = self._vm_progress_table.rowCount(); self._vm_progress_table.insertRow(row)
            self._vm_progress_table.setItem(row, 0, QTableWidgetItem(vm))
            pb = QProgressBar(); pb.setRange(0, 0); pb.setFixedHeight(10); pb.setTextVisible(False)
            pb.setStyleSheet(f"QProgressBar{{background:{DARK_BORDER};border-radius:3px}}"
                             f"QProgressBar::chunk{{background:{DARK_MUTED};border-radius:3px}}")
            self._vm_progress_table.setCellWidget(row, 1, pb)
            st_it = QTableWidgetItem("⏳ Ожидание")
            st_it.setForeground(QColor(DARK_MUTED)); st_it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._vm_progress_table.setItem(row, 2, st_it)
            self._vm_progress_table.setItem(row, 3, QTableWidgetItem("—"))

    def _set_vm_status(self, vm_name: str, status_key: str):
        info = VM_JOB_STATUS.get(status_key, VM_JOB_STATUS["pending"])
        CHUNK_COLORS = {
            "running": DARK_ACCENT,
            "done":    DARK_SUCCESS,
            "error":   DARK_ERROR,
            "skipped": DARK_MUTED,
        }
        for row in range(self._vm_progress_table.rowCount()):
            ni = self._vm_progress_table.item(row, 0)
            if ni and ni.text() == vm_name:
                pb = self._vm_progress_table.cellWidget(row, 1)
                if isinstance(pb, QProgressBar):
                    chunk = CHUNK_COLORS.get(status_key, DARK_MUTED)
                    if status_key == "running":
                        pb.setRange(0, 0)
                    else:
                        pb.setRange(0, 100); pb.setValue(100)
                    pb.setStyleSheet(f"QProgressBar{{background:{DARK_BORDER};border-radius:3px}}"
                                     f"QProgressBar::chunk{{background:{chunk};border-radius:3px}}")
                si = self._vm_progress_table.item(row, 2)
                if si: si.setText(info["text"]); si.setForeground(QColor(info["color"]))
                break

    def _set_vm_folder(self, vm_name: str, folder: str):
        for row in range(self._vm_progress_table.rowCount()):
            ni = self._vm_progress_table.item(row, 0)
            if ni and ni.text() == vm_name:
                fi = self._vm_progress_table.item(row, 3)
                if fi: fi.setText(folder); fi.setForeground(QColor(DARK_MUTED)); break

    def _run_log_append(self, text: str):
        self._run_log.appendPlainText(text)
        self._run_log.verticalScrollBar().setValue(self._run_log.verticalScrollBar().maximum())

    # ── Email-уведомления ─────────────────────────────────────────────────────

    def _get_email_cfg(self) -> dict:
        """Загружает секцию email из конфига."""
        if not os.path.exists(CONFIG_FILE): return {}
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f).get("email", {})
        except Exception:
            return {}

    def _send_email_notification(self, event: str, *,
                                  done: int = 0, err: int = 0, skip: int = 0,
                                  deleted_name: str = ""):
        """Собирает и отправляет email при наступлении события."""
        cfg = self._get_email_cfg()
        if not cfg.get("smtp_host") or not cfg.get("smtp_to"):
            return  # SMTP не настроен

        # Проверяем триггер
        trigger_map = {
            "success": "on_success",
            "error":   "on_error",
            "warning": "on_warning",
            "delete":  "on_delete",
        }
        trigger_key = trigger_map.get(event, "")
        if trigger_key and not cfg.get(trigger_key, True):
            return  # Данный триггер выключен

        # Строим список ВМ для письма
        vm_rows = ""
        if cfg.get("include_vm_list", True):
            for r in range(self._vm_progress_table.rowCount()):
                nm = self._vm_progress_table.item(r, 0)
                st = self._vm_progress_table.item(r, 2)
                if nm and st:
                    row_color = "#1a2e1a" if "Готово" in st.text() else (
                        "#2e1a1a" if "Ошибка" in st.text() else "#1a1a2e")
                    vm_rows += (f'<tr style="background:{row_color};">'
                                f'<td style="padding:8px 12px;border-bottom:1px solid #2a2f45;">{nm.text()}</td>'
                                f'<td style="padding:8px 12px;border-bottom:1px solid #2a2f45;">{st.text()}</td>'
                                f'</tr>')

        path_line = ""
        if cfg.get("include_path", True):
            path_line = f'<p style="color:#94a3b8;font-size:13px;">📁 <b>Папка:</b> {self._backup_export_root}</p>'

        # Размер всех ВМ сессии
        size_line = ""
        if cfg.get("include_size", True) and self._session_sizes:
            total = sum(self._session_sizes.values())
            if total >= 1024**3:   sz = f"{total/1024**3:.2f} ГБ"
            elif total >= 1024**2: sz = f"{total/1024**2:.0f} МБ"
            else:                  sz = f"{total/1024:.0f} КБ"
            size_line = f'<p style="color:#94a3b8;font-size:13px;">📦 <b>Размер:</b> {sz}</p>'

        # Заголовок + цвет по событию
        headers = {
            "success": ("#22c55e", "#14532d", "✅ Бэкап успешно завершён"),
            "error":   ("#ef4444", "#450a0a", "❌ Ошибка при создании бэкапа"),
            "warning": ("#f59e0b", "#451a03", "⚠️ Бэкап завершён с предупреждениями"),
            "delete":  ("#38bdf8", "#0c3045", "🗑  Удаление старых бэкапов"),
        }
        accent, bg, title = headers.get(event, (DARK_ACCENT, "#2e2b6b", "📧 Уведомление"))

        stats_block = ""
        if event != "delete":
            stats_block = f"""
            <div style="display:flex;gap:12px;margin:16px 0;">
              <div style="background:#14532d;border-radius:8px;padding:12px 20px;flex:1;text-align:center;">
                <div style="color:#86efac;font-size:22px;font-weight:700;">{done}</div>
                <div style="color:#86efac;font-size:11px;">ГОТОВО</div>
              </div>
              <div style="background:#450a0a;border-radius:8px;padding:12px 20px;flex:1;text-align:center;">
                <div style="color:#fca5a5;font-size:22px;font-weight:700;">{err}</div>
                <div style="color:#fca5a5;font-size:11px;">ОШИБОК</div>
              </div>
              <div style="background:#1e293b;border-radius:8px;padding:12px 20px;flex:1;text-align:center;">
                <div style="color:#94a3b8;font-size:22px;font-weight:700;">{skip}</div>
                <div style="color:#94a3b8;font-size:11px;">ПРОПУЩЕНО</div>
              </div>
            </div>"""
        else:
            stats_block = f'<p style="color:#94a3b8;">Удалённые бэкапы: <b style="color:#e2e8f0;">{deleted_name}</b></p>'

        vm_table_block = ""
        if vm_rows:
            vm_table_block = f"""
            <table style="width:100%;border-collapse:collapse;margin-top:12px;">
              <thead>
                <tr>
                  <th style="background:#1a1d27;color:#64748b;padding:8px 12px;text-align:left;font-size:11px;">ВМ</th>
                  <th style="background:#1a1d27;color:#64748b;padding:8px 12px;text-align:left;font-size:11px;">СТАТУС</th>
                </tr>
              </thead>
              <tbody>{vm_rows}</tbody>
            </table>"""

        html_body = f"""
        <div style="font-family:'Segoe UI',sans-serif;max-width:600px;background:#0f1117;
                    color:#e2e8f0;border-radius:12px;overflow:hidden;
                    border:1px solid #2a2f45;">
          <div style="background:linear-gradient(135deg,{bg},{accent}22);
                      border-bottom:2px solid {accent};padding:24px 28px;">
            <h2 style="margin:0;color:{accent};font-size:20px;">{title}</h2>
            <p style="margin:6px 0 0;color:#94a3b8;font-size:12px;">
              HyperV Backup Manager Pro v{APP_VERSION} &nbsp;·&nbsp;
              {datetime.now():%d.%m.%Y %H:%M:%S}
            </p>
          </div>
          <div style="padding:24px 28px;">
            {stats_block}
            {path_line}
            {size_line}
            {vm_table_block}
            <p style="color:#475569;font-size:11px;margin-top:20px;border-top:1px solid #1e293b;padding-top:12px;">
              Сервер: {os.environ.get("COMPUTERNAME","—")}
            </p>
          </div>
        </div>"""

        subject_map = {
            "success": f"✅ Бэкап Hyper-V — {done} ВМ ({datetime.now():%d.%m.%Y %H:%M})",
            "error":   f"❌ Ошибка бэкапа Hyper-V — {err} ВМ ({datetime.now():%d.%m.%Y})",
            "warning": f"⚠️ Бэкап Hyper-V — есть ошибки ({datetime.now():%d.%m.%Y})",
            "delete":  f"🗑  Удалены старые бэкапы Hyper-V ({datetime.now():%d.%m.%Y})",
        }
        subject = subject_map.get(event, "📧 Уведомление HyperV Backup")

        # Лог-файл (если включено)
        log_path = None
        if cfg.get("attach_log", True):
            log_path = f"backup_log_{datetime.now():%Y%m%d_%H%M%S}.txt"
            try:
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(f"[{e.ts}][{e.level}] {e.msg}" for e in self._all_logs[-500:]))
            except Exception:
                log_path = None

        worker = EmailWorker(cfg, subject, html_body, attach_log=log_path)
        worker.result.connect(self._on_email_result)
        worker.start()
        self._email_workers.append(worker)

    def _on_email_result(self, ok: bool, msg: str):
        if ok:
            self._log(f"📧 Email отправлен: {msg}", "SUCCESS")
        else:
            self._log(f"📧 Ошибка email: {msg}", "ERROR")

    # ── История ───────────────────────────────────────────────────────────────
    def _create_backups_page(self):
        tab = QWidget(); tab.setStyleSheet(f"background:{DARK_BG};")
        lay = QVBoxLayout(tab); lay.setContentsMargins(16, 12, 16, 12); lay.setSpacing(8)

        ctrl = QHBoxLayout(); ctrl.setSpacing(8)
        rb = QPushButton("  Обновить")
        rb.setIcon(qicon("fa5s.sync-alt", "#c8ccd8", 13))
        rb.clicked.connect(self._refresh_backups_table)
        self._restore_btn = QPushButton("  Восстановить")
        self._restore_btn.setIcon(qicon("fa5s.undo-alt", "#ffffff", 13))
        self._restore_btn.setObjectName("primary"); self._restore_btn.clicked.connect(self._restore_backup)
        self._restore_btn.setEnabled(False)
        self._del_backup_btn = QPushButton("  Удалить")
        self._del_backup_btn.setIcon(qicon("fa5s.trash-alt", "#fca5a5", 13))
        self._del_backup_btn.setObjectName("danger"); self._del_backup_btn.clicked.connect(self._delete_backup)
        self._del_backup_btn.setEnabled(False)
        ctrl.addWidget(rb); ctrl.addStretch(); ctrl.addWidget(self._restore_btn); ctrl.addWidget(self._del_backup_btn)
        lay.addLayout(ctrl)

        self._backups_table = QTableWidget(); self._backups_table.setColumnCount(4)
        self._backups_table.setHorizontalHeaderLabels(["📅 Дата / Время", "💻 ВМ", "📦 Размер", "📁 Путь"])
        for i, m in [
            (0, QHeaderView.ResizeMode.ResizeToContents),
            (1, QHeaderView.ResizeMode.Stretch),
            (2, QHeaderView.ResizeMode.ResizeToContents),
            (3, QHeaderView.ResizeMode.Stretch),
        ]:
            self._backups_table.horizontalHeader().setSectionResizeMode(i, m)
        self._backups_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._backups_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._backups_table.setAlternatingRowColors(True)
        self._backups_table.setStyleSheet(f"QTableWidget{{alternate-background-color:{DARK_SURFACE};}}")
        self._backups_table.itemSelectionChanged.connect(self._on_backup_selected)
        lay.addWidget(self._backups_table)
        QTimer.singleShot(600, self._refresh_backups_table)
        return tab

    def _refresh_backups_table(self):
        path = self._path_edit.text().strip()
        if not os.path.isdir(path): self._backups_table.setRowCount(0); return
        pat = re.compile(r"^\d{4}-\d{2}-\d{2}_\d{6}$"); bs = []
        try:
            for e in os.scandir(path):
                if e.is_dir() and pat.match(e.name):
                    try: vms = [d.name for d in os.scandir(e.path) if d.is_dir()]
                    except: vms = []
                    bs.append({"dt": datetime.strptime(e.name, "%Y-%m-%d_%H%M%S"),
                               "vms": ", ".join(vms) or "—", "path": e.path})
        except Exception as ex: self._log(f"❌ {ex}", "ERROR"); return
        bs.sort(key=lambda x: x["dt"], reverse=True)
        self._backups_table.blockSignals(True); self._backups_table.setRowCount(0)
        for b in bs:
            row = self._backups_table.rowCount(); self._backups_table.insertRow(row)
            self._backups_table.setItem(row, 0, QTableWidgetItem(b["dt"].strftime("%Y-%m-%d  %H:%M:%S")))
            self._backups_table.setItem(row, 1, QTableWidgetItem(b["vms"]))
            si = QTableWidgetItem("⏳ расчёт...")
            si.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            si.setForeground(QColor(DARK_MUTED)); self._backups_table.setItem(row, 2, si)
            pi = QTableWidgetItem(b["path"]); pi.setForeground(QColor(DARK_MUTED))
            self._backups_table.setItem(row, 3, pi)
            w = SizeCalcWorker(b["path"], row)
            w.finished.connect(self._update_backup_size); w.start(); self._size_workers.append(w)
        self._backups_table.blockSignals(False); self._on_backup_selected()

    def _update_backup_size(self, row: int, sz: str, total: int):
        if 0 <= row < self._backups_table.rowCount():
            it = self._backups_table.item(row, 2)
            if it: it.setText(sz); it.setForeground(QColor(DARK_TEXT))

    def _on_backup_selected(self):
        has = bool(self._backups_table.selectedItems())
        self._restore_btn.setEnabled(has); self._del_backup_btn.setEnabled(has)

    def _restore_backup(self):
        row = self._backups_table.currentRow()
        if row < 0: return
        bpath = self._backups_table.item(row, 3).text()
        if QMessageBox.question(self, "Восстановление",
            f"Восстановить ВМ из:\n{bpath}\n\n(Новые ID, конфликтов не будет)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes: return
        self._set_status("Восстановление...", busy=True); self._restore_btn.setEnabled(False)
        ps = f"""$ErrorActionPreference='Stop';$bp='{bpath}';$ok=0;$fail=0
Get-ChildItem $bp -Directory|ForEach-Object{{$vd=$_.FullName;$vf=Join-Path $vd "Virtual Machines";$vc=Get-ChildItem $vf -Filter "*.vmcx" -EA SilentlyContinue;if($vc){{$vmx=$vc[0].FullName;$vn=Split-Path $vd -Leaf;Write-Host "PROC: $vn";try{{Import-VM -Path $vmx -Copy -GenerateNewId -EA Stop;Write-Host "OK: $vn";$ok++}}catch{{Write-Host "ERR: $vn — $($_.Exception.Message)";$fail++}}}}}}
Write-Host "STAT: ok=$ok fail=$fail";Write-Host "DONE:"
"""
        enc = base64.b64encode(ps.encode("utf-16-le")).decode("ascii")
        self._restore_process = QProcess()
        self._restore_process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        self._restore_process.readyReadStandardOutput.connect(self._read_restore_output)
        self._restore_process.finished.connect(self._restore_finished)
        self._restore_process.start("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", enc])

    def _read_restore_output(self):
        proc = self._restore_process
        if not proc: return
        while proc.canReadLine():
            line = re.sub(r"<[^>]+>|_x[0-9A-F]{4}_", "", _safe_decode(proc.readLine().data())).strip()
            if not line or line.isdigit(): continue
            if   line.startswith("OK:"):   self._log(f"✅ {line[3:].strip()}", "SUCCESS")
            elif line.startswith("ERR:"):  self._log(f"❌ {line[4:].strip()}", "ERROR")
            elif line.startswith("PROC:"): self._log(f"⏳ {line[5:].strip()}", "INFO")
            elif line.startswith("STAT:"): self._log(f"📊 {line[5:].strip()}", "INFO")
            elif line.startswith("DONE:"): self._log("✅ Восстановление завершено", "SUCCESS")
            else: self._log(line, "INFO")

    def _restore_finished(self, code, _):
        self._restore_btn.setEnabled(True); self._set_status("Готов к работе")
        if code not in (0, -1): self._log(f"⚠️ Ошибка восстановления (код {code})", "WARNING")
        QTimer.singleShot(300, self._refresh_backups_table)

    def _delete_backup(self):
        row = self._backups_table.currentRow()
        if row < 0: return
        bpath = self._backups_table.item(row, 3).text()
        if QMessageBox.warning(self, "Удаление", f"Удалить:\n{bpath}\n\nНельзя отменить!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes: return
        try:
            shutil.rmtree(bpath)
            self._log(f"🗑️ Удалён: {bpath}", "WARNING")
            self._refresh_backups_table()
        except Exception as e:
            self._log(f"❌ {e}", "ERROR"); QMessageBox.critical(self, "Ошибка", str(e))

    # ── Логи ──────────────────────────────────────────────────────────────────
    def _create_log_page(self):
        tab = QWidget(); tab.setStyleSheet(f"background:{DARK_BG};")
        lay = QVBoxLayout(tab); lay.setContentsMargins(12, 10, 12, 10); lay.setSpacing(8)

        ctrl = QHBoxLayout(); ctrl.setSpacing(6)
        for lbl, ico, fn in [
            ("Очистить",   "fa5s.trash",      self._clear_log),
            ("Копировать", "fa5s.copy",        self._copy_log),
            ("Сохранить",  "fa5s.save",        self._save_log),
        ]:
            b = QPushButton(f"  {lbl}")
            b.setIcon(qicon(ico, "#c8ccd8", 13))
            b.clicked.connect(fn); ctrl.addWidget(b)
        ctrl.addStretch()
        ctrl.addWidget(QLabel("Фильтр:"))
        self._log_filter = QComboBox()
        self._log_filter.addItems(["Все", "Инфо", "Успех", "Ошибка", "Предупреждение", "Система"])
        self._log_filter.currentTextChanged.connect(self._rebuild_log_display)
        ctrl.addWidget(self._log_filter)
        lay.addLayout(ctrl)

        self._log_edit = QPlainTextEdit(); self._log_edit.setReadOnly(True)
        self._log_edit.setMaximumBlockCount(5000)
        lay.addWidget(self._log_edit)
        return tab

    def _log(self, msg, level="INFO"): self._log_worker.add(msg, level)

    def _append_log_entry(self, entry: LogEntry):
        self._all_logs.append(entry)
        if self._passes_filter(entry, self._log_filter.currentText()):
            col = LOG_COLORS.get(entry.level, DARK_TEXT)
            self._log_edit.appendHtml(f'<span style="color:{col}">[{entry.ts}] {entry.msg}</span>')
            sb = self._log_edit.verticalScrollBar(); sb.setValue(sb.maximum())

    @staticmethod
    def _passes_filter(e: LogEntry, ft: str) -> bool:
        m = {"Все": None, "Инфо": "INFO", "Успех": "SUCCESS",
             "Ошибка": "ERROR", "Предупреждение": "WARNING", "Система": "SYSTEM"}
        req = m.get(ft); return req is None or e.level == req

    def _rebuild_log_display(self):
        self._log_edit.clear(); ft = self._log_filter.currentText()
        for e in self._all_logs:
            if self._passes_filter(e, ft):
                col = LOG_COLORS.get(e.level, DARK_TEXT)
                self._log_edit.appendHtml(f'<span style="color:{col}">[{e.ts}] {e.msg}</span>')
        sb = self._log_edit.verticalScrollBar(); sb.setValue(sb.maximum())

    def _clear_log(self): self._all_logs.clear(); self._log_edit.clear(); self._log("📝 Лог очищен", "SYSTEM")
    def _copy_log(self):
        QApplication.clipboard().setText("\n".join(f"[{e.ts}] {e.msg}" for e in self._all_logs))
        self._log("📋 Скопировано", "SYSTEM")
    def _save_log(self):
        p, _ = QFileDialog.getSaveFileName(self, "Сохранить", f"log_{datetime.now():%Y%m%d_%H%M%S}.txt", "*.txt")
        if p:
            with open(p, "w", encoding="utf-8") as f:
                f.write("\n".join(f"[{e.ts}][{e.level}] {e.msg}" for e in self._all_logs))
            self._log(f"💾 {p}", "SYSTEM")

    # ── Настройки ─────────────────────────────────────────────────────────────
    def _create_settings_page(self):
        tab = QWidget(); tab.setStyleSheet(f"background:{DARK_BG};")
        outer = QVBoxLayout(tab); outer.setContentsMargins(16, 12, 16, 12); outer.setSpacing(0)

        tabs = QTabWidget()
        tabs.addTab(self._create_general_settings(), "⚙️  Общие")
        tabs.addTab(self._email_widget, "📧  Email-уведомления")
        outer.addWidget(tabs)
        return tab

    def _create_general_settings(self):
        w = QWidget(); w.setStyleSheet(f"background:{DARK_CARD};")
        outer = QHBoxLayout(w); outer.setContentsMargins(20, 20, 20, 20); outer.setSpacing(16)
        left = QVBoxLayout(); left.setSpacing(12)

        # Папка
        pc = QWidget(); pc.setObjectName("card")
        pg = QGridLayout(pc); pg.setContentsMargins(16, 14, 16, 16); pg.setSpacing(10)
        pg.addWidget(self._section("📁  ПАПКА РЕЗЕРВНЫХ КОПИЙ"), 0, 0, 1, 3)
        pg.addWidget(QLabel("Путь:"), 1, 0)
        self._path_edit = QLineEdit("D:\\HyperV_Backups"); pg.addWidget(self._path_edit, 1, 1)
        bb = QPushButton("📂"); bb.setFixedWidth(36); bb.setToolTip("Выбрать папку")
        bb.clicked.connect(self._browse_dir); pg.addWidget(bb, 1, 2)
        left.addWidget(pc)

        # Параметры
        oc = QWidget(); oc.setObjectName("card")
        og = QGridLayout(oc); og.setContentsMargins(16, 14, 16, 16); og.setSpacing(10)
        og.addWidget(self._section("⚙️  ПАРАМЕТРЫ БЭКАПА"), 0, 0, 1, 2)
        og.addWidget(QLabel("Тип:"), 1, 0)
        self._backup_type_combo = QComboBox()
        self._backup_type_combo.addItems(["Полный экспорт (Export-VM)", "Только конфигурация (-Capture)"])
        og.addWidget(self._backup_type_combo, 1, 1)
        og.addWidget(QLabel("Хранить:"), 2, 0)
        self._retention_spin = QSpinBox(); self._retention_spin.setRange(1, 365)
        self._retention_spin.setValue(7); self._retention_spin.setSuffix(" дн.")
        og.addWidget(self._retention_spin, 2, 1)
        og.addWidget(QLabel("Режим по умолч.:"), 3, 0)
        self._default_mode_combo = QComboBox()
        self._default_mode_combo.addItems(["Вместе (одна папка)", "По одной ВМ"])
        og.addWidget(self._default_mode_combo, 3, 1)
        left.addWidget(oc)

        # Доп.
        nc = QWidget(); nc.setObjectName("card")
        nl = QVBoxLayout(nc); nl.setContentsMargins(16, 14, 16, 14); nl.setSpacing(8)
        nl.addWidget(self._section("🔔  ДОПОЛНИТЕЛЬНО"))
        self._notif_check = QCheckBox("Показывать системные уведомления")
        self._notif_check.setChecked(True); nl.addWidget(self._notif_check)
        self._autostart_check = QCheckBox("Автозапуск с Windows")
        self._autostart_check.stateChanged.connect(self._toggle_autostart)
        nl.addWidget(self._autostart_check)
        left.addWidget(nc); left.addStretch()

        br = QHBoxLayout(); br.setSpacing(8)
        sv = QPushButton("  Сохранить")
        sv.setIcon(qicon("fa5s.save", "#ffffff", 13))
        sv.setObjectName("primary"); sv.clicked.connect(self._save_config)
        rs = QPushButton("  Сбросить")
        rs.setIcon(qicon("fa5s.undo", "#c8ccd8", 13))
        rs.clicked.connect(self._reset_config)
        br.addWidget(sv, 2); br.addWidget(rs, 1)
        left.addLayout(br)
        outer.addLayout(left)
        return w

    def _save_config(self):
        cfg = {}
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f: cfg = json.load(f)
            except Exception: cfg = {}
        cfg.update({
            "backup_path":        self._path_edit.text(),
            "retention_days":     self._retention_spin.value(),
            "backup_type_index":  self._backup_type_combo.currentIndex(),
            "default_mode":       self._default_mode_combo.currentIndex(),
            "notifications":      self._notif_check.isChecked(),
            "autostart":          self._autostart_check.isChecked(),
        })
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(cfg, f, indent=2, ensure_ascii=False)
            self._log("💾 Настройки сохранены", "SUCCESS")
            QMessageBox.information(self, "Успех", "Настройки сохранены")
        except Exception as e: self._log(f"❌ {e}", "ERROR")

    def _reset_config(self):
        self._path_edit.setText("D:\\HyperV_Backups"); self._retention_spin.setValue(7)
        self._backup_type_combo.setCurrentIndex(0); self._default_mode_combo.setCurrentIndex(0)
        self._notif_check.setChecked(True); self._autostart_check.setChecked(False)
        self._log("↺ Настройки сброшены", "SYSTEM")

    def _load_config(self):
        if not os.path.exists(CONFIG_FILE): return
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f: cfg = json.load(f)
            self._path_edit.setText(str(cfg.get("backup_path", "D:\\HyperV_Backups")))
            try: self._retention_spin.setValue(int(cfg.get("retention_days", 7)))
            except: self._retention_spin.setValue(7)
            try: self._backup_type_combo.setCurrentIndex(int(cfg.get("backup_type_index", 0)))
            except: pass
            try: self._default_mode_combo.setCurrentIndex(int(cfg.get("default_mode", 0)))
            except: pass
            rn = cfg.get("notifications", True)
            self._notif_check.setChecked(rn if isinstance(rn, bool) else str(rn).lower() in ("true", "1", "yes"))
            self._autostart_check.blockSignals(True)
            ra = cfg.get("autostart", False)
            self._autostart_check.setChecked(ra if isinstance(ra, bool) else str(ra).lower() in ("true", "1", "yes"))
            self._autostart_check.blockSignals(False)
            try: self._run_mode_combo.setCurrentIndex(int(cfg.get("default_mode", 0)))
            except: pass
            self._log("📂 Настройки загружены", "SYSTEM")
        except json.JSONDecodeError:
            self._log("⚠️ Файл настроек повреждён, сброс", "WARNING"); self._reset_config()
        except Exception as e:
            self._log(f"⚠️ Ошибка загрузки настроек: {e}", "WARNING")

    def _save_schedule(self):
        try:
            with open(SCHEDULE_FILE, "w", encoding="utf-8") as f:
                json.dump(self._scheduled_jobs, f, indent=2, ensure_ascii=False)
        except Exception as e: self._log(f"⚠️ {e}", "WARNING")

    def _load_schedule(self):
        if not os.path.exists(SCHEDULE_FILE): return
        try:
            with open(SCHEDULE_FILE, "r", encoding="utf-8") as f:
                self._scheduled_jobs = json.load(f)
            self._refresh_jobs_list()
            for j in self._scheduled_jobs:
                if j.get("enabled", True): self._register_schedule(j)
            self._log(f"📅 Заданий загружено: {len(self._scheduled_jobs)}", "SYSTEM")
        except Exception as e: self._log(f"⚠️ {e}", "WARNING")

    # ── Утилиты ───────────────────────────────────────────────────────────────
    def _start_scheduler(self):
        self._sched_timer = QTimer()
        self._sched_timer.timeout.connect(schedule.run_pending)
        self._sched_timer.start(1000)

    def _browse_dir(self):
        p = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if p:
            self._path_edit.setText(p); self._log(f"📁 {p}", "INFO")
            QTimer.singleShot(200, self._refresh_backups_table)

    def _check_admin(self):
        try: is_adm = bool(ctypes.windll.shell32.IsUserAnAdmin())
        except: is_adm = False
        if is_adm: self._log("✅ Запущено с правами Администратора", "SUCCESS")
        else:      self._log("⚠️ Рекомендуется запуск от Администратора", "WARNING")

    def _toggle_autostart(self, state: int):
        en = state == Qt.CheckState.Checked.value
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_KEY, 0, winreg.KEY_SET_VALUE)
            if en:
                winreg.SetValueEx(key, AUTOSTART_NAME, 0, winreg.REG_SZ,
                                  f'"{sys.executable}" "{os.path.abspath(__file__)}"')
            else:
                try: winreg.DeleteValue(key, AUTOSTART_NAME)
                except FileNotFoundError: pass
            winreg.CloseKey(key)
        except Exception as e: self._log(f"⚠️ Автозапуск: {e}", "WARNING")

    def _notify(self, title, msg):
        if not self._notif_check.isChecked(): return
        try: self.tray_icon.showMessage(title, msg, QSystemTrayIcon.MessageIcon.Information, 3000)
        except: pass

    def _set_status(self, text, busy=False):
        self.statusBar().showMessage(f"{text}..." if busy else text)

    def _save_last_backup(self):
        try:
            with open(LAST_BACKUP_FILE, "w") as f:
                json.dump({"last_backup": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                           "path": self._path_edit.text(),
                           "type": self._backup_type_combo.currentText()}, f, indent=2)
        except Exception as e: self._log(f"⚠️ {e}", "WARNING")


# ─── Точка входа ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # !! QApplication создаётся ПЕРВЫМ — до любых QPixmap/QIcon/qtawesome !!
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Теперь можно безопасно инициализировать qtawesome
    if not _try_load_qta():
        import subprocess as _sp
        print("Устанавливаю qtawesome...")
        ret = _sp.run([sys.executable, "-m", "pip", "install", "qtawesome"], check=False)
        if ret.returncode == 0:
            _try_load_qta()
        else:
            print("Не удалось установить qtawesome. Иконки будут недоступны.")

    # Тёмная палитра для Fusion
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window,          QColor(DARK_BG))
    palette.setColor(QPalette.ColorRole.WindowText,      QColor(DARK_TEXT))
    palette.setColor(QPalette.ColorRole.Base,            QColor(DARK_SURFACE))
    palette.setColor(QPalette.ColorRole.AlternateBase,   QColor(DARK_CARD))
    palette.setColor(QPalette.ColorRole.ToolTipBase,     QColor(DARK_CARD))
    palette.setColor(QPalette.ColorRole.ToolTipText,     QColor(DARK_TEXT))
    palette.setColor(QPalette.ColorRole.Text,            QColor(DARK_TEXT))
    palette.setColor(QPalette.ColorRole.Button,          QColor(DARK_SURFACE))
    palette.setColor(QPalette.ColorRole.ButtonText,      QColor(DARK_TEXT))
    palette.setColor(QPalette.ColorRole.BrightText,      QColor(DARK_ERROR))
    palette.setColor(QPalette.ColorRole.Link,            QColor(DARK_ACCENT))
    palette.setColor(QPalette.ColorRole.Highlight,       QColor(DARK_ACCENT))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
    app.setPalette(palette)

    window = HyperVBackupApp()
    window.show()
    sys.exit(app.exec())
