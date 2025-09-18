# app.py
# Aplikacja Zamówień — PySide6 (2025 modern)
# Uruchom:  pip install PySide6 pillow pywin32
# python app.py
# EXE (opcjonalnie):  pyinstaller --noconfirm --onefile --windowed --name Zamowienia app.py

from __future__ import annotations
import os, sys, sqlite3, time, re, ast, unicodedata as ud, math, csv
from typing import List, Optional, Tuple, Dict, Set
from datetime import datetime, timedelta

from PySide6.QtCore import (
    Qt, QAbstractTableModel, QModelIndex, QTimer, QByteArray, QPropertyAnimation,
    QEasingCurve, QSize, QSettings
)
from PySide6.QtGui import QAction, QColor, QBrush, QFont, QIcon, QPixmap
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableView,
    QToolBar, QMessageBox, QLineEdit, QCheckBox, QFileDialog, QLabel, QPushButton,
    QComboBox, QDateTimeEdit, QDialog, QFormLayout, QSpinBox, QGridLayout,
    QHeaderView, QMenu, QStyle
)

# ----- Druk / obraz -----
from PIL import Image, ImageDraw, ImageFont, ImageWin
import win32print, win32ui

APP_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_DIR  = os.path.join(APP_DIR, 'data')
DB_PATH   = os.path.join(DATA_DIR, 'db.sqlite3')
PROJ_DIR  = os.path.join(DATA_DIR, 'projekty')
EXPORT_DIR= os.path.join(APP_DIR, 'exporty')
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(PROJ_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)

# ===== UI akcenty / motywy =====
ACCENT = {
    "laser":   "#06b6d4",  # cyjan
    "nadruk":  "#ec4899",  # róż
    "projekt": "#f59e0b",  # bursztyn
}

DARK_QSS = f"""
QWidget {{
    background: #0f1115; color: #e5e7eb; font-size: 12px;
}}
QToolBar {{ background:#111318; border:0; padding:4px; }}
QLineEdit, QComboBox, QDateTimeEdit, QSpinBox, QPlainTextEdit, QTextEdit {{
    background:#1a1d24; border:1px solid #2a2f3a; padding:6px; border-radius:8px;
}}
QTableView {{
    background:#0f1115; gridline-color:#272b35; selection-background-color: #2f3543;
    selection-color:#fff; alternate-background-color:#10131a;
}}
QHeaderView::section {{
    background:#141823; color:#cbd5e1; padding:6px; border:0; border-bottom:1px solid #2a2f3a;
}}
QPushButton {{
    background:#1f2430; border:1px solid #2a2f3a; padding:6px 10px; border-radius:10px;
}}
QPushButton:hover {{ border-color:{ACCENT['laser']}; }}
QPushButton:disabled {{ color:#6b7280; }}
QCheckBox {{ spacing:8px; }}
"""

LIGHT_QSS = f"""
QWidget {{
    background: #f7f8fb; color: #111827; font-size: 12px;
}}
QToolBar {{ background:#ffffff; border:0; padding:4px; }}
QLineEdit, QComboBox, QDateTimeEdit, QSpinBox, QPlainTextEdit, QTextEdit {{
    background:#ffffff; border:1px solid #e5e7eb; padding:6px; border-radius:8px;
}}
QTableView {{
    background:#ffffff; gridline-color:#e5e7eb; selection-background-color: #cde3ff;
    selection-color:#111827; alternate-background-color:#f6f8fc;
}}
QHeaderView::section {{
    background:#f0f3f9; color:#111827; padding:6px; border:0; border-bottom:1px solid #e5e7eb;
}}
QPushButton {{
    background:#ffffff; border:1px solid #e5e7eb; padding:6px 10px; border-radius:10px;
}}
QPushButton:hover {{ border-color:{ACCENT['laser']}; }}
QPushButton:disabled {{ color:#9ca3af; }}
QCheckBox {{ spacing:8px; }}
"""

# ===== Słowniki =====
STYLES   = ["Nadruk", "Laser", "Projekt", "Kubki", "Przypinki", "Tryumf"]
PAYMENTS = ["Niezapłacone", "Zapłacone"]
CALLS    = ["Brak potrzeby", "Dzwonić", "Zadzwonione"]

# ===== Util =====
def now_ts() -> int:            return int(time.time())
def dt_to_ts(dt: datetime) -> int: return int(dt.timestamp())
def ts_to_dt(ts: int) -> datetime: return datetime.fromtimestamp(int(ts))
def midnight(d: datetime) -> datetime: return datetime(d.year, d.month, d.day)

def human_left(now: datetime, deadline: datetime) -> Tuple[str, bool]:
    t = midnight(now); dl = midnight(deadline)
    days = int(round((dl - t).total_seconds() / 86400.0))
    if days == 0: day_txt = "dzisiaj"
    elif days == 1: day_txt = "jutro"
    elif days == 2: day_txt = "2 dni"
    elif days > 2: day_txt = f"{days} dni"
    else: day_txt = f"{days} {'dzień' if abs(days)==1 else 'dni'}"
    delta = deadline - now
    total_min = int(delta.total_seconds() // 60)
    sign = "-" if total_min < 0 else ""
    total_min = abs(total_min)
    hh = total_min // 60
    mm = total_min % 60
    return f"{day_txt} ({sign}{hh}:{mm:02d}h)", (deadline < now)

# Kolory terminów
COLOR_OVERDUE  = QColor(128, 0, 128, 60)   # fiolet
COLOR_TODAY    = QColor(0, 102, 204, 60)   # niebieski
COLOR_TOMORROW = QColor(204, 0, 0, 60)     # czerwony
COLOR_TWO_DAYS = QColor(255, 204, 0, 60)   # żółty
COLOR_LATER    = QColor(0, 128, 0, 40)     # zielony
COLOR_DONE     = QColor(0, 150, 0, 70)     # zielony (zrobione)
COLOR_BLACKLIST= QColor(255, 0, 0, 28)     # leciutki czerwony dla „UWAGA”

def deadline_color(today: datetime, deadline: datetime) -> QColor:
    diff = int(round((midnight(deadline) - midnight(today)).total_seconds() / 86400.0))
    if diff < 0:  return COLOR_OVERDUE
    if diff == 0: return COLOR_TODAY
    if diff == 1: return COLOR_TOMORROW
    if diff == 2: return COLOR_TWO_DAYS
    return COLOR_LATER

def digits_only(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def format_phone(s: str) -> str:
    d = digits_only(s)
    if len(d) <= 3: return d
    if len(d) <= 6: return f"{d[:3]} {d[3:]}"
    return f"{d[:3]} {d[3:6]} {d[6:9]}".strip()

def norm_txt(s: str) -> str:
    s = (s or '').lower()
    s = ud.normalize('NFD', s)
    return ''.join(ch for ch in s if not ud.category(ch).startswith('M'))

# Prosty bezpieczny evaluator
class SafeEval(ast.NodeVisitor):
    allowed = (ast.Expression, ast.BinOp, ast.UnaryOp, ast.Num, ast.Load,
               ast.Add, ast.Sub, ast.Mult, ast.Div, ast.USub, ast.UAdd,
               ast.Pow, ast.Mod, ast.FloorDiv, ast.Constant, ast.Call, ast.Name)
    def __init__(self):
        super().__init__()
        self.env = { 'round': round, 'abs': abs, 'pow': pow, 'min': min, 'max': max }
    def visit(self, node):
        if not isinstance(node, self.allowed):
            raise ValueError("Niedozwolony element w wyrażeniu")
        return super().visit(node)
    def eval(self, expr: str) -> float:
        expr = (expr or "").replace(',', '.')
        if not expr.strip(): return 0.0
        tree = ast.parse(expr, mode='eval')
        return float(self._eval(tree.body))
    def _eval(self, node):
        if isinstance(node, ast.BinOp):
            a = self._eval(node.left); b = self._eval(node.right)
            if isinstance(node.op, ast.Add): return a + b
            if isinstance(node.op, ast.Sub): return a - b
            if isinstance(node.op, ast.Mult): return a * b
            if isinstance(node.op, ast.Div): return a / b
            if isinstance(node.op, ast.Pow): return a ** b
            if isinstance(node.op, ast.Mod): return a % b
            if isinstance(node.op, ast.FloorDiv): return a // b
        if isinstance(node, ast.UnaryOp):
            v = self._eval(node.operand)
            if isinstance(node.op, ast.UAdd): return +v
            if isinstance(node.op, ast.USub): return -v
        if isinstance(node, ast.Num): return node.n
        if isinstance(node, ast.Constant) and isinstance(node.value, (int,float)): return node.value
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in self.env:
            fn = self.env[node.func.id]; args = [self._eval(a) for a in node.args]; return fn(*args)
        if isinstance(node, ast.Name) and node.id in self.env: return self.env[node.id]
        raise ValueError('Wyrażenie nieobsługiwane')

SAFE = SafeEval()
def price_to_str(v: float) -> str:
    return f"{v:,.2f} zł".replace(',', 'X').replace('.', ',').replace('X', '.')

# ===== DB + migracje =====
def init_db():
    con = sqlite3.connect(DB_PATH); cur = con.cursor()
    # główna tabela
    cur.execute("""
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        price_expr TEXT,
        price_value REAL DEFAULT 0,
        project_ready INTEGER DEFAULT 0,
        start_ts INTEGER NOT NULL,
        deadline_ts INTEGER NOT NULL,
        style TEXT NOT NULL,
        payment TEXT DEFAULT 'Niezapłacone',
        gilding INTEGER DEFAULT 0,
        case_extra INTEGER DEFAULT 0,
        call_status TEXT DEFAULT 'Brak potrzeby',
        phone TEXT,
        email TEXT,
        cdr_path TEXT,
        est_minutes INTEGER,
        status TEXT DEFAULT 'Do zrobienia',
        updated_ts INTEGER
    )""")
    # brakujące kolumny (migracje)
    cols = {r[1] for r in cur.execute("PRAGMA table_info(orders)").fetchall()}
    if "first_name" not in cols:
        cur.execute("ALTER TABLE orders ADD COLUMN first_name TEXT")
    if "last_name" not in cols:
        cur.execute("ALTER TABLE orders ADD COLUMN last_name TEXT")
    con.commit()

    # blacklist (UWAGA)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS blacklist (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        phone TEXT,
        comment TEXT,
        created_ts INTEGER
    )""")
    # szablony, logi (opcjonalne)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS templates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        style TEXT,
        est_minutes INTEGER,
        price_expr TEXT,
        payment TEXT,
        gilding INTEGER,
        case_extra INTEGER
    )""")
    cur.execute("""
    CREATE TABLE IF NOT EXISTS logs (
        ts INTEGER,
        action TEXT,
        order_id INTEGER,
        info TEXT
    )""")
    con.commit(); con.close()

def db_all(query: str, params=()) -> List[dict]:
    con = sqlite3.connect(DB_PATH); cur = con.cursor()
    cur.execute(query, params)
    cols = [d[0] for d in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    con.close()
    return rows

def db_exec(query: str, params=()):
    con = sqlite3.connect(DB_PATH); cur = con.cursor()
    cur.execute(query, params)
    con.commit(); con.close()

def log_action(action: str, order_id: int, info: str=""):
    db_exec("INSERT INTO logs (ts,action,order_id,info) VALUES (?,?,?,?)", (now_ts(), action, order_id, info))

def fetch_month_keys() -> List[str]:
    rows = db_all("SELECT start_ts FROM orders")
    keys = {f"{ts_to_dt(r['start_ts']).year}-{ts_to_dt(r['start_ts']).month:02d}" for r in rows}
    if not keys:
        d = datetime.now(); keys.add(f"{d.year}-{d.month:02d}")
    return sorted(list(keys))

def month_range(key: str) -> Tuple[int,int]:
    y, m = map(int, key.split('-'))
    start = datetime(y, m, 1)
    nxt   = datetime(y+1, 1, 1) if m == 12 else datetime(y, m+1, 1)
    return dt_to_ts(start), dt_to_ts(nxt)

def fetch_orders(month_key: str, include_removed=False) -> List[dict]:
    a,b = month_range(month_key)
    if include_removed:
        q = "SELECT * FROM orders WHERE start_ts>=? AND start_ts<? ORDER BY deadline_ts ASC"
        return db_all(q, (a,b))
    else:
        q = "SELECT * FROM orders WHERE status!='Usunięte' AND start_ts>=? AND start_ts<? ORDER BY deadline_ts ASC"
        return db_all(q, (a,b))

def upsert_order(o: dict) -> int:
    con = sqlite3.connect(DB_PATH); cur = con.cursor()
    o = o.copy()
    o['price_value']   = float(o.get('price_value') or 0)
    o['project_ready'] = 1 if o.get('project_ready') else 0
    o['gilding']       = 1 if o.get('gilding') else 0
    o['case_extra']    = 1 if o.get('case_extra') else 0
    o['updated_ts']    = now_ts()
    cols = ['name','price_expr','price_value','project_ready','start_ts','deadline_ts','style','payment',
            'gilding','case_extra','call_status','phone','email','cdr_path','est_minutes','status','updated_ts',
            'first_name','last_name']
    vals = [o.get(c) for c in cols]
    if o.get('id'):
        sets = ','.join([f"{c}=?" for c in cols]); cur.execute(f"UPDATE orders SET {sets} WHERE id=?", vals + [o['id']]); oid = o['id']
    else:
        qmarks = ','.join(['?']*len(cols)); cur.execute(f"INSERT INTO orders ({','.join(cols)}) VALUES ({qmarks})", vals); oid = cur.lastrowid
    con.commit(); con.close()
    log_action("UPSERT", oid, o.get('name',""))
    return oid

def mark_done(order_id: int):
    db_exec("UPDATE orders SET status='Zrobione', updated_ts=? WHERE id=?", (now_ts(), order_id))
    log_action("DONE", order_id)

def mark_undone(order_id: int):
    db_exec("UPDATE orders SET status='Do zrobienia', updated_ts=? WHERE id=?", (now_ts(), order_id))
    log_action("UNDONE", order_id)

def soft_delete(order_id: int):
    db_exec("UPDATE orders SET status='Usunięte', updated_ts=? WHERE id=?", (now_ts(), order_id))
    log_action("DELETE", order_id)

# ===== Blacklist (UWAGA) =====
def get_blacklist() -> Dict[str,str]:
    rows = db_all("SELECT phone, comment FROM blacklist")
    return {digits_only(r['phone']): (r.get('comment') or '') for r in rows if digits_only(r.get('phone'))}

def add_blacklist(phone: str, comment: str):
    db_exec("INSERT INTO blacklist (phone, comment, created_ts) VALUES (?,?,?)", (digits_only(phone), comment.strip(), now_ts()))

def remove_blacklist(phone: str):
    db_exec("DELETE FROM blacklist WHERE phone=?", (digits_only(phone),))

# ===== Model / kolumny =====
COLUMNS = [
    ('price_value','Cena'),
    ('name','Nazwa'),
    ('first_last','Klient'),
    ('project_ready','Projekt'),
    ('start_ts','Start'),
    ('deadline_ts','Deadline'),
    ('left','Pozostało'),
    ('style','Styl'),
    ('payment','Płatność'),
    ('gilding','Złocenie'),
    ('case_extra','Etui'),
    ('phone','Telefon'),
    ('email','Email'),
    ('call_status','Dzwonić'),
    ('cdr_path','CDR'),
    ('print','Drukuj'),        # NOWA kolumna
    ('actions','Zrobione'),
]

class OrdersModel(QAbstractTableModel):
    def __init__(self, rows: List[dict], angry_set: Set[str]):
        super().__init__()
        self.rows = rows
        self.today = datetime.now()
        self.angry_set = angry_set  # zestaw zblacklistowanych numerów (digits)

    def rowCount(self, parent=QModelIndex()): return len(self.rows)
    def columnCount(self, parent=QModelIndex()): return len(COLUMNS)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole: return None
        return COLUMNS[section][1] if orientation == Qt.Horizontal else section + 1

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid(): return None
        r = self.rows[index.row()]
        key = COLUMNS[index.column()][0]

        # delikatne czerwone tło dla wierszy z „UWAGA”
        if role == Qt.BackgroundRole:
            if digits_only(r.get('phone')) in self.angry_set:
                return COLOR_BLACKLIST
            if key == 'left':
                if r.get('status') == 'Zrobione': return COLOR_DONE
                return deadline_color(self.today, ts_to_dt(r['deadline_ts']))
            # subtelne pasy
            if index.row() % 2 == 0:
                return QColor(255,255,255,0)  # tło ustawi QSS
            else:
                return QColor(255,255,255,0)

        if key == 'left' and role == Qt.DisplayRole:
            if r.get('status') == 'Zrobione':
                return 'Zrobione'
            txt, _ = human_left(self.today, ts_to_dt(r['deadline_ts']))
            return txt

        if key == 'price_value':
            if role == Qt.DisplayRole:     return price_to_str(r.get('price_value') or 0)
            if role == Qt.TextAlignmentRole: return Qt.AlignRight | Qt.AlignVCenter

        if key == 'project_ready' and role == Qt.DisplayRole:
            return 'gotowy' if r.get('project_ready') else 'do zrobienia'

        if key in ('start_ts','deadline_ts') and role == Qt.DisplayRole:
            return ts_to_dt(r[key]).strftime('%Y-%m-%d %H:%M')

        if key == 'gilding'    and role == Qt.DisplayRole: return 'tak' if r.get('gilding') else 'nie'
        if key == 'case_extra' and role == Qt.DisplayRole: return 'tak' if r.get('case_extra') else 'nie'
        if key == 'phone' and role == Qt.DisplayRole: return format_phone(r.get('phone') or '')
        if key == 'first_last' and role == Qt.DisplayRole:
            f = (r.get('first_name') or '').strip()
            l = (r.get('last_name') or '').strip()
            return (f"{f} {l}".strip() or '—')

        if key == 'cdr_path':
            if role == Qt.DisplayRole:   return 'OTWÓRZ' if r.get('cdr_path') else '—'
            if role == Qt.ForegroundRole and r.get('cdr_path'): return QBrush(QColor(59,130,246))
            if role == Qt.FontRole and r.get('cdr_path'):
                f = QFont(); f.setUnderline(True); return f

        if key == 'print':
            if role == Qt.DisplayRole: return 'Drukuj/PDF'
            if role == Qt.TextAlignmentRole: return Qt.AlignCenter

        if key == 'actions':
            if role == Qt.DisplayRole:   return 'Przywróć' if r.get('status') == 'Zrobione' else 'Zrobione'
            if role == Qt.TextAlignmentRole: return Qt.AlignCenter

        if role == Qt.DisplayRole: return r.get(key)
        return None

    def sort(self, column, order):
        key = COLUMNS[column][0]
        reverse = order == Qt.DescendingOrder
        if key in ('start_ts','deadline_ts','price_value'):
            self.layoutAboutToBeChanged.emit()
            self.rows.sort(key=lambda r: r.get(key) or 0, reverse=reverse)
            self.layoutChanged.emit()
        else:
            super().sort(column, order)

# ===== Dialog edycji =====
class OrderDialog(QDialog):
    def __init__(self, parent=None, row: Optional[dict]=None):
        super().__init__(parent)
        self.setWindowTitle('Zamówienie')
        self.row = row or {}
        lay = QGridLayout(self)

        self.e_price_expr = QLineEdit(self.row.get('price_expr',''))
        self.lbl_price_val= QLabel(price_to_str(self.row.get('price_value') or 0))
        self.e_name   = QLineEdit(self.row.get('name',''))
        self.e_first  = QLineEdit(self.row.get('first_name','') or '')
        self.e_last   = QLineEdit(self.row.get('last_name','') or '')
        self.chk_ready= QCheckBox('Projekt gotowy'); self.chk_ready.setChecked(bool(self.row.get('project_ready')))
        self.dt_start = QDateTimeEdit(ts_to_dt(self.row.get('start_ts') or now_ts())); self.dt_start.setCalendarPopup(True)
        self.dt_deadline = QDateTimeEdit(ts_to_dt(self.row.get('deadline_ts') or now_ts())); self.dt_deadline.setCalendarPopup(True)

        self.cmb_style= QComboBox(); self.cmb_style.addItems(STYLES)
        if self.row.get('style'): self.cmb_style.setCurrentText(self.row['style'])
        self.spin_est = QSpinBox(); self.spin_est.setRange(0, 24*60); self.spin_est.setValue(int(self.row.get('est_minutes') or (20 if self.cmb_style.currentText()=="Projekt" else 60)))

        self.e_phone  = QLineEdit(format_phone(self.row.get('phone') or ''))
        self.e_email  = QLineEdit(self.row.get('email') or '')
        self.cmb_pay  = QComboBox(); self.cmb_pay.addItems(PAYMENTS); self.cmb_pay.setCurrentText(self.row.get('payment','Niezapłacone'))
        self.chk_gild = QCheckBox('Złocenie'); self.chk_gild.setChecked(bool(self.row.get('gilding')))
        self.chk_case = QCheckBox('Etui');     self.chk_case.setChecked(bool(self.row.get('case_extra')))
        self.cmb_call = QComboBox(); self.cmb_call.addItems(CALLS); self.cmb_call.setCurrentText(self.row.get('call_status','Brak potrzeby'))

        self.e_cdr    = QLineEdit(self.row.get('cdr_path') or '')
        btn_browse = QPushButton('Wybierz CDR'); btn_open = QPushButton('OTWÓRZ')
        btn_browse.clicked.connect(self.choose_cdr); btn_open.clicked.connect(self.open_cdr)

        form = QFormLayout()
        form.addRow('Cena (może mieć nawiasy i + - * /):', self.e_price_expr)
        form.addRow('Wynik:', self.lbl_price_val)
        form.addRow('Nazwa projektu:', self.e_name)
        form.addRow('Imię:', self.e_first)
        form.addRow('Nazwisko:', self.e_last)
        form.addRow('', self.chk_ready)
        form.addRow('Start:', self.dt_start)
        form.addRow('Deadline:', self.dt_deadline)
        form.addRow('Styl robienia:', self.cmb_style)
        form.addRow('Szacowany czas [min]:', self.spin_est)

        form2 = QFormLayout()
        form2.addRow('Telefon:', self.e_phone)
        form2.addRow('Email:', self.e_email)
        form2.addRow('Płatność:', self.cmb_pay)
        form2.addRow('', self.chk_gild)
        form2.addRow('', self.chk_case)
        form2.addRow('Kontakt:', self.cmb_call)
        file_row = QHBoxLayout(); file_row.addWidget(self.e_cdr); file_row.addWidget(btn_browse); file_row.addWidget(btn_open)
        form2.addRow('Plik CDR:', file_row)

        left, right = QWidget(), QWidget()
        left.setLayout(form); right.setLayout(form2)
        lay.addWidget(left, 0,0); lay.addWidget(right, 0,1)

        btn_box = QHBoxLayout()
        btn_save = QPushButton('Zapisz'); btn_cancel = QPushButton('Anuluj')
        btn_box.addStretch(1); btn_box.addWidget(btn_cancel); btn_box.addWidget(btn_save)
        lay.addLayout(btn_box, 1,0,1,2)
        btn_cancel.clicked.connect(self.reject); btn_save.clicked.connect(self.accept)

        self.e_price_expr.textChanged.connect(self.update_price)
        self.cmb_call.currentTextChanged.connect(self.on_call_changed)
        self.resize(980, 560)

    def on_call_changed(self, txt):
        if txt == 'Zadzwonione':
            QMessageBox.information(self, 'Info', 'Status zamówienia zostanie ustawiony na ZROBIONE przy zapisie.')

    def update_price(self):
        try:   val = SAFE.eval(self.e_price_expr.text())
        except Exception: val = 0
        self.lbl_price_val.setText(price_to_str(val))

    def choose_cdr(self):
        path, _ = QFileDialog.getOpenFileName(self, 'Wybierz plik CDR', '', 'CorelDRAW (*.cdr)')
        if path:
            name = os.path.basename(path); dest = os.path.join(PROJ_DIR, name)
            if not os.path.exists(dest):
                try:
                    import shutil; shutil.copy2(path, dest)
                except Exception: dest = path
            self.e_cdr.setText(dest)

    def open_cdr(self):
        path = self.e_cdr.text().strip()
        if path and os.path.exists(path):
            try: os.startfile(path)
            except Exception as e: QMessageBox.warning(self, 'Błąd', f'Nie udało się otworzyć pliku\n{e}')
        else:
            QMessageBox.information(self, 'Info', 'Brak pliku CDR')

    def values(self) -> Optional[dict]:
        name = self.e_name.text().strip()
        if not name:
            QMessageBox.warning(self, 'Błąd', 'Podaj nazwę projektu'); return None
        start = self.dt_start.dateTime().toPython(); deadline = self.dt_deadline.dateTime().toPython()
        if deadline < start:
            QMessageBox.warning(self, 'Błąd', 'Deadline nie może być wcześniejszy niż start'); return None
        try: price_val = SAFE.eval(self.e_price_expr.text())
        except Exception: price_val = 0
        row = dict(
            id=self.row.get('id'), name=name, price_expr=self.e_price_expr.text(),
            price_value=float(round(price_val, 2)), project_ready=bool(self.chk_ready.isChecked()),
            start_ts=dt_to_ts(start), deadline_ts=dt_to_ts(deadline),
            style=self.cmb_style.currentText(), payment=self.cmb_pay.currentText(),
            gilding=bool(self.chk_gild.isChecked()), case_extra=bool(self.chk_case.isChecked()),
            call_status=self.cmb_call.currentText(), phone=digits_only(self.e_phone.text()),
            email=self.e_email.text().strip(), cdr_path=self.e_cdr.text().strip(),
            est_minutes=int(self.spin_est.value()), status=self.row.get('status','Do zrobienia'),
            first_name=self.e_first.text().strip(), last_name=self.e_last.text().strip()
        )
        if row['call_status'] == 'Zadzwonione': row['status'] = 'Zrobione'
        return row

# ===== Manager „UWAGA” =====
class BlacklistDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("UWAGA — czarna lista")
        self.resize(420, 360)
        v = QVBoxLayout(self)
        # listing
        self.list = QTableView()
        self.list.setSelectionBehavior(QTableView.SelectRows)
        self.list.setSelectionMode(QTableView.SingleSelection)
        v.addWidget(self.list)

        # model prościutki z danych
        rows = db_all("SELECT phone, comment, created_ts FROM blacklist ORDER BY created_ts DESC")
        class LM(QAbstractTableModel):
            def rowCount(self, *_): return len(rows)
            def columnCount(self, *_): return 3
            def data(self, idx, role=Qt.DisplayRole):
                if not idx.isValid(): return None
                if role==Qt.DisplayRole:
                    r = rows[idx.row()]
                    return [format_phone(r['phone'] or ''), r['comment'] or '', datetime.fromtimestamp(r['created_ts']).strftime('%Y-%m-%d %H:%M')][idx.column()]
                return None
            def headerData(self, s,o,role=Qt.DisplayRole):
                if role!=Qt.DisplayRole: return None
                return ["Telefon","Komentarz","Dodano"][s] if o==Qt.Horizontal else s+1
        self.list.setModel(LM())
        self.list.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # formularz
        f = QFormLayout()
        self.e_phone = QLineEdit(); self.e_comment = QLineEdit()
        f.addRow("Telefon:", self.e_phone); f.addRow("Komentarz:", self.e_comment)
        v.addLayout(f)
        row = QHBoxLayout()
        b_add = QPushButton("Dodaj"); b_del = QPushButton("Usuń zaznaczony"); b_close=QPushButton("Zamknij")
        row.addWidget(b_add); row.addWidget(b_del); row.addStretch(1); row.addWidget(b_close)
        v.addLayout(row)
        b_close.clicked.connect(self.accept)
        def do_add():
            ph = digits_only(self.e_phone.text()); cm = self.e_comment.text().strip()
            if not ph:
                QMessageBox.information(self,"UWAGA","Podaj numer.")
                return
            add_blacklist(ph, cm); self.accept()
        def do_del():
            idx = self.list.currentIndex()
            if not idx.isValid(): return
            phone_disp = self.list.model().data(self.list.model().index(idx.row(),0))
            remove_blacklist(phone_disp); self.accept()
        b_add.clicked.connect(do_add); b_del.clicked.connect(do_del)

# ===== Wycena =====
MATERIAL_RATES = {
    "sklejka": 0.20,
    "plexi": 0.30,
    "laminat": 0.35,
    "grawerton": 0.25,
    "mosiądz": 0.60,
    "anodowane": 0.30,
    "aluminium": 0.25,
    # specjalny przypadek: szkło — stała cena (35–55 zł zależnie od powierzchni)
    "szkło": 0.0,
}
MATERIALS = list(MATERIAL_RATES.keys())
MODES = ["Grawer z materiałem","Sam materiał","Sam grawer"]
NA_CZYM = ["tworzywo sztuczne","szkło","drewno","skóra","metal normalnie","metal głęboko","srebro/złoto","długopis markowy","długopis zwykły"]
MIN_ORDER_TOTAL = 40
SAM_MATERIAL_UPUST = 0.15
ZNAMIONOWA_FEE = 40

class WycenaDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Wycena")
        f = QFormLayout(self)
        self.cmb_mode = QComboBox(); self.cmb_mode.addItems(MODES)
        self.cmb_mat = QComboBox(); self.cmb_mat.addItems(MATERIALS)
        self.cmb_na = QComboBox(); self.cmb_na.addItems(NA_CZYM)
        self.chk_two = QCheckBox("Grawer 2 strony")
        self.e_h = QLineEdit(); self.e_w = QLineEdit()
        self.e_qty = QLineEdit("1"); self.e_rabat = QLineEdit("0")
        self.cmb_round = QComboBox(); self.cmb_round.addItems(["1","10"])
        self.chk_nameplate = QCheckBox(f"Tabliczka znamionowa (+{ZNAMIONOWA_FEE} zł)")
        self.lbl_sum = QLabel(""); self.lbl_total = QLabel("0 zł"); self.lbl_unit = QLabel("0 zł/szt")
        f.addRow("Tryb:", self.cmb_mode)
        f.addRow("Materiał:", self.cmb_mat)
        f.addRow("Na czym:", self.cmb_na)
        f.addRow("", self.chk_two)
        f.addRow("Wysokość [cm]:", self.e_h)
        f.addRow("Szerokość [cm]:", self.e_w)
        f.addRow("Ilość:", self.e_qty)
        f.addRow("Rabat [%]:", self.e_rabat)
        f.addRow("Zaokrąglanie:", self.cmb_round)
        f.addRow("", self.chk_nameplate)
        f.addRow(QLabel("Podsumowanie:"), self.lbl_sum)
        f.addRow(QLabel("Cena całkowita:"), self.lbl_total)
        f.addRow(QLabel("Cena za sztukę:"), self.lbl_unit)
        # zamknij (bez drukuj)
        row = QHBoxLayout()
        self.btn_close = QPushButton("Zamknij")
        row.addStretch(1); row.addWidget(self.btn_close)
        f.addRow(row)

        # sygnały
        for w in (self.cmb_mode,self.cmb_mat,self.cmb_na,self.chk_two,self.e_h,self.e_w,self.e_qty,self.e_rabat,self.cmb_round,self.chk_nameplate):
            if isinstance(w, (QComboBox,)):
                w.currentIndexChanged.connect(self.recalc)
            elif isinstance(w, (QCheckBox,)):
                w.stateChanged.connect(self.recalc)
            else:
                w.textChanged.connect(self.recalc)
        self.btn_close.clicked.connect(self.accept)
        self.resize(460, 520)
        self._update_visibility()
        self.recalc()
        self.cmb_mode.currentIndexChanged.connect(self._update_visibility)

    def _update_visibility(self):
        mode = self.cmb_mode.currentText()

        # widoczności
        show_mat = (mode in ("Grawer z materiałem","Sam materiał"))
        show_na  = (mode == "Sam grawer")
        show_hw  = (mode in ("Grawer z materiałem","Sam materiał")) or (
            mode=="Sam grawer" and self.cmb_na.currentText() in ("drewno","skóra")
        )
        two_ok = (mode=="Sam grawer" and self.cmb_na.currentText() in ("metal normalnie","srebro/złoto"))

        # Materiał
        for w in (self._lbl_mat, self.cmb_mat):
            if w: w.setVisible(show_mat)

        # Na czym
        for w in (self._lbl_na, self.cmb_na):
            if w: w.setVisible(show_na)

        # Wymiary
        for w in (self._lbl_h, self.e_h, self._lbl_w, self.e_w):
            if w: w.setVisible(show_hw)

        # Grawer 2 strony
        self.chk_two.setVisible(two_ok)


    def parse(self, s: str, default=0.0) -> float:
        try:
            s = (s or "").replace(",", ".")
            return float(s) if s.strip() else default
        except Exception:
            return default

    def calc_sam_grawer_unit(self, rodzaj: str, area: float, qty: int, two: bool) -> float:
        if rodzaj == "szkło": return 45
        if rodzaj in ["drewno","skóra","tworzywo sztuczne"]:
            if area <= 150: return 40
            elif area > 1000: return 55
            else: return 55
        if rodzaj == "metal normalnie": return 60 if two else 40
        if rodzaj == "metal głęboko": return 50
        if rodzaj == "srebro/złoto": return 60 if two else 40
        if rodzaj == "długopis markowy": return 40
        if rodzaj == "długopis zwykły":
            if qty >= 100: return max(3, 1.5)
            elif qty >= 50: return max(5, 1.5)
            elif qty >= 20: return max(10, 1.5)
            else: return 1.5
        return 40

    def recalc(self):
        mode = self.cmb_mode.currentText()
        mat  = self.cmb_mat.currentText()
        na   = self.cmb_na.currentText()
        h = self.parse(self.e_h.text())
        w = self.parse(self.e_w.text())
        qty = max(1, int(self.parse(self.e_qty.text(), 1)))
        rabat = max(0.0, self.parse(self.e_rabat.text()))
        step = max(1, int(self.cmb_round.currentText()))
        area = max(0.0, h*w)
        if mode == "Grawer z materiałem":
            if mat == "szkło":  # stała cena za sztukę
                base = 35 if area<=150 else (55 if area>=1000 else 45)
            else:
                rate = MATERIAL_RATES[mat]; base = area * rate
            summary = f"Grawer z materiałem, {mat}, {qty} szt, {h}x{w} cm"
        elif mode == "Sam materiał":
            if mat == "szkło":
                base = (35 if area<=150 else (55 if area>=1000 else 45)) * (1 - SAM_MATERIAL_UPUST)
            else:
                base = area * MATERIAL_RATES[mat] * (1 - SAM_MATERIAL_UPUST)
            summary = f"Sam materiał, {mat}, aut. −15%, {qty} szt, {h}x{w} cm"
        else:
            two = self.chk_two.isChecked()
            base = self.calc_sam_grawer_unit(na, area, qty, two)
            sides = ", 2 strony" if (na in ["metal normalnie","srebro/złoto"] and two) else ""
            dims  = f", {h}x{w} cm" if na in ["drewno","skóra"] else ""
            summary = f"Sam grawer, {na}{sides}, {qty} szt{dims}"
        discount_unit = base * rabat/100.0
        price_unit = int(round((base - discount_unit)/step))*step
        total = max(MIN_ORDER_TOTAL, price_unit * qty)
        if self.chk_nameplate.isChecked(): total += ZNAMIONOWA_FEE
        self.lbl_sum.setText(summary)
        self.lbl_total.setText(f"{total} zł"); self.lbl_unit.setText(f"{price_unit} zł/szt")

# ===== Stats / Raport =====
class StatsDialog(QDialog):
    def __init__(self, month_key: str, angry_set: Set[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Statystyki — {month_key}")
        v = QVBoxLayout(self)
        rows = fetch_orders(month_key, include_removed=True)
        now = datetime.now()
        # podstawy
        total_value = sum(r.get('price_value') or 0 for r in rows)
        total_count = len(rows)
        done = [r for r in rows if r.get('status') == 'Zrobione']
        open_ = [r for r in rows if r.get('status') != 'Usunięte' and r.get('status') != 'Zrobione']
        overdue = [r for r in rows if ts_to_dt(r['deadline_ts']) < now and r.get('status')!='Zrobione']
        # mediana / średnia
        vals = sorted([(r.get('price_value') or 0) for r in rows])
        avg = (sum(vals)/len(vals)) if vals else 0
        med = (vals[len(vals)//2] if vals else 0) if len(vals)%2==1 else ((vals[len(vals)//2-1]+vals[len(vals)//2])/2 if vals else 0)
        # rozkład stylów
        styles_cnt = {s:0 for s in STYLES}
        styles_val = {s:0.0 for s in STYLES}
        for r in rows:
            s = r.get('style'); styles_cnt[s] = styles_cnt.get(s,0)+1
            styles_val[s] = styles_val.get(s,0.0)+(r.get('price_value') or 0.0)
        # top klient (po telefonie/email)
        groups: Dict[str,float] = {}
        for r in rows:
            key = digits_only(r.get('phone') or '') or (r.get('email') or '').lower()
            groups[key] = groups.get(key,0.0) + (r.get('price_value') or 0.0)
        top_client_val = max(groups.values()) if groups else 0.0
        top_client_key = next((k for k,v in groups.items() if v==top_client_val), "")
        # najdroższe pojedyncze
        if rows:
            max_row = max(rows, key=lambda r:(r.get('price_value') or 0))
        else:
            max_row = None
        # angry count
        angry_count = sum(1 for r in rows if digits_only(r.get('phone')) in angry_set)

        labels = [
            f"Zleceń łącznie: {total_count}",
            f"Zrobione: {len(done)}   •   Otwarte: {len(open_)}   •   Po terminie: {len(overdue)}",
            f"Suma wartości: {price_to_str(total_value)}   •   Średnia: {price_to_str(avg)}   •   Mediana: {price_to_str(med)}",
            "Rozkład stylów (liczba): " + ", ".join([f"{k}: {styles_cnt.get(k,0)}" for k in STYLES]),
            "Rozkład stylów (wartość): " + ", ".join([f"{k}: {price_to_str(styles_val.get(k,0.0))}" for k in STYLES]),
            f"TOP klient: {top_client_key or '—'}  —  {price_to_str(top_client_val)}",
            f"Najdroższe zlecenie: {(max_row.get('name') if max_row else '—')}  —  {price_to_str((max_row.get('price_value') if max_row else 0) or 0)}",
            f"Wkurwionych klientów (wg czarnej listy w tym miesiącu): {angry_count}",
        ]
        for t in labels:
            v.addWidget(QLabel(t))

        btn_pdf = QPushButton("Raport PDF (exporty/)")
        v.addWidget(btn_pdf)
        btn_pdf.clicked.connect(lambda: self.make_pdf(month_key, rows))  # bez angry_count

        btn_close = QPushButton("Zamknij")
        v.addWidget(btn_close)
        btn_close.clicked.connect(self.accept)
        self.resize(640, 420)

    def make_pdf(self, month_key: str, rows: List[dict]):
        W,H = 1240, 1754
        im = Image.new("RGB", (W,H), "white")
        d = ImageDraw.Draw(im)

        def font_try(names, size_px):
            try:
                win_fonts = os.path.join(os.environ.get("WINDIR","C:\\Windows"), "Fonts")
                for n in names:
                    p = os.path.join(win_fonts,n)
                    if os.path.exists(p): return ImageFont.truetype(p, size_px)
            except Exception:
                pass
            return ImageFont.load_default()

        title = font_try(["segoeuib.ttf","arialbd.ttf"], 36)
        f     = font_try(["segoeui.ttf","arial.ttf"], 22)
        y = 60
        d.text((60,y), f"Raport miesiąca: {month_key}", font=title, fill="black"); y+=60

        # metryki
        total_value = sum(r.get('price_value') or 0 for r in rows)
        done = [r for r in rows if r.get('status') == 'Zrobione']
        open_ = [r for r in rows if r.get('status') != 'Usunięte' and r.get('status') != 'Zrobione']
        vals = sorted([(r.get('price_value') or 0) for r in rows])
        avg = (sum(vals)/len(vals)) if vals else 0
        med = (vals[len(vals)//2] if vals else 0) if len(vals)%2==1 else ((vals[len(vals)//2-1]+vals[len(vals)//2])/2 if vals else 0)

        # top klient (po telefonie/email)
        groups = {}
        for r in rows:
            key = (re.sub(r"\\D","", r.get('phone') or "")) or (r.get('email') or "").lower()
            if not key: key = "(brak kontaktu)"
            groups[key] = groups.get(key,0.0) + (r.get('price_value') or 0.0)
        top_client_key, top_client_val = ("—", 0.0)
        if groups:
            top_client_key = max(groups, key=lambda k: groups[k])
            top_client_val = groups[top_client_key]

        # najdroższe pojedyncze
        max_row = max(rows, key=lambda r:(r.get('price_value') or 0)) if rows else None

        # rozkład stylów
        style_cnt = {}
        style_val = {}
        for r in rows:
            s = r.get('style') or "—"
            style_cnt[s] = style_cnt.get(s,0)+1
            style_val[s] = style_val.get(s,0.0)+(r.get('price_value') or 0.0)

        lines = [
            f"Zamówień: {len(rows)}   •   Zrobione: {len(done)}   •   Otwarte: {len(open_)}",
            f"Suma wartości: {price_to_str(total_value)}   •   Średnia: {price_to_str(avg)}   •   Mediana: {price_to_str(med)}",
            f"TOP klient (tel/email): {top_client_key} — {price_to_str(top_client_val)}",
            f"Najdroższe zlecenie: {(max_row.get('name') if max_row else '—')} — {price_to_str((max_row.get('price_value') if max_row else 0) or 0)}",
            "Rozkład stylów (liczba): " + ", ".join([f"{k}: {style_cnt[k]}" for k in sorted(style_cnt)]),
            "Rozkład stylów (wartość): " + ", ".join([f"{k}: {price_to_str(style_val[k])}" for k in sorted(style_val)]),
        ]
        for t in lines:
            d.text((60,y), t, font=f, fill="black"); y+=36

        y += 12
        d.text((60,y), "Top 10 (najbliższe terminy):", font=title, fill="black"); y+=44
        for i, r in enumerate(sorted(rows, key=lambda x: x['deadline_ts'])[:10], start=1):
            d.text((60,y), f"{i}. {r['name']} — {datetime.fromtimestamp(r['deadline_ts']).strftime('%d.%m %H:%M')}", font=f, fill="black"); y+=30

        # logo
        try:
            logo_path = os.path.join(APP_DIR, "logo.png")
            if os.path.exists(logo_path):
                lg = Image.open(logo_path).convert("RGBA")
                tw = 260; scale = tw / lg.width; th = int(lg.height*scale)
                lg = lg.resize((tw,th), Image.Resampling.LANCZOS)
                im.paste(lg, (W-60-tw, 60), lg)
        except Exception:
            pass

        path = os.path.join(EXPORT_DIR, f"raport_{month_key}.pdf")
        im.save(path, "PDF", resolution=150.0)
        QMessageBox.information(self, "Raport PDF", f"Zapisano: {path}")

# ===== MainWindow =====
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Zamówienia')
        init_db()

        # settings
        self.settings = QSettings("MarkolGrawer","ZamowieniaApp")
        theme = self.settings.value("ui/theme","dark")
        self.apply_theme(theme)

        # blacklist cache
        self.blacklist = get_blacklist()
        self.angry_set: Set[str] = set(self.blacklist.keys())

        # miesiące
        self.months = fetch_month_keys(); self.current_month = self.months[-1]

        # toolbar
        self.toolbar = QToolBar(); self.addToolBar(self.toolbar)
        # logo
        self.logo_label = QLabel()
        logo_path = os.path.join(APP_DIR, "logo.png")
        if os.path.exists(logo_path):
            pm = QPixmap(logo_path)
            self.logo_label.setPixmap(pm.scaledToHeight(28, Qt.SmoothTransformation))
            self.toolbar.addWidget(self.logo_label)
        self.toolbar.addSeparator()

        self.cmb_month = QComboBox(); self.cmb_month.addItems(self.months); self.cmb_month.setCurrentText(self.settings.value("ui/month", self.current_month)); self.toolbar.addWidget(self.cmb_month)
        self.search = QLineEdit(); self.search.setPlaceholderText('Szukaj: nazwa, telefon, email'); self.search.setText(self.settings.value("ui/search","")); self.toolbar.addWidget(self.search)
        self.chk_only_open = QCheckBox('Pokaż tylko nie zrobione'); self.chk_only_open.setChecked(self.settings.value("ui/only_open","true")=="true"); self.toolbar.addWidget(self.chk_only_open)

        self.warn_label = QLabel(""); self.toolbar.addWidget(self.warn_label)

        self.toolbar.addSeparator()
        act_new   = QAction('Nowe', self)
        act_edit  = QAction('Edytuj', self)
        act_done  = QAction('Zrobione/Przywróć', self)
        act_delete= QAction('Usuń', self)
        act_export= QAction('Eksport TXT', self)
        act_open_cdr = QAction('Otwórz CDR', self)
        act_stats = QAction('Statystyki', self)
        act_report= QAction('Raport PDF', self)
        act_price = QAction('Wycena', self)
        act_settings = QAction('Ustawienia', self)
        act_theme = QAction('Motyw Jasny/Ciemny', self)
        act_csv_exp = QAction('Eksport CSV (Sheets)', self)
        act_csv_imp = QAction('Import CSV (Sheets)', self)
        act_black   = QAction('UWAGA (czarna lista)', self)
        for a in (act_new, act_edit, act_done, act_delete, act_open_cdr, act_export, act_stats, act_report, act_price, act_csv_exp, act_csv_imp, act_black, act_settings, act_theme):
            self.toolbar.addAction(a)

        central = QWidget(); self.setCentralWidget(central)
        v = QVBoxLayout(central)

        # Tabela
        self.table = QTableView()
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.setSelectionBehavior(QTableView.SelectRows)   # zaznaczanie poziome
        self.table.setSelectionMode(QTableView.SingleSelection)
        v.addWidget(self.table)

        # Sygnaly
        self.cmb_month.currentTextChanged.connect(self.reload)
        self.search.textChanged.connect(self.apply_filter)
        self.chk_only_open.stateChanged.connect(self.apply_filter)
        act_new.triggered.connect(self.add_order)
        act_edit.triggered.connect(self.edit_selected)
        act_done.triggered.connect(self.toggle_selected_done)
        act_delete.triggered.connect(self.delete_selected)
        act_export.triggered.connect(self.export_txt)
        act_open_cdr.triggered.connect(self.open_selected_cdr)
        act_stats.triggered.connect(self.open_stats)
        act_report.triggered.connect(self.open_stats_report)
        act_price.triggered.connect(self.open_wycena)
        act_settings.triggered.connect(self.open_settings)
        act_theme.triggered.connect(self.toggle_theme)
        act_csv_exp.triggered.connect(self.export_csv)
        act_csv_imp.triggered.connect(self.import_csv)
        act_black.triggered.connect(self.manage_blacklist)

        self.table.doubleClicked.connect(self.on_double_click)
        self.table.clicked.connect(self.on_click)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.table_menu)

        # Timer odświeżania (pozostało + popup 1h)
        self.reminded: Set[int] = set()
        self.timer = QTimer(self); self.timer.timeout.connect(self.tick_minute); self.timer.start(60_000)

        self.reload()
        self.resize(1360, 820)
        # przywróć geometrię
        geo = self.settings.value("ui/geom")
        if isinstance(geo, QByteArray): self.restoreGeometry(geo)

    def apply_theme(self, theme: str):
        if theme == "light":
            self.setStyleSheet(LIGHT_QSS)
        else:
            self.setStyleSheet(DARK_QSS)

    def toggle_theme(self):
        cur = self.settings.value("ui/theme","dark")
        new = "light" if cur=="dark" else "dark"
        self.settings.setValue("ui/theme", new)
        self.apply_theme(new)

    # Data loading and filtering
    def reload(self):
        rows = fetch_orders(self.cmb_month.currentText(), include_removed=False)
        if self.chk_only_open.isChecked():
            rows = [r for r in rows if r.get('status') != 'Zrobione']
        q = self.search.text().strip()
        if q:
            nq = norm_txt(q); dq = digits_only(q); out = []
            for r in rows:
                if nq in norm_txt(r.get('name')) or nq in norm_txt(r.get('email')) or nq in norm_txt((r.get('first_name') or '')+" "+(r.get('last_name') or '')) or (dq and dq in (r.get('phone') or '')):
                    out.append(r)
            rows = out
        self.blacklist = get_blacklist()
        self.angry_set = set(self.blacklist.keys())
        self.model = OrdersModel(rows, self.angry_set)
        self.table.setModel(self.model)
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setMinimumSectionSize(70)
        self.update_warning()

    def update_warning(self):
        # Ryzyko spóźnienia: ktoś ma <= 24h
        now = datetime.now()
        rows = self.model.rows
        risky = [r for r in rows if ts_to_dt(r['deadline_ts']) - now <= timedelta(hours=24) and r.get('status')!='Zrobione']
        self.warn_label.setText(f"<span style='color:#f59e0b'>UWAGA ryzyko spóźnienia: {len(risky)}</span>" if risky else "")

    def apply_filter(self): self.reload()

    # Kliknięcia / dwuklik w tabeli
    def on_click(self, index: QModelIndex):
        if not index.isValid(): return
        r = self.model.rows[index.row()]
        key = COLUMNS[index.column()][0]
        if key == 'cdr_path': self.open_cdr_path(r)
        elif key == 'actions': self.toggle_done_row(r)
        elif key == 'print':  self.open_print_dialog(r)

    def on_double_click(self, index: QModelIndex):
        if not index.isValid(): return
        r = self.model.rows[index.row()]
        key = COLUMNS[index.column()][0]
        if key == 'name' or key == 'first_last':
            self.open_dialog_for(r)
        elif key in ('phone','email'):
            txt = format_phone(r.get('phone') or '') if key=='phone' else (r.get('email') or '')
            QApplication.clipboard().setText(txt)
            self.statusBar().showMessage(f'Skopiowano: {txt}', 1500)
        elif key == 'cdr_path':
            self.open_cdr_path(r)
        elif key == 'actions':
            self.toggle_done_row(r)
        elif key == 'project_ready':
            r['project_ready'] = 0 if r.get('project_ready') else 1; upsert_order(r); self.reload()
        elif key == 'gilding':
            r['gilding'] = 0 if r.get('gilding') else 1; upsert_order(r); self.reload()
        elif key == 'case_extra':
            r['case_extra'] = 0 if r.get('case_extra') else 1; upsert_order(r); self.reload()
        elif key == 'call_status':
            seq = ["Brak potrzeby","Dzwonić","Zadzwonione"]
            cur = r.get('call_status') or "Brak potrzeby"
            nxt = seq[(seq.index(cur)+1)%len(seq)]
            r['call_status'] = nxt
            if nxt == "Zadzwonione": r['status']="Zrobione"
            upsert_order(r); self.reload()

    def table_menu(self, pos):
        idx = self.table.indexAt(pos)
        if not idx.isValid(): return
        r = self.model.rows[idx.row()]
        m = QMenu(self)
        m.addAction("Edytuj", lambda: self.open_dialog_for(r))
        m.addAction("Zrobione" if r.get('status')!='Zrobione' else "Przywróć",
                    lambda: self.toggle_done_row(r))
        m.addAction("Otwórz CDR", lambda: self.open_cdr_path(r))
        m.addAction("Otwórz folder", lambda: self.open_cdr_folder(r))
        m.addSeparator()
        m.addAction("Duplikuj", lambda: self.duplicate_row(r))
        m.addAction("Drukuj / PDF", lambda: self.open_print_dialog(r))
        m.exec(self.table.viewport().mapToGlobal(pos))

    def duplicate_row(self, r: dict):
        cp = r.copy(); cp.pop('id', None)
        cp['start_ts'] = now_ts(); cp['status'] = 'Do zrobienia'
        upsert_order(cp); self.reload()

    def open_cdr_path(self, r: dict):
        path = r.get('cdr_path')
        if path and os.path.exists(path):
            os.startfile(path)
            log_action("OPEN_CDR", r['id'], path)
        else: QMessageBox.information(self, 'Info', 'Brak pliku CDR')

    def open_cdr_folder(self, r: dict):
        path = r.get('cdr_path')
        if path and os.path.exists(path):
            os.startfile(os.path.dirname(path))
        else: QMessageBox.information(self, 'Info', 'Brak pliku CDR')

    # Toolbar akcje
    def add_order(self):
        d = OrderDialog(self, row={
            'price_expr':'','price_value':0.0,'project_ready':False,'start_ts': now_ts(),'deadline_ts': now_ts(),
            'style':'Nadruk','payment':'Niezapłacone','gilding':False,'case_extra':False,'call_status':'Brak potrzeby',
            'est_minutes':20,'status':'Do zrobienia','first_name':'','last_name':''
        })
        if d.exec() == QDialog.Accepted:
            row = d.values()
            if not row: return
            upsert_order(row); self.reload()
    def open_selected_cdr(self):
        r = self.get_selected_row()
        if not r:
            QMessageBox.information(self, 'Info', 'Wybierz zamówienie z tabeli')
            return
        self.open_cdr_path(r)


    def get_selected_row(self) -> Optional[dict]:
        idx = self.table.currentIndex()
        if not idx.isValid(): return None
        return self.model.rows[idx.row()]

    def edit_selected(self):
        r = self.get_selected_row()
        if not r: QMessageBox.information(self, 'Info', 'Wybierz zamówienie z tabeli'); return
        self.open_dialog_for(r)

    def open_dialog_for(self, r: dict):
        d = OrderDialog(self, r)
        if d.exec() == QDialog.Accepted:
            row = d.values()
            if not row: return
            row['id'] = r.get('id')
            upsert_order(row)
            self.reload()

    def toggle_selected_done(self):
        r = self.get_selected_row()
        if not r: QMessageBox.information(self, 'Info', 'Wybierz zamówienie z tabeli'); return
        self.toggle_done_row(r)

    def toggle_done_row(self, r: dict):
        if r.get('status') == 'Zrobione': mark_undone(r['id'])
        else:                             mark_done(r['id'])
        self.reload()

    def delete_selected(self):
        r = self.get_selected_row()
        if not r: QMessageBox.information(self, 'Info', 'Wybierz zamówienie z tabeli'); return
        if QMessageBox.question(self, 'Potwierdź', 'Usunąć zamówienie?') == QMessageBox.Yes:
            soft_delete(r['id']); self.reload()

    def export_txt(self):
        m = self.cmb_month.currentText()
        rows = fetch_orders(m, include_removed=True)
        done = [r for r in rows if r.get('status') == 'Zrobione']
        lines = []
        for i, o in enumerate(done, start=1):
            lines.append(f"{i}. Nazwa: {o['name']}")
            lines.append(f"   Cena: {price_to_str(o.get('price_value') or 0)}")
            lines.append(f"   Klient: {(o.get('first_name') or '')} {(o.get('last_name') or '')}".strip())
            lines.append(f"   Projekt gotowy: {'Tak' if o.get('project_ready') else 'Nie'}")
            lines.append(f"   Start: {ts_to_dt(o['start_ts']).strftime('%Y-%m-%d %H:%M')}")
            lines.append(f"   Deadline: {ts_to_dt(o['deadline_ts']).strftime('%Y-%m-%d %H:%M')}")
            lines.append(f"   Styl: {o.get('style')}")
            lines.append(f"   Płatność: {o.get('payment')}")
            lines.append(f"   Złocenie: {'Tak' if o.get('gilding') else 'Nie'}, Etui: {'Tak' if o.get('case_extra') else 'Nie'}")
            lines.append(f"   Dzwonić: {o.get('call_status')}")
            lines.append(f"   Telefon: {format_phone(o.get('phone') or '')}, Email: {o.get('email') or ''}")
            lines.append(f"   Plik: {os.path.basename(o.get('cdr_path') or '') or 'brak'}")
            lines.append("")
        path = os.path.join(EXPORT_DIR, f"export_{m}.txt")
        with open(path, 'w', encoding='utf-8') as f: f.write('\n'.join(lines))
        QMessageBox.information(self, 'Eksport', f'Zapisano do:\n{path}')

    def export_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "Eksport CSV (Google Sheets)", os.path.join(EXPORT_DIR,"orders.csv"), "CSV (*.csv)")
        if not path: return
        rows = fetch_orders(self.cmb_month.currentText(), include_removed=True)
        fields = ["id","name","first_name","last_name","price_value","price_expr","project_ready","start_ts","deadline_ts","style","payment","gilding","case_extra","call_status","phone","email","cdr_path","est_minutes","status"]
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            for r in rows:
                w.writerow({k:r.get(k) for k in fields})
        QMessageBox.information(self, "CSV", f"Zapisano: {path}")

    def import_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import CSV (Google Sheets)", "", "CSV (*.csv)")
        if not path: return
        with open(path, "r", encoding="utf-8") as f:
            rdr = csv.DictReader(f)
            for row in rdr:
                try:
                    # Minimalny mapping i sanityzacja
                    row['project_ready'] = int(row.get('project_ready') or 0)
                    row['gilding'] = int(row.get('gilding') or 0)
                    row['case_extra'] = int(row.get('case_extra') or 0)
                    row['start_ts'] = int(row.get('start_ts') or now_ts())
                    row['deadline_ts'] = int(row.get('deadline_ts') or now_ts())
                    row['price_value'] = float(row.get('price_value') or 0)
                    row['est_minutes'] = int(row.get('est_minutes') or 0)
                    row['phone'] = digits_only(row.get('phone') or "")
                    row['id'] = None  # zawsze jako nowe
                    upsert_order(row)
                except Exception as e:
                    QMessageBox.warning(self, "CSV", f"Pominięto wiersz (błąd): {e}")
        self.reload()
        QMessageBox.information(self, "CSV", "Zaimportowano dane.")

    def open_stats(self):
        StatsDialog(self.cmb_month.currentText(), self.angry_set, self).exec()

    def open_stats_report(self):
        StatsDialog(self.cmb_month.currentText(), self.angry_set, self).make_pdf(
            self.cmb_month.currentText(), fetch_orders(self.cmb_month.currentText(), include_removed=True)
        )

    def open_wycena(self):
        WycenaDialog(self).exec()

    def open_settings(self):
        d = QDialog(self); d.setWindowTitle("Ustawienia")
        f = QFormLayout(d)
        self.chk_hour = QCheckBox("Popup gdy zostaje 1 godzina"); self.chk_hour.setChecked(self.settings.value("notif/hr","true")=="true")
        f.addRow(self.chk_hour)
        row = QHBoxLayout(); b1=QPushButton("Anuluj"); b2=QPushButton("Zapisz"); row.addStretch(1); row.addWidget(b1); row.addWidget(b2)
        f.addRow(row)
        b1.clicked.connect(d.reject)
        def save():
            self.settings.setValue("notif/hr","true" if self.chk_hour.isChecked() else "false")
            d.accept()
        b2.clicked.connect(save)
        d.exec()

    def manage_blacklist(self):
        BlacklistDialog(self).exec()
        # odśwież pamięć
        self.blacklist = get_blacklist()
        self.angry_set = set(self.blacklist.keys())
        self.reload()

    # ========== Druk ==========

    def open_print_dialog(self, r: dict):
        dlg = QDialog(self); dlg.setWindowTitle("Drukuj / Zapisz PDF")
        form = QFormLayout(dlg)
        # lista drukarek
        printers = []
        try:
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            for p in win32print.EnumPrinters(flags):
                printers.append(p[2])
        except Exception:
            pass
        if not printers:
            try: printers = [win32print.GetDefaultPrinter()]
            except Exception: printers = ["Microsoft Print to PDF"]
        cmb = QComboBox(); cmb.addItems(printers)
        form.addRow("Drukarka:", cmb)
        row = QHBoxLayout(); b_pdf = QPushButton("Zapisz PDF"); b_prn = QPushButton("Drukuj"); b_close = QPushButton("Anuluj")
        row.addWidget(b_pdf); row.addWidget(b_prn); row.addStretch(1); row.addWidget(b_close)
        form.addRow(row)
        def go_pdf():
            path, _ = QFileDialog.getSaveFileName(self, "Zapisz kartę jako PDF", os.path.join(EXPORT_DIR, f"karta_{r['id'] or 'nowe'}.pdf"), "PDF (*.pdf)")
            if not path: return
            self.render_card_to_pdf(r, path); dlg.accept()
        def go_print():
            self.render_card_to_printer(r, cmb.currentText()); dlg.accept()
        b_pdf.clicked.connect(go_pdf); b_prn.clicked.connect(go_print); b_close.clicked.connect(dlg.reject)
        dlg.exec()

    def render_card_to_pdf(self, r: dict, path: str):
        # A4 pion
        W,H = 1240, 1754
        page = Image.new("RGB",(W,H),"white")
        d = ImageDraw.Draw(page)
        def font_try(names, pt):
            try:
                fonts = os.path.join(os.environ.get("WINDIR","C:\\Windows"),"Fonts")
                for n in names:
                    p = os.path.join(fonts,n)
                    if os.path.exists(p): return ImageFont.truetype(p, int(pt*2))
            except Exception: ...
            return ImageFont.load_default()
        f_title = font_try(["segoeuib.ttf","arialbd.ttf"], 12)
        f_b = font_try(["segoeuib.ttf","arialbd.ttf"], 9)
        f = font_try(["segoeui.ttf","arial.ttf"], 8)
        y = 80
        def center(t,font, pad=18):
            nonlocal y
            wtxt = d.textbbox((0,0), t, font=font)[2]
            x = (W - wtxt)//2
            d.text((x,y), t, font=font, fill="black")
            y += pad*2
        # logo
        try:
            logo_path = os.path.join(APP_DIR,"logo.png")
            if os.path.exists(logo_path):
                lg = Image.open(logo_path).convert("RGBA")
                tw = 320; sc = tw/lg.width; th=int(lg.height*sc)
                lg = lg.resize((tw,th), Image.Resampling.LANCZOS)
                page.paste(lg,(W//2 - tw//2, 20), lg)
        except Exception: ...
        center("Karta zlecenia", f_title, 14)
        center(r.get('name') or "", f_title, 10)

        left = 120; ln = 34
        def row(label, value, bold=False):
            nonlocal y
            d.text((left,y), f"{label}:", font=f_b if bold else f_b, fill="black")
            d.text((left+320,y), value, font=f, fill="black")
            y += ln

        row("Klient", f"{(r.get('first_name') or '').strip()} {(r.get('last_name') or '').strip()}".strip())
        row("Telefon", format_phone(r.get('phone') or '')); row("Email", r.get('email') or '')
        row("Styl", r.get('style') or ''); row("Płatność", r.get('payment') or '')
        row("Złocenie", "Tak" if r.get('gilding') else "Nie"); row("Etui", "Tak" if r.get('case_extra') else "Nie")
        row("Projekt", "gotowy" if r.get('project_ready') else "do zrobienia")
        row("CDR", os.path.basename(r.get('cdr_path') or '') or '—')
        row("Start", ts_to_dt(r['start_ts']).strftime('%Y-%m-%d %H:%M')); row("Deadline", ts_to_dt(r['deadline_ts']).strftime('%Y-%m-%d %H:%M'))
        row("Szacowany czas", f"{int(r.get('est_minutes') or 0)} min")
        row("Cena", price_to_str(r.get('price_value') or 0), True)

        page.save(path, "PDF", resolution=150.0)
        QMessageBox.information(self, "PDF", f"Zapisano: {path}")

    def render_card_to_printer(self, r: dict, printer: str):
        try:
            hdc = win32ui.CreateDC(); hdc.CreatePrinterDC(printer)
            HORZRES=8; VERTRES=10; LOGPIXELSX=88; LOGPIXELSY=90
            w = hdc.GetDeviceCaps(HORZRES); h = hdc.GetDeviceCaps(VERTRES)
            dpi_x=hdc.GetDeviceCaps(LOGPIXELSX); dpi_y=hdc.GetDeviceCaps(LOGPIXELSY)
            page = Image.new("RGB",(w,h),"white"); d = ImageDraw.Draw(page)
            def px(pt): return max(10, int(pt*dpi_y/72))
            def font_try(names,size):
                fonts = os.path.join(os.environ.get("WINDIR","C:\\Windows"),"Fonts")
                for n in names:
                    p = os.path.join(fonts,n)
                    if os.path.exists(p): return ImageFont.truetype(p, px(size))
                return ImageFont.load_default()
            f_b = font_try(["segoeuib.ttf","arialbd.ttf"], 22)
            f = font_try(["segoeui.ttf","arial.ttf"], 16)
            y = 80
            def center(t,font, pad=8):
                nonlocal y
                wtxt = d.textbbox((0,0), t, font=font)[2]
                x = (w - wtxt)//2
                d.text((x,y), t, font=font, fill="black")
                y += px(pad)
            # logo
            try:
                logo_path = os.path.join(APP_DIR,"logo.png")
                if os.path.exists(logo_path):
                    lg = Image.open(logo_path).convert("RGBA")
                    tw = int(70/25.4*dpi_x); sc = tw/lg.width; th=int(lg.height*sc)
                    lg = lg.resize((tw,th), Image.Resampling.LANCZOS)
                    page.paste(lg,( (w-tw)//2, 10), lg)
                    y += px(6)
            except Exception: ...
            center("Karta zlecenia", f_b, 14)
            center(r.get('name') or "", f_b, 10)

            left = 120; ln = px(10)
            def row(label, value):
                nonlocal y
                d.text((left,y), f"{label}:", font=f, fill="black")
                d.text((left+px(100),y), value, font=f, fill="black")
                y += ln

            row("Klient", f"{(r.get('first_name') or '').strip()} {(r.get('last_name') or '').strip()}".strip())
            row("Telefon", format_phone(r.get('phone') or '')); row("Email", r.get('email') or '')
            row("Styl", r.get('style') or ''); row("Płatność", r.get('payment') or '')
            row("Złocenie", "Tak" if r.get('gilding') else "Nie"); row("Etui", "Tak" if r.get('case_extra') else "Nie")
            row("Projekt", "gotowy" if r.get('project_ready') else "do zrobienia")
            row("CDR", os.path.basename(r.get('cdr_path') or '') or '—')
            row("Start", ts_to_dt(r['start_ts']).strftime('%Y-%m-%d %H:%M')); row("Deadline", ts_to_dt(r['deadline_ts']).strftime('%Y-%m-%d %H:%M'))
            row("Szacowany czas", f"{int(r.get('est_minutes') or 0)} min")
            row("Cena", price_to_str(r.get('price_value') or 0))

            dib = ImageWin.Dib(page)
            hdc.StartDoc("Karta zlecenia"); hdc.StartPage()
            dib.draw(hdc.GetHandleOutput(), (0,0,w,h))
            hdc.EndPage(); hdc.EndDoc(); hdc.DeleteDC()
            QMessageBox.information(self,"Druk","Wysłano do drukarki.")
        except Exception as e:
            QMessageBox.warning(self,"Druk",f"Nie udało się wydrukować.\n{e}")

    # Timer co minutę: odśwież „Pozostało” + popup 1h
    def tick_minute(self):
        self.model.today = datetime.now()
        self.model.dataChanged.emit(self.model.index(0,0),
                                    self.model.index(max(0,self.model.rowCount()-1), self.model.columnCount()-1),
                                    [Qt.DisplayRole, Qt.BackgroundRole])
        if self.settings.value("notif/hr","true")=="true":
            for r in self.model.rows:
                if r.get('status')=='Zrobione': continue
                dl = ts_to_dt(r['deadline_ts']); left = dl - datetime.now()
                if 0 <= left.total_seconds() <= 3600 and r['id'] not in self.reminded:
                    self.reminded.add(r['id'])
                    QMessageBox.information(self, "Uwaga", f"Została godzina do terminu:\n{r['name']} ({dl.strftime('%d.%m %H:%M')})")
        self.update_warning()

    def closeEvent(self, e):
        self.settings.setValue("ui/geom", self.saveGeometry())
        self.settings.setValue("ui/month", self.cmb_month.currentText())
        self.settings.setValue("ui/search", self.search.text())
        self.settings.setValue("ui/only_open", "true" if self.chk_only_open.isChecked() else "false")
        super().closeEvent(e)

# ===== main =====
if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
