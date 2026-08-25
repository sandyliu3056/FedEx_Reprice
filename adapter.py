# -*- coding: utf-8 -*-
"""無頭轉接層:在沒有視窗環境的伺服器上,原封不動地使用桌面版
FedEx_RepricingTool.py 的計價引擎。

做法:先把 tkinter / customtkinter 換成假模組(所有 UI 呼叫都是 no-op),
再 import 桌面版檔案,然後把它的九張費率表換成純資料的 ShimTable。
桌面版檔案本身一行都不用改 —— 它永遠是計價邏輯唯一的來源。
"""
import importlib.util
import json
import os
import sys
import tempfile
import types


# ---------------------------------------------------------------- 假 tkinter
def _make_fake_tk():
    tk = types.ModuleType("tkinter")

    class _Var:
        def __init__(self, *a, **k):
            self._v = k.get("value", "")

        def get(self):
            return self._v

        def set(self, v):
            self._v = v

        def trace_add(self, *a, **k):
            return "t0"

        def trace_remove(self, *a, **k):
            pass

    class _W:
        def __init__(self, *a, **k):
            pass

        # Treeview 會被 get_children / item 走訪,要回可迭代的東西
        def get_children(self, *a, **k):
            return ()

        def item(self, *a, **k):
            return {"values": ()}

        def bbox(self, *a, **k):
            return (0, 0, 0, 0)

        def __getattr__(self, n):
            return lambda *a, **k: None

        def __setitem__(self, k, v):
            pass

        def __getitem__(self, k):
            return ""

    tk.StringVar = tk.BooleanVar = tk.IntVar = tk.DoubleVar = _Var
    tk.Tk = tk.Toplevel = tk.Frame = tk.Canvas = tk.Label = tk.Button = _W
    tk.Entry = tk.Text = tk.Listbox = tk.Scrollbar = tk.Menu = tk.PhotoImage = _W
    tk.Widget = _W
    tk.TclError = type("TclError", (Exception,), {})
    for name in ("END", "INSERT", "W", "E", "N", "S", "NSEW", "BOTH", "X", "Y",
                 "LEFT", "RIGHT", "TOP", "BOTTOM", "CENTER", "DISABLED",
                 "NORMAL", "VERTICAL", "HORIZONTAL", "WORD", "NONE", "SOLID",
                 "ACTIVE", "FLAT", "RAISED", "SUNKEN", "GROOVE", "RIDGE"):
        setattr(tk, name, name.lower())

    ttk = types.ModuleType("tkinter.ttk")
    for name in ("Frame", "Label", "Button", "Entry", "Combobox", "Treeview",
                 "Notebook", "Style", "Scrollbar", "LabelFrame", "Checkbutton",
                 "Radiobutton", "Progressbar", "Separator", "PanedWindow",
                 "Spinbox"):
        setattr(ttk, name, _W)

    filedialog = types.ModuleType("tkinter.filedialog")
    filedialog.asksaveasfilename = lambda *a, **k: ""
    filedialog.askopenfilename = lambda *a, **k: ""
    filedialog.askopenfilenames = lambda *a, **k: ()
    filedialog.askdirectory = lambda *a, **k: ""

    class _MB:
        LAST_ERROR = None

        @staticmethod
        def showinfo(*a, **k):
            pass

        @staticmethod
        def showwarning(*a, **k):
            pass

        @staticmethod
        def showerror(title="", message="", *a, **k):
            _MB.LAST_ERROR = f"{title}: {message}"

        @staticmethod
        def askyesno(*a, **k):
            return False

        @staticmethod
        def askokcancel(*a, **k):
            return False

    messagebox = types.ModuleType("tkinter.messagebox")
    for n in ("showinfo", "showwarning", "showerror", "askyesno", "askokcancel"):
        setattr(messagebox, n, getattr(_MB, n))
    messagebox._MB = _MB

    tk.ttk = ttk
    tk.filedialog = filedialog
    tk.messagebox = messagebox

    ctk = types.ModuleType("customtkinter")
    ctk.set_appearance_mode = lambda *a, **k: None
    ctk.set_default_color_theme = lambda *a, **k: None
    ctk.CTk = _W

    def _ctk_getattr(n):
        # 雙底線屬性(__file__、__path__ 之類)要照實喊沒有,
        # 否則 Python 的 inspect 會拿到假元件而爆掉。
        if n.startswith("__"):
            raise AttributeError(n)
        return _W  # 其他 CTk 元件一律回假元件
    ctk.__getattr__ = _ctk_getattr
    ctk.__file__ = __file__

    return tk, ttk, filedialog, messagebox, ctk, _W


_tk, _ttk, _fd, _mb, _ctk, _W = _make_fake_tk()
for _name, _m in (("tkinter", _tk), ("tkinter.ttk", _ttk),
                  ("tkinter.filedialog", _fd), ("tkinter.messagebox", _mb),
                  ("customtkinter", _ctk)):
    sys.modules[_name] = _m

_HERE = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location(
    "fedex_engine", os.path.join(_HERE, "FedEx_RepricingTool.py"))
mod = importlib.util.module_from_spec(_spec)
sys.modules["fedex_engine"] = mod
_spec.loader.exec_module(mod)


# ---------------------------------------------------------------- 資料表替身
class ShimTable:
    """RateTableFrame 的純資料版:只做 get_rows / load_rows。"""

    def __init__(self, columns):
        self.columns = list(columns)
        self.rows = []

    def get_rows(self):
        return [list(map(str, r)) for r in self.rows]

    def load_rows(self, rows):
        self.rows = [list(map(lambda x: "" if x is None else str(x), r))
                     for r in (rows or [])]

    def clear(self):
        self.rows = []

    def add_row(self, values=None):
        self.rows.append(list(values or [""] * len(self.columns)))

    def save_table(self):
        pass


TABLE_COLUMNS = {
    "base_rate_table": ["Service", "Weight"] + [f"Zone {z}" for z in mod.ZONES],
    "ahs_rate_table": ["Service", "AHS Type", "Zone", "Fee"],
    "das_rate_table": ["Service", "DAS Type", "Zone", "Fee"],
    "oversize_rate_table": ["Service", "Zone", "Fee"],
    "residential_rate_table": ["Service", "Fee"],
    "signature_rate_table": ["Signature Type", "Fee"],
    "general_demand_table": ["Service", "Start Date", "End Date", "Fee"],
    "ahs_demand_table": ["Service", "Start Date", "End Date", "Fee"],
    "oversize_demand_table": ["Service", "Start Date", "End Date", "Fee"],
}

GLOBAL_RULES = [
    # (屬性, 中文標籤, 型別)
    ("dim_factor", "DIM 除數", float),
    ("fuel_percent", "燃油 %", float),
    ("ahs_weight_threshold", "AHS 重量門檻 (lb)", float),
    ("ahs_dimension_lg_limit", "AHS L+G 上限 (in)", float),
    ("ahs_dimension_longest_side", "AHS 最長邊 (in)", float),
    ("ahs_dimension_second_side", "AHS 次長邊 (in)", float),
    ("ahs_dimension_cubic_inches", "AHS 體積 (cu in)", float),
    ("ahs_min_billable_weight", "AHS 最低計費重 (lb)", float),
    ("oversize_longest_side", "Oversize 最長邊 (in)", float),
    ("oversize_actual_weight", "Oversize 實重 (lb)", float),
    ("oversize_cubic_inches", "Oversize 體積 (cu in)", float),
    ("length_girth_limit", "Oversize L+G 上限 (in)", float),
    ("oversize_min_billable_weight", "Oversize 最低計費重 (lb)", float),
]

GLOBAL_FLAGS = [
    ("use_oversize_longest", "Oversize 判定:最長邊"),
    ("use_oversize_weight", "Oversize 判定:實重"),
    ("use_oversize_cubic", "Oversize 判定:體積"),
    ("use_oversize_lg", "Oversize 判定:長+圍長"),
]


def new_app():
    """建立一個無頭引擎實例:桌面版類別 + 資料表替身 + 預設費率。"""
    app = mod.FedExRepricingTool(_W())
    for attr, cols in TABLE_COLUMNS.items():
        setattr(app, attr, ShimTable(cols))
    app.preview = _W()
    app._load_default_tables()      # 把桌面版的預設費率倒進替身表
    return app


def load_config_json(app, text):
    """吃桌面版存出來的 billing_tool_config.json(同一格式)。"""
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False,
                                     encoding="utf-8") as f:
        f.write(text)
        path = f.name
    try:
        app.config_path.set(path)
        app.load_config()
    finally:
        os.unlink(path)


def dump_config_json(app):
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False,
                                     encoding="utf-8") as f:
        path = f.name
    try:
        app.config_path.set(path)
        app.save_config()
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    finally:
        os.unlink(path)


class RunError(Exception):
    pass


def run_repricing(app, billing_bytes, filename):
    """執行桌面版的 run_repricing,回傳 (輸出 xlsx bytes, 摘要文字)。"""
    suffix = ".csv" if filename.lower().endswith(".csv") else ".xlsx"
    with tempfile.NamedTemporaryFile("wb", suffix=suffix, delete=False) as f:
        f.write(billing_bytes)
        in_path = f.name
    out_path = tempfile.mktemp(suffix=".xlsx")
    _mb._MB.LAST_ERROR = None
    old = mod.filedialog.asksaveasfilename
    try:
        app.billing_path.set(in_path)
        mod.filedialog.asksaveasfilename = lambda *a, **k: out_path
        app.run_repricing()
    finally:
        mod.filedialog.asksaveasfilename = old
        os.unlink(in_path)
    if _mb._MB.LAST_ERROR:
        raise RunError(_mb._MB.LAST_ERROR)
    if not os.path.exists(out_path):
        raise RunError(app.status_var.get() or "沒有產生輸出檔")
    with open(out_path, "rb") as f:
        data = f.read()
    os.unlink(out_path)
    summary = ""
    if hasattr(app, "run_summary_var"):
        summary = app.run_summary_var.get() or ""
    return data, summary


def make_billing_sample(app):
    """用桌面版的 export_billing_sample 產生範例帳單檔。"""
    out_path = tempfile.mktemp(suffix=".xlsx")
    old = mod.filedialog.asksaveasfilename
    try:
        mod.filedialog.asksaveasfilename = lambda *a, **k: out_path
        app.export_billing_sample()
    finally:
        mod.filedialog.asksaveasfilename = old
    if not os.path.exists(out_path):
        raise RunError("無法產生範例帳單")
    with open(out_path, "rb") as f:
        data = f.read()
    os.unlink(out_path)
    return data
