#!/usr/bin/env python3
"""
SML Platform Order Importer
นำเข้าคำสั่งซื้อจาก Shopee เข้าระบบ SML ผ่าน REST API
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import csv
import threading
import queue
import os
import sys
import pathlib
import datetime
import requests
import pandas as pd

# ─── App directory (works both dev and PyInstaller) ───────────────────────────
if getattr(sys, "frozen", False):
    APP_DIR = pathlib.Path(sys.executable).parent
else:
    APP_DIR = pathlib.Path(__file__).parent

CONFIG_FILE = APP_DIR / "config.json"

# ─── Constants ────────────────────────────────────────────────────────────────
APP_TITLE = "SML Platform Order Importer v1.0"
REQUEST_TIMEOUT = 15  # seconds

EXCLUDE_STATUSES = {"ที่ต้องจัดส่ง", "ยกเลิกแล้ว"}

# Shopee Excel column mapping (Thai header names)
SHOPEE_COL_MAP = {
    "order_id":      ["หมายเลขคำสั่งซื้อ", "Order ID"],
    "status":        ["สถานะการสั่งซื้อ", "Order Status"],
    "order_date":    ["วันที่ทำการสั่งซื้อ", "Order Creation Date"],
    "product_name":  ["ชื่อสินค้า", "Product Name"],
    "sku":           ["เลขอ้างอิง SKU (SKU Reference No.)", "SKU Reference No.", "SKU"],
    "price":         ["ราคาขาย", "Deal Price"],
    "qty":           ["จำนวน", "Quantity Purchased"],
}

DEFAULT_CONFIG = {
    "server_url":       "http://192.168.2.224:8080",
    "guid":             "SMLX",
    "provider":         "SML1",
    "config_file_name": "SMLConfigSML1.xml",
    "database_name":    "SMLPLOY",
    "doc_format_code":  "IV",
    "sale_code":        "",
    "cust_code":        "",
    "vat_type":         "0",
    "vat_rate":         "7",
    "wh_code":          "",
    "shelf_code":       "",
    "unit_code":        "",
    "doc_time":         "09:00",
}

VAT_TYPE_OPTIONS = ["0 — แยกนอก", "1 — รวมใน", "2 — ศูนย์%"]
VAT_DISPLAY_TO_INT = {"0 — แยกนอก": "0", "1 — รวมใน": "1", "2 — ศูนย์%": "2"}
VAT_INT_TO_DISPLAY = {v: k for k, v in VAT_DISPLAY_TO_INT.items()}


# ─── SML API Client ───────────────────────────────────────────────────────────
class SMLClient:
    def __init__(self, config: dict):
        self.base_url = config["server_url"].rstrip("/")
        self.headers = {
            "guid":           config["guid"],
            "provider":       config["provider"],
            "configFileName": config["config_file_name"],
            "databaseName":   config["database_name"],
            "Content-Type":   "application/json",
        }

    def get_customers(self, size=500):
        url = f"{self.base_url}/SMLJavaRESTService/v3/api/customer"
        r = requests.get(url, headers=self.headers,
                         params={"page": 1, "size": size},
                         timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        return r.json().get("data", [])

    def get_product(self, item_code: str):
        url = f"{self.base_url}/SMLJavaRESTService/v3/api/product/{item_code}"
        r = requests.get(url, headers=self.headers, timeout=REQUEST_TIMEOUT)
        if r.status_code == 200:
            return r.json().get("data", {})
        return None

    def create_invoice(self, payload: dict):
        url = f"{self.base_url}/SMLJavaRESTService/saleinvoice/"
        r = requests.post(url, headers=self.headers,
                          json=payload, timeout=REQUEST_TIMEOUT)
        try:
            resp_json = r.json()
        except Exception:
            resp_json = {"message": r.text}
        return r.status_code, resp_json


# ─── VAT Calculator ───────────────────────────────────────────────────────────
def calc_item_vat(price: float, qty: float, vat_type: int, vat_rate: float):
    """Returns (price_exc, sum_amount, vat_amount, sum_amount_exclude_vat)"""
    rate = vat_rate / 100
    sum_amount = round(price * qty, 2)

    if vat_type == 0:       # แยกนอก
        price_exc = price
        vat_amount = round(sum_amount * rate, 2)
        sum_exc = sum_amount
    elif vat_type == 1:     # รวมใน
        price_exc = round(price / (1 + rate), 6)
        sum_exc = round(price_exc * qty, 2)
        vat_amount = round(sum_amount - sum_exc, 2)
    else:                   # ศูนย์%
        price_exc = price
        vat_amount = 0.0
        sum_exc = sum_amount

    return price_exc, sum_amount, vat_amount, sum_exc


# ─── Shopee Excel Reader ──────────────────────────────────────────────────────
def _find_header_row(df_raw) -> int:
    """Find the row index where Shopee headers appear."""
    candidates = set(SHOPEE_COL_MAP["order_id"])
    for i, row in df_raw.iterrows():
        if any(str(v).strip() in candidates for v in row.values):
            return i
    return 0


def read_shopee_excel(filepath: str):
    """
    Returns:
        orders  : list of order dicts
        warnings: list of human-readable warning strings
    """
    df_raw = pd.read_excel(filepath, header=None, sheet_name=0)
    header_row = _find_header_row(df_raw)
    df = pd.read_excel(filepath, header=header_row, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]

    # Build column mapping: field → actual column name in file
    col = {}
    for field, candidates in SHOPEE_COL_MAP.items():
        for c in candidates:
            if c in df.columns:
                col[field] = c
                break

    missing = [f for f in ["order_id", "status", "order_date", "sku", "price", "qty"]
               if f not in col]
    if missing:
        raise ValueError(
            f"ไม่พบ column ที่จำเป็น: {missing}\n"
            f"Columns ในไฟล์: {list(df.columns[:15])}..."
        )

    warnings = []

    # Filter excluded statuses
    before = len(df)
    df = df[~df[col["status"]].isin(EXCLUDE_STATUSES)].copy()
    dropped = before - len(df)
    if dropped:
        warnings.append(f"กรอง {dropped} แถว (สถานะ: {', '.join(EXCLUDE_STATUSES)})")

    # Group rows by Order ID
    order_dict: dict = {}
    sku_missing_orders: set = set()

    for _, row in df.iterrows():
        order_id = str(row[col["order_id"]]).strip()
        if not order_id or order_id.lower() == "nan":
            continue

        # Parse date
        date_raw = row[col["order_date"]]
        try:
            if isinstance(date_raw, str):
                doc_date = date_raw[:10]
            else:
                doc_date = pd.Timestamp(date_raw).strftime("%Y-%m-%d")
        except Exception:
            doc_date = datetime.date.today().strftime("%Y-%m-%d")

        status = str(row[col["status"]]).strip()

        if order_id not in order_dict:
            order_dict[order_id] = {
                "order_id": order_id,
                "doc_date": doc_date,
                "status":   status,
                "items":    [],
            }

        # SKU
        sku = str(row.get(col["sku"], "")).strip()
        pname = str(row.get(col.get("product_name", ""), "")).strip()
        if not sku or sku.lower() == "nan":
            sku_missing_orders.add(order_id)
            warnings.append(
                f"Order {order_id}: ไม่มีรหัส SKU สำหรับ \"{pname[:40]}\""
                f" — กรุณากำหนด SKU ใน Shopee Seller Center"
            )
            continue

        try:
            price = float(row[col["price"]])
        except (ValueError, TypeError):
            price = 0.0
        try:
            qty = float(row[col["qty"]])
        except (ValueError, TypeError):
            qty = 1.0

        order_dict[order_id]["items"].append({
            "sku":          sku,
            "product_name": pname,
            "price":        price,
            "qty":          qty,
        })

    # Keep only orders with at least 1 valid item
    valid_orders = []
    for o in order_dict.values():
        if o["items"]:
            valid_orders.append(o)
        else:
            if o["order_id"] not in sku_missing_orders:
                warnings.append(f"Order {o['order_id']}: ไม่มีสินค้า — ข้ามไป")

    return valid_orders, warnings


# ─── Payload Builder ──────────────────────────────────────────────────────────
def build_invoice_payload(order: dict, config: dict, product_cache: dict) -> dict:
    vat_type = int(config.get("vat_type", 0))
    vat_rate = float(config.get("vat_rate", 7))

    details = []
    total_value = 0.0
    total_vat   = 0.0
    total_exc   = 0.0

    for i, item in enumerate(order["items"]):
        sku = item["sku"]
        prod = product_cache.get(sku, {}) or {}

        unit_code  = prod.get("start_sale_unit")  or config.get("unit_code")  or ""
        wh_code    = prod.get("start_sale_wh")    or config.get("wh_code")    or ""
        shelf_code = prod.get("start_sale_shelf") or config.get("shelf_code") or ""

        price_exc, sum_amt, vat_amt, sum_exc = calc_item_vat(
            item["price"], item["qty"], vat_type, vat_rate
        )
        total_value += sum_amt
        total_vat   += vat_amt
        total_exc   += sum_exc

        details.append({
            "item_code":             sku,
            "line_number":           i,
            "is_permium":            0,
            "unit_code":             unit_code,
            "wh_code":               wh_code,
            "shelf_code":            shelf_code,
            "qty":                   item["qty"],
            "price":                 round(item["price"], 4),
            "price_exclude_vat":     round(price_exc, 4),
            "discount_amount":       0,
            "sum_amount":            round(sum_amt, 2),
            "vat_amount":            round(vat_amt, 2),
            "tax_type":              0,
            "vat_type":              vat_type,
            "sum_amount_exclude_vat": round(sum_exc, 2),
        })

    total_value = round(total_value, 2)
    total_vat   = round(total_vat,   2)
    total_exc   = round(total_exc,   2)

    if vat_type == 0:       # แยกนอก
        total_before_vat = total_value
        total_after_vat  = round(total_value + total_vat, 2)
        total_amount     = total_after_vat
    elif vat_type == 1:     # รวมใน
        total_before_vat = total_exc
        total_after_vat  = total_value
        total_amount     = total_value
    else:                   # ศูนย์%
        total_before_vat = total_value
        total_after_vat  = total_value
        total_amount     = total_value

    return {
        "doc_no":          order["order_id"],
        "doc_date":        order["doc_date"],
        "doc_time":        config.get("doc_time", "09:00"),
        "doc_format_code": config.get("doc_format_code", ""),
        "cust_code":       config.get("cust_code", ""),
        "sale_code":       config.get("sale_code", ""),
        "sale_type":       0,
        "vat_type":        vat_type,
        "vat_rate":        vat_rate,
        "total_value":     total_value,
        "total_discount":  0,
        "total_before_vat": total_before_vat,
        "total_vat_value":  total_vat,
        "total_except_vat": 0,
        "total_after_vat":  total_after_vat,
        "total_amount":     total_amount,
        "cash_amount":      0,
        "chq_amount":       0,
        "credit_amount":    0,
        "tranfer_amount":   0,
        "details":          details,
        "paydetails":       [],
    }


# ─── Main Application ─────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1000x820")
        self.minsize(800, 600)

        self.config_data   = self._load_config()
        self.orders        = []          # all loaded orders from Excel
        self.product_cache = {}          # sku -> product info dict
        self.stop_flag     = False
        self.msg_queue     = queue.Queue()
        self.excel_path    = None
        self.retry_ids     = None        # set[str] or None

        self._build_ui()
        self._poll_queue()

    # ── Config persistence ────────────────────────────────────────────────────
    def _load_config(self) -> dict:
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                cfg = DEFAULT_CONFIG.copy()
                cfg.update(saved)
                return cfg
            except Exception:
                pass
        return DEFAULT_CONFIG.copy()

    def _save_config(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"บันทึก config ไม่ได้:\n{e}")

    def _collect_config(self):
        """Pull all UI entry values into self.config_data."""
        self.config_data["server_url"]       = self.var_server.get().strip()
        self.config_data["guid"]             = self.var_guid.get().strip()
        self.config_data["provider"]         = self.var_provider.get().strip()
        self.config_data["config_file_name"] = self.var_config_file.get().strip()
        self.config_data["database_name"]    = self.var_db_name.get().strip()
        self.config_data["doc_format_code"]  = self.var_doc_format.get().strip()
        self.config_data["sale_code"]        = self.var_sale_code.get().strip()
        self.config_data["cust_code"]        = self.var_cust_code.get().strip()
        self.config_data["vat_rate"]         = self.var_vat_rate.get().strip()
        self.config_data["wh_code"]          = self.var_wh.get().strip()
        self.config_data["shelf_code"]       = self.var_shelf.get().strip()
        self.config_data["unit_code"]        = self.var_unit.get().strip()
        self.config_data["doc_time"]         = self.var_doc_time.get().strip()
        # vat_type from combobox display
        disp = self.var_vat_display.get()
        self.config_data["vat_type"] = VAT_DISPLAY_TO_INT.get(disp, "0")

    # ── UI Build ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Platform selector bar ──
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10, pady=(8, 4))

        ttk.Label(bar, text="Platform:",
                  font=("TkDefaultFont", 11, "bold")).pack(side="left", padx=(0, 10))

        ttk.Button(bar, text="🟠  Shopee",
                   command=lambda: None).pack(side="left", padx=(0, 4))

        ttk.Button(bar, text="🟣  Lazada",
                   command=self._coming_soon).pack(side="left", padx=(0, 4))

        ttk.Button(bar, text="⬛  TikTok Shop",
                   command=self._coming_soon).pack(side="left")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10, pady=(4, 0))

            # ── Tab buttons ──
        tab_bar = tk.Frame(self, bg="#f0f0f0", pady=4)
        tab_bar.pack(fill="x", padx=10)

        self._tab_btns = {}
        for key, label in [("config", "  ⚙️  ตั้งค่า  "), ("import", "  📥  Import  ")]:
            b = tk.Button(
                tab_bar, text=label,
                font=("TkDefaultFont", 11),
                relief="flat", padx=12, pady=5,
                cursor="hand2",
                command=lambda k=key: self._show_tab(k),
            )
            b.pack(side="left", padx=(0, 2))
            self._tab_btns[key] = b

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10)

        # ── Tab content frames ──
        self._tab_frames = {}
        for key in ("config", "import"):
            f = ttk.Frame(self)
            self._tab_frames[key] = f

        self._build_config_tab(self._tab_frames["config"])
        self._build_import_tab(self._tab_frames["import"])

        self._show_tab("config")

    def _show_tab(self, key: str):
        for k, f in self._tab_frames.items():
            f.pack_forget()
        self._tab_frames[key].pack(fill="both", expand=True, padx=10, pady=8)
        # Highlight active tab button
        for k, b in self._tab_btns.items():
            b.config(relief="sunken" if k == key else "flat")

    def _coming_soon(self):
        messagebox.showinfo(
            "อยู่ระหว่างพัฒนา",
            "ฟีเจอร์นี้อยู่ระหว่างพัฒนา\nปัจจุบันรองรับเฉพาะ Shopee เท่านั้น"
        )

    # ── Config Tab ────────────────────────────────────────────────────────────
    def _build_config_tab(self, parent):
        # Use a Canvas + Scrollbar so content is always scrollable
        canvas = tk.Canvas(parent, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_resize(e):
            canvas.itemconfig(win_id, width=e.width)
        inner.bind("<Configure>", _on_configure)
        canvas.bind("<Configure>", _on_canvas_resize)

        pad = {"padx": 16, "pady": 4}

        # ── Section: การเชื่อมต่อ SML ──
        tk.Label(inner, text="การเชื่อมต่อ SML",
                 font=("TkDefaultFont", 11, "bold")).pack(anchor="w", padx=16, pady=(12, 2))
        tk.Frame(inner, height=1, bg="#cccccc").pack(fill="x", padx=16, pady=(0, 6))

        grid1 = tk.Frame(inner)
        grid1.pack(fill="x", padx=24, pady=(0, 8))
        grid1.columnconfigure(1, weight=1)

        self.var_server      = tk.StringVar(value=self.config_data["server_url"])
        self.var_guid        = tk.StringVar(value=self.config_data["guid"])
        self.var_provider    = tk.StringVar(value=self.config_data["provider"])
        self.var_config_file = tk.StringVar(value=self.config_data["config_file_name"])
        self.var_db_name     = tk.StringVar(value=self.config_data["database_name"])

        def grow(frame, label, var, r, options=None):
            tk.Label(frame, text=label, anchor="w").grid(
                row=r, column=0, sticky="w", padx=(0, 10), pady=4)
            if options:
                w = ttk.Combobox(frame, textvariable=var, values=options,
                                 state="readonly", width=36)
            else:
                w = ttk.Entry(frame, textvariable=var, width=38)
            w.grid(row=r, column=1, sticky="ew", pady=4)

        grow(grid1, "Server URL",     self.var_server,      0)
        grow(grid1, "GUID",           self.var_guid,         1)
        grow(grid1, "Provider",       self.var_provider,     2)
        grow(grid1, "configFileName", self.var_config_file,  3)
        grow(grid1, "databaseName",   self.var_db_name,      4)

        # ── Section: ค่าเริ่มต้นเอกสาร ──
        tk.Label(inner, text="ค่าเริ่มต้นเอกสาร",
                 font=("TkDefaultFont", 11, "bold")).pack(anchor="w", padx=16, pady=(8, 2))
        tk.Frame(inner, height=1, bg="#cccccc").pack(fill="x", padx=16, pady=(0, 6))

        grid2 = tk.Frame(inner)
        grid2.pack(fill="x", padx=24, pady=(0, 8))
        grid2.columnconfigure(1, weight=1)

        self.var_doc_format  = tk.StringVar(value=self.config_data["doc_format_code"])
        self.var_sale_code   = tk.StringVar(value=self.config_data["sale_code"])
        self.var_doc_time    = tk.StringVar(value=self.config_data["doc_time"])
        self.var_vat_display = tk.StringVar(
            value=VAT_INT_TO_DISPLAY.get(str(self.config_data["vat_type"]), "0 — แยกนอก"))
        self.var_vat_rate    = tk.StringVar(value=str(self.config_data["vat_rate"]))
        self.var_wh          = tk.StringVar(value=self.config_data["wh_code"])
        self.var_shelf       = tk.StringVar(value=self.config_data["shelf_code"])
        self.var_unit        = tk.StringVar(value=self.config_data["unit_code"])
        self.var_cust_code   = tk.StringVar(value=self.config_data["cust_code"])

        grow(grid2, "doc_format_code",         self.var_doc_format,  0)
        grow(grid2, "sale_code (รหัสพนักงาน)", self.var_sale_code,   1)
        grow(grid2, "doc_time",                self.var_doc_time,    2)
        grow(grid2, "vat_type", self.var_vat_display, 3, options=VAT_TYPE_OPTIONS)
        grow(grid2, "vat_rate (%)",            self.var_vat_rate,    4)
        grow(grid2, "wh_code (fallback คลัง)", self.var_wh,          5)
        grow(grid2, "shelf_code (fallback)",   self.var_shelf,       6)
        grow(grid2, "unit_code (fallback)",    self.var_unit,        7)

        # cust_code row with load button
        tk.Label(grid2, text="cust_code (รหัสลูกค้า)", anchor="w").grid(
            row=8, column=0, sticky="w", padx=(0, 10), pady=4)
        cust_row = tk.Frame(grid2)
        cust_row.grid(row=8, column=1, sticky="ew", pady=4)
        cust_row.columnconfigure(0, weight=1)
        self.cb_cust = ttk.Combobox(cust_row, textvariable=self.var_cust_code)
        self.cb_cust.grid(row=0, column=0, sticky="ew")
        ttk.Button(cust_row, text="โหลดจาก SML",
                   command=self._load_customers).grid(row=0, column=1, padx=(8, 0))

        # Save button
        btn_row = tk.Frame(inner)
        btn_row.pack(fill="x", padx=16, pady=(4, 16))
        ttk.Button(btn_row, text="บันทึกการตั้งค่า",
                   command=self._on_save_config).pack(side="right")

    # ── Import Tab ────────────────────────────────────────────────────────────
    def _build_import_tab(self, parent):
        parent.columnconfigure(0, weight=1)

        # ── File row ──
        ff = tk.Frame(parent)
        ff.pack(fill="x", padx=10, pady=(10, 4))

        tk.Label(ff, text="ไฟล์ Excel:", font=("TkDefaultFont", 10, "bold")).pack(
            side="left", padx=(0, 6))
        self.var_filepath = tk.StringVar()
        ttk.Entry(ff, textvariable=self.var_filepath, state="readonly",
                  width=45).pack(side="left", fill="x", expand=True)
        ttk.Button(ff, text="เลือกไฟล์",
                   command=self._pick_file).pack(side="left", padx=(6, 0))
        ttk.Button(ff, text="โหลด Log (Retry)",
                   command=self._load_old_log).pack(side="left", padx=(6, 0))

        # ── Summary label ──
        self.lbl_summary = tk.Label(parent, text="ยังไม่ได้เลือกไฟล์", anchor="w")
        self.lbl_summary.pack(fill="x", padx=10, pady=(0, 2))

        # ── Preview table ──
        tree_frame = tk.Frame(parent)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(0, 4))

        cols = ("order_id", "doc_date", "items", "total", "status")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=8)
        for col, text, w in [
            ("order_id", "Order ID",     180),
            ("doc_date", "วันที่",        100),
            ("items",    "# สินค้า",      70),
            ("total",    "ยอดรวม (฿)",  110),
            ("status",   "สถานะ",         250),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=w, minwidth=40)

        vsb2 = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb2.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb2.pack(side="right", fill="y")

        # ── Progress ──
        pg = tk.Frame(parent)
        pg.pack(fill="x", padx=10, pady=4)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(pg, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x")
        self.lbl_progress = tk.Label(pg, text="", anchor="w")
        self.lbl_progress.pack(fill="x")

        # ── Action buttons ──
        bf = tk.Frame(parent)
        bf.pack(fill="x", padx=10, pady=(0, 4))
        self.btn_import = ttk.Button(bf, text="Import",
                                     command=self._start_import)
        self.btn_import.pack(side="left", padx=(0, 6))
        self.btn_stop = ttk.Button(bf, text="หยุด",
                                   command=self._stop_import, state="disabled")
        self.btn_stop.pack(side="left", padx=(0, 6))
        self.btn_retry = ttk.Button(bf, text="Retry Failed",
                                    command=self._start_import, state="disabled")
        self.btn_retry.pack(side="left")

        # ── Log panel ──
        tk.Label(parent, text="Log", font=("TkDefaultFont", 10, "bold"), anchor="w").pack(
            fill="x", padx=10, pady=(4, 0))

        log_btn = tk.Frame(parent)
        log_btn.pack(fill="x", padx=10, pady=(2, 2))
        ttk.Button(log_btn, text="Copy Log", command=self._copy_log).pack(
            side="left", padx=(0, 6))
        ttk.Button(log_btn, text="Clear",    command=self._clear_log).pack(side="left")

        self.log_text = scrolledtext.ScrolledText(
            parent, height=8, state="disabled",
            font=("Courier", 10), wrap="word"
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_text.tag_configure("success", foreground="#1a7a1a")
        self.log_text.tag_configure("error",   foreground="#cc0000")
        self.log_text.tag_configure("warn",    foreground="#cc7700")
        self.log_text.tag_configure("info",    foreground="#333333")

    # ── Config actions ────────────────────────────────────────────────────────
    def _on_save_config(self):
        self._collect_config()
        self._save_config()
        messagebox.showinfo("บันทึกแล้ว", "บันทึกการตั้งค่าเรียบร้อยแล้ว")

    def _load_customers(self):
        self._collect_config()
        try:
            client = SMLClient(self.config_data)
            customers = client.get_customers()
            codes = [c["code"] for c in customers if c.get("code")]
            self.cb_cust["values"] = codes
            messagebox.showinfo("โหลดสำเร็จ",
                                f"พบลูกค้าทั้งหมด {len(codes)} รายการ")
        except Exception as e:
            messagebox.showerror("Error", f"โหลดลูกค้าไม่ได้:\n{e}")

    # ── File actions ──────────────────────────────────────────────────────────
    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="เลือกไฟล์ Excel Shopee",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")]
        )
        if not path:
            return
        self.excel_path = path
        self.var_filepath.set(path)
        self.retry_ids = None
        self.btn_retry.config(state="disabled")
        self._load_excel(path)

    def _load_excel(self, path: str):
        self._log("info", f"อ่านไฟล์: {pathlib.Path(path).name}")
        try:
            orders, warnings = read_shopee_excel(path)
            for w in warnings:
                self._log("warn", f"⚠  {w}")
            self.orders = orders
            self._populate_preview(orders)
            self._log("info", f"พบ {len(orders)} orders พร้อม import")
        except Exception as e:
            messagebox.showerror("อ่านไฟล์ไม่ได้", str(e))
            self._log("error", f"✗ {e}")

    def _populate_preview(self, orders: list):
        self.tree.delete(*self.tree.get_children())
        for o in orders:
            total = sum(it["price"] * it["qty"] for it in o["items"])
            self.tree.insert("", "end", values=(
                o["order_id"],
                o["doc_date"],
                len(o["items"]),
                f"{total:,.2f}",
                o["status"][:60],
            ))
        self.lbl_summary.config(text=f"จำนวน {len(orders)} orders พร้อม import")

    def _load_old_log(self):
        path = filedialog.askopenfilename(
            title="เลือกไฟล์ Log CSV",
            filetypes=[("CSV", "*.csv"), ("All", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                error_ids = {row["order_id"] for row in reader
                             if row.get("status") == "error"}
            if not error_ids:
                messagebox.showinfo("ไม่มี Error",
                                    "ไม่พบ orders ที่ status=error ในไฟล์ log นี้")
                return
            self.retry_ids = error_ids
            self._log("info",
                      f"โหลด Log: พบ {len(error_ids)} orders ที่ต้อง Retry")
            if self.orders:
                retry_orders = [o for o in self.orders
                                if o["order_id"] in error_ids]
                self._populate_preview(retry_orders)
                self.lbl_summary.config(
                    text=f"Retry mode: {len(retry_orders)} orders ที่ล้มเหลว")
            else:
                self._log("warn", "⚠  กรุณาเลือกไฟล์ Excel ก่อน แล้วค่อย Retry")
            self.btn_retry.config(state="normal")
        except Exception as e:
            messagebox.showerror("Error", f"อ่าน log ไม่ได้:\n{e}")

    # ── Import ────────────────────────────────────────────────────────────────
    def _start_import(self):
        if not self.excel_path:
            messagebox.showwarning("ยังไม่มีข้อมูล",
                                   "กรุณาเลือกไฟล์ Excel ก่อน")
            return

        self._collect_config()

        # Determine which orders to import
        target = self.orders
        if self.retry_ids is not None:
            target = [o for o in self.orders if o["order_id"] in self.retry_ids]

        if not target:
            messagebox.showwarning("ไม่มี orders", "ไม่มี orders ที่จะ import")
            return

        if not self.config_data.get("cust_code", "").strip():
            if not messagebox.askyesno(
                "cust_code ว่าง",
                "ยังไม่ได้ระบุ cust_code (รหัสลูกค้า)\nต้องการ import ต่อหรือไม่?"
            ):
                return

        # Setup log file next to Excel
        excel_dir = pathlib.Path(self.excel_path).parent
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = excel_dir / f"import_log_{ts}.csv"

        self.stop_flag = False
        self.btn_import.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_retry.config(state="disabled")
        self.progress_var.set(0)
        self.lbl_progress.config(text="")

        self._log("info", f"\n{'─'*55}")
        self._log("info", f"เริ่ม import {len(target)} orders — {ts}")
        self._log("info", f"Log file: {log_path.name}")
        self._log("info", f"{'─'*55}")

        t = threading.Thread(
            target=self._import_worker,
            args=(target, self.config_data.copy(), log_path),
            daemon=True
        )
        t.start()

    def _stop_import(self):
        self.stop_flag = True
        self._log("warn", "⏹  กำลังหยุด... รอ order ปัจจุบันเสร็จก่อน")
        self.btn_stop.config(state="disabled")

    # ── Import worker (runs in background thread) ─────────────────────────────
    def _import_worker(self, orders: list, config: dict, log_path: pathlib.Path):
        client = SMLClient(config)
        total  = len(orders)

        # Load product cache for unique SKUs
        all_skus = list({item["sku"] for o in orders for item in o["items"]})
        if all_skus:
            self.msg_queue.put(("info",
                                f"โหลดข้อมูลสินค้า {len(all_skus)} รายการ..."))
            for sku in all_skus:
                if sku not in self.product_cache:
                    try:
                        prod = client.get_product(sku)
                        self.product_cache[sku] = prod or {}
                    except Exception:
                        self.product_cache[sku] = {}

        success_count = 0
        error_count   = 0

        with open(log_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow([
                "order_id", "doc_date", "item_count",
                "total_amount", "status", "message", "timestamp"
            ])

            for idx, order in enumerate(orders, 1):
                if self.stop_flag:
                    self.msg_queue.put((
                        "warn",
                        f"⏹  หยุดแล้ว — import ไป {idx-1}/{total} orders"
                    ))
                    break

                ts_now   = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                total_amt = sum(it["price"] * it["qty"] for it in order["items"])
                item_cnt  = len(order["items"])

                try:
                    payload = build_invoice_payload(order, config, self.product_cache)
                    status_code, resp = client.create_invoice(payload)

                    if status_code == 201 or resp.get("status") == "success":
                        msg = resp.get("message", "create success")
                        writer.writerow([order["order_id"], order["doc_date"],
                                         item_cnt, total_amt, "success", msg, ts_now])
                        f.flush()
                        success_count += 1
                        self.msg_queue.put((
                            "success",
                            f"[{ts_now}] ✓  {order['order_id']} — {msg}"
                        ))
                    else:
                        msg = resp.get("message", str(resp))
                        writer.writerow([order["order_id"], order["doc_date"],
                                         item_cnt, total_amt, "error", msg, ts_now])
                        f.flush()
                        error_count += 1
                        self.msg_queue.put((
                            "error",
                            f"[{ts_now}] ✗  {order['order_id']} — ERROR: {msg}"
                        ))

                except requests.exceptions.ConnectionError:
                    msg = "Connection Error — ไม่สามารถเชื่อมต่อ server ได้"
                    writer.writerow([order["order_id"], order["doc_date"],
                                     item_cnt, total_amt, "error", msg, ts_now])
                    f.flush()
                    error_count += 1
                    self.msg_queue.put((
                        "error",
                        f"[{ts_now}] ✗  {order['order_id']} — {msg}"
                    ))

                except requests.exceptions.Timeout:
                    msg = f"Timeout — server ไม่ตอบภายใน {REQUEST_TIMEOUT} วินาที"
                    writer.writerow([order["order_id"], order["doc_date"],
                                     item_cnt, total_amt, "error", msg, ts_now])
                    f.flush()
                    error_count += 1
                    self.msg_queue.put((
                        "error",
                        f"[{ts_now}] ✗  {order['order_id']} — {msg}"
                    ))

                except Exception as e:
                    msg = str(e)
                    writer.writerow([order["order_id"], order["doc_date"],
                                     item_cnt, total_amt, "error", msg, ts_now])
                    f.flush()
                    error_count += 1
                    self.msg_queue.put((
                        "error",
                        f"[{ts_now}] ✗  {order['order_id']} — ERROR: {msg}"
                    ))

                pct = (idx / total) * 100
                self.msg_queue.put((
                    "progress",
                    (pct, f"{idx}/{total}  ✓ {success_count}  ✗ {error_count}")
                ))

        summary = (
            f"เสร็จสิ้น — ✓ {success_count} สำเร็จ  /  ✗ {error_count} ล้มเหลว\n"
            f"Log: {log_path}"
        )
        self.msg_queue.put(("done", summary))

    # ── Queue polling (runs on main thread via after()) ───────────────────────
    def _poll_queue(self):
        try:
            while True:
                kind, data = self.msg_queue.get_nowait()
                if kind == "progress":
                    pct, label = data
                    self.progress_var.set(pct)
                    self.lbl_progress.config(text=label)
                elif kind == "done":
                    self._log("info", f"\n{'='*55}\n{data}\n{'='*55}\n")
                    self.btn_import.config(state="normal")
                    self.btn_stop.config(state="disabled")
                    self.btn_retry.config(state="normal")
                    messagebox.showinfo("เสร็จสิ้น", data)
                else:
                    self._log(kind, data)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    # ── Log helpers ───────────────────────────────────────────────────────────
    def _log(self, tag: str, text: str):
        self.log_text.config(state="normal")
        self.log_text.insert("end", text + "\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _copy_log(self):
        content = self.log_text.get("1.0", "end")
        self.clipboard_clear()
        self.clipboard_append(content)
        messagebox.showinfo("Copy", "คัดลอก Log ไปยัง Clipboard แล้ว")

    def _clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")


# ─── Entry point ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
