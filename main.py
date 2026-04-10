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

try:
    import sv_ttk
    _HAS_SV_TTK = True
except ImportError:
    _HAS_SV_TTK = False

# ─── App directory (works both dev and PyInstaller) ───────────────────────────
if getattr(sys, "frozen", False):
    APP_DIR = pathlib.Path(sys.executable).parent
else:
    APP_DIR = pathlib.Path(__file__).parent

CONFIG_FILE = APP_DIR / "config.json"

# ─── Constants ────────────────────────────────────────────────────────────────
APP_TITLE   = "SML Platform Importer"
APP_VERSION = "v1.0"
REQUEST_TIMEOUT = 15  # seconds

EXCLUDE_STATUSES = {"ที่ต้องจัดส่ง", "ยกเลิกแล้ว"}

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

VAT_TYPE_OPTIONS   = ["0 — แยกนอก", "1 — รวมใน", "2 — ศูนย์%"]
VAT_DISPLAY_TO_INT = {"0 — แยกนอก": "0", "1 — รวมใน": "1", "2 — ศูนย์%": "2"}
VAT_INT_TO_DISPLAY = {v: k for k, v in VAT_DISPLAY_TO_INT.items()}

# ─── UI Design Tokens ─────────────────────────────────────────────────────────
C = {
    "bg":          "#EEF2F7",
    "header":      "#1A253A",
    "card":        "#FFFFFF",
    "border":      "#D1D9E6",
    "primary":     "#2563EB",
    "primary_hv":  "#1D4ED8",
    "shopee":      "#EE4D2D",
    "shopee_hv":   "#CC3B1F",
    "lazada":      "#0F1DC5",
    "tiktok":      "#111111",
    "success":     "#15803D",
    "success_lt":  "#DCFCE7",
    "error":       "#B91C1C",
    "error_lt":    "#FEE2E2",
    "warn":        "#92400E",
    "warn_lt":     "#FEF3C7",
    "text":        "#0F172A",
    "muted":       "#64748B",
    "row_a":       "#F8FAFC",
    "row_b":       "#FFFFFF",
    "log_bg":      "#0D1117",
    "log_fg":      "#C9D1D9",
    "log_ok":      "#3FB950",
    "log_err":     "#F85149",
    "log_warn":    "#D29922",
    "log_info":    "#8B949E",
    "stop":        "#DC2626",
    "stop_hv":     "#B91C1C",
    "retry":       "#D97706",
    "retry_hv":    "#B45309",
    "slate":       "#475569",
    "slate_hv":    "#334155",
}

FN  = ("Segoe UI", 10)          # normal
FB  = ("Segoe UI", 10, "bold")  # bold
FH  = ("Segoe UI", 11, "bold")  # heading
FS  = ("Segoe UI", 9)           # small
FM  = ("Consolas", 9)           # mono


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

    if vat_type == 0:
        price_exc = price
        vat_amount = round(sum_amount * rate, 2)
        sum_exc = sum_amount
    elif vat_type == 1:
        price_exc = round(price / (1 + rate), 6)
        sum_exc = round(price_exc * qty, 2)
        vat_amount = round(sum_amount - sum_exc, 2)
    else:
        price_exc = price
        vat_amount = 0.0
        sum_exc = sum_amount

    return price_exc, sum_amount, vat_amount, sum_exc


# ─── Shopee Excel Reader ──────────────────────────────────────────────────────
def _find_header_row(df_raw) -> int:
    candidates = set(SHOPEE_COL_MAP["order_id"])
    for i, row in df_raw.iterrows():
        if any(str(v).strip() in candidates for v in row.values):
            return i
    return 0


def read_shopee_excel(filepath: str):
    df_raw = pd.read_excel(filepath, header=None, sheet_name=0)
    header_row = _find_header_row(df_raw)
    df = pd.read_excel(filepath, header=header_row, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]

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

    before = len(df)
    df = df[~df[col["status"]].isin(EXCLUDE_STATUSES)].copy()
    dropped = before - len(df)
    if dropped:
        warnings.append(f"กรอง {dropped} แถว (สถานะ: {', '.join(EXCLUDE_STATUSES)})")

    order_dict: dict = {}
    sku_missing_orders: set = set()

    for _, row in df.iterrows():
        order_id = str(row[col["order_id"]]).strip()
        if not order_id or order_id.lower() == "nan":
            continue

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

        sku   = str(row.get(col["sku"], "")).strip()
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

    details     = []
    total_value = 0.0
    total_vat   = 0.0
    total_exc   = 0.0

    for i, item in enumerate(order["items"]):
        sku  = item["sku"]
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
            "item_code":              sku,
            "line_number":            i,
            "is_permium":             0,
            "unit_code":              unit_code,
            "wh_code":                wh_code,
            "shelf_code":             shelf_code,
            "qty":                    item["qty"],
            "price":                  round(item["price"], 4),
            "price_exclude_vat":      round(price_exc, 4),
            "discount_amount":        0,
            "sum_amount":             round(sum_amt, 2),
            "vat_amount":             round(vat_amt, 2),
            "tax_type":               0,
            "vat_type":               vat_type,
            "sum_amount_exclude_vat": round(sum_exc, 2),
        })

    total_value = round(total_value, 2)
    total_vat   = round(total_vat,   2)
    total_exc   = round(total_exc,   2)

    if vat_type == 0:
        total_before_vat = total_value
        total_after_vat  = round(total_value + total_vat, 2)
        total_amount     = total_after_vat
    elif vat_type == 1:
        total_before_vat = total_exc
        total_after_vat  = total_value
        total_amount     = total_value
    else:
        total_before_vat = total_value
        total_after_vat  = total_value
        total_amount     = total_value

    return {
        "doc_no":           order["order_id"],
        "doc_date":         order["doc_date"],
        "doc_time":         config.get("doc_time", "09:00"),
        "doc_format_code":  config.get("doc_format_code", ""),
        "cust_code":        config.get("cust_code", ""),
        "sale_code":        config.get("sale_code", ""),
        "sale_type":        0,
        "vat_type":         vat_type,
        "vat_rate":         vat_rate,
        "total_value":      total_value,
        "total_discount":   0,
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
        self.geometry("1050x860")
        self.minsize(820, 640)
        self.configure(bg=C["bg"])

        if _HAS_SV_TTK:
            sv_ttk.set_theme("light")

        self.config_data   = self._load_config()
        self.orders        = []
        self.product_cache = {}
        self.stop_flag     = False
        self.msg_queue     = queue.Queue()
        self.excel_path    = None
        self.retry_ids     = None

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
        disp = self.var_vat_display.get()
        self.config_data["vat_type"] = VAT_DISPLAY_TO_INT.get(disp, "0")

    # ── UI helpers ────────────────────────────────────────────────────────────
    def _btn(self, parent, text, bg, hv, cmd, state="normal", **kw):
        b = tk.Button(
            parent, text=text,
            bg=bg, fg="#FFFFFF",
            disabledforeground="#A0A0A0",
            font=FB, relief="flat", bd=0,
            padx=16, pady=8,
            cursor="hand2",
            activebackground=hv,
            activeforeground="#FFFFFF",
            command=cmd, state=state, **kw
        )
        return b

    def _small_btn(self, parent, text, bg, hv, cmd, **kw):
        return tk.Button(
            parent, text=text,
            bg=bg, fg="#FFFFFF",
            font=FS, relief="flat", bd=0,
            padx=10, pady=4,
            cursor="hand2",
            activebackground=hv,
            activeforeground="#FFFFFF",
            command=cmd, **kw
        )

    # ── Top-level UI ──────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header bar ──
        hdr = tk.Frame(self, bg=C["header"], height=58)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="  \u25fc  SML Platform Importer",
                 bg=C["header"], fg="#F8FAFC",
                 font=("Segoe UI", 15, "bold")).pack(side="left", padx=(12, 0), pady=14)
        tk.Label(hdr, text=f"  {APP_VERSION}",
                 bg=C["header"], fg="#64748B",
                 font=FS).pack(side="left", pady=14)

        # ── Platform selector row ──
        prow = tk.Frame(self, bg=C["bg"])
        prow.pack(fill="x", padx=20, pady=(14, 6))

        tk.Label(prow, text="Platform:", bg=C["bg"], fg=C["muted"],
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 10))

        for label, bg, hv, cmd, active in [
            ("  Shopee  ",  C["shopee"], C["shopee_hv"], lambda: None,       True),
            ("  Lazada  ",  C["lazada"], "#0B16A0",      self._coming_soon,   False),
            ("  TikTok  ",  C["tiktok"], "#333333",      self._coming_soon,   False),
        ]:
            b = tk.Button(
                prow, text=label,
                bg=bg if active else "#CBD5E1",
                fg="#FFFFFF" if active else "#475569",
                font=("Segoe UI", 9, "bold"),
                relief="flat", bd=0, padx=2, pady=6,
                cursor="hand2",
                activebackground=hv if active else "#94A3B8",
                activeforeground="#FFFFFF",
                command=cmd,
            )
            b.pack(side="left", padx=(0, 6))

        # ── Tab bar ──
        tabrow = tk.Frame(self, bg=C["bg"])
        tabrow.pack(fill="x", padx=20, pady=(2, 0))

        self._tab_btns = {}
        for key, icon, label in [
            ("config", "\u2699", "  ตั้งค่า  "),
            ("import", "\u2193", "  Import  "),
        ]:
            b = tk.Button(
                tabrow, text=f" {icon}  {label}",
                font=FN, relief="flat", bd=0,
                padx=10, pady=9,
                cursor="hand2",
                command=lambda k=key: self._show_tab(k),
            )
            b.pack(side="left")
            self._tab_btns[key] = b

        # Active-tab underline
        tk.Frame(self, bg=C["border"], height=2).pack(fill="x")

        # ── Tab content ──
        self._tab_frames = {}
        for key in ("config", "import"):
            f = tk.Frame(self, bg=C["bg"])
            self._tab_frames[key] = f

        self._build_config_tab(self._tab_frames["config"])
        self._build_import_tab(self._tab_frames["import"])

        self._show_tab("config")

    def _show_tab(self, key: str):
        for k, f in self._tab_frames.items():
            f.pack_forget()
        self._tab_frames[key].pack(fill="both", expand=True)
        for k, b in self._tab_btns.items():
            if k == key:
                b.config(bg=C["card"], fg=C["primary"],
                         font=("Segoe UI", 10, "bold"))
            else:
                b.config(bg=C["bg"], fg=C["muted"], font=FN)

    def _coming_soon(self):
        messagebox.showinfo(
            "อยู่ระหว่างพัฒนา",
            "ฟีเจอร์นี้อยู่ระหว่างพัฒนา\nปัจจุบันรองรับเฉพาะ Shopee เท่านั้น"
        )

    # ── Config Tab ────────────────────────────────────────────────────────────
    def _build_config_tab(self, parent):
        canvas = tk.Canvas(parent, bg=C["bg"], borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=C["bg"])
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _resize(_): canvas.configure(scrollregion=canvas.bbox("all"))
        def _width(e):  canvas.itemconfig(win_id, width=e.width)
        inner.bind("<Configure>", _resize)
        canvas.bind("<Configure>", _width)

        def card_section(title):
            """Return the body Frame of a card section."""
            outer = tk.Frame(inner, bg=C["card"],
                             highlightbackground=C["border"],
                             highlightthickness=1)
            outer.pack(fill="x", padx=20, pady=(14, 0))

            hdr_bar = tk.Frame(outer, bg=C["border"], height=36)
            hdr_bar.pack(fill="x")
            hdr_bar.pack_propagate(False)
            tk.Label(hdr_bar, text=f"  {title}",
                     bg=C["border"], fg=C["text"],
                     font=FB).pack(side="left", padx=8, pady=6)

            body = tk.Frame(outer, bg=C["card"])
            body.pack(fill="x", padx=18, pady=(10, 14))
            body.columnconfigure(1, weight=1)
            return body

        def field(grid, label, var, row, options=None):
            tk.Label(grid, text=label, bg=C["card"], fg=C["text"],
                     font=FN, anchor="w", width=30).grid(
                row=row, column=0, sticky="w", padx=(0, 14), pady=6)
            if options:
                w = ttk.Combobox(grid, textvariable=var, values=options,
                                 state="readonly", font=FN)
            else:
                w = ttk.Entry(grid, textvariable=var, font=FN)
            w.grid(row=row, column=1, sticky="ew", pady=6)

        # ── Section 1 ──
        g1 = card_section("การเชื่อมต่อ SML Server")

        self.var_server      = tk.StringVar(value=self.config_data["server_url"])
        self.var_guid        = tk.StringVar(value=self.config_data["guid"])
        self.var_provider    = tk.StringVar(value=self.config_data["provider"])
        self.var_config_file = tk.StringVar(value=self.config_data["config_file_name"])
        self.var_db_name     = tk.StringVar(value=self.config_data["database_name"])

        field(g1, "Server URL",        self.var_server,      0)
        field(g1, "GUID",              self.var_guid,         1)
        field(g1, "Provider",          self.var_provider,     2)
        field(g1, "Config File Name",  self.var_config_file,  3)
        field(g1, "Database Name",     self.var_db_name,      4)

        # ── Section 2 ──
        g2 = card_section("ค่าเริ่มต้นเอกสาร")

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

        field(g2, "รูปแบบเอกสาร (doc_format_code)",  self.var_doc_format,  0)
        field(g2, "รหัสพนักงานขาย (sale_code)",       self.var_sale_code,   1)
        field(g2, "เวลาเอกสาร (doc_time)",            self.var_doc_time,    2)
        field(g2, "ประเภทภาษี (vat_type)",            self.var_vat_display, 3, options=VAT_TYPE_OPTIONS)
        field(g2, "อัตราภาษี % (vat_rate)",           self.var_vat_rate,    4)
        field(g2, "รหัสคลัง fallback (wh_code)",      self.var_wh,          5)
        field(g2, "รหัส Shelf fallback (shelf_code)",  self.var_shelf,       6)
        field(g2, "รหัสหน่วย fallback (unit_code)",   self.var_unit,        7)

        # cust_code row with load button
        tk.Label(g2, text="รหัสลูกค้า (cust_code)", bg=C["card"], fg=C["text"],
                 font=FN, anchor="w", width=30).grid(
            row=8, column=0, sticky="w", padx=(0, 14), pady=6)

        cust_row = tk.Frame(g2, bg=C["card"])
        cust_row.grid(row=8, column=1, sticky="ew", pady=6)
        cust_row.columnconfigure(0, weight=1)

        self.cb_cust = ttk.Combobox(cust_row, textvariable=self.var_cust_code, font=FN)
        self.cb_cust.grid(row=0, column=0, sticky="ew")

        self._small_btn(cust_row, "โหลดจาก SML",
                        C["primary"], C["primary_hv"],
                        self._load_customers).grid(row=0, column=1, padx=(8, 0))

        # ── Save button ──
        foot = tk.Frame(inner, bg=C["bg"])
        foot.pack(fill="x", padx=20, pady=(16, 20))

        self._btn(foot, "  บันทึกการตั้งค่า  ",
                  C["primary"], C["primary_hv"],
                  self._on_save_config).pack(side="right")

    # ── Import Tab ────────────────────────────────────────────────────────────
    def _build_import_tab(self, parent):
        parent.configure(bg=C["bg"])

        # ── File picker card ──
        file_card = tk.Frame(parent, bg=C["card"],
                             highlightbackground=C["border"],
                             highlightthickness=1)
        file_card.pack(fill="x", padx=20, pady=(16, 0))

        ff = tk.Frame(file_card, bg=C["card"])
        ff.pack(fill="x", padx=16, pady=12)

        tk.Label(ff, text="ไฟล์ Excel:", bg=C["card"], fg=C["text"],
                 font=FB).pack(side="left", padx=(0, 8))

        self.var_filepath = tk.StringVar()
        ttk.Entry(ff, textvariable=self.var_filepath, state="readonly",
                  font=FN).pack(side="left", fill="x", expand=True, padx=(0, 8))

        self._small_btn(ff, "เลือกไฟล์",
                        C["primary"], C["primary_hv"],
                        self._pick_file).pack(side="left", padx=(0, 6))

        self._small_btn(ff, "Retry จาก Log",
                        C["slate"], C["slate_hv"],
                        self._load_old_log).pack(side="left")

        # ── Summary ──
        sumrow = tk.Frame(parent, bg=C["bg"])
        sumrow.pack(fill="x", padx=20, pady=(8, 4))

        self.lbl_summary = tk.Label(
            sumrow, text="ยังไม่ได้เลือกไฟล์",
            bg=C["bg"], fg=C["muted"], font=FS, anchor="w"
        )
        self.lbl_summary.pack(side="left")

        # ── Orders preview table ──
        tree_card = tk.Frame(parent, bg=C["card"],
                             highlightbackground=C["border"],
                             highlightthickness=1)
        tree_card.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        style = ttk.Style()
        style.configure("T.Treeview",
                        background=C["card"],
                        fieldbackground=C["card"],
                        rowheight=28,
                        font=FN)
        style.configure("T.Treeview.Heading",
                        background=C["border"],
                        foreground=C["text"],
                        font=FB, relief="flat")
        style.map("T.Treeview",
                  background=[("selected", C["primary"])],
                  foreground=[("selected", "#FFFFFF")])

        cols = ("order_id", "doc_date", "items", "total", "status")
        self.tree = ttk.Treeview(tree_card, columns=cols, show="headings",
                                 height=8, style="T.Treeview")
        for col, text, w, anc in [
            ("order_id", "Order ID",    180, "w"),
            ("doc_date", "วันที่",       100, "center"),
            ("items",    "# สินค้า",     70,  "center"),
            ("total",    "ยอดรวม (฿)",  120, "e"),
            ("status",   "สถานะ",        240, "w"),
        ]:
            self.tree.heading(col, text=text, anchor=anc)
            self.tree.column(col, width=w, minwidth=40, anchor=anc)

        self.tree.tag_configure("odd",      background=C["row_a"])
        self.tree.tag_configure("even",     background=C["row_b"])
        self.tree.tag_configure("done_ok",  background=C["success_lt"],
                                foreground=C["success"])
        self.tree.tag_configure("done_err", background=C["error_lt"],
                                foreground=C["error"])

        vsb2 = ttk.Scrollbar(tree_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb2.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb2.pack(side="right", fill="y")

        # ── Progress card ──
        prog_card = tk.Frame(parent, bg=C["card"],
                             highlightbackground=C["border"],
                             highlightthickness=1)
        prog_card.pack(fill="x", padx=20, pady=(0, 8))

        pg = tk.Frame(prog_card, bg=C["card"])
        pg.pack(fill="x", padx=16, pady=10)

        self.progress_var = tk.DoubleVar()
        style.configure("Blue.Horizontal.TProgressbar",
                        troughcolor=C["border"],
                        background=C["primary"],
                        thickness=10)
        self.progress_bar = ttk.Progressbar(
            pg, variable=self.progress_var, maximum=100,
            style="Blue.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill="x", pady=(0, 4))

        self.lbl_progress = tk.Label(pg, text="", bg=C["card"], fg=C["muted"],
                                     font=FS, anchor="w")
        self.lbl_progress.pack(fill="x")

        # ── Action buttons ──
        btn_area = tk.Frame(parent, bg=C["bg"])
        btn_area.pack(fill="x", padx=20, pady=(0, 8))

        self.btn_import = self._btn(btn_area, "  \u25b6  Import",
                                    C["primary"], C["primary_hv"],
                                    self._start_import)
        self.btn_import.pack(side="left", padx=(0, 8))

        self.btn_stop = self._btn(btn_area, "  \u25a0  หยุด",
                                  C["stop"], C["stop_hv"],
                                  self._stop_import, state="disabled")
        self.btn_stop.pack(side="left", padx=(0, 8))

        self.btn_retry = self._btn(btn_area, "  \u21ba  Retry Failed",
                                   C["retry"], C["retry_hv"],
                                   self._start_import, state="disabled")
        self.btn_retry.pack(side="left")

        # ── Log panel ──
        log_hdr = tk.Frame(parent, bg=C["bg"])
        log_hdr.pack(fill="x", padx=20, pady=(4, 4))

        tk.Label(log_hdr, text="Log", bg=C["bg"], fg=C["text"],
                 font=FB).pack(side="left")

        self._small_btn(log_hdr, "Clear", C["slate"], C["slate_hv"],
                        self._clear_log).pack(side="right", padx=(6, 0))
        self._small_btn(log_hdr, "Copy",  C["slate"], C["slate_hv"],
                        self._copy_log).pack(side="right")

        self.log_text = scrolledtext.ScrolledText(
            parent, height=10, state="disabled",
            bg=C["log_bg"], fg=C["log_fg"],
            font=FM,
            insertbackground=C["log_fg"],
            selectbackground="#264F78",
            wrap="word",
            bd=0, relief="flat",
            padx=14, pady=12,
        )
        self.log_text.pack(fill="both", expand=True, padx=20, pady=(0, 16))
        self.log_text.tag_configure("success", foreground=C["log_ok"])
        self.log_text.tag_configure("error",   foreground=C["log_err"])
        self.log_text.tag_configure("warn",    foreground=C["log_warn"])
        self.log_text.tag_configure("info",    foreground=C["log_info"])

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
            messagebox.showinfo("โหลดสำเร็จ", f"พบลูกค้าทั้งหมด {len(codes)} รายการ")
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
                self._log("warn", f"  {w}")
            self.orders = orders
            self._populate_preview(orders)
            self._log("info", f"พบ {len(orders)} orders พร้อม import")
        except Exception as e:
            messagebox.showerror("อ่านไฟล์ไม่ได้", str(e))
            self._log("error", f"  {e}")

    def _populate_preview(self, orders: list):
        self.tree.delete(*self.tree.get_children())
        for i, o in enumerate(orders):
            total = sum(it["price"] * it["qty"] for it in o["items"])
            tag   = "odd" if i % 2 == 0 else "even"
            self.tree.insert("", "end", iid=o["order_id"], values=(
                o["order_id"],
                o["doc_date"],
                len(o["items"]),
                f"{total:,.2f}",
                o["status"][:60],
            ), tags=(tag,))
        self.lbl_summary.config(
            text=f"โหลดแล้ว  {len(orders)} orders  พร้อม import",
            fg=C["text"]
        )

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
            self._log("info", f"โหลด Log: พบ {len(error_ids)} orders ที่ต้อง Retry")
            if self.orders:
                retry_orders = [o for o in self.orders if o["order_id"] in error_ids]
                self._populate_preview(retry_orders)
                self.lbl_summary.config(
                    text=f"Retry mode: {len(retry_orders)} orders ที่ล้มเหลว",
                    fg=C["retry"]
                )
            else:
                self._log("warn", "  กรุณาเลือกไฟล์ Excel ก่อน แล้วค่อย Retry")
            self.btn_retry.config(state="normal")
        except Exception as e:
            messagebox.showerror("Error", f"อ่าน log ไม่ได้:\n{e}")

    # ── Import ────────────────────────────────────────────────────────────────
    def _start_import(self):
        if not self.excel_path:
            messagebox.showwarning("ยังไม่มีข้อมูล", "กรุณาเลือกไฟล์ Excel ก่อน")
            return

        self._collect_config()
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
        self._log("warn", "  กำลังหยุด... รอ order ปัจจุบันเสร็จก่อน")
        self.btn_stop.config(state="disabled")

    # ── Import worker (background thread) ────────────────────────────────────
    def _import_worker(self, orders: list, config: dict, log_path: pathlib.Path):
        client = SMLClient(config)
        total  = len(orders)

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
                        f"  หยุดแล้ว — import ไป {idx-1}/{total} orders"
                    ))
                    break

                ts_now    = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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
                            f"[{ts_now}]  {order['order_id']} — {msg}"
                        ))
                        self.msg_queue.put(("tree_ok", order["order_id"]))
                    else:
                        msg = resp.get("message", str(resp))
                        writer.writerow([order["order_id"], order["doc_date"],
                                         item_cnt, total_amt, "error", msg, ts_now])
                        f.flush()
                        error_count += 1
                        self.msg_queue.put((
                            "error",
                            f"[{ts_now}]  {order['order_id']} — ERROR: {msg}"
                        ))
                        self.msg_queue.put(("tree_err", order["order_id"]))

                except requests.exceptions.ConnectionError:
                    msg = "Connection Error — ไม่สามารถเชื่อมต่อ server ได้"
                    writer.writerow([order["order_id"], order["doc_date"],
                                     item_cnt, total_amt, "error", msg, ts_now])
                    f.flush()
                    error_count += 1
                    self.msg_queue.put((
                        "error", f"[{ts_now}]  {order['order_id']} — {msg}"
                    ))
                    self.msg_queue.put(("tree_err", order["order_id"]))

                except requests.exceptions.Timeout:
                    msg = f"Timeout — server ไม่ตอบภายใน {REQUEST_TIMEOUT} วินาที"
                    writer.writerow([order["order_id"], order["doc_date"],
                                     item_cnt, total_amt, "error", msg, ts_now])
                    f.flush()
                    error_count += 1
                    self.msg_queue.put((
                        "error", f"[{ts_now}]  {order['order_id']} — {msg}"
                    ))
                    self.msg_queue.put(("tree_err", order["order_id"]))

                except Exception as e:
                    msg = str(e)
                    writer.writerow([order["order_id"], order["doc_date"],
                                     item_cnt, total_amt, "error", msg, ts_now])
                    f.flush()
                    error_count += 1
                    self.msg_queue.put((
                        "error", f"[{ts_now}]  {order['order_id']} — ERROR: {msg}"
                    ))
                    self.msg_queue.put(("tree_err", order["order_id"]))

                pct = (idx / total) * 100
                self.msg_queue.put((
                    "progress",
                    (pct, f"{idx} / {total}    \u2713 {success_count}    \u2717 {error_count}")
                ))

        summary = (
            f"เสร็จสิ้น — \u2713 {success_count} สำเร็จ  /  \u2717 {error_count} ล้มเหลว\n"
            f"Log: {log_path}"
        )
        self.msg_queue.put(("done", summary))

    # ── Queue polling ─────────────────────────────────────────────────────────
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
                elif kind == "tree_ok":
                    try:
                        self.tree.item(data, tags=("done_ok",))
                    except Exception:
                        pass
                elif kind == "tree_err":
                    try:
                        self.tree.item(data, tags=("done_err",))
                    except Exception:
                        pass
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
