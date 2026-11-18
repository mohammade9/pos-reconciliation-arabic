from __future__ import annotations

import os
import sys

# Prevent Intel MKL / numexpr / bottleneck conflicts on macOS & Linux
os.environ["MKL_SERVICE_FORCE_INTEL"] = "1"
sys.modules["numexpr"] = None
sys.modules["bottleneck"] = None

import base64
import copy
import io
import json
import re
import zipfile

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
EXAMPLE_CONFIG_PATH = os.path.join(APP_DIR, "config.example.json")

# ---------------------------------------------------------------------------
# Payment Networks (Primary: Mastercard -> Mada -> Visa; Rare: 0 Variants)
# ---------------------------------------------------------------------------
STANDARD_NETWORKS: list[tuple[str, str]] = [
    ("master", "Mastercard"),
    ("mada", "Mada"),
    ("visa", "Visa"),
]

RARE_NETWORKS: list[tuple[str, str]] = [
    ("gccnet", "GCCNET"),
    ("amex", "Amex"),
    ("discover", "Discover"),
    ("unionpay", "UnionPay"),
    ("maestro", "Maestro"),
]

ALL_NETWORKS: list[tuple[str, str]] = STANDARD_NETWORKS + RARE_NETWORKS
ALL_NETWORK_KEYS: list[str] = [k for k, _ in ALL_NETWORKS]
NETWORK_LABELS: dict[str, str] = dict(ALL_NETWORKS)

# ---------------------------------------------------------------------------
# Default Template Configuration (Clean Open-Source Placeholders)
# ---------------------------------------------------------------------------
DEFAULT_CONFIG: dict = {
    "company_name_ar": "الشركة النموذجية للتجارة",
    "company_name_en": "Sample Retail Co.",
    "accent_color": "#1D54B5",
    "accent2_color": "#5BC2E7",
    "logo_data_uri": "",
    "currency": "SAR",
    "shortage_account": "SHORTAGE-ACC",
    "description_template": "إيداع بنكي: {deposit_acc}-{network}, {currency}. 000001",
    "clearing_accounts": {
        "master": "CLEARING-MASTERCARD",
        "mada": "CLEARING-MADA",
        "visa": "CLEARING-VISA",
        "gccnet": "CLEARING-GCCNET",
        "amex": "CLEARING-AMEX",
        "discover": "CLEARING-DISCOVER",
        "unionpay": "CLEARING-UNIONPAY",
        "maestro": "CLEARING-MAESTRO",
    },
    "branch_groups": [
        {
            "id": "GROUP-A",
            "name": "Branch Group A",
            "company_code": "GROUP-A",
            "multi_file": False,
            "append_branch_code_to_gl": False,
            "accent_color": "#1D54B5",
            "accent2_color": "#5BC2E7",
            "shortage_account": "SHORTAGE-ACC",
            "branches": [
                {
                    "name": "Branch 01",
                    "code": "BR-01",
                    "bank_account": "BANK-001",
                    "deposit_account": "BANK-001",
                    "network_deposit_accounts": {},
                },
                {
                    "name": "Branch 02",
                    "code": "BR-02",
                    "bank_account": "BANK-002",
                    "deposit_account": "BANK-002",
                    "network_deposit_accounts": {},
                },
            ],
        },
        {
            "id": "GROUP-B",
            "name": "Branch Group B",
            "company_code": "GROUP-B",
            "multi_file": True,
            "append_branch_code_to_gl": True,
            "accent_color": "#0E7490",
            "accent2_color": "#38BDF8",
            "shortage_account": "SHORTAGE-ACC",
            "branches": [
                {
                    "name": "Branch 03",
                    "code": "BR-03",
                    "bank_account": "BANK-003",
                    "deposit_account": "DEP-003",
                    "network_deposit_accounts": {
                        "master": "DEP-003-MC",
                        "mada": "DEP-003-MADA",
                        "visa": "DEP-003-VISA",
                    },
                },
                {
                    "name": "Branch 04",
                    "code": "BR-04",
                    "bank_account": "BANK-004",
                    "deposit_account": "DEP-004",
                    "network_deposit_accounts": {},
                },
            ],
        },
    ],
}

# ---------------------------------------------------------------------------
# Excel Cell Fills & Numeric Constants
# ---------------------------------------------------------------------------
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
GRAY_FILL = PatternFill(start_color="A6A6A6", end_color="A6A6A6", fill_type="solid")
RED_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

TOLERANCE = 0.01
ARABIC_DIGITS = str.maketrans(
    "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹٫٬",
    "01234567890123456789.,",
)


# ---------------------------------------------------------------------------
# Configuration Load / Save Helpers
# ---------------------------------------------------------------------------
def load_app_config() -> tuple[dict, bool]:
    """Return (config_dict, is_configured). Reads config.json or falls back to config.example.json."""
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            return _merge_config_defaults(data), True
        except Exception:
            pass

    if os.path.exists(EXAMPLE_CONFIG_PATH):
        try:
            with open(EXAMPLE_CONFIG_PATH, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            return _merge_config_defaults(data), False
        except Exception:
            pass

    return copy.deepcopy(DEFAULT_CONFIG), False


def _merge_config_defaults(data: dict) -> dict:
    merged = copy.deepcopy(DEFAULT_CONFIG)
    if not isinstance(data, dict):
        return merged
    for k, v in data.items():
        if k == "clearing_accounts" and isinstance(v, dict):
            merged["clearing_accounts"].update(v)
        elif k == "branch_groups" and isinstance(v, list) and v:
            merged["branch_groups"] = v
        else:
            merged[k] = v
    return merged


def save_app_config(cfg: dict) -> None:
    """Persist configuration to config.json."""
    with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)


def hex_to_rgba(hex_color: str, alpha: float = 0.16) -> str:
    h = (hex_color or "#1D54B5").lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return f"rgba({r}, {g}, {b}, {alpha})"
    except Exception:
        return f"rgba(29, 84, 181, {alpha})"


# ---------------------------------------------------------------------------
# Parsing & Normalization Utilities
# ---------------------------------------------------------------------------
def normalize_amount(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    text = str(value).translate(ARABIC_DIGITS)
    text = re.sub(r"(?i)sar|ر\.?س|ريال", "", text)
    text = text.replace(",", "").replace(" ", "").replace("\xa0", "").strip()
    if not text:
        return None
    try:
        return round(float(text), 2)
    except ValueError:
        return None


def parse_drive_amounts(raw_text: str) -> list[float]:
    if not (raw_text and raw_text.strip()):
        return []
    amounts: list[float] = []
    for token in raw_text.translate(ARABIC_DIGITS).replace("+", " ").split():
        val = normalize_amount(token)
        if val is None:
            raise ValueError(f"قيمة غير صحيحة: {token}")
        amounts.append(val)
    return amounts


def shift_suffix(shift_value) -> str:
    """Extract unique digits after the colon ':' (e.g. '920' from '03047309:920')."""
    text = str(shift_value or "").translate(ARABIC_DIGITS).strip()
    if ":" in text:
        text = text.split(":")[-1]
    text = re.sub(r"\D", "", text)
    return text.lstrip("0") or ("0" if text else "")


def fmt_signed(value: float) -> str:
    if abs(value) < TOLERANCE:
        return "0.00"
    label = "زيادة" if value > 0 else "عجز"
    return f"{value:+,.2f} ({label})"


def is_cash_payment(payment_text: str) -> bool:
    return "cash" in payment_text or "نقد" in payment_text


def detect_payment_network_key(payment_type_lower: str) -> str | None:
    if "mada" in payment_type_lower or "مدى" in payment_type_lower:
        return "mada"
    if "master" in payment_type_lower:
        return "master"
    if "visa" in payment_type_lower:
        return "visa"
    if "gcc" in payment_type_lower:
        return "gccnet"
    if "amex" in payment_type_lower or "american" in payment_type_lower:
        return "amex"
    if "discover" in payment_type_lower:
        return "discover"
    if "union" in payment_type_lower:
        return "unionpay"
    if "maestro" in payment_type_lower:
        return "maestro"
    return None


FILENAME_STOPWORDS = {
    "عجز",
    "زيادة",
    "زياده",
    "مبلغ",
    "بنكي",
    "الفرق",
    "فرق",
    "نقد",
    "نقدا",
    "كاش",
    "اسم",
    "الكاشير",
    "كاشير",
    "ريال",
    "sar",
    "cash",
    "copy",
    "scan",
    "نسخة",
    "وردية",
    "الوردية",
}


def _format_num_str(val: float) -> str:
    if abs(val - round(val)) < 1e-9:
        return str(int(round(val)))
    return f"{val:.2f}".rstrip("0").rstrip(".")


def parse_receipt_filename(
    filename: str,
    known_cash_values: list[float] | None = None,
    known_shift_suffixes: set[str] | None = None,
) -> dict:
    """
    Auto-extract Cashier Name (اسم الكاشير), Cash or -Shift (مبلغ الكاش / الوردية -),
    and Signed Variance (مبلغ الفرق: عجز / زيادة) from the uploaded receipt's filename.
    """
    stem = filename.rsplit(".", 1)[0] if "." in filename else filename
    text = stem.translate(ARABIC_DIGITS)
    text = re.sub(r"[_،;|]+", " ", text)

    raw_nums = re.findall(r"\d+(?:[.,]\d{1,2})?", text)
    if len(raw_nums) >= 3:
        text_no_idx = re.sub(
            r"^\s*\d{1,2}\s*[-.)]+\s*(?=[A-Za-z\u0600-\u06FF])",
            "",
            text,
        )
        if text_no_idx != text:
            text = text_no_idx

    numbers: list[tuple[float, str, int, int, str]] = []
    for m in re.finditer(r"([-+−]?)(\d+(?:[.,]\d{1,2})?)", text):
        sign_char = m.group(1)
        start = m.start()
        if sign_char and start > 0 and not text[start - 1].isspace():
            sign_char = ""
        raw_digits = m.group(2)
        value = float(raw_digits.replace(",", "."))
        numbers.append((value, sign_char, m.start(), m.end(), raw_digits))

    result = {
        "name": "",
        "cash": 0.0,
        "cash_input": "",
        "variance": 0.0,
        "sign": "-",
        "found": False,
    }

    if numbers:
        cash_idx = 0
        is_shift_token = False

        # Check if any token is a -Shift lookup token (e.g., -920)
        for idx_n, (val_n, sign_n, st_n, en_n, raw_d) in enumerate(numbers):
            if sign_n in ("-", "−") and "." not in raw_d and "," not in raw_d:
                cand_sfx = shift_suffix(raw_d)
                before_tok = text[max(0, st_n - 12) : st_n]
                after_tok = text[en_n : min(len(text), en_n + 12)]
                has_var_word_near = (
                    "عجز" in before_tok
                    or "عجز" in after_tok
                    or "زياد" in before_tok
                    or "زياد" in after_tok
                )
                if known_shift_suffixes and cand_sfx in known_shift_suffixes and not has_var_word_near:
                    cash_idx = idx_n
                    is_shift_token = True
                    break
                if idx_n == 0 and (len(numbers) > 1 or not has_var_word_near):
                    cash_idx = idx_n
                    is_shift_token = True
                    break

        if not is_shift_token and known_cash_values and len(numbers) > 1:
            for idx_n, (val_n, _s, _st, _en, _rd) in enumerate(numbers):
                if any(abs(val_n - kc) < TOLERANCE for kc in known_cash_values):
                    cash_idx = idx_n
                    break

        if is_shift_token:
            raw_d = numbers[cash_idx][4]
            sfx_digits = shift_suffix(raw_d) or raw_d
            result["cash"] = 0.0
            result["cash_input"] = f"-{sfx_digits}"
            result["found"] = True
        else:
            cash_val = round(numbers[cash_idx][0], 2)
            result["cash"] = cash_val
            result["cash_input"] = _format_num_str(cash_val) if cash_val > 0 else ""
            result["found"] = True

        remaining = [n for i, n in enumerate(numbers) if i != cash_idx]
        if remaining:
            var_value, var_sign, var_start, var_end, _ = remaining[0]
            before = text[max(0, var_start - 14) : var_start]
            after = text[var_end : min(len(text), var_end + 14)]
            if "زياد" in before or "زياد" in after or var_sign == "+":
                sign = "+"
            elif "عجز" in before or "عجز" in after or var_sign in ("-", "−"):
                sign = "-"
            elif "زياد" in text:
                sign = "+"
            else:
                sign = "-"
            result["sign"] = sign
            result["variance"] = round(var_value if sign == "+" else -var_value, 2)
    elif "زياد" in text:
        result["sign"] = "+"

    name_text = re.sub(r"[-+−]?\d+(?:[.,]\d+)?", " ", text)
    name_text = re.sub(r"[\-\(\)\[\]{}#=:]+", " ", name_text)
    words = [w for w in name_text.split() if w.lower() not in FILENAME_STOPWORDS]
    result["name"] = " ".join(words).strip()
    if result["name"]:
        result["found"] = True
    return result


# ---------------------------------------------------------------------------
# System Workbook (ملف الحركة) Parsing
# ---------------------------------------------------------------------------
def detect_worksheet_columns(ws) -> tuple[dict, int]:
    found: dict[str, int | None] = {
        "bank_amt": None,
        "movement_amt": None,
        "payment_name": None,
        "payment_method": None,
        "treasury_amt": None,
        "calculated_col": None,
        "diff": None,
        "shift_id": None,
    }
    header_row = 1
    for row_idx in range(1, 6):
        for cell in ws[row_idx]:
            if cell.value is None:
                continue
            header = str(cell.value).strip()
            if ("مبلغ بنكي" in header or "مبلغ البنك" in header) and found["bank_amt"] is None:
                found["bank_amt"] = cell.column
            elif "مبلغ الحركة" in header and found["movement_amt"] is None:
                found["movement_amt"] = cell.column
            elif "اسم طريقة الدفع" in header and found["payment_name"] is None:
                found["payment_name"] = cell.column
            elif header == "طريقة الدفع" and found["payment_method"] is None:
                found["payment_method"] = cell.column
            elif "مبلغ الخزينة" in header and found["treasury_amt"] is None:
                found["treasury_amt"] = cell.column
            elif "المحسوب" in header and found["calculated_col"] is None:
                found["calculated_col"] = cell.column
            elif "مبلغ الفرق" in header and found["diff"] is None:
                found["diff"] = cell.column
            elif "الوردية" in header and found["shift_id"] is None:
                found["shift_id"] = cell.column
        if found["bank_amt"]:
            header_row = row_idx
            break

    fallbacks = {
        "bank_amt": 7,        # Column G: مبلغ بنكي
        "movement_amt": 6,    # Column F: مبلغ الحركة
        "payment_name": 3,    # Column C: اسم طريقة الدفع
        "payment_method": 2,  # Column B: طريقة الدفع
        "treasury_amt": 8,    # Column H: مبلغ الخزينة
        "calculated_col": 9,  # Column I: المحسوب
        "diff": 10,           # Column J: مبلغ الفرق
        "shift_id": 1,        # Column A: الوردية
    }
    for key, default_col in fallbacks.items():
        if found[key] is None:
            found[key] = default_col

    return found, header_row


def parse_system_sheet(ws) -> dict:
    cols, header_row = detect_worksheet_columns(ws)
    c_bank = cols["bank_amt"]
    c_mov = cols["movement_amt"]
    c_pay = cols["payment_name"]
    c_method = cols["payment_method"]
    c_treasury = cols["treasury_amt"]
    c_shift = cols["shift_id"]
    max_row = ws.max_row

    cash_rows: list[dict] = []
    for row in range(header_row + 1, max_row + 1):
        pay_raw = str(ws.cell(row=row, column=c_pay).value or "").strip()
        if not is_cash_payment(pay_raw.lower()):
            continue

        bank_val = normalize_amount(ws.cell(row=row, column=c_bank).value) or 0.0
        mov_val = normalize_amount(ws.cell(row=row, column=c_mov).value) or 0.0
        treasury_val = normalize_amount(ws.cell(row=row, column=c_treasury).value) or 0.0
        method_raw = str(ws.cell(row=row, column=c_method).value or "").strip()
        shift_raw = str(ws.cell(row=row, column=c_shift).value or "").strip()

        app_amounts: dict[str, float] = {k: 0.0 for k in ALL_NETWORK_KEYS}
        network_rows: dict[str, int] = {}

        for sub in range(row + 1, min(row + 10, max_row + 1)):
            sub_pay = str(ws.cell(row=sub, column=c_pay).value or "").lower()
            if is_cash_payment(sub_pay):
                break
            sub_shift = str(ws.cell(row=sub, column=c_shift).value or "").strip()
            if shift_raw and sub_shift and sub_shift != shift_raw:
                break
            net_key = detect_payment_network_key(sub_pay)
            if net_key:
                app_amounts[net_key] = round(
                    app_amounts[net_key]
                    + (normalize_amount(ws.cell(row=sub, column=c_mov).value) or 0.0),
                    2,
                )
                network_rows[net_key] = sub

        cash_rows.append(
            {
                "row": row,
                "cash": round(bank_val, 2),
                "cash_movement": round(mov_val, 2),
                "treasury_amount": round(treasury_val, 2),
                "payment_method": method_raw,
                "payment_name": pay_raw,
                "shift_id": shift_raw,
                "shift_suffix": shift_suffix(shift_raw),
                "app_amounts": app_amounts,
                "network_rows": network_rows,
            }
        )

    headers_ok = any(
        str(c.value or "").strip() in ("مبلغ بنكي", "مبلغ البنك")
        for c in ws[header_row]
    )
    return {
        "columns": cols,
        "header_row": header_row,
        "cash_rows": cash_rows,
        "headers_ok": headers_ok,
    }


@st.cache_data(show_spinner=False)
def parse_system_bytes(file_bytes: bytes) -> dict:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    return parse_system_sheet(wb.active)


# ---------------------------------------------------------------------------
# Dynamic Accounting & Double-Entry Journal Logic
# ---------------------------------------------------------------------------
def get_group_config(config: dict, group_id: str) -> dict:
    groups = config.get("branch_groups", [])
    for g in groups:
        if g.get("id") == group_id:
            return g
    return groups[0] if groups else DEFAULT_CONFIG["branch_groups"][0]


def get_branch_config(group_cfg: dict, branch_name: str) -> dict:
    branches = group_cfg.get("branches", [])
    for b in branches:
        if b.get("name") == branch_name:
            return b
    return branches[0] if branches else {
        "name": "Branch 01",
        "code": "BR-01",
        "bank_account": "BANK-001",
        "deposit_account": "BANK-001",
        "network_deposit_accounts": {},
    }


def get_network_accounting_config(
    config: dict,
    group_id: str,
    net_key: str,
    branch_name: str = "",
) -> dict:
    group_cfg = get_group_config(config, group_id)
    branch_cfg = get_branch_config(group_cfg, branch_name)

    net_label = net_key.upper()
    code = str(branch_cfg.get("code", "")).strip()
    bank_acc = str(branch_cfg.get("bank_account", "BANK-001")).strip()

    clearing_map = {
        **config.get("clearing_accounts", {}),
        **group_cfg.get("clearing_accounts", {}),
    }
    base_middle = str(clearing_map.get(net_key, f"CLEARING-{net_label}")).strip()
    base_shortage = str(
        group_cfg.get("shortage_account")
        or config.get("shortage_account", "SHORTAGE-ACC")
    ).strip()

    if group_cfg.get("append_branch_code_to_gl", False) and code:
        middle_acc = f"{base_middle}-{code}"
        shortage_acc = f"{base_shortage}-{code}"
    else:
        middle_acc = base_middle
        shortage_acc = base_shortage

    net_dep_map = branch_cfg.get("network_deposit_accounts") or {}
    deposit_acc = (
        str(net_dep_map.get(net_key, "")).strip()
        or str(branch_cfg.get("deposit_account", "")).strip()
        or bank_acc
    )

    currency = str(config.get("currency", "SAR")).strip() or "SAR"
    desc_template = (
        group_cfg.get("description_template")
        or config.get("description_template")
        or "إيداع بنكي: {deposit_acc}-{network}, {currency}. 000001"
    )
    try:
        desc = desc_template.format(
            deposit_acc=deposit_acc,
            bank_acc=bank_acc,
            network=net_label,
            branch_code=code,
            branch_name=branch_cfg.get("name", ""),
            currency=currency,
        )
    except Exception:
        desc = f"إيداع بنكي: {deposit_acc}-{net_label}, {currency}. 000001"

    company = str(group_cfg.get("company_code") or group_cfg.get("id", "COMPANY")).strip()
    return {
        "bank_acc": bank_acc,
        "middle_acc": middle_acc,
        "shortage_acc": shortage_acc,
        "desc": desc,
        "company": company,
        "currency": currency,
    }


def journal_row(
    acc_type: str,
    account: str,
    desc: str,
    debit: float,
    credit: float,
    company: str,
    offset_account: str = "",
    currency: str = "SAR",
) -> list:
    return [
        acc_type,
        account,
        desc,
        debit,
        credit,
        company,
        "دفتر الأستاذ",
        offset_account,
        currency,
        1,
    ]


def process_network(
    config: dict,
    net_key: str,
    drive_amounts: list[float],
    app_amt: float,
    group_id: str,
    branch_name: str = "",
) -> tuple[list[list], float, float]:
    rows: list[list] = []
    clean_drives = [
        round(float(x), 2) for x in (drive_amounts or []) if round(float(x), 2) > 0
    ]
    drive_total = round(sum(clean_drives), 2)
    app_val = round(float(app_amt or 0.0), 2)

    if drive_total < TOLERANCE and abs(app_val) < TOLERANCE:
        return rows, 0.0, 0.0

    variance = round(drive_total - app_val, 2)
    cfg = get_network_accounting_config(config, group_id, net_key, branch_name)
    bank = cfg["bank_acc"]
    middle = cfg["middle_acc"]
    shortage = cfg["shortage_acc"]
    desc = cfg["desc"]
    co = cfg["company"]
    curr = cfg["currency"]

    # Missing Network Single-Line Rule:
    # If the network does not exist (or is 0) in ملف الحركة and user entered > 0,
    # generate one single double-entry line directly between البنك and Shortage (عجز).
    if abs(app_val) < TOLERANCE:
        if drive_total >= TOLERANCE:
            rows.append(
                journal_row("البنك", bank, desc, drive_total, 0, co, shortage, curr)
            )
        return rows, variance, drive_total

    if drive_total < TOLERANCE:
        return rows, variance, 0.0

    if len(clean_drives) == 1 and abs(variance) < TOLERANCE:
        rows.append(
            journal_row("البنك", bank, desc, clean_drives[0], 0, co, middle, curr)
        )
        return rows, variance, drive_total

    for amt in clean_drives:
        rows.append(journal_row("البنك", bank, desc, amt, 0, co, "", curr))

    rows.append(journal_row("دفتر الأستاذ", middle, desc, 0, app_val, co, "", curr))
    if variance >= TOLERANCE:
        rows.append(journal_row("دفتر الأستاذ", shortage, desc, 0, variance, co, "", curr))
    elif variance <= -TOLERANCE:
        rows.append(
            journal_row("دفتر الأستاذ", shortage, desc, abs(variance), 0, co, "", curr)
        )
    return rows, variance, drive_total


def compute_calculated(
    cash_val: float,
    cash_movement: float,
    user_totals: dict,
    app_amounts: dict,
    discount_val: float,
    target_variance: float | None = None,
) -> dict:
    cash_diff = round(cash_val - cash_movement, 2)
    net_diff = round(
        sum(
            user_totals.get(k, 0.0) - app_amounts.get(k, 0.0)
            for k in ALL_NETWORK_KEYS
        ),
        2,
    )
    base_diff = round(cash_diff + net_diff, 2)
    if abs(discount_val) < TOLERANCE:
        calculated = base_diff
    else:
        cand_add = round(base_diff + discount_val, 2)
        cand_sub = round(base_diff - discount_val, 2)
        if (
            target_variance is not None
            and abs(cand_sub - target_variance) <= TOLERANCE + 1e-9
        ):
            calculated = cand_sub
        elif (
            target_variance is not None
            and abs(cand_add - target_variance) <= TOLERANCE + 1e-9
        ):
            calculated = cand_add
        else:
            calculated = cand_add if base_diff <= 0 else cand_sub
    return {"cash_diff": cash_diff, "net_diff": net_diff, "calculated": calculated}


def find_cash_record(
    parsed: dict,
    cash: float,
    shift_filter: str,
    used_rows: set,
    lookup_by_shift: bool = False,
) -> dict | None:
    target_sfx = shift_suffix(shift_filter) if shift_filter else ""
    for rec in parsed["cash_rows"]:
        if rec["row"] in used_rows:
            continue
        if lookup_by_shift and target_sfx:
            if rec["shift_suffix"] != target_sfx:
                continue
            return rec
        if abs(rec["cash"] - cash) >= TOLERANCE:
            continue
        if target_sfx and rec["shift_suffix"] != target_sfx:
            continue
        return rec
    return None


def process_single_workbook(
    config: dict,
    uploaded_bytes: bytes,
    cashiers: list[dict],
    group_id: str,
    default_branch_name: str = "",
) -> tuple[list[list], list[list], bytes, list[str]]:
    """Return (journal_rows, summary_rows, formatted_workbook_bytes, warnings) for one ملف الحركة."""
    wb_data = load_workbook(io.BytesIO(uploaded_bytes), data_only=True)
    parsed = parse_system_sheet(wb_data.active)
    cols = parsed["columns"]
    c_bank, c_diff = cols["bank_amt"], cols["diff"]
    c_name = c_diff + 2

    wb_format = load_workbook(io.BytesIO(uploaded_bytes), data_only=False)
    ws_format = wb_format.active
    by_row = {r["row"]: r for r in parsed["cash_rows"]}

    group_cfg = get_group_config(config, group_id)
    include_branch_col = bool(group_cfg.get("multi_file", False))

    journal_rows: list[list] = []
    summary_rows: list[list] = []
    warnings: list[str] = []
    used_rows: set[int] = set()

    for entry in cashiers:
        name = entry["name"]
        cash = float(entry["cash"])
        variance_input = float(entry["variance"])
        discount_val = float(entry.get("discount", 0.0))
        branch_name = entry.get("sub_branch") or default_branch_name
        lookup_by_shift = bool(entry.get("lookup_by_shift", False))
        network_lists = entry.get("network_lists", {})
        network_totals_saved = entry.get("network_totals", {})
        network_raw_saved = entry.get("network_raw", {})

        rec = by_row.get(entry.get("row"))
        if (
            rec is None
            or rec["row"] in used_rows
            or (not lookup_by_shift and abs(rec["cash"] - cash) >= TOLERANCE)
        ):
            rec = find_cash_record(
                parsed,
                cash,
                entry.get("shift_id", ""),
                used_rows,
                lookup_by_shift=lookup_by_shift,
            )

        if rec is None:
            warnings.append(
                f"مبلغ بنكي {cash:,.2f} للكاشير '{name}' غير موجود في ملف الحركة — تم تجاهله."
            )
            continue

        used_rows.add(rec["row"])
        cash = round(float(rec["cash"]), 2)
        drive_totals: dict[str, float] = {}

        for net_key in ALL_NETWORK_KEYS:
            drive_list = list(network_lists.get(net_key, []))
            raw_str = str(network_raw_saved.get(net_key, "")).strip()
            if not drive_list and raw_str:
                try:
                    drive_list = parse_drive_amounts(raw_str)
                except ValueError:
                    drive_list = []
            if (
                not drive_list
                and round(float(network_totals_saved.get(net_key, 0.0)), 2) > 0
            ):
                drive_list = [round(float(network_totals_saved[net_key]), 2)]

            sys_app_amt = round(float(rec["app_amounts"].get(net_key, 0.0)), 2)
            user_drive_sum = round(sum(drive_list), 2)
            if user_drive_sum < TOLERANCE and abs(sys_app_amt) < TOLERANCE:
                drive_totals[net_key] = 0.0
                continue

            rows, _var, total = process_network(
                config,
                net_key,
                drive_list,
                sys_app_amt,
                group_id,
                branch_name,
            )
            journal_rows.extend(rows)
            drive_totals[net_key] = total

        calc = compute_calculated(
            cash,
            rec["cash_movement"],
            drive_totals,
            rec["app_amounts"],
            discount_val,
            variance_input,
        )
        is_match = abs(calc["calculated"] - variance_input) <= TOLERANCE + 1e-9

        ws_format.cell(row=rec["row"], column=c_bank).fill = YELLOW_FILL
        name_cell = ws_format.cell(row=rec["row"], column=c_name)
        name_cell.value = name
        name_cell.fill = YELLOW_FILL

        var_cell = ws_format.cell(row=rec["row"], column=c_name + 1)
        var_cell.value = variance_input
        var_cell.fill = YELLOW_FILL

        missing_net_col = c_name + 2
        for net_key in ALL_NETWORK_KEYS:
            total = drive_totals.get(net_key, 0.0)
            if total <= 0:
                continue
            if net_key in rec["network_rows"]:
                r = rec["network_rows"][net_key]
                bank_cell = ws_format.cell(row=r, column=c_bank)
                bank_cell.value = total
                bank_cell.fill = YELLOW_FILL
                ws_format.cell(row=r, column=c_diff).fill = GRAY_FILL
            else:
                extra_cell = ws_format.cell(row=rec["row"], column=missing_net_col)
                extra_cell.value = f"{NETWORK_LABELS.get(net_key, net_key.upper())}: {total:.2f}"
                extra_cell.fill = YELLOW_FILL
                missing_net_col += 1

        ws_format.cell(row=rec["row"], column=c_diff).fill = (
            GRAY_FILL if is_match else RED_FILL
        )

        shortage = abs(variance_input) if variance_input < 0 else 0.0
        extra = variance_input if variance_input > 0 else 0.0
        row_vals = [
            name,
            cash,
            shortage,
            extra,
            discount_val,
            calc["net_diff"],
            calc["calculated"],
            variance_input,
            "مطابق" if is_match else "غير مطابق",
        ]
        summary_rows.append(
            [branch_name] + row_vals if include_branch_col else row_vals
        )

    buf = io.BytesIO()
    wb_format.save(buf)
    return journal_rows, summary_rows, buf.getvalue(), warnings


SUMMARY_COLUMNS = [
    "اسم الكاشير",
    "مبلغ بنكي",
    "عجز",
    "زيادة",
    "خصم",
    "فرق الشبكات",
    "المحسوب",
    "مبلغ الفرق",
    "الحالة",
]


def rows_to_excel_bytes(rows: list, columns: list[str] | None = None) -> bytes:
    buf = io.BytesIO()
    if columns is not None:
        pd.DataFrame(rows, columns=columns).to_excel(buf, index=False)
    else:
        pd.DataFrame(rows).to_excel(buf, index=False, header=False)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Neutral Brand Header & Logo Helpers
# ---------------------------------------------------------------------------
def default_badge_svg(
    company_en: str,
    company_ar: str,
    accent: str,
    accent2: str,
) -> str:
    initial = (company_en.strip()[:1] or "P").upper()
    safe_en = (
        company_en.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    )
    safe_ar = (
        company_ar.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    )
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" width="380" height="110" viewBox="0 0 380 110">'
        "<defs>"
        '<linearGradient id="g" x1="0" y1="0" x2="1" y2="1">'
        f'<stop offset="0" stop-color="{accent}"/>'
        f'<stop offset="1" stop-color="{accent2}"/>'
        "</linearGradient>"
        "</defs>"
        '<rect x="4" y="9" width="92" height="92" rx="24" fill="url(#g)"/>'
        f'<text x="50" y="76" font-family="Helvetica, Arial, sans-serif" font-size="54" font-weight="800" fill="#fff" text-anchor="middle">{initial}</text>'
        f'<text x="114" y="54" font-family="Helvetica, Arial, sans-serif" font-size="28" font-weight="800" fill="{accent}">{safe_en[:20]}</text>'
        f'<text x="114" y="90" font-family="Geeza Pro, Tahoma, Arial, sans-serif" font-size="22" font-weight="700" fill="{accent2}">{safe_ar[:26]}</text>'
        "</svg>"
    )


def get_header_logo_uri(config: dict, accent: str, accent2: str) -> str:
    custom_uri = str(config.get("logo_data_uri") or "").strip()
    if custom_uri.startswith("data:image/"):
        return custom_uri
    svg = default_badge_svg(
        config.get("company_name_en", "POS Reconciliation"),
        config.get("company_name_ar", "الشركة النموذجية"),
        accent,
        accent2,
    ).encode("utf-8")
    return "data:image/svg+xml;base64," + base64.b64encode(svg).decode()


# ---------------------------------------------------------------------------
# RTL Styling & Multi-Page Hover Magnifier Viewer (No OCR)
# ---------------------------------------------------------------------------
BASE_CSS = """
<style>
.block-container{padding-top:2.5rem !important;padding-bottom:2rem;max-width:1500px;}
[data-testid="stAppViewContainer"],[data-testid="stSidebar"],[data-testid="stHeader"]{direction:rtl;}
[data-testid="stMarkdownContainer"],[data-testid="stWidgetLabel"],label,p,h1,h2,h3,h4,h5,h6,
[data-testid="stCaptionContainer"],[data-testid="stAlert"]{text-align:right;}
[data-testid="stTextInput"] input,[data-testid="stNumberInput"] input{direction:ltr;text-align:right;}
[data-testid="stDataFrame"]{direction:ltr;}
[data-testid="stFileUploaderDropzone"]{padding:.6rem .8rem;}
.app-header{display:flex;align-items:center;justify-content:space-between;gap:1.4rem;padding:.95rem 1.35rem;border-radius:16px;margin-bottom:1rem;}
.app-header img.brand-logo{width:185px;max-width:200px;min-width:170px;height:auto;background:transparent !important;display:block;object-fit:contain;filter:drop-shadow(0 2px 10px rgba(0,0,0,.18));}
.app-title{font-size:1.65rem;font-weight:800;line-height:1.3;}
.app-sub{font-size:.95rem;opacity:.82;margin-top:.15rem;}
.section-title{font-size:1.08rem;font-weight:700;margin:.4rem 0 .6rem 0;padding-right:.6rem;}
.chip{display:inline-block;padding:.12rem .6rem;border-radius:999px;font-size:.82rem;font-weight:600;margin-left:.3rem;}
.calc-box{display:flex;gap:.5rem;flex-wrap:wrap;margin:.3rem 0 .4rem 0;}
.calc-item{flex:1;min-width:120px;border-radius:12px;padding:.45rem .7rem;background:color-mix(in srgb,var(--accent) 9%,transparent);border:1px solid color-mix(in srgb,var(--accent) 28%,transparent);}
.calc-item .k{font-size:.78rem;opacity:.75;}
.calc-item .v{font-size:1.05rem;font-weight:700;direction:ltr;text-align:right;}
</style>
"""

THEME_CSS_TEMPLATE = """
<style>
:root{--accent:__ACCENT__;--accent2:__ACCENT2__;--soft:__SOFT__;}
.app-header{background:linear-gradient(135deg,color-mix(in srgb,var(--accent) 18%,transparent),color-mix(in srgb,var(--accent2) 6%,transparent));border:1px solid color-mix(in srgb,var(--accent) 35%,transparent);}
.app-title{color:var(--accent2);}
.section-title{border-right:4px solid var(--accent);color:var(--accent2);}
.chip{background:var(--soft);color:var(--accent2);border:1px solid color-mix(in srgb,var(--accent) 35%,transparent);}
button[kind="primary"],[data-testid="stBaseButton-primary"],[data-testid="stDownloadButton"] button{
background:linear-gradient(135deg,var(--accent),var(--accent2)) !important;border:none !important;color:#fff !important;}
button[kind="primary"]:hover,[data-testid="stBaseButton-primary"]:hover,[data-testid="stDownloadButton"] button:hover{filter:brightness(1.07);}
button[kind="primary"]:disabled,[data-testid="stBaseButton-primary"]:disabled{opacity:.45;}
button[kind="secondary"]:hover,[data-testid="stBaseButton-secondary"]:hover{border-color:var(--accent) !important;color:var(--accent2) !important;}
button:focus:not(:active){box-shadow:0 0 0 .15rem color-mix(in srgb,var(--accent) 35%,transparent) !important;}
[data-baseweb="input"]:focus-within,[data-baseweb="base-input"]:focus-within,[data-baseweb="select"] > div:focus-within{
border-color:var(--accent) !important;box-shadow:0 0 0 1px var(--accent) !important;}
[data-testid*="segmented_controlActive"],[data-testid*="pillsActive"]{
background:var(--accent) !important;border-color:var(--accent) !important;color:#fff !important;}
[data-testid*="segmented_controlActive"] p,[data-testid*="pillsActive"] p{color:#fff !important;}
[data-testid="stFileUploaderDropzone"]{border:1px dashed var(--accent2) !important;background:var(--soft) !important;}
[data-testid="stExpander"] details{border-color:color-mix(in srgb,var(--accent) 30%,transparent) !important;}
[data-testid="stExpander"] summary:hover,[data-testid="stExpander"] summary:hover p{color:var(--accent2) !important;}
[data-testid="stProgress"] [role="progressbar"] > div > div{background-color:var(--accent) !important;}
a{color:var(--accent2) !important;}
</style>
"""


@st.cache_data(show_spinner=False, max_entries=20)
def render_document_pages(file_bytes: bytes, file_name: str) -> list[str]:
    """Return a list of data-URIs (one per page) for the continuous white paper stream viewer."""
    ext = file_name.lower().rsplit(".", 1)[-1] if "." in file_name else ""
    if ext in ("png", "jpg", "jpeg"):
        mime = "image/png" if ext == "png" else "image/jpeg"
        return [f"data:{mime};base64," + base64.b64encode(file_bytes).decode()]
    if ext == "pdf" and fitz is not None:
        pages: list[str] = []
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            try:
                img_bytes, mime = pix.tobytes("jpg", jpg_quality=88), "image/jpeg"
            except Exception:
                img_bytes, mime = pix.tobytes("png"), "image/png"
            pages.append(f"data:{mime};base64," + base64.b64encode(img_bytes).decode())
        doc.close()
        return pages
    return []


VIEWER_HTML = """<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
html,body{margin:0;padding:0;background:transparent;direction:ltr;font-family:-apple-system,"Segoe UI",Tahoma,sans-serif;}
#wrap{position:relative;height:__H__px;}
#viewer{height:100%;overflow-y:auto;overflow-x:hidden;background:#fff;border:1px solid #e6e8eb;border-radius:14px;}
#stream{background:#fff;line-height:0;}
#stream img{display:block;width:100%;height:auto;margin:0;padding:0;border:0;background:#fff;user-select:none;-webkit-user-drag:none;}
#pill{position:absolute;top:10px;left:10px;z-index:5;display:flex;gap:2px;padding:3px;border-radius:999px;
background:rgba(255,255,255,.92);box-shadow:0 2px 10px rgba(0,0,0,.15);backdrop-filter:blur(6px);}
#pill button{border:0;cursor:pointer;font-size:12.5px;font-weight:600;padding:5px 11px;border-radius:999px;background:transparent;color:#333;}
#pill #btnOn{background:__ACCENT__;color:#fff;direction:rtl;}
#pill #btnOn.off{background:#e9eaee;color:#555;}
#pill #btnZ{color:__ACCENT__;min-width:44px;}
#lens{position:fixed;width:__L__px;height:__L__px;border-radius:50%;border:3px solid __ACCENT__;pointer-events:none;
display:none;background-repeat:no-repeat;background-color:#fff;box-shadow:0 6px 22px rgba(0,0,0,.28);z-index:10;}
</style></head><body>
<div id="wrap"><div id="pill"><button id="btnOn"></button><button id="btnZ"></button></div>
<div id="viewer"><div id="stream">__IMAGES__</div></div></div>
<div id="lens"></div>
<script>
const lens=document.getElementById('lens'),btnOn=document.getElementById('btnOn'),btnZ=document.getElementById('btnZ');
const viewer=document.getElementById('viewer');const L=__L__;const zooms=[2,2.5,3.5];
let on=true,zi=1,current=null;
try{const s=localStorage.getItem('pos_mag_on');if(s!==null)on=(s==='1');const z=localStorage.getItem('pos_mag_z');if(z!==null&&zooms[+z])zi=+z;}catch(e){}
function paint(){btnOn.textContent='🔍 مكبر: '+(on?'تشغيل':'إيقاف');btnOn.classList.toggle('off',!on);btnZ.textContent=zooms[zi]+'x';if(!on)lens.style.display='none';}
btnOn.onclick=()=>{on=!on;try{localStorage.setItem('pos_mag_on',on?'1':'0');}catch(e){}paint();};
btnZ.onclick=()=>{zi=(zi+1)%zooms.length;try{localStorage.setItem('pos_mag_z',String(zi));}catch(e){}paint();};
function move(e){
  if(!on)return;const t=e.target;if(!t||t.tagName!=='IMG')return;
  const r=t.getBoundingClientRect();const x=e.clientX-r.left,y=e.clientY-r.top;
  if(x<0||y<0||x>r.width||y>r.height){lens.style.display='none';return;}
  if(current!==t){current=t;lens.style.backgroundImage='url("'+t.src+'")';}
  const z=zooms[zi];
  lens.style.backgroundSize=(r.width*z)+'px '+(r.height*z)+'px';
  lens.style.backgroundPosition=(-(x*z-L/2))+'px '+(-(y*z-L/2))+'px';
  lens.style.left=(e.clientX-L/2)+'px';lens.style.top=(e.clientY-L/2)+'px';lens.style.display='block';
}
document.querySelectorAll('#stream img').forEach(img=>{img.addEventListener('mousemove',move);img.addEventListener('mouseleave',()=>{lens.style.display='none';});});
viewer.addEventListener('scroll',()=>{lens.style.display='none';});
paint();
</script></body></html>"""


def render_viewer(
    pages: list[str],
    accent: str,
    height: int = 735,
    lens_size: int = 142,
) -> None:
    images = "".join(
        f'<img src="{src}" alt="page {i + 1}">' for i, src in enumerate(pages)
    )
    html = (
        VIEWER_HTML.replace("__IMAGES__", images)
        .replace("__ACCENT__", accent)
        .replace("__H__", str(height))
        .replace("__L__", str(lens_size))
    )
    components.html(html, height=height + 4, scrolling=False)


def choice_control(
    label: str,
    options: list,
    key: str,
    default=None,
    format_func=str,
    label_visibility: str = "visible",
):
    if hasattr(st, "segmented_control"):
        return st.segmented_control(
            label,
            options,
            default=default,
            format_func=format_func,
            key=key,
            selection_mode="single",
            label_visibility=label_visibility,
        )
    idx = options.index(default) if default in options else 0
    return st.radio(
        label,
        options,
        index=idx,
        format_func=format_func,
        key=key,
        horizontal=True,
        label_visibility=label_visibility,
    )


def section(title: str) -> None:
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# First-Time Setup / Settings UI (إعدادات الشركة والفروع)
# ---------------------------------------------------------------------------
def render_settings_view(config: dict, is_configured: bool) -> None:
    section("⚙️ إعدادات الشركة والفروع ودليل الحسابات")
    if not is_configured:
        st.info(
            "👋 مرحباً بك! لم يتم العثور على ملف `config.json` بعد. يمكنك مراجعة أو تعديل إعدادات شركتك وفروعك وأرقام الحسابات أدناه ثم الضغط على **حفظ الإعدادات** للبدء."
        )

    c1, c2, c3, c4 = st.columns([1.4, 1.4, 0.8, 0.8])
    with c1:
        company_ar = st.text_input(
            "اسم الشركة (بالعربية)",
            value=config.get("company_name_ar", "الشركة النموذجية للتجارة"),
        )
    with c2:
        company_en = st.text_input(
            "اسم الشركة (بالإنجليزية)",
            value=config.get("company_name_en", "Sample Retail Co."),
        )
    with c3:
        accent_color = st.color_picker(
            "اللون الرئيسي",
            value=config.get("accent_color", "#1D54B5"),
        )
    with c4:
        accent2_color = st.color_picker(
            "اللون الثانوي",
            value=config.get("accent2_color", "#5BC2E7"),
        )

    logo_col1, logo_col2 = st.columns([2, 1])
    with logo_col1:
        uploaded_logo = st.file_uploader(
            "شعار الشركة الاختياري (.png, .jpg, .svg)",
            type=["png", "jpg", "jpeg", "svg"],
            key="settings_logo_uploader",
        )
    logo_data_uri = config.get("logo_data_uri", "")
    if uploaded_logo is not None:
        ext = uploaded_logo.name.lower().rsplit(".", 1)[-1]
        mime = (
            "image/svg+xml"
            if ext == "svg"
            else ("image/png" if ext == "png" else "image/jpeg")
        )
        logo_data_uri = (
            f"data:{mime};base64,"
            + base64.b64encode(uploaded_logo.getvalue()).decode()
        )
    with logo_col2:
        if logo_data_uri:
            if st.button("🗑️ إزالة الشعار المرفوع واستخدام الشعار الافتراضي"):
                logo_data_uri = ""
                config["logo_data_uri"] = ""
                save_app_config(config)
                st.rerun()

    st.divider()
    section("الحسابات العامة وقالب الوصف")
    g1, g2, g3 = st.columns([1, 1, 2])
    with g1:
        shortage_acc = st.text_input(
            "حساب العجز والزيادة (Shortage GL)",
            value=config.get("shortage_account", "SHORTAGE-ACC"),
        )
    with g2:
        currency = st.text_input(
            "العملة",
            value=config.get("currency", "SAR"),
        )
    with g3:
        desc_template = st.text_input(
            "قالب وصف القيد اليومي",
            value=config.get(
                "description_template",
                "إيداع بنكي: {deposit_acc}-{network}, {currency}. 000001",
            ),
            help="المتغيرات المتاحة: {deposit_acc}, {bank_acc}, {network}, {branch_code}, {branch_name}, {currency}",
        )

    st.caption("حسابات التسوية الوسيطة للشبكات (Clearing Accounts):")
    clearing_cfg = dict(config.get("clearing_accounts", {}))
    cl_cols = st.columns(4)
    new_clearing: dict[str, str] = {}
    for idx, (net_key, net_label) in enumerate(ALL_NETWORKS):
        with cl_cols[idx % 4]:
            new_clearing[net_key] = st.text_input(
                f"حساب {net_label}",
                value=clearing_cfg.get(net_key, f"CLEARING-{net_key.upper()}"),
                key=f"cfg_clearing_{net_key}",
            ).strip()

    st.divider()
    section("مجموعات الفروع والفروع التابعة")

    groups_state_key = "settings_branch_groups"
    if groups_state_key not in st.session_state:
        st.session_state[groups_state_key] = copy.deepcopy(
            config.get("branch_groups", DEFAULT_CONFIG["branch_groups"])
        )

    groups_list: list[dict] = st.session_state[groups_state_key]
    updated_groups: list[dict] = []

    for g_idx, grp in enumerate(groups_list):
        with st.expander(
            f"📁 مجموعة الفروع {g_idx + 1}: {grp.get('name', grp.get('id', ''))}",
            expanded=(g_idx == 0),
        ):
            gc1, gc2, gc3, gc4 = st.columns([1, 1.4, 1, 1])
            with gc1:
                g_id = st.text_input(
                    "معرّف المجموعة (ID)",
                    value=grp.get("id", f"GROUP-{g_idx + 1}"),
                    key=f"grp_id_{g_idx}",
                ).strip()
            with gc2:
                g_name = st.text_input(
                    "اسم المجموعة الظاهر",
                    value=grp.get("name", f"Branch Group {g_idx + 1}"),
                    key=f"grp_name_{g_idx}",
                ).strip()
            with gc3:
                g_co = st.text_input(
                    "رمز الشركة في القيد",
                    value=grp.get("company_code", g_id),
                    key=f"grp_co_{g_idx}",
                ).strip()
            with gc4:
                g_shortage = st.text_input(
                    "حساب العجز للمجموعة",
                    value=grp.get("shortage_account", shortage_acc),
                    key=f"grp_sh_{g_idx}",
                ).strip()

            go1, go2, go3, go4 = st.columns([1.2, 1.2, 0.8, 0.8])
            with go1:
                g_multi = st.checkbox(
                    "رفع ملفات حركة متعددة (.xlsx) وتصدير ZIP",
                    value=bool(grp.get("multi_file", False)),
                    key=f"grp_multi_{g_idx}",
                )
            with go2:
                g_append_code = st.checkbox(
                    "إلحاق رمز الفرع بحسابات التسوية والعجز (-CODE)",
                    value=bool(grp.get("append_branch_code_to_gl", False)),
                    key=f"grp_appcode_{g_idx}",
                )
            with go3:
                g_acc1 = st.color_picker(
                    "لون المجموعة",
                    value=grp.get("accent_color", accent_color),
                    key=f"grp_c1_{g_idx}",
                )
            with go4:
                g_acc2 = st.color_picker(
                    "اللون الثانوي للمجموعة",
                    value=grp.get("accent2_color", accent2_color),
                    key=f"grp_c2_{g_idx}",
                )

            branch_rows = []
            for b in grp.get("branches", []):
                net_deps = b.get("network_deposit_accounts") or {}
                branch_rows.append(
                    {
                        "name": b.get("name", ""),
                        "code": b.get("code", ""),
                        "bank_account": b.get("bank_account", ""),
                        "deposit_account": b.get("deposit_account", ""),
                        "master_deposit": net_deps.get("master", ""),
                        "mada_deposit": net_deps.get("mada", ""),
                        "visa_deposit": net_deps.get("visa", ""),
                    }
                )
            if not branch_rows:
                branch_rows = [
                    {
                        "name": "Branch 01",
                        "code": "BR-01",
                        "bank_account": "BANK-001",
                        "deposit_account": "BANK-001",
                        "master_deposit": "",
                        "mada_deposit": "",
                        "visa_deposit": "",
                    }
                ]

            st.caption("جدول الفروع التابعة للمجموعة (يمكنك إضافة أو تعديل أو حذف الصفوف):")
            edited_df = st.data_editor(
                pd.DataFrame(branch_rows),
                num_rows="dynamic",
                use_container_width=True,
                key=f"branches_editor_{g_idx}",
                column_config={
                    "name": st.column_config.TextColumn("اسم الفرع", required=True),
                    "code": st.column_config.TextColumn("رمز الفرع (Code)", required=True),
                    "bank_account": st.column_config.TextColumn("حساب البنك", required=True),
                    "deposit_account": st.column_config.TextColumn("حساب الإيداع الافتراضي"),
                    "master_deposit": st.column_config.TextColumn("إيداع Mastercard (اختياري)"),
                    "mada_deposit": st.column_config.TextColumn("إيداع Mada (اختياري)"),
                    "visa_deposit": st.column_config.TextColumn("إيداع Visa (اختياري)"),
                },
            )

            parsed_branches = []
            for _, r in edited_df.iterrows():
                b_name = str(r.get("name") or "").strip()
                if not b_name:
                    continue
                b_code = str(r.get("code") or "").strip()
                b_bank = str(r.get("bank_account") or "").strip()
                b_dep = str(r.get("deposit_account") or "").strip() or b_bank
                net_deps = {}
                for nk, col_k in [
                    ("master", "master_deposit"),
                    ("mada", "mada_deposit"),
                    ("visa", "visa_deposit"),
                ]:
                    val_d = str(r.get(col_k) or "").strip()
                    if val_d:
                        net_deps[nk] = val_d
                parsed_branches.append(
                    {
                        "name": b_name,
                        "code": b_code,
                        "bank_account": b_bank,
                        "deposit_account": b_dep,
                        "network_deposit_accounts": net_deps,
                    }
                )

            updated_groups.append(
                {
                    "id": g_id or f"GROUP-{g_idx + 1}",
                    "name": g_name or f"Branch Group {g_idx + 1}",
                    "company_code": g_co or g_id or f"GROUP-{g_idx + 1}",
                    "multi_file": g_multi,
                    "append_branch_code_to_gl": g_append_code,
                    "accent_color": g_acc1,
                    "accent2_color": g_acc2,
                    "shortage_account": g_shortage or shortage_acc,
                    "branches": parsed_branches
                    or [
                        {
                            "name": "Branch 01",
                            "code": "BR-01",
                            "bank_account": "BANK-001",
                            "deposit_account": "BANK-001",
                            "network_deposit_accounts": {},
                        }
                    ],
                }
            )

    btn_add, btn_rem, _ = st.columns([1, 1, 2])
    with btn_add:
        if st.button("➕ إضافة مجموعة فروع جديدة", use_container_width=True):
            n_idx = len(groups_list) + 1
            groups_list.append(
                {
                    "id": f"GROUP-{n_idx}",
                    "name": f"Branch Group {n_idx}",
                    "company_code": f"GROUP-{n_idx}",
                    "multi_file": False,
                    "append_branch_code_to_gl": False,
                    "accent_color": accent_color,
                    "accent2_color": accent2_color,
                    "shortage_account": shortage_acc,
                    "branches": [
                        {
                            "name": f"Branch 0{n_idx}",
                            "code": f"BR-0{n_idx}",
                            "bank_account": f"BANK-00{n_idx}",
                            "deposit_account": f"BANK-00{n_idx}",
                            "network_deposit_accounts": {},
                        }
                    ],
                }
            )
            st.rerun()
    with btn_rem:
        if len(groups_list) > 1 and st.button(
            "➖ حذف آخر مجموعة فروع", use_container_width=True
        ):
            groups_list.pop()
            st.rerun()

    st.divider()
    save_col1, save_col2 = st.columns([1.5, 1])
    with save_col1:
        if st.button(
            "💾 حفظ الإعدادات في config.json والانتقال للمطابقة",
            type="primary",
            use_container_width=True,
        ):
            new_cfg = {
                "company_name_ar": company_ar.strip() or "الشركة النموذجية للتجارة",
                "company_name_en": company_en.strip() or "Sample Retail Co.",
                "accent_color": accent_color,
                "accent2_color": accent2_color,
                "logo_data_uri": logo_data_uri,
                "currency": currency.strip() or "SAR",
                "shortage_account": shortage_acc.strip() or "SHORTAGE-ACC",
                "description_template": desc_template.strip()
                or "إيداع بنكي: {deposit_acc}-{network}, {currency}. 000001",
                "clearing_accounts": new_clearing,
                "branch_groups": updated_groups,
            }
            save_app_config(new_cfg)
            st.session_state[groups_state_key] = copy.deepcopy(updated_groups)
            st.session_state["nav_mode"] = "reconcile"
            st.success("✅ تم حفظ الإعدادات بنجاح في `config.json`!")
            st.rerun()
    with save_col2:
        st.download_button(
            "📥 تحميل نسخة من config.json",
            data=json.dumps(
                {
                    "company_name_ar": company_ar,
                    "company_name_en": company_en,
                    "accent_color": accent_color,
                    "accent2_color": accent2_color,
                    "logo_data_uri": logo_data_uri,
                    "currency": currency,
                    "shortage_account": shortage_acc,
                    "description_template": desc_template,
                    "clearing_accounts": new_clearing,
                    "branch_groups": updated_groups,
                },
                ensure_ascii=False,
                indent=2,
            ).encode("utf-8"),
            file_name="config.json",
            mime="application/json",
            use_container_width=True,
        )


# ---------------------------------------------------------------------------
# Main Streamlit Application
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="نظام مطابقة نقاط البيع",
    page_icon="🧾",
    layout="wide",
)
st.markdown(BASE_CSS, unsafe_allow_html=True)

app_config, is_configured = load_app_config()
branch_groups = app_config.get("branch_groups") or DEFAULT_CONFIG["branch_groups"]
group_ids = [g["id"] for g in branch_groups]
group_labels = {g["id"]: g.get("name", g["id"]) for g in branch_groups}

for _gid in group_ids:
    st.session_state.setdefault(f"cashiers_{_gid.lower()}", [])

st.session_state.setdefault("generated_outputs", {})
st.session_state.setdefault("current_doc_index", 0)
st.session_state.setdefault("form_reset_counter", 0)
st.session_state.setdefault("last_doc_batch_signature", ())
st.session_state.setdefault("active_group", group_ids[0])
st.session_state.setdefault("nav_mode", "reconcile" if is_configured else "settings")

if st.session_state.active_group not in group_ids:
    st.session_state.active_group = group_ids[0]

active_group_cfg = get_group_config(app_config, st.session_state.active_group)
active_accent = active_group_cfg.get("accent_color") or app_config.get(
    "accent_color", "#1D54B5"
)
active_accent2 = active_group_cfg.get("accent2_color") or app_config.get(
    "accent2_color", "#5BC2E7"
)
active_soft = hex_to_rgba(active_accent, 0.16)

st.markdown(
    THEME_CSS_TEMPLATE.replace("__ACCENT__", active_accent)
    .replace("__ACCENT2__", active_accent2)
    .replace("__SOFT__", active_soft),
    unsafe_allow_html=True,
)

header_slot = st.empty()
logo_uri = get_header_logo_uri(app_config, active_accent, active_accent2)
header_slot.markdown(
    f'<div class="app-header">'
    f'<img class="brand-logo" src="{logo_uri}" alt="{app_config.get("company_name_en", "Company")}">'
    f"<div>"
    f'<div class="app-title">نظام مطابقة نقاط البيع الآلي</div>'
    f'<div class="app-sub">{app_config.get("company_name_ar", "الشركة النموذجية")} · مطابقة ملف الحركة مع ملف البنك وإنشاء القيود اليومية</div>'
    f"</div></div>",
    unsafe_allow_html=True,
)

top_bar_col1, top_bar_col2 = st.columns([3, 1.2])
with top_bar_col2:
    picked_mode = choice_control(
        "الوضع",
        ["reconcile", "settings"],
        key="nav_mode_selector",
        default=st.session_state.nav_mode,
        format_func=lambda m: (
            "🧾 مطابقة نقاط البيع" if m == "reconcile" else "⚙️ إعدادات الشركة والفروع"
        ),
        label_visibility="collapsed",
    )
    if picked_mode and picked_mode != st.session_state.nav_mode:
        st.session_state.nav_mode = picked_mode
        st.rerun()

if st.session_state.nav_mode == "settings":
    render_settings_view(app_config, is_configured)
    st.stop()

with top_bar_col1:
    picked_group = choice_control(
        "مجموعة الفروع",
        group_ids,
        key="branch_group_choice",
        default=st.session_state.active_group,
        format_func=lambda g: group_labels.get(g, g),
    )
    if picked_group and picked_group != st.session_state.active_group:
        st.session_state.active_group = picked_group
        st.rerun()

group = st.session_state.active_group
group_cfg = get_group_config(app_config, group)
is_multi_file = bool(group_cfg.get("multi_file", False))
branches_in_group = group_cfg.get("branches") or []
branch_names = [b["name"] for b in branches_in_group] or ["Branch 01"]

single_file = None
multi_files = []
active_file = None
selected_multi_file = ""
sub_branch = branch_names[0]

up_col1, up_col2, up_col3 = st.columns([1.25, 1.25, 1])
with up_col1:
    if is_multi_file:
        multi_files = (
            st.file_uploader(
                "ملف الحركة (ملفات متعددة .xlsx)",
                type=["xlsx"],
                accept_multiple_files=True,
                key=f"{group.lower()}_multi_uploader",
            )
            or []
        )
    else:
        single_file = st.file_uploader(
            "ملف الحركة (.xlsx)",
            type=["xlsx"],
            key=f"{group.lower()}_uploader",
        )
        active_file = single_file

with up_col2:
    uploaded_docs = (
        st.file_uploader(
            "ملف البنك — إيصالات المطابقة (.pdf, .png, .jpg)",
            type=["pdf", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key="batch_doc_uploader",
        )
        or []
    )

with up_col3:
    if is_multi_file:
        file_names = [f.name for f in multi_files]
        if file_names:
            selected_multi_file = st.selectbox("الملف الحالي", file_names)
            active_file = next(
                (f for f in multi_files if f.name == selected_multi_file), None
            )
        default_idx = next(
            (
                i
                for i, b in enumerate(branch_names)
                if selected_multi_file and b in selected_multi_file
            ),
            0,
        )
        sub_branch = st.selectbox(
            "الفرع",
            branch_names,
            index=default_idx,
            key=f"branch_{group}_{selected_multi_file}",
        )
    else:
        sub_branch = st.selectbox(
            "الفرع",
            branch_names,
            key=f"branch_{group}",
        )

    b_cfg = get_branch_config(group_cfg, sub_branch)
    shortage_gl = group_cfg.get("shortage_account") or app_config.get(
        "shortage_account", "SHORTAGE-ACC"
    )
    st.caption(
        f'`{b_cfg.get("code", "")}` · البنك `{b_cfg.get("bank_account", "")}` · العجز `{shortage_gl}`'
    )

parsed_system = None
if active_file is not None:
    try:
        parsed_system = parse_system_bytes(active_file.getvalue())
        if not parsed_system["headers_ok"]:
            st.caption(
                "⚠️ لم يُعثر على عنوان «مبلغ بنكي» — تم استخدام العمود G افتراضياً."
            )
    except Exception as exc:
        st.error(f"تعذر قراءة ملف الحركة: {exc}")

queue_key = f"cashiers_{group.lower()}"
queue: list[dict] = st.session_state.setdefault(queue_key, [])
active_file_name = active_file.name if active_file is not None else ""

batch_sig = tuple(d.name for d in uploaded_docs)
if batch_sig != st.session_state.last_doc_batch_signature:
    st.session_state.last_doc_batch_signature = batch_sig
    st.session_state.current_doc_index = 0

total_docs = len(uploaded_docs)
doc_idx = st.session_state.current_doc_index
active_doc = uploaded_docs[doc_idx] if (0 <= doc_idx < total_docs) else None

known_cash_list = (
    [r["cash"] for r in parsed_system["cash_rows"]] if parsed_system else None
)
known_shift_set = (
    {r["shift_suffix"] for r in parsed_system["cash_rows"] if r.get("shift_suffix")}
    if parsed_system
    else None
)
prefill = (
    parse_receipt_filename(active_doc.name, known_cash_list, known_shift_set)
    if active_doc
    else {
        "name": "",
        "cash": 0.0,
        "cash_input": "",
        "variance": 0.0,
        "sign": "-",
        "found": False,
    }
)

st.divider()
view_col, form_col = st.columns([1.2, 1], gap="large")

with view_col:
    if total_docs:
        nav1, nav2, nav3 = st.columns([1, 2.2, 1])
        with nav1:
            if st.button(
                "→ السابق",
                disabled=(doc_idx <= 0),
                use_container_width=True,
            ):
                st.session_state.current_doc_index = max(0, doc_idx - 1)
                st.session_state.form_reset_counter += 1
                st.rerun()
        with nav2:
            if active_doc:
                st.markdown(
                    f'<div style="text-align:center;padding-top:.45rem;">'
                    f'<span class="chip">الإيصال {doc_idx + 1} من {total_docs}</span>'
                    f'<span style="font-size:.85rem;opacity:.7;">{active_doc.name}</span>'
                    f"</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    '<div style="text-align:center;padding-top:.45rem;"><span class="chip">تمت مراجعة جميع الإيصالات ✓</span></div>',
                    unsafe_allow_html=True,
                )
        with nav3:
            if st.button(
                "التالي ←",
                disabled=(doc_idx >= total_docs - 1),
                use_container_width=True,
            ):
                st.session_state.current_doc_index = min(total_docs - 1, doc_idx + 1)
                st.session_state.form_reset_counter += 1
                st.rerun()

    with st.container(height=780, border=False):
        if active_doc is not None:
            doc_bytes = active_doc.getvalue()
            pages = render_document_pages(doc_bytes, active_doc.name)
            if pages:
                render_viewer(pages, active_accent, height=735)
            else:
                b64 = base64.b64encode(doc_bytes).decode()
                st.markdown(
                    f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="725" style="border:1px solid #e6e8eb;border-radius:14px;"></iframe>',
                    unsafe_allow_html=True,
                )
                st.caption(
                    "لتفعيل المكبّر وعرض الصفحات المتعددة ثبّت PyMuPDF: `pip install pymupdf`"
                )
        elif total_docs:
            st.success("✅ تمت مراجعة جميع الإيصالات.")
            if st.button("العودة إلى الإيصال الأول"):
                st.session_state.current_doc_index = 0
                st.session_state.form_reset_counter += 1
                st.rerun()
        else:
            st.info(
                "ارفع إيصالات المطابقة (ملف البنك) لعرضها هنا بجانب نموذج الإدخال."
            )

with form_col:
    section(f"إدخال بيانات الكاشير · {group_labels.get(group, group)}")
    if prefill["found"]:
        st.caption(
            "📎 تمت التعبئة تلقائياً من اسم الملف — راجع القيم قبل الإعتماد."
        )

    sfx = f"{group}_{doc_idx}_{st.session_state.form_reset_counter}"
    var_key = f"var_{sfx}"
    sign_key = f"sign_{sfx}"
    if var_key not in st.session_state:
        st.session_state[var_key] = float(prefill["variance"])
    if sign_key not in st.session_state:
        st.session_state[sign_key] = prefill["sign"]

    def _on_var_input_change():
        val = round(float(st.session_state.get(var_key, 0.0)), 2)
        cur_sign = st.session_state.get(sign_key, "-") or "-"
        if val < 0:
            st.session_state[sign_key] = "-"
            st.session_state[var_key] = val
        elif val > 0 and cur_sign == "-":
            st.session_state[var_key] = -val
        else:
            st.session_state[var_key] = val

    def _on_sign_pill_change():
        picked_sign = st.session_state.get(sign_key) or "-"
        mag = abs(round(float(st.session_state.get(var_key, 0.0)), 2))
        st.session_state[var_key] = -mag if picked_sign == "-" else mag

    name_input = st.text_input(
        "اسم الكاشير", value=prefill["name"], key=f"name_{sfx}"
    )

    f1, f2, f3 = st.columns([1.15, 1, 0.85])
    with f1:
        cash_raw_input = st.text_input(
            "مبلغ الكاش / الوردية -",
            value=str(prefill.get("cash_input", "")),
            placeholder="مثال: 2780 أو -920",
            help="أدخل مبلغ الكاش (مبلغ بنكي)، أو ابدأ بعلامة ناقص (-) متبوعة برقم الوردية بعد النقطتين (مثال: -920) للبحث بالوردية وقراءة الكاش من العمود G حتى لو كان 0.00",
            key=f"cash_{sfx}",
        )
    with f2:
        st.number_input(
            "مبلغ الفرق",
            step=0.01,
            format="%.2f",
            key=var_key,
            on_change=_on_var_input_change,
            help="سالب = عجز (-6.90) ، موجب = زيادة (6.90)",
        )
    with f3:
        if hasattr(st, "segmented_control"):
            st.segmented_control(
                "النوع",
                ["-", "+"],
                key=sign_key,
                format_func=lambda s: "− عجز" if s == "-" else "+ زيادة",
                selection_mode="single",
                label_visibility="hidden",
                on_change=_on_sign_pill_change,
            )
        else:
            st.radio(
                "النوع",
                ["-", "+"],
                key=sign_key,
                format_func=lambda s: "− عجز" if s == "-" else "+ زيادة",
                horizontal=True,
                label_visibility="hidden",
                on_change=_on_sign_pill_change,
            )

    variance_input = round(float(st.session_state[var_key]), 2)
    cash_str = str(cash_raw_input or "").translate(ARABIC_DIGITS).strip()
    is_shift_lookup = cash_str.startswith("-") or cash_str.startswith("−")

    used_rows = {
        c.get("row") for c in queue if c.get("file_name") == active_file_name
    }
    matched_rec = None
    shift_input = ""
    can_approve = False
    cash_val = 0.0

    if parsed_system is None:
        st.info("ارفع ملف الحركة للتحقق المباشر من مبلغ بنكي أو رقم الوردية.")
    elif not cash_str:
        st.info("أدخل مبلغ الكاش / الوردية - للتحقق من ملف الحركة.")
    elif is_shift_lookup:
        target_shift = shift_suffix(cash_str[1:])
        if not target_shift:
            st.info("أدخل رقم الوردية بعد علامة الناقص (مثال: -920).")
        else:
            candidates = [
                r
                for r in parsed_system["cash_rows"]
                if r["shift_suffix"] == target_shift
            ]
            fresh = [r for r in candidates if r["row"] not in used_rows]
            if not candidates:
                st.error(
                    f"الوردية ({target_shift}) غير موجودة في ملف الحركة! تم إيقاف الإعتماد."
                )
            elif not fresh:
                st.warning(
                    f"تم اعتماد الوردية ({target_shift}) مسبقاً في ملف الحركة."
                )
            else:
                matched_rec = fresh[0]
                cash_val = round(float(matched_rec["cash"]), 2)
                can_approve = True
                st.success(
                    f'تم تحديد الوردية {matched_rec["shift_id"]} · مبلغ بنكي: {cash_val:,.2f}'
                )
                st.caption(
                    f'طريقة الدفع: `{matched_rec["payment_method"] or "1"}` · '
                    f'اسم طريقة الدفع: `{matched_rec["payment_name"]}` · '
                    f'مبلغ الحركة: `{matched_rec["cash_movement"]:,.2f}` · '
                    f'مبلغ الخزينة: `{matched_rec["treasury_amount"]:,.2f}`'
                )
    else:
        parsed_cash = normalize_amount(cash_str)
        if parsed_cash is None:
            st.error("قيمة غير صحيحة في حقل «مبلغ الكاش / الوردية -».")
        elif parsed_cash <= 0:
            st.info(
                "أدخل مبلغ بنكي أكبر من صفر، أو استخدم (-الوردية) للبحث بالوردية عندما يكون الكاش 0.00."
            )
        else:
            cash_val = parsed_cash
            candidates = [
                r
                for r in parsed_system["cash_rows"]
                if abs(r["cash"] - cash_val) < TOLERANCE
            ]
            fresh = [r for r in candidates if r["row"] not in used_rows]
            if not candidates:
                st.error("مبلغ بنكي غير موجود في ملف الحركة! تم إيقاف الإعتماد.")
            elif not fresh:
                st.warning(
                    "تم اعتماد هذا المبلغ مسبقاً لكل الورديات المطابقة في ملف الحركة."
                )
            elif len(fresh) == 1:
                matched_rec = fresh[0]
                can_approve = True
                st.success(
                    f'مبلغ بنكي مطابق في ملف الحركة · الوردية {matched_rec["shift_suffix"] or "—"}'
                )
                st.caption(
                    f'طريقة الدفع: `{matched_rec["payment_method"] or "1"}` · '
                    f'اسم طريقة الدفع: `{matched_rec["payment_name"]}` · '
                    f'مبلغ الحركة: `{matched_rec["cash_movement"]:,.2f}` · '
                    f'مبلغ الخزينة: `{matched_rec["treasury_amount"]:,.2f}`'
                )
            else:
                st.warning(
                    f"مبلغ بنكي مكرر في ملف الحركة ({len(fresh)} ورديات). أدخل رقم الوردية."
                )
                shift_input = st.text_input(
                    "الوردية",
                    key=f"shift_{sfx}",
                    placeholder="الأرقام بعد النقطتين ، مثال: 920",
                    help="الورديات المطابقة: "
                    + "، ".join(r["shift_suffix"] for r in fresh),
                )
                st.caption(
                    "الورديات المطابقة: "
                    + " · ".join(f'`{r["shift_suffix"]}`' for r in fresh)
                )
                if shift_input.strip():
                    hits = [
                        r
                        for r in fresh
                        if r["shift_suffix"] == shift_suffix(shift_input)
                    ]
                    if len(hits) == 1:
                        matched_rec = hits[0]
                        can_approve = True
                        st.success(f'تم تحديد الوردية {matched_rec["shift_id"]}.')
                        st.caption(
                            f'طريقة الدفع: `{matched_rec["payment_method"] or "1"}` · '
                            f'اسم طريقة الدفع: `{matched_rec["payment_name"]}` · '
                            f'مبلغ الحركة: `{matched_rec["cash_movement"]:,.2f}` · '
                            f'مبلغ الخزينة: `{matched_rec["treasury_amount"]:,.2f}`'
                        )
                    else:
                        st.error("رقم الوردية لا يطابق أي وردية لهذا المبلغ.")

    raw_inputs: dict[str, str] = {}
    live_totals: dict[str, float] = {}
    parse_error = False

    def network_input(net_key: str, label: str) -> None:
        global parse_error
        raw = st.text_input(
            label,
            key=f"net_{net_key}_{sfx}",
            placeholder="أرقام مفصولة بمسافة ، مثال: 1992 1203",
        )
        raw_inputs[net_key] = raw
        try:
            live_totals[net_key] = round(sum(parse_drive_amounts(raw)), 2)
        except ValueError:
            live_totals[net_key] = 0.0
            parse_error = True
            st.caption(f"⚠️ قيمة غير صحيحة في {label}")

    for key, label in STANDARD_NETWORKS:
        network_input(key, label)

    with st.expander(
        "خيارات إضافية: الخصم والشبكات النادرة (0)", expanded=False
    ):
        discount_input = st.number_input(
            "خصم",
            value=0.0,
            min_value=0.0,
            step=0.01,
            format="%.2f",
            key=f"disc_{sfx}",
        )
        rc1, rc2 = st.columns(2)
        for i, (key, label) in enumerate(RARE_NETWORKS):
            with rc1 if i % 2 == 0 else rc2:
                network_input(key, label)

    is_match = True
    calc = None
    if matched_rec is not None and not parse_error:
        calc = compute_calculated(
            cash_val,
            matched_rec["cash_movement"],
            live_totals,
            matched_rec["app_amounts"],
            float(discount_input),
            variance_input,
        )
        is_match = abs(calc["calculated"] - variance_input) <= TOLERANCE + 1e-9
        st.markdown(
            f'<div class="calc-box">'
            f'<div class="calc-item"><div class="k">المحسوب</div><div class="v">{fmt_signed(calc["calculated"])}</div></div>'
            f'<div class="calc-item"><div class="k">مبلغ الفرق</div><div class="v">{fmt_signed(variance_input)}</div></div>'
            f'<div class="calc-item"><div class="k">فرق النقد · الشبكات</div><div class="v">{calc["cash_diff"]:+,.2f} · {calc["net_diff"]:+,.2f}</div></div>'
            f"</div>",
            unsafe_allow_html=True,
        )
        missing = [
            NETWORK_LABELS[k]
            for k in ALL_NETWORK_KEYS
            if live_totals.get(k, 0) > 0
            and abs(matched_rec["app_amounts"].get(k, 0.0)) < TOLERANCE
        ]
        if missing:
            st.info(
                f'الشبكة ({"، ".join(missing)}) غير موجودة في ملف الحركة — سيُسجَّل كامل مبلغها في قيد واحد بين البنك وحساب العجز.'
            )
        if is_match:
            st.success("المحسوب يطابق مبلغ الفرق ✓")
        else:
            st.warning(
                "⚠️ المحسوب لا يطابق مبلغ الفرق — سيتم تلوين خلية مبلغ الفرق باللون الأحمر."
            )

    if st.button(
        "اعتماد الكاشير",
        type="primary",
        use_container_width=True,
        disabled=(not can_approve or parse_error),
    ):
        if not name_input.strip():
            st.error("يرجى إدخال اسم الكاشير.")
        else:
            lists = {
                k: parse_drive_amounts(raw_inputs.get(k, ""))
                for k in ALL_NETWORK_KEYS
            }
            queue.append(
                {
                    "file_name": active_file_name,
                    "sub_branch": sub_branch,
                    "name": name_input.strip(),
                    "cash": cash_val,
                    "lookup_by_shift": is_shift_lookup,
                    "variance": variance_input,
                    "discount": float(discount_input),
                    "shift_id": matched_rec["shift_id"],
                    "row": matched_rec["row"],
                    "network_raw": {k: v.strip() for k, v in raw_inputs.items()},
                    "network_lists": lists,
                    "network_totals": {
                        k: round(sum(v), 2) for k, v in lists.items()
                    },
                    "calculated": calc["calculated"] if calc else None,
                    "match": is_match,
                }
            )
            st.session_state.generated_outputs.pop(group, None)
            st.session_state.current_doc_index += 1
            st.session_state.form_reset_counter += 1
            st.rerun()

st.divider()
section(f"المعاينة المباشرة · {len(queue)} كاشير")
if queue:
    preview = []
    for i, c in enumerate(queue, start=1):
        row = {"#": i}
        if len(branch_names) > 1 or is_multi_file:
            row["الفرع"] = c["sub_branch"]
        row["اسم الكاشير"] = c["name"]
        row["مبلغ بنكي"] = c["cash"]
        row["الوردية"] = shift_suffix(c.get("shift_id", ""))
        for k, label in STANDARD_NETWORKS:
            row[label] = c["network_totals"].get(k, 0.0)
        rare_total = round(
            sum(c["network_totals"].get(k, 0.0) for k, _ in RARE_NETWORKS), 2
        )
        if rare_total:
            row["شبكات نادرة"] = rare_total
        row["خصم"] = c.get("discount", 0.0)
        row["المحسوب"] = c.get("calculated")
        row["مبلغ الفرق"] = c["variance"]
        row["الحالة"] = "✅ مطابق" if c.get("match") else "🔴 غير مطابق"
        preview.append(row)

    df = pd.DataFrame(preview)
    df = df[list(reversed(df.columns))]
    st.dataframe(df, hide_index=True, use_container_width=True)

    b1, b2, _sp = st.columns([1, 1, 3])
    with b1:
        if st.button("حذف آخر كاشير", use_container_width=True):
            queue.pop()
            st.session_state.current_doc_index = max(
                0, st.session_state.current_doc_index - 1
            )
            st.session_state.form_reset_counter += 1
            st.session_state.generated_outputs.pop(group, None)
            st.rerun()
    with b2:
        if st.button("مسح القائمة", use_container_width=True):
            queue.clear()
            st.session_state.current_doc_index = 0
            st.session_state.form_reset_counter += 1
            st.session_state.generated_outputs.pop(group, None)
            st.rerun()
else:
    st.caption("لا توجد بيانات معتمدة بعد.")

st.divider()
can_generate = bool(queue) and (
    (not is_multi_file and single_file is not None)
    or (is_multi_file and bool(multi_files))
)
if st.button(
    "إنشاء الملفات وتحميلها",
    type="primary",
    use_container_width=True,
    disabled=(not can_generate),
):
    try:
        if not is_multi_file:
            j_rows, s_rows, fmt_bytes, warns = process_single_workbook(
                app_config,
                single_file.getvalue(),
                queue,
                group,
                sub_branch,
            )
            st.session_state.generated_outputs[group] = {
                "journal": rows_to_excel_bytes(j_rows),
                "summary": rows_to_excel_bytes(s_rows, SUMMARY_COLUMNS),
                "formatted": fmt_bytes,
                "warnings": warns,
            }
        else:
            all_j, all_s, all_w = [], [], []
            zip_buf = io.BytesIO()
            done = set()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in multi_files:
                    file_cashiers = [
                        c for c in queue if c["file_name"] == f.name
                    ]
                    if not file_cashiers:
                        continue
                    done.add(f.name)
                    branch_for_file = file_cashiers[0]["sub_branch"]
                    j_rows, s_rows, fmt_bytes, warns = process_single_workbook(
                        app_config,
                        f.getvalue(),
                        file_cashiers,
                        group,
                        branch_for_file,
                    )
                    all_j.extend(j_rows)
                    all_s.extend(s_rows)
                    all_w.extend(warns)
                    zf.writestr(
                        f"Formatted_{branch_for_file}_{f.name}", fmt_bytes
                    )
            for c in queue:
                if c["file_name"] not in done:
                    all_w.append(
                        f"ملف الحركة '{c['file_name']}' الخاص بالكاشير '{c['name']}' غير مرفوع حالياً."
                    )
            st.session_state.generated_outputs[group] = {
                "journal": rows_to_excel_bytes(all_j),
                "summary": rows_to_excel_bytes(
                    all_s, ["الفرع"] + SUMMARY_COLUMNS
                ),
                "zip": zip_buf.getvalue(),
                "warnings": all_w,
            }
    except Exception as exc:
        st.error(f"حدث خطأ أثناء معالجة الملفات: {exc}")

if not can_generate:
    st.caption(
        "ارفع ملف الحركة واعتمد كاشيراً واحداً على الأقل لتفعيل الإنشاء."
    )

out = st.session_state.generated_outputs.get(group)
if out:
    for w in out.get("warnings", []):
        st.warning(w)
    XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            "📥 ملف القيد اليومي",
            out["journal"],
            f"Journal_Entry_Clean_{group}.xlsx",
            XLSX,
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "📥 تقرير الكاشير",
            out["summary"],
            f"Cashier_Summary_{group}.xlsx",
            XLSX,
            use_container_width=True,
        )
    with d3:
        if is_multi_file:
            st.download_button(
                "📦 ملفات الحركة المنسقة (.zip)",
                out["zip"],
                f"Formatted_{group}_App_Files.zip",
                "application/zip",
                use_container_width=True,
            )
        else:
            st.download_button(
                "📥 ملف الحركة المنسق",
                out["formatted"],
                f"Formatted_App_Data_{group}.xlsx",
                XLSX,
                use_container_width=True,
            )
