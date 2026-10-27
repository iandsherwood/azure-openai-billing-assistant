#!/usr/bin/env python3
"""
Utility Billing Automation with Azure OpenAI
=============================================

Pipeline:
    Gmail bills → Azure OpenAI extracts amounts → Spreadsheet stores raw totals
    → Script reads spreadsheet formulas to compute tenant shares
    → Azure OpenAI analyzes month-over-month changes + anomalies
    → Synopsis email lands in your inbox (you on the loop, not in the way)
    → Gmail draft to tenant sits in Drafts until you hit send

Usage:
    python3 run_billing.py                  # Read Gmail for last month, full pipeline
    python3 run_billing.py --month 2026-04  # Specific billing month
    python3 run_billing.py --manual         # Skip Gmail, enter amounts by hand
    python3 run_billing.py --auth           # One-time Gmail OAuth setup (run first)
    python3 run_billing.py --dry-run        # Skip ledger write + drafts; preview only
    python3 run_billing.py --no-synopsis    # Skip the synopsis email to self
    python3 run_billing.py --no-tenant-draft # Skip the Gmail draft to tenant

Dependencies:
    pip3 install reportlab openpyxl google-api-python-client google-auth-oauthlib openai
"""

import argparse
import ast
import base64
import json
import mimetypes
import operator
import os
import pickle
import re
import statistics
import subprocess
import sys
import traceback
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path

try:
    from . import llm
except ImportError:  # python billing_automation/run_billing.py
    import llm

# ── Third-party ──────────────────────────────────────────────────────────────
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.colors import HexColor, white, black
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import openpyxl
except ImportError as e:
    sys.exit(f"Missing dependency: {e}\nRun: pip3 install reportlab openpyxl")

# ── Paths ────────────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent.resolve()
CONFIG_FILE = SCRIPT_DIR / "billing_config.json"
TOKEN_FILE  = SCRIPT_DIR / "gmail_token.pickle"
CREDS_FILE  = SCRIPT_DIR / "credentials.json"

# ── Brand ────────────────────────────────────────────────────────────────────
BRAND_COLOR  = HexColor("#3D6B7A")
ROW_ALT_BG   = HexColor("#EAF3F6")
DARK_GRAY    = HexColor("#404040")
MID_GRAY     = HexColor("#888888")

PAGE_W, PAGE_H = letter
MARGIN = 54


# ══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════════════

def load_config() -> dict:
    if not CONFIG_FILE.exists():
        sys.exit(f"Config not found: {CONFIG_FILE}")
    with open(CONFIG_FILE) as f:
        return json.load(f)


def resolve_path(p: str) -> Path:
    return Path(os.path.expandvars(os.path.expanduser(p)))


# ══════════════════════════════════════════════════════════════════════════════
# FONTS
# ══════════════════════════════════════════════════════════════════════════════

def setup_fonts() -> tuple[str, str]:
    """Try Gill Sans MT (matches Pages template); fall back to Helvetica."""
    candidate_dirs = [
        "/System/Library/Fonts/Supplemental",
        "/Library/Fonts",
        str(Path.home() / "Library/Fonts"),
    ]
    for d in candidate_dirs:
        reg = Path(d) / "Gill Sans MT.ttf"
        bold = Path(d) / "Gill Sans MT Bold.ttf"
        if reg.exists():
            try:
                pdfmetrics.registerFont(TTFont("GillSans", str(reg)))
                if bold.exists():
                    pdfmetrics.registerFont(TTFont("GillSans-Bold", str(bold)))
                    return "GillSans", "GillSans-Bold"
                return "GillSans", "Helvetica-Bold"
            except Exception:
                pass
    return "Helvetica", "Helvetica-Bold"


# ══════════════════════════════════════════════════════════════════════════════
# SPREADSHEET, write raw amounts, evaluate formulas to compute tenant shares
# ══════════════════════════════════════════════════════════════════════════════

# Safe formula evaluator: parses arithmetic expressions from spreadsheet
# formula text, substitutes cell references with concrete numbers, and
# evaluates using a whitelisted set of operators + MIN/MAX.
# This is what lets the spreadsheet stay the single source of truth on splits
# without us needing Excel itself to recompute.

_AST_OPS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.USub: operator.neg,
    ast.UAdd: operator.pos,
}

def _eval_node(node) -> float:
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return float(node.value)
    if isinstance(node, ast.BinOp) and type(node.op) in _AST_OPS:
        return _AST_OPS[type(node.op)](_eval_node(node.left), _eval_node(node.right))
    if isinstance(node, ast.UnaryOp) and type(node.op) in _AST_OPS:
        return _AST_OPS[type(node.op)](_eval_node(node.operand))
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
        fn = node.func.id.lower()
        args = [_eval_node(a) for a in node.args]
        if fn == "min":
            return min(args)
        if fn == "max":
            return max(args)
        if fn == "sum":
            return sum(args)
        raise ValueError(f"Unsupported function: {fn}")
    raise ValueError(f"Unsupported expression node: {ast.dump(node)}")


def evaluate_excel_formula(formula: str, cell_values: dict, worksheet=None) -> float:
    """
    Evaluate an Excel formula string (without leading '=') after substituting
    cell references with concrete values from cell_values dict.

    If a referenced cell is not in cell_values and a worksheet is provided,
    read it directly from the sheet (blank/non-numeric cells default to 0).
    Lets the script handle credits in column I, etc., without pre-listing them.

    Supports: + - * /, parentheses, numbers, MIN, MAX, SUM, cell refs like B7.
    """
    expr = formula.lstrip("=").strip()
    def _sub_ref(m):
        ref = m.group(0)
        if ref in cell_values:
            return repr(float(cell_values[ref]))
        if worksheet is not None:
            try:
                v = worksheet[ref].value
                if isinstance(v, (int, float)):
                    return repr(float(v))
                # If the cell contains its own formula, we can't recursively eval
                # safely here without risking cycles. Treat as 0.
            except Exception:
                pass
            return "0.0"
        raise KeyError(f"Cell {ref} referenced in formula but no value provided.")
    expr = re.sub(r"\b[A-Z]+\d+\b", _sub_ref, expr)
    # MIN/MAX are case-sensitive in AST, but Excel users write either case.
    # Lowercase function names so our AST walker handles them.
    expr = re.sub(r"\bMIN\b", "min", expr)
    expr = re.sub(r"\bMAX\b", "max", expr)
    expr = re.sub(r"\bSUM\b", "sum", expr)
    tree = ast.parse(expr, mode="eval")
    return _eval_node(tree.body)


def write_bills_to_spreadsheet(bills: dict, billing_month: date, config: dict) -> tuple[Path, str, int] | None:
    """
    Write raw bill amounts into Master_Monthly_Utilities_Template.
    Returns (workbook_path, sheet_name, data_row) on success, or None.
    """
    path = resolve_path(config.get("spreadsheet_path", ""))
    if not path.exists():
        print(f"  ⚠  Spreadsheet not found: {path}")
        return None

    try:
        wb = openpyxl.load_workbook(str(path))
    except Exception as e:
        print(f"  ⚠  Could not open spreadsheet: {e}")
        return None

    sheet_name = f"Monthly Utilities {billing_month.year}"
    if sheet_name not in wb.sheetnames:
        print(f"  ⚠  Sheet '{sheet_name}' not found.")
        return None

    ws = wb[sheet_name]
    month_name = billing_month.strftime("%B").lower()

    target_row = None
    for row in ws.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        if isinstance(cell.value, str) and cell.value.strip().lower() == month_name:
            target_row = cell.row
            break

    if target_row is None:
        print(f"  ⚠  Could not find row for '{billing_month.strftime('%B')}' in sheet.")
        return None

    cols = config["spreadsheet_columns"]
    for key, col_letter in cols.items():
        if key in bills and bills[key] is not None and bills[key] > 0:
            ws[f"{col_letter}{target_row}"].value = bills[key]

    try:
        wb.save(str(path))
        print(f"  ✓ Wrote raw amounts to row {target_row} ({billing_month.strftime('%B')})")
        return path, sheet_name, target_row
    except Exception as e:
        print(f"  ⚠  Save failed: {e}")
        return None


def read_tenant_shares(
    workbook_path: Path,
    sheet_name: str,
    data_row: int,
    bills: dict,
    config: dict,
) -> dict:
    """
    Reopen workbook (formulas not values), read each tenant-share-row formula,
    substitute raw amounts we just wrote, evaluate. Returns:
        {
            "per_utility": {utility_key: tenant_share_amount},
            "total":       float,
            "raw_amounts": {utility_key: raw_amount},
        }
    """
    wb = openpyxl.load_workbook(str(workbook_path))
    ws = wb[sheet_name]
    tenant_share_row = data_row + 1   # Tenant Share row is always directly below data row

    cols = config["spreadsheet_columns"]
    # Build cell_values map for the data row.
    # Pre-populate ALL utility cells (B-G) at 0.0 so formulas referencing cells
    # we didn't explicitly write (e.g., pet_waste when service is cancelled) still
    # evaluate cleanly. Then let actual bill values override.
    cell_values = {}
    for key in config["utility_keys_in_order"]:
        col_letter = cols.get(key)
        if col_letter:
            cell_values[f"{col_letter}{data_row}"] = 0.0
    for key, col_letter in cols.items():
        if key in bills and bills[key] is not None:
            cell_values[f"{col_letter}{data_row}"] = float(bills[key])

    per_utility = {}
    for utility_key in config["utility_keys_in_order"]:
        col_letter = cols.get(utility_key)
        if not col_letter:
            continue
        formula_cell = ws[f"{col_letter}{tenant_share_row}"]
        formula = formula_cell.value
        if formula is None:
            per_utility[utility_key] = 0.0
            continue
        if isinstance(formula, (int, float)):
            # Hardcoded value in tenant share row (e.g., G4 = 80, capped manually)
            per_utility[utility_key] = float(formula)
            continue
        if not isinstance(formula, str) or not formula.startswith("="):
            per_utility[utility_key] = 0.0
            continue
        try:
            per_utility[utility_key] = round(evaluate_excel_formula(formula, cell_values), 2)
        except Exception as e:
            print(f"  ⚠  Could not evaluate formula for {utility_key}: {formula} ({e})")
            per_utility[utility_key] = 0.0

    # Total: evaluate the H column formula on the data row
    total_col = cols.get("tenant_owes", "H")
    total_cell = ws[f"{total_col}{data_row}"]
    total_formula = total_cell.value
    total = sum(per_utility.values())  # fallback
    if isinstance(total_formula, str) and total_formula.startswith("="):
        try:
            total = round(evaluate_excel_formula(total_formula, cell_values), 2)
        except Exception as e:
            print(f"  ⚠  Could not evaluate total formula: {e}")

    return {
        "per_utility": per_utility,
        "total": total,
        "raw_amounts": {k: bills.get(k, 0) for k in config["utility_keys_in_order"]},
    }


def read_history(workbook_path: Path, sheet_name: str, data_row: int, config: dict, months_back: int = 6) -> list[dict]:
    """
    Read the prior N months of bills from the spreadsheet (rows above data_row).
    Returns list of {"month": str, "row": int, "raw": {utility_key: amount}, "tenant_total": float}.
    Used by the analysis layer to compute %change and anomaly metrics.
    """
    wb = openpyxl.load_workbook(str(workbook_path), data_only=True)
    ws = wb[sheet_name]
    cols = config["spreadsheet_columns"]
    history = []

    # Step backwards 2 rows at a time (data row → tenant share row → next data row)
    for delta in range(2, months_back * 2 + 1, 2):
        r = data_row - delta
        if r < 3:
            break
        month_cell = ws.cell(row=r, column=1).value
        if not isinstance(month_cell, str):
            continue
        raw = {}
        for utility_key in config["utility_keys_in_order"]:
            col_letter = cols.get(utility_key)
            v = ws[f"{col_letter}{r}"].value if col_letter else None
            raw[utility_key] = float(v) if isinstance(v, (int, float)) else 0.0
        tenant_owes = ws[f"{cols.get('tenant_owes', 'H')}{r}"].value
        history.append({
            "month": month_cell.strip(),
            "row": r,
            "raw": raw,
            "tenant_total": float(tenant_owes) if isinstance(tenant_owes, (int, float)) else None,
        })

    # Reverse so most recent is last
    history.reverse()
    return history


# ══════════════════════════════════════════════════════════════════════════════
# ANALYSIS, anomaly detection + LLM commentary
# ══════════════════════════════════════════════════════════════════════════════

def analyze_bills(current: dict, history: list[dict], config: dict) -> dict:
    """
    Compute per-utility %change vs. prior month, Z-score vs. history window,
    plus flags. Returns structured analysis ready for both the synopsis email
    and the Azure LLM commentary step.
    """
    thresholds = config.get("analysis", {})
    mom_alert = thresholds.get("mom_change_alert_pct", 25)
    z_alert = thresholds.get("zscore_alert_threshold", 2.0)

    flags = []
    per_utility = []
    labels = config["utility_labels"]
    raw_now = current["raw_amounts"]

    prior_month = history[-1] if history else None

    for key in config["utility_keys_in_order"]:
        label = labels.get(key, key)
        cur = float(raw_now.get(key, 0) or 0)
        prior = float(prior_month["raw"].get(key, 0) or 0) if prior_month else 0.0

        # %change
        if prior > 0:
            pct = (cur - prior) / prior * 100
        elif cur > 0:
            pct = None  # appearing for first time
        else:
            pct = 0.0

        # Z-score across history window (raw amounts)
        series = [m["raw"].get(key, 0) or 0 for m in history if (m["raw"].get(key, 0) or 0) > 0]
        zscore = None
        if len(series) >= 3:
            mean = statistics.mean(series)
            stdev = statistics.stdev(series) if len(series) >= 2 else 0
            if stdev > 0:
                zscore = (cur - mean) / stdev

        # Flags
        if pct is not None and abs(pct) >= mom_alert and cur > 0:
            flags.append({
                "utility": label,
                "kind": "mom_change",
                "detail": f"{label} changed {pct:+.1f}% vs prior month ({prior:.2f} → {cur:.2f})",
            })
        if zscore is not None and abs(zscore) >= z_alert:
            flags.append({
                "utility": label,
                "kind": "outlier",
                "detail": f"{label} is {zscore:+.2f} stdev from {len(series)}-month mean (${cur:.2f} vs avg ${statistics.mean(series):.2f})",
            })
        if prior > 0 and cur == 0:
            flags.append({
                "utility": label,
                "kind": "missing",
                "detail": f"{label} bill not found this month (last month was ${prior:.2f})",
            })

        per_utility.append({
            "utility_key": key,
            "label": label,
            "current": cur,
            "prior": prior,
            "pct_change": pct,
            "zscore": zscore,
        })

    return {
        "per_utility": per_utility,
        "flags": flags,
        "tenant_total_now": current["total"],
        "tenant_total_prior": prior_month["tenant_total"] if prior_month else None,
        "history_months": [m["month"] for m in history],
    }


def azure_commentary(analysis: dict, billing_month: date, config: dict) -> str:
    """Generate commentary with the selected provider (legacy function name)."""
    provider = llm.get_provider(config)
    client, deployment = _get_llm_client(config)
    if not client:
        return ""
    try:
        response = llm.complete(
            client, provider, deployment, llm.commentary_prompt(analysis, billing_month),
            max_tokens=llm.COMMENTARY_MAX_TOKENS,
            temperature=llm.COMMENTARY_TEMPERATURE,
        )
        return response.text.strip()
    except Exception as e:
        name = "Azure" if provider == "azure_openai" else "Claude Foundry"
        # SDK error bodies may echo private bill content; keep diagnostics bounded.
        return f"({name} commentary failed: {type(e).__name__})"


# ══════════════════════════════════════════════════════════════════════════════
# PDF INVOICE
# ══════════════════════════════════════════════════════════════════════════════

def generate_invoice_pdf(
    per_utility: dict,
    total: float,
    billing_month: date,
    send_date: date,
    config: dict,
    output_dir: Path,
) -> Path:
    """Generate the tenant invoice PDF using spreadsheet-computed values."""
    font_reg, font_bold = setup_fonts()

    month_upper = billing_month.strftime("%B").upper()
    month_long  = billing_month.strftime("%B %Y")
    date_str    = send_date.strftime("%m/%d/%Y")

    tenant_name    = config.get("tenant_name", "Tenant")
    landlord_email = config.get("landlord_email", "")
    prop_addr      = config.get("property_billing_address", "")
    contact_addr   = config.get("landlord_contact_address", "")
    labels         = config.get("utility_labels", {})

    prop_id = config.get("property_short_id", "Property")
    fname = f"{prop_id}_-_{billing_month.strftime('%B_%Y')}_Utilities_Invoice.pdf"
    out_path = output_dir / fname

    c = rl_canvas.Canvas(str(out_path), pagesize=letter)
    W, H = PAGE_W, PAGE_H
    M = MARGIN
    UW = W - 2 * M

    # HEADER
    c.setFillColor(BRAND_COLOR)
    c.rect(M, H - M - 4, UW, 4, fill=1, stroke=0)
    c.setFont(font_bold, 16)
    invoice_header = config.get("invoice_header", config.get("landlord_name", "")).upper()
    c.drawCentredString(W / 2, H - M - 22, invoice_header)
    c.rect(M, H - M - 27, UW, 1.2, fill=1, stroke=0)
    c.setFont(font_bold, 11)
    c.drawString(M, H - M - 43, prop_addr.upper())
    c.drawString(M, H - M - 59, f"{month_upper} UTILITIES")

    # BILL-TO BLOCK
    bt_x = M + 220
    bt_y = H - M - 43
    def label_value(y, label, value, label_font, value_font, size=10):
        c.setFont(label_font, size); c.setFillColor(black)
        c.drawString(bt_x, y, label)
        lw = c.stringWidth(label, label_font, size)
        c.setFont(value_font, size)
        c.drawString(bt_x + lw, y, value)
    label_value(bt_y,      "To:  ",             tenant_name, font_bold, font_reg)
    label_value(bt_y - 16, "Billing Period:  ", month_long,  font_bold, font_reg)
    label_value(bt_y - 32, "Date:  ",           date_str,    font_bold, font_reg)

    # INVOICE SIDEBAR
    inv_y = H - M - 105
    c.setFont(font_bold, 42); c.setFillColor(DARK_GRAY)
    c.drawString(M, inv_y, "INVOICE")
    c.setFont(font_reg, 8); c.setFillColor(MID_GRAY)
    c.drawString(M, inv_y - 16, contact_addr)
    c.drawString(M, inv_y - 28, landlord_email)

    # TABLE
    table_top = H - M - 160
    hdr_h = 24
    row_h = 22
    col1_x = M + 16
    col2_x = M + UW - 10

    # Show only utilities with > $0 tenant share
    rows = []
    for key in config["utility_keys_in_order"]:
        amt = per_utility.get(key, 0)
        if amt > 0:
            rows.append((labels.get(key, key), amt))

    c.setFillColor(BRAND_COLOR)
    c.rect(M, table_top - hdr_h, UW, hdr_h, fill=1, stroke=0)
    c.setFont(font_bold, 10); c.setFillColor(white)
    c.drawString(col1_x, table_top - hdr_h + 8, "Utility")
    c.drawRightString(col2_x, table_top - hdr_h + 8, "Cost")

    row_base = table_top - hdr_h
    for i, (name, amount) in enumerate(rows):
        ry = row_base - (i + 1) * row_h
        if i % 2 == 1:
            c.setFillColor(ROW_ALT_BG)
            c.rect(M, ry, UW, row_h, fill=1, stroke=0)
        c.setFont(font_reg, 10); c.setFillColor(black)
        c.drawString(col1_x, ry + 7, name)
        c.drawRightString(col2_x, ry + 7, f"${amount:,.2f}")

    c.setStrokeColor(HexColor("#CCDDDF"))
    c.setLineWidth(0.4)
    c.setDash(2, 3)
    for i in range(1, len(rows)):
        dy = row_base - i * row_h
        c.line(M + 8, dy, M + UW - 8, dy)
    c.setDash()

    n = len(rows)
    total_y = row_base - (n + 1) * row_h
    c.setStrokeColor(BRAND_COLOR); c.setLineWidth(1.5)
    c.line(M, total_y + row_h, M + UW, total_y + row_h)
    c.setFont(font_bold, 10); c.setFillColor(black)
    c.drawString(col1_x, total_y + 7, "Total")
    c.drawRightString(col2_x, total_y + 7, f"${total:,.2f}")

    # FOOTER
    footer_y = M + 24
    c.setFillColor(BRAND_COLOR)
    c.rect(M, footer_y, UW, 1.5, fill=1, stroke=0)
    c.setFont(font_reg, 8); c.setFillColor(MID_GRAY)
    c.drawRightString(M + UW, footer_y - 11, "Page 1")

    c.save()
    return out_path


# ══════════════════════════════════════════════════════════════════════════════
# GMAIL, read bills (Azure-first) + create drafts (replaces Apple Mail)
# ══════════════════════════════════════════════════════════════════════════════

def _decode_body(payload: dict) -> str:
    """
    Walk a Gmail payload tree and return ALL text/plain and text/html parts
    concatenated. Modern marketing emails put the dollar amount in HTML only;
    older transactional emails put it in text/plain only. Concatenating both
    gives the LLM a fair shot regardless.
    """
    def _collect(p, kind):
        out = []
        if p.get("mimeType", "") == kind:
            data = p.get("body", {}).get("data", "")
            if data:
                out.append(base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore"))
        for sub in p.get("parts", []):
            out.extend(_collect(sub, kind))
        return out

    plain = _collect(payload, "text/plain")
    html_parts = _collect(payload, "text/html")
    pieces = []
    if plain:
        pieces.append("=== PLAIN ===\n" + "\n".join(plain))
    if html_parts:
        pieces.append("=== HTML ===\n" + "\n".join(html_parts))
    return "\n\n".join(pieces)


def _get_azure_client(config: dict):
    """Compatibility wrapper for callers that specifically need Azure OpenAI."""
    return llm.get_client(config, provider="azure_openai")


def _get_llm_client(config: dict):
    if llm.get_provider(config) == "azure_openai":
        return _get_azure_client(config)
    return llm.get_client(config)


def _extract_amount_azure(body: str, label: str, client, deployment: str,
                          snippet: str = "", provider: str = "azure_openai") -> float | None:
    """Extract using the shared prompt/parser (legacy function name)."""
    try:
        response = llm.complete(
            client, provider, deployment, llm.extraction_prompt(body, label, snippet),
            max_tokens=llm.EXTRACTION_MAX_TOKENS,
            temperature=llm.EXTRACTION_TEMPERATURE,
        )
        return llm.parse_amount(response.text)
    except Exception as e:
        name = "Azure" if provider == "azure_openai" else "Claude Foundry"
        print(f"  ⚠  {name} parse failed for {label}: {type(e).__name__}")
        return None


def _extract_amount_regex(body: str, pattern: str | None) -> float | None:
    if pattern:
        m = re.search(pattern, body, re.IGNORECASE)
        if m:
            try:
                return float(m.group(1).replace(",", ""))
            except (ValueError, IndexError):
                pass
    amounts = re.findall(r"\$\s*([\d,]+\.\d{2})", body)
    if amounts:
        return max(float(a.replace(",", "")) for a in amounts)
    return None


def _gmail_service(config: dict):
    """Return an authenticated Gmail API service, or None."""
    try:
        from googleapiclient.discovery import build
        from google.auth.transport.requests import Request
    except ImportError:
        print("  ⚠  google-api-python-client not installed.")
        return None
    if not TOKEN_FILE.exists():
        print("  ⚠  Gmail not authenticated. Run: python3 run_billing.py --auth")
        return None
    with open(TOKEN_FILE, "rb") as f:
        creds = pickle.load(f)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return build("gmail", "v1", credentials=creds)


def read_bills_from_gmail(billing_month: date, config: dict) -> dict | None:
    service = _gmail_service(config)
    if not service:
        return None

    after = (billing_month.replace(day=1) - timedelta(days=5)).strftime("%Y/%m/%d")
    before = (billing_month.replace(day=28) + timedelta(days=15)).strftime("%Y/%m/%d")
    labels = config.get("utility_labels", {})

    provider = llm.get_provider(config)
    provider_name = "Azure OpenAI" if provider == "azure_openai" else "Claude Foundry"
    az_client, az_deployment = _get_llm_client(config)
    if az_client:
        print(f"  Using {provider_name} for bill parsing.")
    else:
        print(f"  {provider_name} not configured, using regex fallback.")

    bills = {}
    for key, search_cfg in config.get("gmail_searches", {}).items():
        if key.startswith("_") or search_cfg.get("_skip"):
            continue
        sender               = search_cfg.get("sender", "")
        pattern              = search_cfg.get("amount_pattern")
        must_contain         = search_cfg.get("must_contain")
        subject_must_contain = search_cfg.get("subject_must_contain")
        label                = labels.get(key, key)

        # Build query with optional content/subject filters for accounts with
        # multiple matching emails per cycle (e.g., reminders + confirmations)
        # or multiple accounts at the same sender (e.g., two NW Natural accounts).
        query_parts = [f"from:{sender}", f"after:{after}", f"before:{before}"]
        if must_contain:
            query_parts.append(f'"{must_contain}"')
        if subject_must_contain:
            query_parts.append(f'subject:"{subject_must_contain}"')
        query = " ".join(query_parts)

        try:
            result = service.users().messages().list(
                userId="me", q=query, maxResults=5
            ).execute()
            messages = result.get("messages", [])
        except Exception as e:
            print(f"  ⚠  Gmail search failed for {label}: {e}")
            bills[key] = None
            continue

        if not messages:
            print(f"  ⚠  No email found for {label}  (query: {query})")
            bills[key] = None
            continue

        try:
            msg = service.users().messages().get(
                userId="me", id=messages[0]["id"], format="full"
            ).execute()
            body = _decode_body(msg.get("payload", {}))
            snippet = msg.get("snippet", "")
            hdrs = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
        except Exception as e:
            print(f"  ⚠  Could not fetch {label} email: {e}")
            bills[key] = None
            continue

        if az_client:
            amount = _extract_amount_azure(body, label, az_client, az_deployment,
                                           snippet=snippet, provider=provider)
            method = "Azure" if provider == "azure_openai" else "Claude Foundry"
        else:
            # Try snippet first, then full body, snippet is high-signal
            amount = _extract_amount_regex(snippet, pattern) if snippet else None
            if amount is None:
                amount = _extract_amount_regex(body, pattern)
            method = "regex"

        if amount is None:
            print(f"  ⚠  Could not extract amount from {label} ({method})")
            # Print diagnostic info so we can see what was tried
            print(f"     Subject: {hdrs.get('Subject','?')[:100]}")
            print(f"     Snippet: {snippet[:200]}")
        else:
            print(f"  ✓ {label}: ${amount:.2f}  [{method}]  ({hdrs.get('Subject','?')[:60]})")
        bills[key] = amount

    return bills


def create_gmail_draft(
    service,
    to: str | list[str],
    subject: str,
    body: str,
    attachments: list[Path] = None,
    cc: str = None,
    from_addr: str = None,
) -> str | None:
    """
    Create a draft in the authenticated Gmail account. Returns draft ID.
    """
    msg = EmailMessage()
    msg["To"] = ", ".join(to) if isinstance(to, list) else to
    if cc:
        msg["Cc"] = cc
    if from_addr:
        msg["From"] = from_addr
    msg["Subject"] = subject
    msg.set_content(body)

    for att in attachments or []:
        att = Path(att)
        if not att.exists():
            print(f"  ⚠  Attachment missing: {att}")
            continue
        ctype, _ = mimetypes.guess_type(str(att))
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        with open(att, "rb") as f:
            msg.add_attachment(
                f.read(), maintype=maintype, subtype=subtype, filename=att.name
            )

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    try:
        result = service.users().drafts().create(
            userId="me", body={"message": {"raw": raw}}
        ).execute()
        return result.get("id")
    except Exception as e:
        print(f"  ⚠  Gmail draft creation failed: {e}")
        return None


# ══════════════════════════════════════════════════════════════════════════════
# SYNOPSIS EMAIL TO SELF
# ══════════════════════════════════════════════════════════════════════════════

def build_synopsis_body(
    analysis: dict,
    commentary: str,
    billing_month: date,
    shares: dict,
    config: dict,
) -> str:
    """Format the synopsis email body (plain text)."""
    labels = config["utility_labels"]
    lines = []
    lines.append(f"Utility Billing Synopsis, {billing_month.strftime('%B %Y')}")
    lines.append("=" * 60)
    lines.append("")
    lines.append(f"TENANT OWES THIS MONTH: ${shares['total']:.2f}")
    if analysis["tenant_total_prior"] is not None:
        prior = analysis["tenant_total_prior"]
        delta = shares["total"] - prior
        pct = (delta / prior * 100) if prior else 0
        lines.append(f"  (prior month: ${prior:.2f}  ·  {delta:+.2f} = {pct:+.1f}%)")
    lines.append("")

    # Per-utility table
    lines.append("PER-UTILITY BREAKDOWN")
    lines.append("-" * 60)
    lines.append(f"{'Utility':22s}  {'Raw $':>10s}  {'Prior $':>10s}  {'%Δ':>8s}")
    for u in analysis["per_utility"]:
        label = u["label"]
        cur = u["current"]
        prior = u["prior"]
        pct = f"{u['pct_change']:+.1f}%" if u["pct_change"] is not None else "  new"
        lines.append(f"{label:22s}  {cur:>10.2f}  {prior:>10.2f}  {pct:>8s}")
    lines.append("")

    # Flags
    if analysis["flags"]:
        lines.append("FLAGS")
        lines.append("-" * 60)
        for f in analysis["flags"]:
            lines.append(f"  · [{f['kind']}] {f['detail']}")
        lines.append("")
    else:
        lines.append("No anomalies flagged.")
        lines.append("")

    # AI commentary
    if commentary:
        lines.append("AI ANALYSIS")
        lines.append("-" * 60)
        lines.append(commentary)
        lines.append("")

    # Tenant shares (the actual invoice line items)
    lines.append("TENANT INVOICE LINE ITEMS")
    lines.append("-" * 60)
    for key in config["utility_keys_in_order"]:
        amt = shares["per_utility"].get(key, 0)
        if amt > 0:
            lines.append(f"  {labels.get(key, key):22s}  ${amt:>8.2f}")
    lines.append(f"  {'TOTAL':22s}  ${shares['total']:>8.2f}")
    lines.append("")
    lines.append("Next step: open your Gmail Drafts folder. The tenant invoice")
    lines.append("is already composed, review and hit Send when you're ready.")
    return "\n".join(lines)


def send_synopsis_email(
    service,
    body: str,
    invoice_path: Path,
    billing_month: date,
    config: dict,
    as_draft: bool = False,
) -> bool:
    """
    Send (or draft) the synopsis email to yourself. as_draft=True makes it a
    draft instead of sending it.
    """
    to = config.get("synopsis_recipient") or config.get("gmail_account")
    if not to:
        print("  ⚠  No synopsis_recipient configured.")
        return False

    subject = f"📋 Billing synopsis, {billing_month.strftime('%B %Y')} (REVIEW)"

    msg = EmailMessage()
    msg["To"] = to
    msg["From"] = config.get("gmail_account", to)
    msg["Subject"] = subject
    msg.set_content(body)

    if invoice_path and Path(invoice_path).exists():
        with open(invoice_path, "rb") as f:
            msg.add_attachment(
                f.read(), maintype="application", subtype="pdf",
                filename=Path(invoice_path).name,
            )

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    try:
        if as_draft:
            service.users().drafts().create(
                userId="me", body={"message": {"raw": raw}}
            ).execute()
            print(f"  ✓ Synopsis saved as draft (will not send): {to}")
        else:
            service.users().messages().send(
                userId="me", body={"raw": raw}
            ).execute()
            print(f"  ✓ Synopsis sent to {to}")
        return True
    except Exception as e:
        print(f"  ⚠  Synopsis send failed: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# MANUAL INPUT FALLBACK
# ══════════════════════════════════════════════════════════════════════════════

def prompt_for_bills(config: dict) -> dict:
    print("\n" + "─" * 50)
    print("  Enter raw bill totals (press Enter to skip/zero)")
    print("─" * 50)
    bills = {}
    labels = config["utility_labels"]
    for key in config["utility_keys_in_order"]:
        # Skip pet_waste by default since it's cancelled
        if config.get("gmail_searches", {}).get(key, {}).get("_skip"):
            bills[key] = 0.0
            continue
        while True:
            raw = input(f"  {labels.get(key, key)}: $").strip()
            if raw == "":
                bills[key] = 0.0
                break
            try:
                bills[key] = float(raw.replace("$", "").replace(",", ""))
                break
            except ValueError:
                print("    Please enter a number, e.g. 57.18")
    return bills


def is_interactive() -> bool:
    """
    True only if there is a real human on stdin.

    Under a scheduler (launchd/cron), stdin is /dev/null. Calling input() there
    raises EOFError and kills the run, silently, with no notification. This
    exact failure took down a scheduled run in production. Never prompt unless
    this returns True.
    """
    try:
        return sys.stdin is not None and sys.stdin.isatty()
    except Exception:
        return False


def fill_missing(bills: dict, config: dict, interactive: bool = True) -> tuple[dict, list[str]]:
    """
    Resolve bills that could not be auto-read.

    Interactive  → prompt the human, as before.
    Unattended   → do NOT prompt. Return the list of unresolved keys so the
                   caller can abort cleanly and send an alert instead of
                   crashing or, worse, invoicing the tenant a wrong total.

    Returns (bills, unresolved_keys).
    """
    labels = config["utility_labels"]
    missing = [k for k, v in bills.items() if v is None]
    if not missing:
        return bills, []

    if not interactive:
        print("\n  ⚠  Could not auto-read: "
              + ", ".join(labels.get(k, k) for k in missing))
        print("  ⚠  Unattended run, not prompting. Aborting before any writes.")
        return bills, missing

    print("\nCould not auto-read these bills:")
    for key in missing:
        while True:
            raw = input(f"  {labels.get(key, key)} (enter amount or 0 to skip): $").strip()
            if raw == "":
                bills[key] = 0.0
                break
            try:
                bills[key] = float(raw.replace("$", "").replace(",", ""))
                break
            except ValueError:
                print("    Please enter a number, e.g. 57.18")
    return bills, []


def send_alert_email(config: dict, subject: str, body: str) -> bool:
    """
    Send a plain alert to the operator. Used when an unattended run cannot
    complete. This is the safety net: a failed run must never be silent.
    """
    to = config.get("synopsis_recipient") or config.get("gmail_account")
    if not to:
        print("  ⚠  No synopsis_recipient configured, cannot send alert.")
        return False
    try:
        service = _gmail_service(config)
        if service is None:
            print("  ⚠  Gmail unavailable, cannot send alert.")
            return False
        msg = EmailMessage()
        msg["To"] = to
        msg["From"] = config.get("gmail_account", to)
        msg["Subject"] = subject
        msg.set_content(body)
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        print(f"  ✓ Alert email sent to {to}")
        return True
    except Exception as e:
        print(f"  ⚠  Alert email failed: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# GMAIL OAUTH (one-time)
# ══════════════════════════════════════════════════════════════════════════════

def run_gmail_auth() -> None:
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        sys.exit("Run: pip3 install google-auth-oauthlib")
    if not CREDS_FILE.exists():
        sys.exit(
            f"\ncredentials.json not found at: {CREDS_FILE}\n\n"
            "Setup steps:\n"
            "  1. https://console.cloud.google.com\n"
            "  2. Create project → Enable Gmail API\n"
            "  3. Credentials → Create OAuth 2.0 Client ID (Desktop app)\n"
            "  4. Download JSON → rename to credentials.json\n"
            f"  5. Move to: {SCRIPT_DIR}\n"
            "  6. Re-run: python3 run_billing.py --auth\n"
        )
    SCOPES = [
        "https://www.googleapis.com/auth/gmail.readonly",
        "https://www.googleapis.com/auth/gmail.compose",
        "https://www.googleapis.com/auth/gmail.send",
    ]
    print("Opening browser for Google sign-in...")
    print("Sign in as the Gmail account that receives your utility bills.\n")
    flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_FILE), SCOPES)
    creds = flow.run_local_server(port=0)
    with open(TOKEN_FILE, "wb") as f:
        pickle.dump(creds, f)
    print(f"\n✓ Gmail authenticated. Token saved: {TOKEN_FILE}")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(description="Utility Billing Automation with Azure OpenAI")
    p.add_argument("--month", type=str, help="Billing month YYYY-MM (default: last month)")
    p.add_argument("--manual", action="store_true", help="Skip Gmail, enter amounts manually")
    p.add_argument("--auth", action="store_true", help="One-time Gmail OAuth setup")
    p.add_argument("--dry-run", action="store_true", help="Preview only; no spreadsheet write, no drafts")
    p.add_argument("--no-synopsis", action="store_true", help="Skip synopsis email to self")
    p.add_argument("--no-tenant-draft", action="store_true", help="Skip Gmail draft to tenant")
    p.add_argument("--synopsis-as-draft", action="store_true",
                   help="Save synopsis as a draft instead of sending immediately")
    p.add_argument("--non-interactive", action="store_true",
                   help="Never prompt, even from a terminal. Abort + email on any "
                        "unreadable bill. (Auto-enabled when stdin is not a TTY, "
                        "e.g. under launchd/cron.)")
    args = p.parse_args()

    if args.auth:
        run_gmail_auth()
        return

    config = load_config()

    # Billing month
    if args.month:
        try:
            billing_month = datetime.strptime(args.month, "%Y-%m").date().replace(day=1)
        except ValueError:
            sys.exit("Invalid --month. Use YYYY-MM, e.g. 2026-04")
    else:
        today = date.today()
        billing_month = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
    send_date = date.today()

    print()
    print("═" * 60)
    print(f"  Utility Billing, {config.get('property_billing_address', 'Property')}")
    print(f"  Billing Period : {billing_month.strftime('%B %Y')}")
    print(f"  Send Date      : {send_date.strftime('%m/%d/%Y')}")
    if args.dry_run:    print("  Mode           : DRY RUN")
    print("═" * 60)

    interactive = is_interactive() and not args.non_interactive
    if not interactive:
        print("  Mode           : UNATTENDED (no stdin, will not prompt)")

    # 1. Read bills
    bills = None
    if not args.manual:
        print("\n[1/6] Reading bills from Gmail...")
        bills = read_bills_from_gmail(billing_month, config)

    if bills is None:
        # Gmail itself failed (auth, network, etc.)
        if not interactive:
            detail = (
                f"The {billing_month.strftime('%B %Y')} billing run could not read "
                f"any bills from Gmail.\n\n"
                f"Likely causes: expired OAuth token, or no network at run time.\n\n"
                f"Nothing was written to the spreadsheet. No invoice was generated.\n\n"
                f"Fix: cd to the billing_automation folder and run\n"
                f"    python3 run_billing.py --auth\n"
                f"then re-run:\n"
                f"    python3 run_billing.py --month {billing_month.strftime('%Y-%m')}\n"
            )
            print("\n  ⚠  Gmail returned nothing. Unattended, aborting and alerting.")
            send_alert_email(
                config,
                f"🔴 Billing run FAILED, {billing_month.strftime('%B %Y')} (Gmail unreadable)",
                detail,
            )
            sys.exit(1)
        bills = prompt_for_bills(config)
        unresolved = []
    else:
        bills, unresolved = fill_missing(bills, config, interactive=interactive)

    # 1b. Unattended run with a bill we could not read → stop BEFORE any writes.
    #     Invoicing the tenant a total that is missing a utility is worse than
    #     being a day late, and a half-written spreadsheet row corrupts the
    #     history that the anomaly analysis depends on.
    if unresolved:
        labels = config["utility_labels"]
        found_lines = "\n".join(
            f"    {labels.get(k, k):22s} ${v:>8.2f}"
            for k, v in bills.items() if v is not None and v > 0
        )
        missing_lines = "\n".join(f"    {labels.get(k, k)}" for k in unresolved)
        detail = (
            f"The {billing_month.strftime('%B %Y')} billing run ran on schedule but "
            f"could not read every bill, so it STOPPED before touching the "
            f"spreadsheet or drafting anything to the tenant.\n\n"
            f"Read successfully:\n{found_lines}\n\n"
            f"COULD NOT READ:\n{missing_lines}\n\n"
            f"This usually means a provider changed its sender address or email "
            f"template. Check the Gmail search config for the missing utility in "
            f"billing_config.json.\n\n"
            f"Nothing was written. Nothing was sent. To finish the month manually:\n"
            f"    python3 run_billing.py --month {billing_month.strftime('%Y-%m')}\n"
            f"(run from a terminal, it will prompt you for the missing amount)\n"
        )
        send_alert_email(
            config,
            f"🟠 Billing run STOPPED, {billing_month.strftime('%B %Y')} "
            f"({len(unresolved)} bill(s) unreadable)",
            detail,
        )
        sys.exit(1)

    # 2. Dry run stops here (no writes, no drafts)
    if args.dry_run:
        print("\n[DRY RUN] Bills:")
        for k, v in bills.items():
            print(f"  {config['utility_labels'].get(k, k)}: ${v or 0:.2f}")
        print("\nDry run complete, no spreadsheet writes, no drafts.")
        return

    # 3. Write to spreadsheet
    print("\n[2/6] Writing raw amounts to spreadsheet...")
    write_result = write_bills_to_spreadsheet(bills, billing_month, config)
    if not write_result:
        sys.exit("Could not write to spreadsheet. Aborting.")
    workbook_path, sheet_name, data_row = write_result

    # 4. Compute tenant shares from spreadsheet formulas
    print("\n[3/6] Evaluating spreadsheet formulas for tenant shares...")
    shares = read_tenant_shares(workbook_path, sheet_name, data_row, bills, config)
    print(f"  ✓ Tenant total: ${shares['total']:.2f}")
    for k, v in shares["per_utility"].items():
        if v > 0:
            print(f"    {config['utility_labels'].get(k, k):22s}  ${v:>8.2f}")

    # 5. Analyze vs. history + selected provider commentary
    commentary_name = "Azure" if llm.get_provider(config) == "azure_openai" else "Claude Foundry"
    print(f"\n[4/6] Running anomaly analysis + {commentary_name} commentary...")
    history = read_history(workbook_path, sheet_name, data_row, config,
                           months_back=config.get("analysis", {}).get("history_window_months", 6))
    current = {
        "raw_amounts": {k: bills.get(k, 0) or 0 for k in config["utility_keys_in_order"]},
        "total": shares["total"],
    }
    analysis = analyze_bills(current, history, config)
    if analysis["flags"]:
        print(f"  ⚠  {len(analysis['flags'])} anomaly flag(s):")
        for fl in analysis["flags"]:
            print(f"    · {fl['detail']}")
    else:
        print("  ✓ No anomalies flagged.")
    commentary = azure_commentary(analysis, billing_month, config)
    if commentary:
        print(f"  ✓ {commentary_name} commentary obtained ({len(commentary)} chars)")

    # 6. PDF invoice
    print("\n[5/6] Generating invoice PDF...")
    output_dir = resolve_path(config.get("invoice_output_dir", "~"))
    invoice_path = generate_invoice_pdf(
        shares["per_utility"], shares["total"], billing_month, send_date, config, output_dir
    )
    print(f"  ✓ Saved: {invoice_path}")

    # 7. Gmail drafts (tenant + synopsis)
    service = _gmail_service(config)
    if service is None:
        print("\n  ⚠  Gmail not authenticated, skipping drafts. Run --auth.")
        return

    if not args.no_tenant_draft:
        print("\n[6/6a] Creating Gmail draft to tenant...")
        tenant_email = config.get("tenant_email_to", "")
        if "TENANT_EMAIL_HERE" in tenant_email or "@" not in tenant_email:
            print(f"  ⚠  tenant_email_to not configured ('{tenant_email}'). Skipping tenant draft.")
        else:
            month_long = billing_month.strftime("%B %Y")
            month_name = billing_month.strftime("%B")
            # All copy is config-driven so this script ships as a clean template.
            greeting = config.get("tenant_email_greeting", "Hi,")
            body_tmpl = config.get(
                "tenant_email_body_template",
                "Here are {month}'s utilities. Thank you!"
            )
            signoff_lines = config.get(
                "tenant_email_signoff_lines",
                ["Best,", config.get("landlord_name", "")]
            )
            body = (
                f"{greeting}\n\n"
                f"{body_tmpl.format(month=month_name)}\n\n"
                + "\n".join(signoff_lines)
            )
            prop_addr = config.get("property_billing_address", "Property")
            cc = config.get("tenant_email_cc") or None
            draft_id = create_gmail_draft(
                service,
                to=tenant_email,
                subject=f"{month_long} Utilities Invoice, {prop_addr}",
                body=body,
                attachments=[invoice_path],
                cc=cc,
            )
            if draft_id:
                print(f"  ✓ Tenant draft created (ID: {draft_id[:12]}...). Review + send when ready.")

    if not args.no_synopsis:
        print("\n[6/6b] Sending synopsis email to yourself...")
        synopsis_body = build_synopsis_body(analysis, commentary, billing_month, shares, config)
        send_synopsis_email(
            service, synopsis_body, invoice_path, billing_month, config,
            as_draft=args.synopsis_as_draft,
        )

    print()
    print("═" * 60)
    print(f"  ✓ Invoice generated: {invoice_path.name}")
    print(f"  ✓ Tenant total: ${shares['total']:.2f}")
    print(f"  → Review synopsis in your inbox.")
    print(f"  → Review tenant draft in Gmail → Drafts.")
    print(f"  → Hit send when ready.")
    print("═" * 60)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        # sys.exit() from our own code, already handled + alerted. Pass through.
        raise
    except Exception:
        # Last-resort net. Any unhandled crash gets emailed instead of dying
        # quietly in a log file nobody reads. An EOFError under the scheduler
        # once killed a run and the only symptom was an email that never arrived.
        tb = traceback.format_exc()
        print(tb, file=sys.stderr)
        try:
            cfg = load_config()
            send_alert_email(
                cfg,
                "🔴 Billing run CRASHED, unhandled error",
                "The utility billing script crashed with an unhandled exception.\n"
                "Nothing was sent to the tenant.\n\n"
                f"{tb}\n",
            )
        except Exception:
            pass  # if even the alert fails, the stderr log is all we have
        sys.exit(1)
