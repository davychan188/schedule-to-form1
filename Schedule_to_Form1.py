#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Schedule to Form 1
專為有 Zone / TRMZ 的 Tree Assessment Schedule 設計
（例如 Queen Mary Hospital、Kowloon Hospital）

排除（不計入樹群）：
  - Removed
  - Duplicate / not in use
  - Out of LOT
  - Within construction / Construction area
  - Outside boundary
"""

import os
import re
import sys
import threading
from pathlib import Path
from datetime import datetime, timedelta

def _bootstrap_tcl_tk():
    if not getattr(sys, "frozen", False):
        return
    base = getattr(sys, "_MEIPASS", None)
    if not base:
        return
    for root, dirs, files in os.walk(base):
        if "init.tcl" in files:
            os.environ["TCL_LIBRARY"] = root
            break

_bootstrap_tcl_tk()

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    HAS_TK = True
except Exception:
    HAS_TK = False

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT_FILENAME = "Schedule_to_Form1_Output.xlsx"
WORD_FILENAME = "Schedule_to_Form1_Cover.docx"

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    HAS_DOCX = True
except Exception:
    HAS_DOCX = False

# ── styles ──
title_f = Font(name="Arial", bold=True, size=14)
bold_f = Font(name="Arial", bold=True, size=10)
normal_f = Font(name="Arial", size=10)
italic_f = Font(name="Arial", italic=True, size=9)
header_fill = PatternFill("solid", fgColor="1F4E79")
header_font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
yellow_f = PatternFill("solid", fgColor="FFFF00")
orange_f = PatternFill("solid", fgColor="FFA500")
red_f = PatternFill("solid", fgColor="FF0000")
black_f = PatternFill("solid", fgColor="333333")
thin = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def style_header(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        cell.border = thin


def auto_width(ws, min_w=8, max_w=40):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        length = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[letter].width = min(max(length + 2, min_w), max_w)


def _mode_or(series, default="Fair"):
    """Return most common value; never raise IndexError on empty."""
    try:
        s = series.dropna().astype(str)
        s = s[~s.str.lower().isin(["", "nan", "none", "n.a.", "n.a", "-"])]
        if len(s) == 0:
            return default
        vc = s.value_counts()
        return vc.index[0] if len(vc) else default
    except Exception:
        return default


def overall_cond(row):
    order = {"Normal": 0, "Fair": 1, "Poor": 2, "Very Poor": 3, "Dead": 4}
    h = str(row.get("Health", "")).strip()
    s = str(row.get("Structural", "")).strip()
    scores = [order[x] for x in (h, s) if x in order]
    if not scores:
        return "Fair"
    inv = {v: k for k, v in order.items()}
    return inv[max(scores)]


def _looks_like_tree_no(val) -> bool:
    """自動識別樹編號：不限定前綴／連字號次數。"""
    s = str(val).strip().replace("\u3000", " ").replace("\xa0", " ")
    s = re.sub(r"\s+", "", s)
    sl = s.lower().rstrip(".")
    junk = {
        "nan", "none", "nat", "-", "n.a.", "n.a", "n/a",
        "treeno", "treenumber", "treeid", "item", "sn", "s/n", "no",
    }
    if not s or sl in junk:
        return False
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        return False
    if len(s) > 32:
        return False
    if not re.search(r"\d", s):
        return False
    return bool(re.match(r"^[A-Za-z0-9][A-Za-z0-9\-_./]*[A-Za-z0-9]$", s) or re.match(r"^\d+[A-Za-z]?$", s))


def load_schedule(path: str) -> pd.DataFrame:
    """讀取有 Zone 的 Schedule（.xlsx / .xls）。"""
    path_l = str(path).lower()
    if path_l.endswith(".xls") and not path_l.endswith(".xlsx"):
        try:
            xl = pd.ExcelFile(path, engine="xlrd")
        except ImportError:
            raise ImportError(
                "此檔案是 .xls。請：\n"
                "  pip install xlrd\n"
                "或用 Excel 另存為 .xlsx 再試。"
            )
    else:
        try:
            xl = pd.ExcelFile(path, engine="openpyxl")
        except Exception:
            xl = pd.ExcelFile(path)

    raw = None
    used_sheet = None
    for sh in xl.sheet_names:
        tmp = pd.read_excel(path, sheet_name=sh, header=None)
        for i in range(min(25, len(tmp))):
            row_vals = [
                str(v).strip().lower().replace("\n", " ") if pd.notna(v) else ""
                for v in tmp.iloc[i].tolist()
            ]
            joined = " ".join(row_vals)
            first = row_vals[0] if row_vals else ""
            has_tree_id = any(
                k in joined
                for k in ("tree no", "tree id", "tree number", "dept. tree", "樹木編號", "樹編號")
            )
            has_species = "scientific" in joined or "species" in joined or "樹種" in joined
            # 真正表頭：Tree No / Tree ID + 品種欄，且不要誤匹配 legend
            if (
                has_tree_id
                and "removed trees" not in joined
                and (has_species or "dbh" in joined or "zone" in joined or first in ("trmz", "item", "no.", "s/n", "sn"))
            ):
                raw = tmp
                used_sheet = sh
                header_row = i
                break
            if first == "trmz" and "zone" in joined and has_tree_id:
                raw = tmp
                used_sheet = sh
                header_row = i
                break
            if has_tree_id and has_species:
                raw = tmp
                used_sheet = sh
                header_row = i
                break
        if raw is not None:
            break
    if raw is None:
        raw = pd.read_excel(path, sheet_name=xl.sheet_names[0], header=None)
        header_row = 9
        used_sheet = xl.sheet_names[0]

    # subheader
    subheader_row = None
    data_start = header_row + 1
    if header_row + 1 < len(raw):
        next_vals = [
            str(v).strip().lower().replace("\n", " ") if pd.notna(v) else ""
            for v in raw.iloc[header_row + 1].tolist()
        ]
        if any(("scientific" in v) or ("chinese" in v) or ("dbh" in v) for v in next_vals):
            subheader_row = header_row + 1
            data_start = header_row + 2

    h1 = [
        str(v).strip().lower().replace("\n", " ") if pd.notna(v) else ""
        for v in raw.iloc[header_row].tolist()
    ]
    h2 = (
        [
            str(v).strip().lower().replace("\n", " ") if pd.notna(v) else ""
            for v in raw.iloc[subheader_row].tolist()
        ]
        if subheader_row is not None
        else [""] * len(h1)
    )
    # 上一列常有 Remarks / Mitigation（Parkvale：表頭分兩列，Remarks 在 Tree No. 上一列）
    h_prev = [""] * len(h1)
    if header_row > 0:
        h_prev = [
            str(v).strip().lower().replace("\n", " ") if pd.notna(v) else ""
            for v in raw.iloc[header_row - 1].tolist()
        ]
    n_cols = max(len(h1), len(h2), len(h_prev))
    h1 += [""] * (n_cols - len(h1))
    h2 += [""] * (n_cols - len(h2))
    h_prev += [""] * (n_cols - len(h_prev))
    combined = []
    for i in range(n_cols):
        parts = [h_prev[i], h1[i], h2[i]]
        combined.append(" ".join(p for p in parts if p).strip())

    def _find(*kws, exclude=None):
        exclude = exclude or []
        for idx, t in enumerate(combined):
            if any(ex in t for ex in exclude):
                continue
            if all(kw in t for kw in kws):
                return idx
        for idx, t in enumerate(combined):
            if any(ex in t for ex in exclude):
                continue
            if any(kw in t for kw in kws):
                return idx
        return None

    col_map = {
        "TRMZ": _find("trmz"),
        "Zone": _find("tree group") or _find("zone"),
        "Tree_No": (
            _find("tree no") or _find("tree id") or _find("tree number")
            or _find("dept. tree") or _find("樹木編號") or _find("樹編號")
        ),
        "Scientific_Name": _find("scientific"),
        "Chinese_Name": _find("chinese"),
        "DBH_mm": _find("dbh"),
        "Height_m": _find("height"),
        "Spread_m": _find("spread") or _find("crown"),
        "Health": _find("health"),
        "Structural": _find("structural"),
        "Tree_Form": _find("form condition") or _find("form"),
        "Triage_Color": _find("triage"),
        "Category": _find("category"),
        "Lean": _find("lean"),
        "Remarks": _find("remark", exclude=["remarks*", "suitability"]),
        "Recommendations": _find("mitigation") or _find("proposed") or _find("recommend"),
        "Completion_Date": _find("completion") or _find("anticipated"),
        "Form1_Ref": _find("form 1") or _find("form1"),
    }

    # fallbacks for QMH-style: TRMZ Zone TreeNo Sci Chi DBH Ht Spr Health Struct Form Triage ...
    if col_map["Tree_No"] is None:
        col_map["Tree_No"] = 2 if col_map.get("Zone") is not None else 0
    if col_map["Scientific_Name"] is None:
        col_map["Scientific_Name"] = col_map["Tree_No"] + 1
    if col_map["Chinese_Name"] is None:
        col_map["Chinese_Name"] = col_map["Scientific_Name"] + 1
    if col_map["DBH_mm"] is None:
        col_map["DBH_mm"] = col_map["Chinese_Name"] + 1
    if col_map["Height_m"] is None:
        col_map["Height_m"] = col_map["DBH_mm"] + 1
    if col_map["Spread_m"] is None:
        col_map["Spread_m"] = col_map["Height_m"] + 1
    if col_map["Health"] is None:
        col_map["Health"] = col_map["Spread_m"] + 1
    if col_map["Structural"] is None:
        col_map["Structural"] = col_map["Health"] + 1
    if col_map["Tree_Form"] is None:
        col_map["Tree_Form"] = col_map["Structural"] + 1
    if col_map["Triage_Color"] is None:
        col_map["Triage_Color"] = col_map["Tree_Form"] + 1
    if col_map["Remarks"] is None:
        # after Lean/Category possibly
        col_map["Remarks"] = (col_map.get("Lean") or col_map.get("Category") or col_map["Triage_Color"]) + 1
    if col_map["Recommendations"] is None:
        col_map["Recommendations"] = col_map["Remarks"] + 1
    if col_map["Completion_Date"] is None:
        col_map["Completion_Date"] = col_map["Recommendations"] + 1

    body = raw.iloc[data_start:].copy().reset_index(drop=True)
    body = body.dropna(how="all")

    def _get(key):
        idx = col_map.get(key)
        if idx is None or idx >= body.shape[1]:
            return pd.Series([""] * len(body))
        return body.iloc[:, idx]

    out = pd.DataFrame({
        "TRMZ": _get("TRMZ") if col_map.get("TRMZ") is not None else "",
        "Zone": _get("Zone") if col_map.get("Zone") is not None else "",
        "Tree_No": _get("Tree_No"),
        "Scientific_Name": _get("Scientific_Name"),
        "Chinese_Name": _get("Chinese_Name"),
        "DBH_mm": _get("DBH_mm"),
        "Height_m": _get("Height_m"),
        "Spread_m": _get("Spread_m"),
        "Health": _get("Health"),
        "Structural": _get("Structural"),
        "Tree_Form": _get("Tree_Form"),
        "Triage_Color": _get("Triage_Color"),
        "Category": _get("Category") if col_map.get("Category") is not None else "",
        "Lean": _get("Lean") if col_map.get("Lean") is not None else "",
        "Remarks": _get("Remarks"),
        "Recommendations": _get("Recommendations"),
        "Completion_Date": _get("Completion_Date"),
        "Form1_Ref": _get("Form1_Ref") if col_map.get("Form1_Ref") is not None else "",
    })

    out["Tree_No"] = out["Tree_No"].astype(str).str.strip().str.replace(r"\s+", "", regex=True)
    kept = out[out["Tree_No"].apply(_looks_like_tree_no)].copy()
    if len(kept) == 0:
        # 自動掃描其他欄，找最多「像樹編號」的一欄
        best_idx, best_n = None, 0
        for i in range(min(body.shape[1], 20)):
            n = int(body.iloc[:, i].astype(str).apply(_looks_like_tree_no).sum())
            if n > best_n:
                best_n, best_idx = n, i
        if best_idx is not None and best_n >= 3:
            out["Tree_No"] = body.iloc[:, best_idx].astype(str).str.strip().str.replace(r"\s+", "", regex=True)
            kept = out[out["Tree_No"].apply(_looks_like_tree_no)].copy()
        if len(kept) == 0:
            samples = [str(v) for v in out["Tree_No"].head(12).tolist()]
            raise ValueError(
                "找不到有效樹編號。請確認 Tree No. / Tree ID 欄。\n"
                f"工作表: {used_sheet}　表頭列: {header_row + 1}\n"
                f"目前讀到的欄位樣本: {samples}"
            )
    out = kept

    for col in [
        "TRMZ", "Zone", "Scientific_Name", "Chinese_Name", "Triage_Color",
        "Health", "Structural", "Tree_Form", "Category", "Lean",
        "Remarks", "Recommendations", "Form1_Ref",
    ]:
        out[col] = (
            out[col].astype(str)
            .str.replace(r"[\r\n]+", " ", regex=True)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
            .replace({"nan": "N.A.", "None": "N.A.", "NaT": "N.A."})
        )

    # TPRP / Initial Survey：G/A/P → Normal/Fair/Poor
    gap = {"G": "Normal", "A": "Fair", "P": "Poor", "Good": "Normal", "Average": "Fair"}
    for col in ("Health", "Structural", "Tree_Form"):
        out[col] = out[col].apply(lambda v: gap.get(str(v).strip(), v))

    def _fmt_date(v):
        s = str(v).strip()
        if s in ("-", "", "nan", "N.A.", "n.a.", "None", "NaT"):
            return "-"
        try:
            n = float(s)
            if n > 30000:
                d = datetime(1899, 12, 30) + timedelta(days=n)
                return d.strftime("%d/%m/%Y")
        except Exception:
            pass
        if len(s) >= 10 and s[4] == "-" and s[7] == "-":
            try:
                y, m, d = s[:10].split("-")
                return f"{d}/{m}/{y}"
            except Exception:
                pass
        return s[:10] if len(s) > 12 else s

    out["Completion_Date"] = out["Completion_Date"].astype(str).str.strip().apply(_fmt_date)
    out.attrs["sheet"] = used_sheet
    return out.reset_index(drop=True)


def classify_exclude(row) -> str:
    """回傳排除原因；空字串 = 計入樹群。"""
    remarks = str(row.get("Remarks", "")).strip().lower()
    rec = str(row.get("Recommendations", "")).strip().lower()
    zone = str(row.get("Zone", "")).strip().lower()
    combo = remarks + " " + rec

    # 只有整棵樹 Removed 才排除（Remarks 為 Removed / Removed in 2026 等）
    # 「Main trunk removed」等描述性字句則保留
    if re.search(r"^removed(\s|$|,|;|/)", remarks) or remarks in ("removed",):
        return "Removed"
    if "removed" in remarks and "trunk" not in remarks and "branch" not in remarks:
        return "Removed"
    if any(k in remarks for k in ("duplicate", "not in use", "number not in use")):
        return "Duplicate"
    if "out of lot" in remarks or "out of\u00a0lot" in remarks:
        return "Out of LOT"
    if any(k in combo for k in (
        "construction area", "within construction", "within constrution",
        "within contruction", "construction site", "contruction site",
    )):
        return "Construction"
    if any(k in combo for k in ("outside", "boundary")):
        return "Outside Boundary"
    # blank health + blank size often = removed row without word
    def blank(s):
        return str(s).strip().lower() in ("", "-", "nan", "none", "n.a.", "n.a", "nat")
    if blank(row.get("Scientific_Name", "")) and blank(row.get("Health", "")):
        return "Blank/Invalid"
    if blank(row.get("Scientific_Name", "")) and blank(row.get("DBH_mm", "")):
        return "Blank/Invalid"
    return ""


def generate(df: pd.DataFrame, output_path: str, log_fn=None, dead_as="Unidentified", options=None):
    def log(msg):
        if log_fn:
            log_fn(msg)

    opt = {
        "summary": True,
        "section5": True,
        "section6": True,
        "form1_triage": True,
        "form1_removal": True,
        "form1_copy": True,
        "zone_sheets": True,
        "full": True,
        "word": True,
    }
    if options:
        opt.update(options)

    full = df.copy()

    # Dead tree：可選 Unidentified 或保留 Dead tree（仍計入樹群）
    dead_as = "Dead tree" if str(dead_as).strip().lower().startswith("dead") else "Unidentified"
    chi = "死樹" if dead_as == "Dead tree" else "Unidentified"

    def _is_dead_row(row):
        rem = str(row.get("Remarks", "")).strip().lower()
        sci = str(row.get("Scientific_Name", "")).strip().lower()
        return (
            "dead tree" in rem or rem in ("dead", "dead tree.")
            or "dead tree" in sci or sci == "dead"
        )

    dead_mask = full.apply(_is_dead_row, axis=1)
    if dead_mask.any():
        full.loc[dead_mask, "Scientific_Name"] = dead_as
        full.loc[dead_mask, "Chinese_Name"] = chi
        full.loc[dead_mask, "Health"] = "Dead"

    full["Exclude_Reason"] = full.apply(classify_exclude, axis=1)
    excluded = full[full["Exclude_Reason"] != ""].copy()
    active = full[full["Exclude_Reason"] == ""].copy()

    n_dead = int(dead_mask.sum())
    log(f"讀取樹位: {len(full)}")
    if n_dead:
        log(f"  Dead tree → findings={dead_as}: {n_dead} 棵（仍計入樹群）")
    for reason, g in excluded.groupby("Exclude_Reason"):
        log(f"  排除 [{reason}]: {len(g)}")
    log(f"有效總數（計入樹群）: {len(active)}")

    total = len(active)
    species_df = (
        active.groupby(["Scientific_Name", "Chinese_Name"]).size()
        .reset_index(name="Count").sort_values("Count", ascending=False).reset_index(drop=True)
    )
    species_df.insert(0, "No.", range(1, len(species_df) + 1))
    if len(species_df):
        top = species_df.iloc[0]
        dominant = f"{top['Scientific_Name']} {top['Chinese_Name']} ({int(top['Count'])})"
    else:
        dominant = "N.A."

    def triage_n(color):
        return int((active["Triage_Color"].astype(str).str.strip() == color).sum())

    yellow_n, orange_n, red_n, black_n = map(triage_n, ["Yellow", "Orange", "Red", "Black"])
    yellow_ids = active.loc[active["Triage_Color"].astype(str).str.strip() == "Yellow", "Tree_No"].tolist()
    orange_ids = active.loc[active["Triage_Color"].astype(str).str.strip() == "Orange", "Tree_No"].tolist()
    red_ids = active.loc[active["Triage_Color"].astype(str).str.strip() == "Red", "Tree_No"].tolist()
    black_ids = active.loc[active["Triage_Color"].astype(str).str.strip() == "Black", "Tree_No"].tolist()

    def has_remedial(rec):
        s = str(rec).strip()
        if s in ("", "-", "N.A.", "nan", "None", "No mitigation is required", "No mitigation is requried"):
            return False
        return True

    form1_mask = (
        active["Triage_Color"].astype(str).str.strip().isin(["Yellow", "Orange", "Red", "Black"])
        | active["Recommendations"].apply(has_remedial)
    )
    form1_A = active.loc[form1_mask].copy()
    form1_A["Overall"] = form1_A.apply(overall_cond, axis=1)
    form1_B = active.loc[~form1_mask].copy()
    b_species = (
        form1_B.groupby(["Scientific_Name", "Chinese_Name"])
        .agg(
            Qty=("Tree_No", "count"),
            Height_From=("Height_m", lambda s: pd.to_numeric(s, errors="coerce").min()),
            Height_To=("Height_m", lambda s: pd.to_numeric(s, errors="coerce").max()),
            Overall=("Health", lambda s: _mode_or(s)),
        )
        .reset_index()
    )

    # mitigation counts
    rec = active["Recommendations"].astype(str)
    # Pruning 類：crown 相關 + removal of defective/dead branch + selective pruning + thinning 等
    PRUNE_PAT = (
        r"Crown cleaning|Crown reduction|Crown raising|Crown thinning|Crown restoration|"
        r"Regular crown size control|"
        r"pruning|selective pruning|selective branch|"
        r"Removal of defective branch|Removal of dead (?:trunk|branch)|dead trunk|"
        r"Removal of one of the co-dominant|co-dominant branch|"
        r"epicormic"
    )
    prune_mask = rec.str.contains(PRUNE_PAT, case=False, na=False, regex=True)
    removal_mask = rec.str.contains(r"Tree removal|\bfell(ing)?\b", case=False, na=False, regex=True)
    reinspect_mask = rec.str.contains(r"Re-inspection", case=False, na=False)
    form2_mask = rec.str.contains(r"Form 2", case=False, na=False)
    has_rec_mask = active["Recommendations"].apply(has_remedial)
    others_mask = has_rec_mask & ~prune_mask & ~removal_mask
    both_prune_other = prune_mask & (reinspect_mask | form2_mask)
    pruning_n = int(prune_mask.sum())
    removal_n = int(removal_mask.sum())
    reinspect_n = int(reinspect_mask.sum())
    form2_n = int(form2_mask.sum())
    others_n = int(others_mask.sum())
    rec_trees_n = int(has_rec_mask.sum())
    both_n = int(both_prune_other.sum())

    log(f"Dominant: {dominant}")
    log(f"TRIAGE Yellow:{yellow_n} Orange:{orange_n} Red:{red_n} Black:{black_n}")
    log(f"Form1 A:{len(form1_A)} B species rows:{len(b_species)}")

    wb = Workbook()

    # Dashboard
    ws = wb.active
    ws.title = "Summary_Dashboard"
    ws["A1"] = "Schedule to Form 1 – Summary Dashboard"
    ws["A1"].font = title_f
    rows = [
        ("Total Trees (effective)", total),
        ("Total Species", len(species_df)),
        ("Dominant Species", dominant),
        ("", ""),
        ("Black TRIAGE", black_n),
        ("Red TRIAGE", red_n),
        ("Orange TRIAGE", orange_n),
        ("Yellow TRIAGE", yellow_n),
        ("", ""),
        ("Excluded – Removed", int((excluded["Exclude_Reason"] == "Removed").sum())),
        ("Excluded – Construction", int((excluded["Exclude_Reason"] == "Construction").sum())),
        ("Excluded – Outside Boundary", int((excluded["Exclude_Reason"] == "Outside Boundary").sum())),
        ("Excluded – Duplicate", int((excluded["Exclude_Reason"] == "Duplicate").sum())),
        ("Excluded – Out of LOT", int((excluded["Exclude_Reason"] == "Out of LOT").sum())),
        ("Excluded – Other/Blank", int((excluded["Exclude_Reason"] == "Blank/Invalid").sum())),
        ("", ""),
        ("Trees with any recommendation (unique trees)", rec_trees_n),
        ("  of which pruning (crown cleaning / reduction / etc.)", pruning_n),
        ("  of which others only (not pruning / not removal)", others_n),
        ("  of which pruning AND others on same tree", both_n),
        ("Tree removal / fell", removal_n),
        ("Re-inspection (incl. 3/6/9 months)", reinspect_n),
        ("Form 2 assessment", form2_n),
        ("Note", "Pruning + others on same tree counted once in unique total"),
        ("Form1 Table A (triage/remedial)", len(form1_A)),
        ("Form1 Table B (other trees)", len(form1_B)),
    ]
    for i, (a, b) in enumerate(rows, 3):
        ws.cell(row=i, column=1, value=a).font = bold_f if a else normal_f
        ws.cell(row=i, column=2, value=b).font = normal_f

    r = 3 + len(rows) + 1
    for label, ids, fill in [
        ("Black TRIAGE Tree List", black_ids, black_f),
        ("Red TRIAGE Tree List", red_ids, red_f),
        ("Orange TRIAGE Tree List", orange_ids, orange_f),
        ("Yellow TRIAGE Tree List", yellow_ids, yellow_f),
    ]:
        cell = ws.cell(row=r, column=1, value=label)
        cell.fill = fill
        cell.font = Font(name="Arial", bold=True, size=10, color="FFFFFF" if fill == black_f else "000000")
        ws.cell(row=r + 1, column=1, value=", ".join(ids) if ids else "/").font = normal_f
        r += 3
    auto_width(ws)

    # Zone Summary
    wsZ = wb.create_sheet("Zone_Summary")
    wsZ["A1"] = "Zone Summary（已排除 Removed / Construction / Outside / Duplicate）"
    wsZ["A1"].font = title_f
    z_headers = [
        "Zone", "TRMZ", "No. of Trees",
        "Yellow", "Orange", "Red", "Black", "Triage Total (Y+O+R+B)",
        "Species", "Dominant Species",
    ]
    for i, h in enumerate(z_headers, 1):
        wsZ.cell(row=3, column=i, value=h)
    style_header(wsZ, 3, len(z_headers))
    z_row = 4
    zone_col = active["Zone"].astype(str).str.strip().replace({"nan": "-", "": "-"})
    for zone_name, zdf in active.groupby(zone_col, dropna=False):
        trmz_vals = zdf["TRMZ"].astype(str).str.strip().replace({"nan": "", "N.A.": ""}).unique()
        trmz_str = ", ".join(sorted(v for v in trmz_vals if v and v not in ("-", "")))
        zy = int((zdf["Triage_Color"].astype(str).str.strip() == "Yellow").sum())
        zo = int((zdf["Triage_Color"].astype(str).str.strip() == "Orange").sum())
        zr = int((zdf["Triage_Color"].astype(str).str.strip() == "Red").sum())
        zb = int((zdf["Triage_Color"].astype(str).str.strip() == "Black").sum())
        sp = zdf.groupby(["Scientific_Name", "Chinese_Name"]).size()
        dom = ""
        if len(sp):
            t = sp.sort_values(ascending=False)
            dom = f"{t.index[0][0]} {t.index[0][1]} ({int(t.iloc[0])})"
        vals = [zone_name, trmz_str, len(zdf), zy, zo, zr, zb, zy + zo + zr + zb, int(sp.shape[0]), dom]
        for c, v in enumerate(vals, 1):
            cell = wsZ.cell(row=z_row, column=c, value=v)
            cell.font = normal_f
            cell.border = thin
        z_row += 1
    triage_all = yellow_n + orange_n + red_n + black_n
    wsZ.cell(row=z_row, column=1, value="TOTAL").font = bold_f
    wsZ.cell(row=z_row, column=3, value=total).font = bold_f
    wsZ.cell(row=z_row, column=4, value=yellow_n).font = bold_f
    wsZ.cell(row=z_row, column=5, value=orange_n).font = bold_f
    wsZ.cell(row=z_row, column=6, value=red_n).font = bold_f
    wsZ.cell(row=z_row, column=7, value=black_n).font = bold_f
    wsZ.cell(row=z_row, column=8, value=triage_all).font = bold_f
    z_row += 2
    wsZ.cell(row=z_row, column=1, value="TRIAGE colour totals (all zones)").font = bold_f
    z_row += 1
    for i, h in enumerate(["Colour", "Number of trees"], 1):
        wsZ.cell(row=z_row, column=i, value=h)
    style_header(wsZ, z_row, 2)
    z_row += 1
    for color, cnt, fill in [
        ("Yellow", yellow_n, yellow_f),
        ("Orange", orange_n, orange_f),
        ("Red", red_n, red_f),
        ("Black", black_n, black_f),
        ("Triage Total", triage_all, None),
    ]:
        c1 = wsZ.cell(row=z_row, column=1, value=color)
        c2 = wsZ.cell(row=z_row, column=2, value=cnt)
        c1.font = bold_f if color == "Triage Total" else normal_f
        c2.font = bold_f if color == "Triage Total" else normal_f
        c1.border = thin
        c2.border = thin
        if fill:
            c1.fill = fill
            if color == "Black":
                c1.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        z_row += 1
    auto_width(wsZ)

    # 每個 Zone 各自的 Table A / Table B
    def _write_zone_ab(zname, zdf):
        safe = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in str(zname))[:28]
        z_mask = (
            zdf["Triage_Color"].astype(str).str.strip().isin(["Yellow", "Orange", "Red", "Black"])
            | zdf["Recommendations"].apply(has_remedial)
        )
        za = zdf.loc[z_mask].copy()
        za["Overall"] = za.apply(overall_cond, axis=1)
        zb = zdf.loc[~z_mask].copy()
        zb_sp = (
            zb.groupby(["Scientific_Name", "Chinese_Name"])
            .agg(
                Qty=("Tree_No", "count"),
                Height_From=("Height_m", lambda s: pd.to_numeric(s, errors="coerce").min()),
                Height_To=("Height_m", lambda s: pd.to_numeric(s, errors="coerce").max()),
                Overall=("Health", lambda s: _mode_or(s)),
            )
            .reset_index()
        )

        ws = wb.create_sheet(f"Z_{safe}"[:31])
        ws["A1"] = f"Zone {zname} – Form 1 Table (A) & Table (B)"
        ws["A1"].font = title_f
        ws["A2"] = f"Zone trees: {len(zdf)} | Table A: {len(za)} | Table B other trees: {len(zb)}"
        ws["A2"].font = normal_f

        # ── Table A ──
        ws["A4"] = "Table (A) Triage Trees & Trees Require Remedial Actions"
        ws["A4"].font = bold_f
        headers = [
            "Dept. Tree ID", "Zone", "Tree Species", "DBH (mm)", "Height (m)", "Spread (m)",
            "Tree Status", "Overall", "Triage", "Remedial Action", "Completion Date",
        ]
        for i, h in enumerate(headers, 1):
            ws.cell(row=5, column=i, value=h)
        style_header(ws, 5, len(headers))
        omap = {"Poor": "Poor 差", "Very Poor": "Very Poor 極差", "Fair": "Fair 一般", "Normal": "Normal 正常"}
        tmap = {"Yellow": "Yellow 黃", "Orange": "Orange 橙", "Red": "Red 紅", "Black": "Black 黑"}
        for i, (_, row) in enumerate(za.iterrows()):
            r = i + 6
            tc = str(row["Triage_Color"]).strip()
            try:
                dbh = float(pd.to_numeric(row["DBH_mm"], errors="coerce") or 0)
                ht = float(pd.to_numeric(row["Height_m"], errors="coerce") or 0)
            except Exception:
                dbh, ht = 0, 0
            status = "Large Tree 大樹" if (dbh >= 500 or ht >= 9) else "Other Trees 其他樹木"
            action = str(row["Recommendations"]).strip()
            if action in ("", "-", "N.A.", "nan"):
                action = "No mitigation is required"
            date = row["Completion_Date"] if str(row["Completion_Date"]) not in ("-", "nan", "") else ""
            vals = [
                row["Tree_No"], zname, f"{row['Scientific_Name']}{row['Chinese_Name']}",
                row["DBH_mm"], row["Height_m"], row["Spread_m"], status,
                omap.get(row["Overall"], row["Overall"]), tmap.get(tc, "No 無"), action, date,
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = normal_f
                cell.border = thin
            if tc == "Yellow":
                ws.cell(row=r, column=9).fill = yellow_f
            elif tc == "Orange":
                ws.cell(row=r, column=9).fill = orange_f
            elif tc == "Red":
                ws.cell(row=r, column=9).fill = red_f
            elif tc == "Black":
                ws.cell(row=r, column=9).fill = black_f
        a_end = 6 + len(za)
        ws.cell(row=a_end, column=1, value=f"Sub-total Table (A): {len(za)}").font = bold_f

        # ── Table B (same sheet, below A) ──
        b_title = a_end + 2
        ws.cell(row=b_title, column=1, value="Table (B) Other Trees").font = bold_f
        b_header = b_title + 1
        for i, h in enumerate(["Tree Species", "Qty", "Height From (m)", "To (m)", "Overall Conditions"], 1):
            ws.cell(row=b_header, column=i, value=h)
        style_header(ws, b_header, 5)
        omap2 = {"Normal": "Normal 正常", "Fair": "Fair 一般", "Poor": "Poor 差", "Very Poor": "Very Poor 極差"}
        for i, (_, row) in enumerate(zb_sp.iterrows()):
            r = b_header + 1 + i
            h_from = round(row["Height_From"], 1) if pd.notna(row["Height_From"]) else "-"
            h_to = round(row["Height_To"], 1) if pd.notna(row["Height_To"]) else "-"
            for c, v in enumerate([
                f"{row['Scientific_Name']}{row['Chinese_Name']}", int(row["Qty"]),
                h_from, h_to, omap2.get(str(row["Overall"]), str(row["Overall"])),
            ], 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = normal_f
                cell.border = thin
        b_end = b_header + 1 + len(zb_sp)
        ws.cell(row=b_end, column=1,
                value=f"Sub-total Table (B): {len(zb)} trees | {len(zb_sp)} species").font = bold_f
        auto_width(ws)

    for zone_name, zdf in active.groupby(zone_col, dropna=False):
        trmz_vals = zdf["TRMZ"].astype(str).str.strip().replace({"nan": "", "N.A.": ""}).unique()
        trmz_str = ", ".join(sorted(v for v in trmz_vals if v and v not in ("-", "")))
        zy = int((zdf["Triage_Color"].astype(str).str.strip() == "Yellow").sum())
        zo = int((zdf["Triage_Color"].astype(str).str.strip() == "Orange").sum())
        zr = int((zdf["Triage_Color"].astype(str).str.strip() == "Red").sum())
        zb = int((zdf["Triage_Color"].astype(str).str.strip() == "Black").sum())
        sp = zdf.groupby(["Scientific_Name", "Chinese_Name"]).size()
        dom = ""
        if len(sp):
            t = sp.sort_values(ascending=False)
            dom = f"{t.index[0][0]} {t.index[0][1]} ({int(t.iloc[0])})"
        for c, v in enumerate([zone_name, trmz_str, len(zdf), zy, zo, zr, zb, int(sp.shape[0]), dom], 1):
            cell = wsZ.cell(row=z_row, column=c, value=v)
            cell.font = normal_f
            cell.border = thin
        z_row += 1
    wsZ.cell(row=z_row, column=1, value="TOTAL").font = bold_f
    wsZ.cell(row=z_row, column=3, value=total).font = bold_f
    wsZ.cell(row=z_row, column=4, value=yellow_n).font = bold_f
    wsZ.cell(row=z_row, column=5, value=orange_n).font = bold_f
    wsZ.cell(row=z_row, column=6, value=red_n).font = bold_f
    wsZ.cell(row=z_row, column=7, value=black_n).font = bold_f
    auto_width(wsZ)

    # 每個 Zone 各自的 Table A / Table B
    for zone_name, zdf in active.groupby(zone_col, dropna=False):
        if str(zone_name).strip() in ("", "-", "nan", "N.A."):
            continue
        _write_zone_ab(zone_name, zdf)
    log(f"已為各 Zone 產生 Table A / Table B")

    # Section 5 – 全站（所有有效樹）品種 + TRIAGE 摘要
    ws5 = wb.create_sheet("Section5_Survey_Findings")
    ws5["A1"] = "Section 5 - Tree Survey Findings (All Trees)"
    ws5["A1"].font = title_f
    ws5["A3"] = (
        f"A total of {total} trees and {len(species_df)} species were recorded within the site "
        f"(excluding Removed / Construction / Outside Boundary / Duplicate / Out of LOT). "
        f"The dominant species is {dominant}."
    )
    ws5["A3"].font = normal_f
    ws5["A5"] = "Tree Species Summary"
    ws5["A5"].font = bold_f
    for i, h in enumerate(["No.", "Scientific Name", "Chinese Name", "Number of trees"], 1):
        ws5.cell(row=6, column=i, value=h)
    style_header(ws5, 6, 4)
    for i, row in species_df.sort_values("Scientific_Name").reset_index(drop=True).iterrows():
        for c, v in enumerate([i + 1, row["Scientific_Name"], row["Chinese_Name"], int(row["Count"])], 1):
            cell = ws5.cell(row=7 + i, column=c, value=v)
            cell.font = Font(name="Arial", size=10, italic=True) if c == 2 else normal_f
            cell.border = thin
    # TRIAGE colour summary（全站）
    tri_start = 7 + len(species_df) + 2
    ws5.cell(row=tri_start, column=1, value="TRIAGE colour summary (All Trees)").font = bold_f
    for i, h in enumerate(["Colour", "Tree no.", "Number of trees"], 1):
        ws5.cell(row=tri_start + 1, column=i, value=h)
    style_header(ws5, tri_start + 1, 3)
    for idx, (color, ids, cnt, fill) in enumerate([
        ("Black", black_ids, black_n, black_f),
        ("Red", red_ids, red_n, red_f),
        ("Orange", orange_ids, orange_n, orange_f),
        ("Yellow", yellow_ids, yellow_n, yellow_f),
    ]):
        r = tri_start + 2 + idx
        cell = ws5.cell(row=r, column=1, value=color)
        cell.border = thin
        if fill:
            cell.fill = fill
            if color == "Black":
                cell.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        ws5.cell(row=r, column=2, value=", ".join(ids) if ids else "/").border = thin
        ws5.cell(row=r, column=2).alignment = Alignment(wrap_text=True)
        ws5.cell(row=r, column=3, value=cnt).border = thin
    auto_width(ws5)
    ws5.column_dimensions["B"].width = 55

    # Section 6 – 全站緩減措施（所有有效樹）
    ws6 = wb.create_sheet("Section6_Risk_Findings")
    ws6["A1"] = "Section 6 - Tree Risk Assessment Findings (All Trees)"
    ws6["A1"].font = title_f
    action_total = pruning_n + removal_n + reinspect_n + form2_n
    ws6["A3"] = (
        f"The general health and structural condition of the trees on the site are fair. "
        f"A total of {len(form1_A)} trees were recommended for action or triage monitoring "
        f"(see Table 1). Counts are based on all effective trees across all zones."
    )
    ws6["A3"].font = normal_f
    ws6["A5"] = "Table 1. Summary of proposed mitigation measures"
    ws6["A5"].font = bold_f
    for i, h in enumerate(["Recommendations", "No. of trees"], 1):
        ws6.cell(row=6, column=i, value=h)
    style_header(ws6, 6, 2)
    summary_rows = [
        ("Tree pruning (incl. defective branch / selective / thinning etc.)", pruning_n),
        ("Tree removal", removal_n),
        ("Re-inspection", reinspect_n),
        ("Form 2 assessment", form2_n),
        ("Yellow TRIAGE", yellow_n),
        ("Orange TRIAGE", orange_n),
        ("Red TRIAGE", red_n),
        ("Black TRIAGE", black_n),
    ]
    for i, (lab, val) in enumerate(summary_rows):
        ws6.cell(row=7 + i, column=1, value=lab).border = thin
        ws6.cell(row=7 + i, column=1).font = normal_f
        cell = ws6.cell(row=7 + i, column=2, value=val)
        cell.border = thin
        cell.alignment = Alignment(horizontal="center")
        cell.font = normal_f

    # Detailed lists
    rec = active["Recommendations"].astype(str)
    cleaning = active[rec.str.contains(r"Crown cleaning", case=False, na=False)]["Tree_No"].tolist()
    reduction = active[rec.str.contains(r"Crown reduction|Regular crown size control", case=False, na=False)]["Tree_No"].tolist()
    raising = active[rec.str.contains(r"Crown raising", case=False, na=False)]["Tree_No"].tolist()
    thinning = active[rec.str.contains(r"Crown thinning|Crown restoration", case=False, na=False)]["Tree_No"].tolist()
    selective = active[rec.str.contains(r"selective pruning|selective branch|epicormic", case=False, na=False)]["Tree_No"].tolist()
    defective = active[rec.str.contains(
        r"Removal of defective branch|Removal of dead (?:trunk|branch)|dead trunk|"
        r"Removal of one of the co-dominant|co-dominant branch",
        case=False, na=False, regex=True
    )]["Tree_No"].tolist()
    removals = active[rec.str.contains(r"Tree removal", case=False, na=False)]["Tree_No"].tolist()
    form2 = active[rec.str.contains(r"Form 2", case=False, na=False)]["Tree_No"].tolist()
    reinspect = active[rec.str.contains(r"Re-inspection", case=False, na=False)]
    re3 = reinspect[reinspect["Recommendations"].str.contains("3 months", case=False, na=False)]["Tree_No"].tolist()
    re6 = reinspect[reinspect["Recommendations"].str.contains("6 months", case=False, na=False)]["Tree_No"].tolist()
    re9 = reinspect[reinspect["Recommendations"].str.contains("9 months", case=False, na=False)]["Tree_No"].tolist()
    both = set(cleaning) & set(reduction)

    row = 7 + len(summary_rows) + 2
    ws6.cell(row=row, column=1, value="Table 2. Detailed mitigation measures").font = bold_f
    row += 2

    def _write_list(title, tree_ids, mark_both=False):
        nonlocal row
        ws6.cell(row=row, column=1, value=title).font = bold_f
        row += 1
        if tree_ids:
            text_ids = ", ".join(
                [t + ("*" if mark_both and t in both else "") for t in tree_ids]
            )
        else:
            text_ids = "/"
        ws6.cell(row=row, column=1, value=text_ids).alignment = Alignment(wrap_text=True)
        ws6.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        ws6.row_dimensions[row].height = max(18, 14 * (1 + len(text_ids) // 80))
        row += 2

    _write_list("Crown cleaning:", cleaning, mark_both=True)
    _write_list("Crown reduction / Regular crown size control:", reduction, mark_both=True)
    if raising:
        _write_list("Crown raising:", raising)
    if thinning:
        _write_list("Crown thinning / restoration:", thinning)
    if selective:
        _write_list("Selective pruning / epicormics:", selective)
    if defective:
        _write_list("Removal of defective / dead branch / co-dominant branch:", defective)
    ws6.cell(row=row, column=1, value="* The tree requires both treatments.").font = Font(
        name="Arial", italic=True, size=9
    )
    row += 2
    _write_list("Tree Removal:", removals)
    _write_list("Form 2 assessment:", form2)
    _write_list("Re-inspection after 3 months:", re3)
    _write_list("Re-inspection after 6 months:", re6)
    _write_list("Re-inspection after 9 months:", re9)
    _write_list("Yellow TRIAGE:", yellow_ids)
    _write_list("Orange TRIAGE:", orange_ids)
    _write_list("Red TRIAGE:", red_ids)
    _write_list("Black TRIAGE:", black_ids)
    auto_width(ws6)
    ws6.column_dimensions["A"].width = 90

    # Form1 A
    wsA = wb.create_sheet("Form1_TableA_Remedial")
    wsA["A1"] = "Form 1 - Table (A) Triage Trees & Trees Require Remedial Actions"
    wsA["A1"].font = title_f
    headers_A = [
        "Dept. Tree ID", "Zone", "Tree Species (Scientific + Chinese)", "DBH (mm)",
        "Est. Height (m)", "Est. Crown Spread (m)", "Tree Status", "Overall Tree Conditions",
        "Triage Colour", "Remedial Action", "Anticipated Completion Date",
    ]
    for i, h in enumerate(headers_A, 1):
        wsA.cell(row=4, column=i, value=h)
    style_header(wsA, 4, len(headers_A))
    omap = {"Poor": "Poor 差", "Very Poor": "Very Poor 極差", "Fair": "Fair 一般", "Normal": "Normal 正常"}
    triage_map = {
        "Yellow": "Yellow 黃", "Orange": "Orange 橙", "Red": "Red 紅", "Black": "Black 黑",
    }
    for i, (_, row) in enumerate(form1_A.iterrows()):
        r = i + 5
        tn = row["Tree_No"]
        species = f"{row['Scientific_Name']}{row['Chinese_Name']}"
        tc = str(row["Triage_Color"]).strip()
        triage = triage_map.get(tc, "No 無")
        action = str(row["Recommendations"]).strip()
        if action in ("", "-", "N.A.", "nan"):
            action = "No mitigation is required"
        try:
            dbh = float(pd.to_numeric(row["DBH_mm"], errors="coerce") or 0)
            ht = float(pd.to_numeric(row["Height_m"], errors="coerce") or 0)
        except Exception:
            dbh, ht = 0, 0
        status = "Large Tree 大樹" if (dbh >= 500 or ht >= 9) else "Other Trees 其他樹木"
        date = row["Completion_Date"] if str(row["Completion_Date"]) not in ("-", "nan", "") else ""
        zone = str(row.get("Zone", "")).strip()
        if zone in ("N.A.", "nan", "-"):
            zone = ""
        vals = [
            tn, zone, species, row["DBH_mm"], row["Height_m"], row["Spread_m"],
            status, omap.get(row["Overall"], row["Overall"]), triage, action, date,
        ]
        for c, v in enumerate(vals, 1):
            cell = wsA.cell(row=r, column=c, value=v)
            cell.font = normal_f
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        if tc == "Yellow":
            wsA.cell(row=r, column=9).fill = yellow_f
        elif tc == "Orange":
            wsA.cell(row=r, column=9).fill = orange_f
        elif tc == "Red":
            wsA.cell(row=r, column=9).fill = red_f
        elif tc == "Black":
            wsA.cell(row=r, column=9).fill = black_f
            wsA.cell(row=r, column=9).font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    wsA.cell(row=5 + len(form1_A) + 1, column=1,
             value=f"Sub-total No. of Trees in Table (A): {len(form1_A)}").font = bold_f
    auto_width(wsA)

    # Form1 B
    wsB = wb.create_sheet("Form1_TableB_OtherTrees")
    wsB["A1"] = "Form 1 - Table (B) Other Trees"
    wsB["A1"].font = title_f
    for i, h in enumerate(["Tree Species 樹種", "App. Quantity of Trees", "Height From (m)", "To (m)", "Overall Conditions"], 1):
        wsB.cell(row=4, column=i, value=h)
    style_header(wsB, 4, 5)
    omap2 = {"Normal": "Normal 正常", "Fair": "Fair 一般", "Poor": "Poor 差", "Very Poor": "Very Poor 極差"}
    for i, (_, row) in enumerate(b_species.iterrows()):
        r = i + 5
        species = f"{row['Scientific_Name']}{row['Chinese_Name']}"
        h_from = round(row["Height_From"], 1) if pd.notna(row["Height_From"]) else "-"
        h_to = round(row["Height_To"], 1) if pd.notna(row["Height_To"]) else "-"
        for c, v in enumerate([species, int(row["Qty"]), h_from, h_to, omap2.get(str(row["Overall"]), str(row["Overall"]))], 1):
            cell = wsB.cell(row=r, column=c, value=v)
            cell.font = normal_f
            cell.border = thin
    auto_width(wsB)

    # Form 1（Removal / Fell）— 欄位對齊官方 Form_1.xlsm 的 Sheet1 / Sheet2，方便整表複製
    def _is_felling(rec):
        s = str(rec).strip().lower()
        if re.search(r"\bfell(ing)?\b", s):
            return True
        if "tree removal" in s:
            return True
        if re.search(r"^removal\s*$", s):
            return True
        return False

    fell_mask = active["Recommendations"].apply(_is_felling)
    fell_df = active.loc[fell_mask].copy()
    other_df = active.loc[~fell_mask].copy()
    fell_df["Overall"] = fell_df.apply(overall_cond, axis=1)

    def _sheet1_overall(row):
        remarks = str(row.get("Remarks", "")).strip().lower()
        if "dead tree" in remarks or str(row.get("Health", "")).strip().lower() == "dead":
            return "Dead"
        return row["Overall"]

    fell_df["Sheet1_Overall"] = fell_df.apply(_sheet1_overall, axis=1)

    other_sp = (
        other_df.assign(_ov=other_df.apply(overall_cond, axis=1))
        .groupby(["Scientific_Name", "Chinese_Name"])
        .agg(
            Qty=("Tree_No", "count"),
            Height_From=("Height_m", lambda s: pd.to_numeric(s, errors="coerce").min()),
            Height_To=("Height_m", lambda s: pd.to_numeric(s, errors="coerce").max()),
            Overall=("_ov", lambda s: _mode_or(s)),
        )
        .reset_index()
    )

    # ---- 官方 Sheet1 格式（直接 copy 到 Form_1.xlsm!Sheet1）----
    wsS1 = wb.create_sheet("Form1_Copy_Sheet1")
    wsS1["A1"] = "Copy this sheet into official Form_1.xlsm → Sheet1 (from A1, no header)"
    wsS1["A1"].font = italic_f
    # 官方 Sheet1 無表頭，由 A1 開始：TreeNo, Sci, Chi, DBH, Ht, Spread, Health, Structural, Form, Triage, Remarks, Rec, Date
    s1_headers = [
        "Dept Tree ID", "Scientific Name", "Chinese Name", "DBH (mm)", "Height (m)", "Spread (m)",
        "Health", "Structural", "Tree Form", "Triage Colour", "Remarks", "Recommendations", "Completion Date",
    ]
    for i, h in enumerate(s1_headers, 1):
        wsS1.cell(row=2, column=i, value=h)
    style_header(wsS1, 2, len(s1_headers))
    for i, (_, row) in enumerate(fell_df.iterrows()):
        r = i + 3
        tc = str(row["Triage_Color"]).strip()
        if tc.lower() in ("n.a.", "n.a", "nan", "-", ""):
            tc = "N.A."
        rec = str(row["Recommendations"]).strip()
        date = row["Completion_Date"]
        if str(date) in ("-", "nan", "", "None"):
            date = ""
        vals = [
            row["Tree_No"],
            row["Scientific_Name"],
            row["Chinese_Name"],
            row["DBH_mm"],
            row["Height_m"],
            row["Spread_m"],
            row["Health"] if str(row["Health"]) not in ("-", "nan") else row["Sheet1_Overall"],
            row["Structural"] if str(row["Structural"]) not in ("-", "nan") else "",
            row["Tree_Form"] if str(row["Tree_Form"]) not in ("-", "nan") else "",
            tc,
            row["Remarks"],
            rec,
            date,
        ]
        for c, v in enumerate(vals, 1):
            cell = wsS1.cell(row=r, column=c, value=v)
            cell.font = normal_f
            cell.border = thin
        if tc == "Yellow":
            wsS1.cell(row=r, column=10).fill = yellow_f
        elif tc == "Orange":
            wsS1.cell(row=r, column=10).fill = orange_f
        elif tc == "Red":
            wsS1.cell(row=r, column=10).fill = red_f
        elif tc == "Black":
            wsS1.cell(row=r, column=10).fill = black_f
    wsS1.cell(row=3 + len(fell_df) + 1, column=1,
              value=f"Paste data rows A3:M{2+len(fell_df)} into Form_1.xlsm Sheet1!A1 (without this header row)").font = italic_f
    auto_width(wsS1)

    # ---- 官方 Sheet2 格式（直接 copy 到 Form_1.xlsm!Sheet2）----
    wsS2 = wb.create_sheet("Form1_Copy_Sheet2")
    wsS2["A1"] = "Copy this sheet into official Form_1.xlsm → Sheet2 (from A1, no header)"
    wsS2["A1"].font = italic_f
    for i, h in enumerate(["Tree Species 樹種", "Qty", "Height From (m)", "Height To (m)", "Overall Conditions"], 1):
        wsS2.cell(row=2, column=i, value=h)
    style_header(wsS2, 2, 5)
    omap2 = {"Normal": "Normal", "Fair": "Fair", "Poor": "Poor", "Very Poor": "Very Poor", "Dead": "Dead"}
    for i, (_, row) in enumerate(other_sp.iterrows()):
        r = i + 3
        species = f"{row['Scientific_Name']}{row['Chinese_Name']}".replace("\n", "")
        h_from = round(row["Height_From"], 1) if pd.notna(row["Height_From"]) else ""
        h_to = round(row["Height_To"], 1) if pd.notna(row["Height_To"]) else ""
        ov = omap2.get(str(row["Overall"]), str(row["Overall"]))
        for c, v in enumerate([species, int(row["Qty"]), h_from, h_to, ov], 1):
            cell = wsS2.cell(row=r, column=c, value=v)
            cell.font = normal_f
            cell.border = thin
    wsS2.cell(row=3 + len(other_sp) + 1, column=1,
              value=f"Paste data rows A3:E{2+len(other_sp)} into Form_1.xlsm Sheet2!A1 (without this header row)").font = italic_f
    auto_width(wsS2)

    # ---- 預覽頁（中英對照，方便核對，不必複製這頁）----
    wsR = wb.create_sheet("Form1_Removal_Fell")
    wsR["A1"] = "Form 1 preview — Table (A) Removal/Fell   |   Table (B) Other Trees"
    wsR["A1"].font = title_f
    wsR["A2"] = (
        f"Table A: {len(fell_df)} trees (copy from Form1_Copy_Sheet1)    "
        f"Table B: {len(other_df)} trees / {len(other_sp)} species (copy from Form1_Copy_Sheet2)"
    )
    wsR["A2"].font = normal_f
    wsR["A4"] = "(A) Tree Removal / Felling  — same columns as official Form 1 Table A"
    wsR["A4"].font = bold_f
    headers_R = [
        "Dept. Tree ID 部門樹木編號", "Tree Species 樹種", "DBH (mm) 胸徑",
        "Height (m) 樹高", "Spread (m) 樹冠", "Tree Status 樹木類別",
        "Overall Conditions 整體狀況", "Triage Colour 分流顏色",
        "Remedial Action 緩減措施", "Completion Date 預計完成日期",
    ]
    for i, h in enumerate(headers_R, 1):
        wsR.cell(row=5, column=i, value=h)
    style_header(wsR, 5, len(headers_R))
    omap = {
        "Poor": "Poor 差", "Very Poor": "Very Poor 十分差", "Fair": "Fair 一般",
        "Normal": "Normal 正常", "Dead": "Dead Tree 死樹",
    }
    tmap = {"Yellow": "Yellow 黃", "Orange": "Orange 橙", "Red": "Red 紅", "Black": "Black 黑", "N.A.": "No 無"}
    for i, (_, row) in enumerate(fell_df.iterrows()):
        r = i + 6
        tc = str(row["Triage_Color"]).strip()
        if tc.lower() in ("n.a.", "n.a", "nan", "-", ""):
            tc_key = "N.A."
        else:
            tc_key = tc
        try:
            dbh = float(pd.to_numeric(row["DBH_mm"], errors="coerce") or 0)
            ht = float(pd.to_numeric(row["Height_m"], errors="coerce") or 0)
        except Exception:
            dbh, ht = 0, 0
        status = "Large Tree 大樹" if (dbh >= 500 or ht >= 9) else "Other Trees 其他樹木"
        ov = row["Sheet1_Overall"]
        action = str(row["Recommendations"]).strip()
        if _is_felling(action):
            action = "Removal 移除"
        date = row["Completion_Date"] if str(row["Completion_Date"]) not in ("-", "nan", "") else ""
        vals = [
            row["Tree_No"],
            f"{row['Scientific_Name']}{row['Chinese_Name']}".replace("\n", ""),
            row["DBH_mm"], row["Height_m"], row["Spread_m"], status,
            omap.get(ov, ov), tmap.get(tc_key, "No 無"), action, date,
        ]
        for c, v in enumerate(vals, 1):
            cell = wsR.cell(row=r, column=c, value=v)
            cell.font = normal_f
            cell.border = thin
        if tc_key == "Yellow":
            wsR.cell(row=r, column=8).fill = yellow_f
        elif tc_key == "Orange":
            wsR.cell(row=r, column=8).fill = orange_f
        elif tc_key == "Red":
            wsR.cell(row=r, column=8).fill = red_f
        elif tc_key == "Black":
            wsR.cell(row=r, column=8).fill = black_f
    a_end = 6 + len(fell_df)
    wsR.cell(row=a_end, column=1, value=f"Sub-total Table (A): {len(fell_df)}").font = bold_f

    b_title = a_end + 2
    wsR.cell(row=b_title, column=1, value="(B) Other Trees — same columns as official Form 1 Table B").font = bold_f
    b_header = b_title + 1
    for i, h in enumerate(["Tree Species 樹種", "App. Quantity 數量", "Height From (m) 由", "To (m) 至", "Overall Conditions 整體狀況"], 1):
        wsR.cell(row=b_header, column=i, value=h)
    style_header(wsR, b_header, 5)
    omap_b = {"Normal": "Normal 正常", "Fair": "Fair 一般", "Poor": "Poor 差", "Very Poor": "Very Poor 十分差", "Dead": "Dead Tree 死樹"}
    for i, (_, row) in enumerate(other_sp.iterrows()):
        r = b_header + 1 + i
        h_from = round(row["Height_From"], 1) if pd.notna(row["Height_From"]) else ""
        h_to = round(row["Height_To"], 1) if pd.notna(row["Height_To"]) else ""
        for c, v in enumerate([
            f"{row['Scientific_Name']}{row['Chinese_Name']}".replace("\n", ""),
            int(row["Qty"]), h_from, h_to,
            omap_b.get(str(row["Overall"]), str(row["Overall"])),
        ], 1):
            cell = wsR.cell(row=r, column=c, value=v)
            cell.font = normal_f
            cell.border = thin
    b_end = b_header + 1 + len(other_sp)
    wsR.cell(row=b_end, column=1,
             value=f"Sub-total Table (B): {len(other_df)} trees | {len(other_sp)} species").font = bold_f
    auto_width(wsR)
    log(f"Form1 Removal/Fell Table A: {len(fell_df)} | Table B: {len(other_df)}")

    # 不輸出詳細 excluded 清單（只在 Dashboard 顯示分類合計）

    # Full processed
    wsF = wb.create_sheet("Full_Schedule_Processed")
    wsF["A1"] = "Full Schedule (with Exclude_Reason)"
    wsF["A1"].font = title_f
    cols = list(full.columns)
    for i, h in enumerate(cols, 1):
        wsF.cell(row=3, column=i, value=h)
    style_header(wsF, 3, len(cols))
    for i, (_, row) in enumerate(full.iterrows()):
        for c, col in enumerate(cols, 1):
            wsF.cell(row=4 + i, column=c, value=row[col]).font = normal_f
    auto_width(wsF)

    keep = set()
    if opt.get("summary", True):
        keep.add("Summary_Dashboard")
    if opt.get("zone_sheets", True):
        keep.add("Zone_Summary")
        keep.update(n for n in wb.sheetnames if n.startswith("Z_"))
    if opt.get("section5", True):
        keep.add("Section5_Survey_Findings")
    if opt.get("section6", True):
        keep.add("Section6_Risk_Findings")
    if opt.get("form1_triage", True):
        keep.update(["Form1_TableA_Remedial", "Form1_TableB_OtherTrees"])
    if opt.get("form1_copy", True):
        keep.update(["Form1_Copy_Sheet1", "Form1_Copy_Sheet2"])
    if opt.get("form1_removal", True):
        keep.add("Form1_Removal_Fell")
    if opt.get("full", True):
        keep.add("Full_Schedule_Processed")
    keep &= set(wb.sheetnames)
    if not keep:
        keep.add(wb.sheetnames[0])
    for name in list(wb.sheetnames):
        if name not in keep and len(wb.sheetnames) > 1:
            del wb[name]
    log("輸出工作表: " + ", ".join(wb.sheetnames))

    wb.save(output_path)
    log(f"已儲存: {output_path}")

    word_path = ""
    if opt.get("word", True):
        word_path = str(Path(output_path).with_name(WORD_FILENAME))
        try:
            word_path = write_cover_word(
                word_path,
                active=active,
                species_df=species_df,
                total=total,
                dominant=dominant,
                yellow_n=yellow_n, orange_n=orange_n, red_n=red_n, black_n=black_n,
                yellow_ids=yellow_ids, orange_ids=orange_ids, red_ids=red_ids, black_ids=black_ids,
                pruning_n=pruning_n, removal_n=removal_n, rec_trees_n=rec_trees_n,
                log_fn=log,
            )
            log(f"已儲存 Word: {word_path}")
        except Exception as e:
            log(f"Word 未能產生: {e}")
            word_path = ""

    return {
        "total": total,
        "species": len(species_df),
        "yellow": yellow_n,
        "orange": orange_n,
        "red": red_n,
        "black": black_n,
        "dominant": dominant,
        "excluded": len(excluded),
        "form1_a": len(form1_A),
        "path": output_path,
        "word_path": word_path,
    }


def write_cover_word(path, active, species_df, total, dominant,
                    yellow_n, orange_n, red_n, black_n,
                    yellow_ids, orange_ids, red_ids, black_ids,
                    pruning_n, removal_n, rec_trees_n, log_fn=None):
    """Generate Cover-style Word with Section 5 / 6 filled from Schedule."""
    if not HAS_DOCX:
        raise RuntimeError("請安裝 python-docx：pip install python-docx")

    rec = active["Recommendations"].astype(str)
    PRUNE_PAT = (
        r"Crown cleaning|Crown reduction|Crown raising|Crown thinning|Crown restoration|"
        r"Regular crown size control|pruning|selective pruning|selective branch|"
        r"Removal of defective branch|Removal of dead (?:trunk|branch)|dead trunk|"
        r"Removal of one of the co-dominant|co-dominant branch|epicormic"
    )
    cleaning = active[rec.str.contains(r"Crown cleaning", case=False, na=False)]["Tree_No"].tolist()
    reduction = active[rec.str.contains(r"Crown reduction", case=False, na=False)]["Tree_No"].tolist()
    raising = active[rec.str.contains(r"Crown raising", case=False, na=False)]["Tree_No"].tolist()
    removals = active[rec.str.contains(r"Tree removal|\bfell(ing)?\b", case=False, na=False)]["Tree_No"].tolist()
    reinspect = active[rec.str.contains(r"Re-inspection", case=False, na=False)]["Tree_No"].tolist()

    def _ids(lst):
        return ", ".join(str(x) for x in lst) if lst else "/"

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.8)

    def _set_run_font(run, name="Times New Roman", size=12, bold=False):
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = name
        r = run._element
        rPr = r.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:ascii"), name)
        rFonts.set(qn("w:hAnsi"), name)
        rFonts.set(qn("w:eastAsia"), "PMingLiU")

    def add_heading_line(text, size=16, center=True):
        p = doc.add_paragraph()
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        _set_run_font(run, size=size, bold=True)
        return p

    def add_body(text, size=12):
        p = doc.add_paragraph()
        run = p.add_run(text)
        _set_run_font(run, size=size)
        p.paragraph_format.space_after = Pt(8)
        return p

    def shade_header(cell, color="1F4E79"):
        tc = cell._tePr if hasattr(cell, "_tePr") else cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), color)
        shd.set(qn("w:val"), "clear")
        tcPr.append(shd)

    def add_table(headers, rows):
        tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
        tbl.style = "Table Grid"
        for i, h in enumerate(headers):
            cell = tbl.rows[0].cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(h))
            _set_run_font(run, size=11, bold=True)
            run.font.color.rgb = RGBColor(255, 255, 255)
            shade_header(cell)
        for r_i, row in enumerate(rows):
            for c_i, val in enumerate(row):
                cell = tbl.rows[r_i + 1].cells[c_i]
                cell.text = ""
                run = cell.paragraphs[0].add_run("" if val is None else str(val))
                _set_run_font(run, size=11)
        doc.add_paragraph()
        return tbl

    def add_species_line(sci, chi, size=12):
        p = doc.add_paragraph()
        run = p.add_run("A total of ")
        _set_run_font(run, size=size)
        run = p.add_run(str(total))
        _set_run_font(run, size=size)
        run = p.add_run(f" trees and {len(species_df)} species were recorded. The dominant species is ")
        _set_run_font(run, size=size)
        run = p.add_run(str(sci))
        _set_run_font(run, size=size)
        run.italic = True
        if chi:
            run = p.add_run(" " + str(chi))
            _set_run_font(run, size=size)
        if "(" in str(dominant):
            tail = str(dominant)[str(dominant).rfind("("):]
            run = p.add_run(" " + tail)
            _set_run_font(run, size=size)
        run = p.add_run(".")
        _set_run_font(run, size=size)
        p.paragraph_format.space_after = Pt(8)

    def add_species_table(headers, rows):
        tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
        tbl.style = "Table Grid"
        for i, h in enumerate(headers):
            cell = tbl.rows[0].cells[i]
            cell.text = ""
            run = cell.paragraphs[0].add_run(str(h))
            _set_run_font(run, size=11, bold=True)
            run.font.color.rgb = RGBColor(255, 255, 255)
            shade_header(cell)
        for r_i, row in enumerate(rows):
            sci, chi, qty = row
            cell = tbl.rows[r_i + 1].cells[0]
            cell.text = ""
            p = cell.paragraphs[0]
            if sci:
                run = p.add_run(str(sci))
                _set_run_font(run, size=11)
                run.italic = True
            if chi:
                run = p.add_run((" " if sci else "") + str(chi))
                _set_run_font(run, size=11)
            cell2 = tbl.rows[r_i + 1].cells[1]
            cell2.text = ""
            run = cell2.paragraphs[0].add_run("" if qty is None else str(qty))
            _set_run_font(run, size=11)
        doc.add_paragraph()
        return tbl

    add_heading_line("Tree Risk Assessment Report", 18)
    add_heading_line("Schedule to Form 1", 13)
    add_heading_line(datetime.now().strftime("%B %Y"), 12)
    doc.add_paragraph()

    add_heading_line("Section 5", 16, center=False)
    add_heading_line("Tree Survey Findings", 14, center=False)
    top = species_df.iloc[0] if len(species_df) else None
    add_species_line(
        top["Scientific_Name"] if top is not None else "",
        top["Chinese_Name"] if top is not None else "",
    )
    species_alpha = species_df.sort_values("Scientific_Name", key=lambda s: s.astype(str).str.lower()).reset_index(drop=True)
    sp_rows = [
        (row["Scientific_Name"], row["Chinese_Name"], int(row["Count"]))
        for _, row in species_alpha.iterrows()
    ]
    sp_rows.append(("", "Total number of trees", total))
    add_species_table(["Tree Species", "Quantity"], sp_rows)

    add_body("TRIAGE colour summary")
    add_table(
        ["Colour", "No. of trees", "Tree No."],
        [
            ["Black", black_n, _ids(black_ids)],
            ["Red", red_n, _ids(red_ids)],
            ["Orange", orange_n, _ids(orange_ids)],
            ["Yellow", yellow_n, _ids(yellow_ids)],
        ],
    )

    add_heading_line("Section 6", 16, center=False)
    add_heading_line("Tree Risk Assessment Findings", 14, center=False)
    add_body(
        f"A total of {rec_trees_n} trees were recommended for further action, including tree pruning and tree removal. "
        f"(See Table 1). The proposed mitigations for specific trees are summarised in the table (See Table 2). "
        f"For details, please refer to Section 7 – Appendix D (Tree inventory)."
    )
    measure_rows = []
    detail_blocks = []
    if cleaning:
        measure_rows.append(["Crown cleaning", len(set(cleaning))])
        detail_blocks.append(("Crown cleaning:", cleaning))
    if reduction:
        measure_rows.append(["Crown reduction", len(set(reduction))])
        detail_blocks.append(("Crown reduction:", reduction))
    if raising:
        measure_rows.append(["Crown raising", len(set(raising))])
        detail_blocks.append(("Crown raising:", raising))
    thinning = active[rec.str.contains(r"Crown thinning|Crown restoration", case=False, na=False)]["Tree_No"].tolist()
    selective = active[rec.str.contains(r"selective pruning|selective branch|epicormic", case=False, na=False)]["Tree_No"].tolist()
    defective = active[rec.str.contains(
        r"Removal of defective branch|Removal of dead (?:trunk|branch)|dead trunk|Removal of one of the co-dominant|co-dominant branch",
        case=False, na=False,
    )]["Tree_No"].tolist()
    if thinning:
        measure_rows.append(["Crown thinning / restoration", len(set(thinning))])
        detail_blocks.append(("Crown thinning / restoration:", thinning))
    if selective:
        measure_rows.append(["Selective pruning", len(set(selective))])
        detail_blocks.append(("Selective pruning:", selective))
    if defective:
        measure_rows.append(["Removal of defective branches", len(set(defective))])
        detail_blocks.append(("Removal of defective branches:", defective))
    if removals:
        measure_rows.append(["Tree removal", len(set(removals))])
        detail_blocks.append(("Tree removal:", removals))
    if reinspect:
        measure_rows.append(["Re-inspection", len(set(reinspect))])
        detail_blocks.append(("Re-inspection:", reinspect))
    if not measure_rows:
        measure_rows.append(["No mitigation recommended", 0])
    add_body("Table 1. The summary of tree amounts with proposed mitigation measures")
    add_table(["Recommendations", "No. of trees"], measure_rows)
    add_body("Table 2. The summary of tree amounts with proposed mitigation measures")
    if detail_blocks:
        for title, ids in detail_blocks:
            add_body(f"{title}\n{_ids(ids)}")
    else:
        add_body("No mitigation measures were recommended.")

    add_heading_line("Appendices", 14, center=False)
    add_body("Appendix A – Form 1")
    add_body("Appendix B – Location Plan")
    add_body("Appendix C – Tree Inventory")
    add_body("Appendix D – Photographic record")

    doc.save(path)
    return path


# ── GUI ──
class ZoneApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Schedule to Form 1")
        self.root.geometry("780x560")
        self._build()

    def _build(self):
        pad = {"padx": 8, "pady": 4}
        frm = ttk.Frame(self.root, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Schedule to Form 1", font=("", 14, "bold")).grid(row=0, column=0, sticky="w", **pad)
        ttk.Label(frm, text="由 Tree Assessment Schedule 產生 Cover / Form 1；可揀輸出部份").grid(row=1, column=0, sticky="w", **pad)

        ttk.Label(frm, text="1. Tree Assessment Schedule（.xlsx / .xls）").grid(row=2, column=0, sticky="w", **pad)
        self.var_src = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_src, width=70).grid(row=3, column=0, sticky="ew", **pad)
        ttk.Button(frm, text="瀏覽...", command=self._browse_src).grid(row=3, column=1, **pad)

        ttk.Label(frm, text="2. 輸出資料夾").grid(row=4, column=0, sticky="w", **pad)
        self.var_out = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_out, width=70).grid(row=5, column=0, sticky="ew", **pad)
        ttk.Button(frm, text="瀏覽...", command=self._browse_out).grid(row=5, column=1, **pad)

        ttk.Label(frm, text="3. Dead tree 顯示為").grid(row=6, column=0, sticky="w", **pad)
        self.var_dead = tk.StringVar(value="Unidentified")
        opt = ttk.Frame(frm)
        opt.grid(row=7, column=0, sticky="w", **pad)
        ttk.Radiobutton(opt, text="Unidentified", variable=self.var_dead, value="Unidentified").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(opt, text="Dead tree 死樹", variable=self.var_dead, value="Dead tree").pack(side="left")

        ttk.Label(frm, text="4. 輸出部份（可多選）").grid(row=8, column=0, sticky="w", **pad)
        parts = ttk.Frame(frm)
        parts.grid(row=9, column=0, columnspan=2, sticky="w", **pad)
        self.opt_vars = {}
        checks = [
            ("summary", "Summary", True),
            ("section5", "Section 5", True),
            ("section6", "Section 6", True),
            ("form1_triage", "Form1 A/B（Triage／緩減）", True),
            ("form1_removal", "Form1 A=Removal／B=其他", True),
            ("form1_copy", "Form1 官方複製欄", True),
            ("zone_sheets", "Zone Summary＋每 Zone 一頁", True),
            ("full", "Full Schedule", False),
            ("word", "Word Cover（Section 5/6）", True),
        ]
        for i, (key, label, default) in enumerate(checks):
            v = tk.BooleanVar(value=default)
            self.opt_vars[key] = v
            ttk.Checkbutton(parts, text=label, variable=v).grid(row=i // 4, column=i % 4, sticky="w", padx=6, pady=2)

        preset = ttk.Frame(frm)
        preset.grid(row=10, column=0, sticky="w", **pad)
        ttk.Label(preset, text="預設：").pack(side="left")
        ttk.Button(preset, text="全部", command=lambda: self._set_parts(True)).pack(side="left", padx=4)
        ttk.Button(preset, text="只要 Cover+Form1", command=self._preset_cover).pack(side="left", padx=4)
        ttk.Button(preset, text="只要 Zone", command=self._preset_zone).pack(side="left", padx=4)

        self.btn = ttk.Button(frm, text="▶ 產生報告", command=self._run)
        self.btn.grid(row=11, column=0, sticky="w", pady=10)

        self.txt = scrolledtext.ScrolledText(frm, height=14, width=90, bg="#1e1e1e", fg="#d4d4d4")
        self.txt.grid(row=12, column=0, columnspan=2, sticky="nsew", **pad)
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(12, weight=1)

        self.log("排除：Removed / Construction / Outside Boundary / Duplicate / Out of LOT")
        self.log("可揀輸出部份；Dead tree 可選 Unidentified 或 Dead tree")

    def _set_parts(self, val: bool):
        for v in self.opt_vars.values():
            v.set(val)

    def _preset_cover(self):
        want = {"summary", "section5", "section6", "form1_triage", "form1_removal"}
        for k, v in self.opt_vars.items():
            v.set(k in want)

    def _preset_zone(self):
        want = {"summary", "zone_sheets", "section5", "section6"}
        for k, v in self.opt_vars.items():
            v.set(k in want)

    def log(self, msg):
        self.txt.insert(tk.END, msg + "\n")
        self.txt.see(tk.END)
        self.root.update_idletasks()

    def _safe_initial_dir(self):
        src = self.var_src.get().strip()
        if src and os.path.isfile(src):
            return os.path.dirname(src)
        home = os.path.expanduser("~")
        docs = os.path.join(home, "Documents")
        if os.path.isdir(docs):
            return docs
        return home

    def _browse_src(self):
        try:
            p = filedialog.askopenfilename(
                title="選擇 Tree Assessment Schedule",
                filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")],
                initialdir=self._safe_initial_dir(),
            )
        except Exception as e:
            self.log(f"選檔對話框錯誤: {e}")
            p = ""
        if p:
            self.var_src.set(p)
            if not self.var_out.get().strip():
                self.var_out.set(os.path.dirname(p))

    def _browse_out(self):
        init = self._safe_initial_dir()
        p = ""
        try:
            p = filedialog.askdirectory(title="選擇輸出資料夾", initialdir=init, mustexist=True)
        except Exception as e:
            self.log(f"選資料夾失敗（可改在輸入框直接貼路徑）: {e}")
            p = ""
        if not p:
            # OneDrive / 桌面 moniker 失敗時：選該資料夾內任何檔案，取其資料夾
            try:
                f = filedialog.askopenfilename(
                    title="無法開資料夾時：請選輸出資料夾內任何一個檔案",
                    initialdir=init,
                    filetypes=[("All", "*.*")],
                )
                if f:
                    p = os.path.dirname(f)
            except Exception as e:
                self.log(f"備選選檔亦失敗: {e}")
        if p:
            self.var_out.set(p)

    def _run(self):
        src = self.var_src.get().strip()
        out = self.var_out.get().strip()
        if not src or not os.path.isfile(src):
            messagebox.showerror("錯誤", "請選擇 Schedule 檔案")
            return
        if not out:
            # 預設與 Schedule 同一資料夾
            out = os.path.dirname(src)
            self.var_out.set(out)
        try:
            os.makedirs(out, exist_ok=True)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法建立輸出資料夾:\n{out}\n{e}\n\n可直接在「輸出資料夾」欄貼上路徑，例如:\nC:\\Users\\davy.chan\\Documents")
            return
        self.btn.config(state=tk.DISABLED)
        threading.Thread(target=self._worker, args=(src, out), daemon=True).start()

    def _worker(self, src, out):
        try:
            self.log("=" * 50)
            self.log(f"讀取: {src}")
            df = load_schedule(src)
            self.log(f"工作表資料列: {len(df)}")
            path = str(Path(out) / OUTPUT_FILENAME)
            opts = {k: v.get() for k, v in self.opt_vars.items()}
            # Auto：無真實 Zone 就唔出每 Zone 分頁（仍可手動勾）
            if opts.get("zone_sheets"):
                z = df["Zone"].astype(str).str.strip()
                real_z = z[~z.isin(["", "-", "nan", "N.A.", "N.A"])]
                if len(real_z.unique()) <= 1 and "TRMZ" in df.columns:
                    tr = df["TRMZ"].astype(str).str.strip()
                    real_z = tr[~tr.isin(["", "-", "nan", "N.A.", "N.A"])]
                if len(real_z.unique()) <= 1:
                    self.log("未偵測到多個 Zone，已略過每 Zone 分頁（Summary 仍計）")
                    opts["zone_sheets"] = False
            r = generate(df, path, log_fn=self.log, dead_as=self.var_dead.get(), options=opts)
            self.log(f"完成 → {r}")
            extra = f"\nWord: {r.get('word_path')}" if r.get("word_path") else ""
            messagebox.showinfo("完成", f"有效總數: {r['total']}\nYellow: {r['yellow']}  Orange: {r['orange']}\nExcel: {r['path']}{extra}")
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            self.log(f"錯誤: {e}\n{tb}")
            messagebox.showerror("錯誤", f"{e}\n\n{tb[-800:]}")
        finally:
            self.btn.config(state=tk.NORMAL)

    def run(self):
        self.root.mainloop()


def main():
    if not HAS_TK:
        print("需要 tkinter")
        return
    ZoneApp().run()


if __name__ == "__main__":
    main()
