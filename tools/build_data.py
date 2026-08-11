#!/usr/bin/env python3
"""
UWANT price-monitor report -> site/data/data.json

Run this every time a new UWANT price report .xlsx is produced.
It reads the report, rebuilds the current snapshot (matrix / stats / extras)
and appends a new point to the price-history series - it never touches any
of the site's HTML/CSS/JS.

Usage:
    py build_data.py "C:\\path\\to\\UWANT_price_report_YYYY-MM-DD.xlsx"

The report must follow the standard 4-sheet layout used by the UWANT price
monitor ("price_report" family):
    sheet 0: model x site comparison matrix
    sheet 1: full catalog (one row per listing, with direct product links)
    sheet 2: coverage / quality statistics per site
    sheet 3: extra UWANT products (accessories / models outside the active list)

Excluded models (per product spec, never shown even if present in the source
report): V500 Plus, Y100B, Y200B.
"""
import sys
import re
import json
import shutil
import datetime
from pathlib import Path

import openpyxl

SITE_DIR = Path(__file__).resolve().parent.parent
DATA_PATH = SITE_DIR / "data" / "data.json"
REPORT_COPY_PATH = SITE_DIR / "data" / "UWANT_price_report_latest.xlsx"

ACTIVE_MODELS = [
    "U300 V25", "T300", "U400", "D500", "D600", "D700", "DX800",
    "V500", "V600", "V800", "M400", "M500", "M600",
    "Y100", "Y200", "Y100 Steam", "Y200 Steam",
]
EXCLUDED_MODELS = {"V500 PLUS", "V500 Plus", "Y100B", "Y200B"}

SITE_ID_MAP = {
    "KSP": "ksp", "Walla Shops": "walla", "NetoNeto": "netoneto", "CWC": "cwc",
    "LastPrice": "lastprice", "ALM": "alm", "Traklin": "traklin",
    "City Deal": "citydeal", "Multi Store": "multistore", "UWANT IL": "uwant",
}


def parse_hyperlink_cell(v):
    """Cells were written with a HYPERLINK() formula that this workbook
    engine could not evaluate, so the cached value is left as a descriptive
    placeholder string. Recover (display_text, url) from it; for a plain
    string/None, return (v, None)."""
    if isinstance(v, str) and v.startswith("HYPERLINK is not implemented"):
        m = re.search(r"linkLocation=(.*?), friendlyName=(.*)$", v)
        if m:
            return m.group(2).strip(), m.group(1)
    return v, None


def norm_model(name):
    if not name:
        return None
    s = str(name).replace("UWANT", "").strip()
    s = re.sub(r"\s+", " ", s)
    return s


def to_number(v):
    if v in (None, "-", ""):
        return None
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


def extract_date(title, path):
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", title or "")
    if m:
        d, mo, y = m.groups()
        return f"{y}-{mo}-{d}"
    return datetime.date.fromtimestamp(path.stat().st_mtime).isoformat()


def header_index(header_row, *synonym_groups):
    """Map each requested field (a group of acceptable header-text synonyms,
    matched by prefix so parenthetical suffixes don't break the match) to its
    column index. Report layouts have drifted release to release (columns
    added/reordered), so every sheet is read by header name, never position."""
    idx = {}
    for col, cell in enumerate(header_row):
        if not isinstance(cell, str):
            continue
        idx[cell.strip()] = col
    result = []
    for synonyms in synonym_groups:
        found = None
        for syn in synonyms:
            for header_text, col in idx.items():
                if header_text.startswith(syn):
                    found = col
                    break
            if found is not None:
                break
        result.append(found)
    return result


def get(row, col):
    if col is None or col >= len(row):
        return None
    return row[col]


def build(report_path: Path):
    wb = openpyxl.load_workbook(report_path, data_only=True, read_only=True)
    sheet_names = wb.sheetnames
    main_ws = wb[sheet_names[0]]
    catalog_ws = wb[sheet_names[1]]
    stats_ws = wb[sheet_names[2]]
    extras_ws = wb[sheet_names[3]]

    # ---- meta + site list (row 2 of the main sheet) ----
    main_rows = list(main_ws.iter_rows(values_only=True))
    title = main_rows[0][0] if main_rows else ""
    checked_at = title
    report_date = extract_date(title, report_path)

    header = main_rows[1]
    sites = []
    for v in header[1:]:
        name, url = parse_hyperlink_cell(v)
        if not url:
            continue  # only real per-site hyperlink columns; skip summary columns
        sites.append({"id": SITE_ID_MAP.get(name, re.sub(r"\W+", "-", name.lower())),
                       "name": name, "uwant_url": url})
    site_names = [s["name"] for s in sites]

    # ---- full catalog: one row per listing ----
    cat_rows = list(catalog_ws.iter_rows(values_only=True))
    c_site, c_model, c_name, c_price, c_regular, c_club, c_stock, c_kind, c_url, c_verify, c_note, c_checked = \
        header_index(cat_rows[0],
                     ["אתר"], ["דגם ממופה"], ["שם מוצר"], ["מחיר נוכחי"], ["מחיר רגיל"],
                     ["מחיר מועדון"], ["מלאי"], ["סוג"], ["קישור"], ["אימות"], ["הערה"], ["מועד בדיקה"])
    catalog = []
    for row in cat_rows[1:]:
        if not row or get(row, c_site) is None:
            continue
        _, url = parse_hyperlink_cell(get(row, c_url))
        model_n = norm_model(get(row, c_model))
        if model_n in EXCLUDED_MODELS:
            continue
        catalog.append({
            "site": get(row, c_site), "model": model_n, "name": get(row, c_name),
            "price": to_number(get(row, c_price)), "regular": to_number(get(row, c_regular)),
            "club": to_number(get(row, c_club)),
            "stock": get(row, c_stock), "kind": get(row, c_kind), "url": url,
            "verify": get(row, c_verify), "note": get(row, c_note), "checked": get(row, c_checked),
        })

    # ---- comparison matrix: cheapest in-catalog "device" listing per model+site ----
    matrix = {m: {s: None for s in site_names} for m in ACTIVE_MODELS}
    for m in ACTIVE_MODELS:
        for s in site_names:
            candidates = [c for c in catalog if c["model"] == m and c["site"] == s
                          and c["kind"] == "מכשיר" and c["price"] is not None]
            if not candidates:
                continue
            best = min(candidates, key=lambda c: c["price"])
            matrix[m][s] = {
                "price": best["price"], "regular": best["regular"], "club": best["club"],
                "stock": best["stock"], "url": best["url"], "note": best["note"],
            }

    # ---- stats sheet (coverage stats are derived straight from the catalog below,
    #      since the sheet's own column set has drifted across report vintages;
    #      only the free-text note column is worth carrying over) ----
    stats_rows = list(stats_ws.iter_rows(values_only=True))
    header_row_idx = next((i for i, r in enumerate(stats_rows) if r and r[0] == "אתר"), 2)
    s_site, s_note = header_index(stats_rows[header_row_idx], ["אתר"], ["הערה", "סטטוס"])
    notes_by_site = {}
    for row in stats_rows[header_row_idx + 1:]:
        if not row or get(row, s_site) is None:
            continue
        notes_by_site[get(row, s_site)] = get(row, s_note)

    stats = []
    for s in site_names:
        site_catalog = [c for c in catalog if c["site"] == s]
        variants = sum(1 for c in site_catalog if c["kind"] == "מכשיר")
        total_records = len(site_catalog)
        price_coverage = len({c["model"] for c in site_catalog
                               if c["kind"] == "מכשיר" and c["model"] in ACTIVE_MODELS and c["price"] is not None})
        coverage_pct = round(100 * price_coverage / len(ACTIVE_MODELS))
        stats.append({
            "site": s, "variants": variants, "total_records": total_records,
            "price_coverage": price_coverage, "coverage_pct": coverage_pct,
            "note": notes_by_site.get(s),
        })

    # ---- extras: dedicated sheet + any catalog rows outside the active-model / device scope ----
    extras_rows = list(extras_ws.iter_rows(values_only=True))
    e_site, e_model, e_name, e_price, e_stock, e_kind, e_url = \
        header_index(extras_rows[0], ["אתר"], ["דגם"], ["שם מדויק"], ["מחיר"], ["מלאי"], ["סוג"], ["קישור"])
    extras = []
    seen_urls = set()
    for row in extras_rows[1:]:
        if not row or get(row, e_site) is None:
            continue
        _, url = parse_hyperlink_cell(get(row, e_url))
        model_n = norm_model(get(row, e_model))
        if model_n in EXCLUDED_MODELS:
            continue
        extras.append({"site": get(row, e_site), "model": model_n, "name": get(row, e_name),
                        "price": to_number(get(row, e_price)),
                        "stock": get(row, e_stock), "kind": get(row, e_kind), "url": url})
        if url:
            seen_urls.add(url)
    for c in catalog:
        if c["model"] not in ACTIVE_MODELS or c["kind"] == "אביזר":
            if c["url"] and c["url"] not in seen_urls:
                extras.append({"site": c["site"], "model": c["model"], "name": c["name"],
                                "price": c["price"], "stock": c["stock"], "kind": c["kind"], "url": c["url"]})
                seen_urls.add(c["url"])

    return {
        "title": title,
        "checked_at": checked_at,
        "report_date": report_date,
        "sites": sites,
        "models": ACTIVE_MODELS,
        "matrix": matrix,
        "stats": stats,
        "extras": extras,
    }


def merge_history(existing_history, snapshot):
    history = existing_history or {}
    for model, per_site in snapshot["matrix"].items():
        for site, cell in per_site.items():
            if not cell or cell.get("price") is None:
                continue
            series = history.setdefault(model, {}).setdefault(site, [])
            series = [p for p in series if p["date"] != snapshot["report_date"]]
            series.append({"date": snapshot["report_date"], "price": cell["price"]})
            series.sort(key=lambda p: p["date"])
            history[model][site] = series
    return history


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    report_path = Path(sys.argv[1]).resolve()
    if not report_path.exists():
        print(f"File not found: {report_path}")
        sys.exit(1)

    snapshot = build(report_path)

    existing = {}
    if DATA_PATH.exists():
        existing = json.loads(DATA_PATH.read_text(encoding="utf-8"))

    history = merge_history(existing.get("history"), snapshot)

    out = {
        "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "source_file": report_path.name,
        **snapshot,
        "history": history,
    }

    DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
    DATA_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    shutil.copyfile(report_path, REPORT_COPY_PATH)
    print(f"Wrote {DATA_PATH} (report date {snapshot['report_date']}, "
          f"{sum(len(v) for v in history.values())} models with history)")
    print(f"Copied report to {REPORT_COPY_PATH} (used by the site's XLSX download button)")


if __name__ == "__main__":
    main()
