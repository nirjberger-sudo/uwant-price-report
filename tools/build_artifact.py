#!/usr/bin/env python3
"""
Bundle the static site (index.html + styles.css + app.js + data/data.json +
data/UWANT_price_report_latest.xlsx) into a single self-contained HTML file
suitable for the Artifacts publisher (which accepts exactly one file, no
external requests). Run this after build_data.py whenever data.json changes.

Usage:
    py build_artifact.py
Writes: site/dist/uwant-price-report.artifact.html
"""
import base64
import json
import re
from pathlib import Path

SITE_DIR = Path(__file__).resolve().parent.parent
DIST_DIR = SITE_DIR / "dist"
OUT_PATH = DIST_DIR / "uwant-price-report.artifact.html"


def main():
    html = (SITE_DIR / "index.html").read_text(encoding="utf-8")
    css = (SITE_DIR / "styles.css").read_text(encoding="utf-8")
    js = (SITE_DIR / "app.js").read_text(encoding="utf-8")
    data = json.loads((SITE_DIR / "data" / "data.json").read_text(encoding="utf-8"))

    xlsx_path = SITE_DIR / "data" / "UWANT_price_report_latest.xlsx"
    xlsx_b64 = base64.b64encode(xlsx_path.read_bytes()).decode("ascii")
    xlsx_data_uri = (
        "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + xlsx_b64
    )

    logo_path = SITE_DIR / "assets" / "uwant-logo-black.png"
    logo_b64 = base64.b64encode(logo_path.read_bytes()).decode("ascii")
    logo_data_uri = "data:image/png;base64," + logo_b64
    html = html.replace('src="./assets/uwant-logo-black.png"', f'src="{logo_data_uri}"')

    for fname in ("heebo-hebrew.woff2", "heebo-latin.woff2"):
        font_bytes = (SITE_DIR / "assets" / fname).read_bytes()
        font_b64 = base64.b64encode(font_bytes).decode("ascii")
        css = css.replace(
            f'url("./assets/{fname}") format("woff2")',
            f'url("data:font/woff2;base64,{font_b64}") format("woff2")',
        )

    # app.js downloads via a relative file path; the inlined build downloads via a data: URI instead.
    js_inline = js.replace(
        'a.href = "./data/UWANT_price_report_latest.xlsx";',
        "a.href = window.__UWANT_XLSX_DATA_URI__;",
    )

    html = html.replace(
        '<link rel="stylesheet" href="./styles.css">',
        f"<style>\n{css}\n</style>",
    )
    data_script = (
        f"<script>window.__UWANT_DATA__ = {json.dumps(data, ensure_ascii=False)};\n"
        f'window.__UWANT_XLSX_DATA_URI__ = "{xlsx_data_uri}";</script>'
    )
    html = html.replace(
        '<script src="./app.js"></script>',
        f"{data_script}\n<script>\n{js_inline}\n</script>",
    )

    DIST_DIR.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(html, encoding="utf-8")
    print(f"Wrote {OUT_PATH} ({OUT_PATH.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
