#!/usr/bin/env python3
"""
Live price-check for the UWANT price-comparison site.

Reads data/data.json for the exact product URL already on file for every
(model, site) cell and every extra, re-fetches each one right now, and
rewrites data/data.json with fresh prices/stock + a new history point for
today. Runs inside GitHub Actions (see .github/workflows/update-prices.yml),
triggered by the "עדכן דוח" button on the site.

Per-site extraction rules below were reverse-engineered by hand, one site at
a time, against the live pages — see the comment on each extractor for what
it's working around.

Three sites — KSP, CWC, ALM — are deliberately NOT checked here. All three
render price client-side and all three block or silently withhold pricing
data from cloud/datacenter IPs specifically (confirmed by hand: KSP and CWC
return an explicit 403 "access blocked" page from GitHub Actions' IP range
even via a real headless browser with anti-detection tweaks; ALM renders the
page shell but omits the price entirely). This is deliberate anti-bot
behavior targeting exactly this kind of automation, not a bug to work around
with a better selector or a longer wait. Their cells simply keep whatever
value they last had; check them by hand in a live session when needed (a
real interactive browser, not a cloud IP, gets past this fine — that's how
they were verified originally).

A site that fails (blocked, timeout, page changed) keeps its previous value
in data.json and is listed in the run summary — this must never crash the
whole run or silently null out data that was fine yesterday.
"""
import json
import re
import sys
import time
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parent.parent
DATA_PATH = ROOT / "data" / "data.json"

ACTIVE_MODELS = [
    "U300 V25", "T300", "U400", "D500", "D600", "D700", "DX800",
    "V500", "V600", "V800", "M400", "M500", "M600",
    "Y100", "Y200", "Y100 Steam", "Y200 Steam",
]
SKIPPED_SITES = {"KSP", "CWC", "ALM"}  # blocked from cloud IPs — see module docstring

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "he-IL,he;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


def normalize_price(s):
    if s is None:
        return None
    s = re.sub(r"[^\d.]", "", str(s).replace(",", ""))
    if not s:
        return None
    try:
        return round(float(s))
    except ValueError:
        return None


def fetch_html(url, session, retries=2, delay=1.5):
    last_err = None
    for attempt in range(retries + 1):
        try:
            r = session.get(url, headers=HEADERS, timeout=20)
            r.raise_for_status()
            return r.text
        except Exception as e:  # noqa: BLE001 - genuinely want to retry on anything
            last_err = e
            if attempt < retries:
                time.sleep(delay)
    raise last_err


# ---------------- per-site extractors: (html) -> (price:int|None, stock:'in'|'out'|None) ----------------

def extract_woocommerce(html):
    """CWC, Multi Store: WooCommerce. The price block can be a plain price, or
    a <del>regular</del> immediately followed by the real (sale) price with no
    wrapping <ins> — take the first .woocommerce-Price-amount NOT inside a
    <del>, not just the first one on the page (that was the original bug: the
    literal first match is the crossed-out regular price on a sale item)."""
    soup = BeautifulSoup(html, "html.parser")
    container = soup.select_one(".summary .price, .entry-summary .price")
    price = None
    if container:
        amounts = container.select(".woocommerce-Price-amount bdi") or container.select(".woocommerce-Price-amount")
        target = next((el for el in amounts if not el.find_parent("del")), None) or (amounts[0] if amounts else None)
        if target:
            price = target.get_text(strip=True)
    add_to_cart = soup.select_one('button[name="add-to-cart"], .single_add_to_cart_button')
    out_of_stock = any(w in html for w in ("אזל מהמלאי", "אזל המלאי", "אין במלאי"))
    stock = "out" if out_of_stock else ("in" if add_to_cart else None)
    return normalize_price(price), stock


def extract_lastprice(html):
    """LastPrice shows the live buy price as 'קנה עכשיו ב- X ₪' or (promo
    copy) 'שלמו רק X ₪' / 'הצעה פנטסטית רק X ₪' — never a fixed label, so
    match on the "רק"/"ב-" that always precedes the real price."""
    m = re.search(r"(?:רק|ב-)\s*([\d,]+)\s*₪", html)
    price = m.group(1) if m else None
    out_of_stock = any(w in html for w in ("אזל מהמלאי", "אין במלאי"))
    return normalize_price(price), ("out" if out_of_stock else ("in" if price else None))


def extract_traklin(html):
    """The price sits inside <strong>590</strong> right after "לרכישה ב",
    with an HTML tag between the label and the digits and the ₪ as an
    &#8362; entity — match on cleaned text (tags stripped, entities decoded),
    not the raw markup."""
    text = BeautifulSoup(html, "html.parser").get_text(" ")
    m = re.search(r"לרכישה\s*ב\s*([\d,]+)", text)
    price = m.group(1) if m else None
    return normalize_price(price), ("in" if price else None)


def find_json_ld_price(soup):
    """Pull price + stock out of a page's JSON-LD Product schema, wherever it
    is nested: a bare Product object, an ItemPage wrapping one in
    `mainEntity` (Magento does this — caught the NetoNeto bug), or a
    `@graph` array of nodes. This is far more reliable than scanning the page
    for the first ₪-adjacent number, which happily matches a shipping
    threshold or a stray JS comment before it ever reaches the real price."""
    def offers_of(node):
        if not isinstance(node, dict):
            return []
        if node.get("@type") == "Product" and node.get("offers"):
            o = node["offers"]
            return o if isinstance(o, list) else [o]
        return []

    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "{}")
        except Exception:  # noqa: BLE001
            continue
        nodes = data if isinstance(data, list) else [data]
        candidates = []
        for node in nodes:
            candidates += offers_of(node)
            if isinstance(node, dict):
                candidates += offers_of(node.get("mainEntity"))
                for g in node.get("@graph", []) or []:
                    candidates += offers_of(g)
        for o in candidates:
            if isinstance(o, dict) and o.get("price"):
                avail = str(o.get("availability", ""))
                stock = "out" if "OutOfStock" in avail else ("in" if avail else None)
                return o["price"], stock
    return None, None


def extract_netoneto(html):
    soup = BeautifulSoup(html, "html.parser")
    price, stock = find_json_ld_price(soup)
    if price is not None:
        return normalize_price(price), stock
    text = soup.get_text(" ")
    m = re.search(r"([\d,]+)\s*₪", text)
    price = m.group(1) if m else None
    out_of_stock = "אזל" in text and "הוספה לסל" not in text
    return normalize_price(price), ("out" if out_of_stock else ("in" if price else None))


def extract_uwant_il(html):
    """UWANT's own Shopify store: read the JSON-LD Product schema first (most
    robust against theme changes); fall back to a decimal-price regex
    ('X.00 ₪' — Shopify formats prices with two decimals) if that's absent."""
    soup = BeautifulSoup(html, "html.parser")
    price, stock = find_json_ld_price(soup)
    if price is not None:
        return normalize_price(price), stock
    m = re.search(r"([\d,]+\.\d{2})\s*₪", html)
    price = m.group(1) if m else None
    out_of_stock = "אזל מהמלאי" in html
    return normalize_price(price), ("out" if out_of_stock else ("in" if price else None))


def extract_citydeal(html):
    """City Deal (Next.js) streams the product object as an escaped JSON
    string fragment (React Server Components payload) near the end of the
    document — \"price\":N, \"inStock\":true/false. It's the first "price"
    key in the page, well before any 'related products' JSON further down,
    so taking the first match is reliable. Fall back to a currency regex on
    cleaned text if the payload shape ever changes."""
    m = re.search(r'\\?"price\\?"\s*:\s*(\d+)', html)
    if m:
        price = m.group(1)
        stock_m = re.search(r'\\?"inStock\\?"\s*:\s*(true|false)', html)
        stock = "in" if (stock_m and stock_m.group(1) == "true") else ("out" if stock_m else None)
        return normalize_price(price), stock
    text = BeautifulSoup(html, "html.parser").get_text(" ")
    m = re.search(r"([\d,]+)\s*(?:ILS|₪|ש\"ח)", text)
    price = m.group(1) if m else None
    return normalize_price(price), ("in" if price else None)


def extract_walla(html):
    """Walla Shops embeds a "recommended items" carousel ABOVE the actual
    product on many pages, sharing the same .item_price_money class as the
    real price — a naive first-match grabs the carousel's price instead.
    Anchor on the "הוספה לעגלה" (add-to-cart) button, which only exists once,
    inside the real product's own container, and read the price from there."""
    soup = BeautifulSoup(html, "html.parser")
    btn = None
    for el in soup.find_all(True):
        if not el.find(True, recursive=False) and el.get_text(strip=True) == "הוספה לעגלה":
            btn = el
            break
    price = None
    if btn is not None:
        node = btn
        for _ in range(8):
            if node.parent is None:
                break
            node = node.parent
        price_el = node.select_one(".item_price_money") if node else None
        if price_el:
            price = price_el.get_text(strip=True)
    return normalize_price(price), ("in" if price else None)


SITE_EXTRACTORS = {
    "NetoNeto": extract_netoneto,
    "City Deal": extract_citydeal,
    "UWANT IL": extract_uwant_il,
    "Multi Store": extract_woocommerce,
    "LastPrice": extract_lastprice,
    "Walla Shops": extract_walla,
    "Traklin": extract_traklin,
}


def check_simple_site(site, entries, log):
    """entries: list of (kind, key, url) where kind is 'matrix' or 'extras'."""
    extractor = SITE_EXTRACTORS[site]
    session = requests.Session()
    results = {}
    for kind, key, url in entries:
        try:
            html = fetch_html(url, session)
            price, stock = extractor(html)
            results[(kind, key)] = (price, stock)
            if price is None:
                log.append(f"[{site}] {key}: fetched OK but no price extracted (len={len(html)})")
            time.sleep(0.4)  # be polite to the same host; City Deal 429s on bursts
        except Exception as e:  # noqa: BLE001
            log.append(f"[{site}] {key}: FAILED ({e}) - keeping previous value")
            results[(kind, key)] = (None, None)
    return site, results


def main():
    log = []
    data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    site_names = [s["name"] for s in data["sites"]]

    # Build the checklist from URLs already on file (see build_data.py / the
    # button's own README for why we never rediscover products from scratch).
    by_site = {}
    for model in ACTIVE_MODELS:
        for site in site_names:
            if site in SKIPPED_SITES:
                continue
            cell = data["matrix"].get(model, {}).get(site)
            if cell and cell.get("url"):
                by_site.setdefault(site, []).append(("matrix", (model, site), cell["url"]))
    for i, x in enumerate(data.get("extras", [])):
        if x.get("url") and x.get("site") not in SKIPPED_SITES:
            by_site.setdefault(x["site"], []).append(("extras", i, x["url"]))

    if SKIPPED_SITES:
        log.append(f"Skipped entirely (blocked from cloud IPs, kept previous values): {', '.join(sorted(SKIPPED_SITES))}")

    all_results = {}
    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = [pool.submit(check_simple_site, site, entries, log) for site, entries in by_site.items()]
        for fut in as_completed(futures):
            site, results = fut.result()
            all_results[site] = results

    # ---- rebuild matrix ----
    checked_count = 0
    for model in ACTIVE_MODELS:
        for site in site_names:
            r = all_results.get(site, {}).get(("matrix", (model, site)))
            if r and r[0] is not None:
                old_cell = data["matrix"].get(model, {}).get(site) or {}
                data["matrix"].setdefault(model, {})[site] = {
                    "price": r[0], "regular": None, "club": None,
                    "stock": "אזל" if r[1] == "out" else "במלאי",
                    "url": old_cell.get("url"), "note": None,
                }
                checked_count += 1
            # else: leave whatever was already there (failed fetch, or never tracked)

    # ---- rebuild extras ----
    for i, x in enumerate(data.get("extras", [])):
        r = all_results.get(x.get("site"), {}).get(("extras", i))
        if r and r[0] is not None:
            x["price"] = r[0]
            x["stock"] = "אזל" if r[1] == "out" else "במלאי"
            checked_count += 1

    # ---- stats ----
    stats = []
    for site in site_names:
        coverage = sum(1 for m in ACTIVE_MODELS if data["matrix"].get(m, {}).get(site))
        stats.append({
            "site": site, "variants": coverage, "total_records": coverage,
            "price_coverage": coverage,
            "coverage_pct": round(100 * coverage / len(ACTIVE_MODELS)),
            "note": None,
        })
    data["stats"] = stats

    # ---- history: append today, never overwrite older dates ----
    now = datetime.now(ZoneInfo("Asia/Jerusalem"))
    report_date = now.date().isoformat()
    history = data.setdefault("history", {})
    for model in ACTIVE_MODELS:
        for site in site_names:
            cell = data["matrix"].get(model, {}).get(site)
            if not cell or cell.get("price") is None:
                continue
            series = history.setdefault(model, {}).setdefault(site, [])
            series[:] = [p for p in series if p["date"] != report_date]
            series.append({"date": report_date, "price": cell["price"]})
            series.sort(key=lambda p: p["date"])

    title = f"דוח ניטור מחירי UWANT — {now.strftime('%d.%m.%Y %H:%M')} (Asia/Jerusalem)"
    data["generated_at"] = now.isoformat(timespec="seconds")
    data["source_file"] = f"live-update-{report_date}"
    data["title"] = title
    data["checked_at"] = title
    data["report_date"] = report_date

    DATA_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")

    print(f"Checked {checked_count}/{sum(len(v) for v in by_site.values())} cells successfully.")
    if log:
        print("Failures (kept previous value):")
        for line in log:
            print(" -", line)
    else:
        print("No failures.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(1)
