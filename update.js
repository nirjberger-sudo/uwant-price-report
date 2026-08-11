(() => {
  "use strict";

  // Mirrors site/tools/build_data.py exactly (same constants, same header-by-name
  // parsing, same merge rules) so a report uploaded here produces identical output
  // to running the Python script — just without anyone needing to run Python.

  const GITHUB_OWNER = "nirjberger-sudo";
  const GITHUB_REPO = "uwant-price-report";
  const GITHUB_BRANCH = "main";
  const TOKEN_KEY = "uwant_gh_token";

  const ACTIVE_MODELS = [
    "U300 V25", "T300", "U400", "D500", "D600", "D700", "DX800",
    "V500", "V600", "V800", "M400", "M500", "M600",
    "Y100", "Y200", "Y100 Steam", "Y200 Steam",
  ];
  const EXCLUDED_MODELS = new Set(["V500 PLUS", "V500 Plus", "Y100B", "Y200B"]);
  const SITE_ID_MAP = {
    "KSP": "ksp", "Walla Shops": "walla", "NetoNeto": "netoneto", "CWC": "cwc",
    "LastPrice": "lastprice", "ALM": "alm", "Traklin": "traklin",
    "City Deal": "citydeal", "Multi Store": "multistore", "UWANT IL": "uwant",
  };

  function parseHyperlinkCell(v) {
    if (typeof v === "string" && v.startsWith("HYPERLINK is not implemented")) {
      const m = v.match(/linkLocation=(.*?), friendlyName=(.*)$/);
      if (m) return [m[2].trim(), m[1]];
    }
    return [v, null];
  }

  function normModel(name) {
    if (!name) return null;
    return String(name).replace(/UWANT/g, "").trim().replace(/\s+/g, " ");
  }

  function toNumber(v) {
    if (v === null || v === undefined || v === "-" || v === "") return null;
    if (typeof v === "number") return v;
    const n = Number(String(v).replace(/,/g, "").trim());
    return Number.isNaN(n) ? null : n;
  }

  function extractDate(title, fallbackDate) {
    const m = (title || "").match(/(\d{2})\.(\d{2})\.(\d{4})/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    return fallbackDate.toISOString().slice(0, 10);
  }

  // Field -> column index, matched by header-text prefix (never by position — report
  // layouts drift release to release).
  function headerIndex(headerRow, synonymGroups) {
    const idx = [];
    headerRow.forEach((cell, col) => {
      if (typeof cell === "string") idx.push([cell.trim(), col]);
    });
    return synonymGroups.map((synonyms) => {
      for (const syn of synonyms) {
        for (const [text, col] of idx) {
          if (text.startsWith(syn)) return col;
        }
      }
      return null;
    });
  }

  function get(row, col) {
    if (col === null || col === undefined || col >= row.length) return null;
    const v = row[col];
    return v === undefined ? null : v;
  }

  // openpyxl reads these workbooks' HYPERLINK() cells as a literal cached string
  // ("HYPERLINK is not implemented. linkLocation=..., friendlyName=...") because
  // whatever tool wrote the file left that placeholder as the cached formula
  // result. SheetJS instead surfaces the cell as an error type (t:"e") with no
  // .v — the formula text and display string survive on .f/.w. Reconstruct the
  // same placeholder string from .f here so parseHyperlinkCell works unchanged.
  function cellValue(cell) {
    if (!cell) return null;
    if (cell.f) {
      const m = cell.f.match(/HYPERLINK\(\s*"([^"]*)"\s*(?:,\s*"([^"]*)")?\s*\)/i);
      if (m) {
        const url = m[1];
        const label = m[2] !== undefined ? m[2] : url;
        return `HYPERLINK is not implemented. linkLocation=${url}, friendlyName=${label}`;
      }
    }
    if (cell.v !== undefined) return cell.v;
    if (cell.w !== undefined) return cell.w;
    return null;
  }

  function sheetRows(ws) {
    if (!ws["!ref"]) return [];
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const rows = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        row.push(cellValue(ws[XLSX.utils.encode_cell({ r, c })]));
      }
      rows.push(row);
    }
    return rows;
  }

  function buildSnapshot(workbook, fileDate) {
    const sheetNames = workbook.SheetNames;
    const mainRows = sheetRows(workbook.Sheets[sheetNames[0]]);
    const catRows = sheetRows(workbook.Sheets[sheetNames[1]]);
    const extrasRows = sheetRows(workbook.Sheets[sheetNames[3]]);

    const title = mainRows[0] ? mainRows[0][0] : "";
    const reportDate = extractDate(title, fileDate);

    const header = mainRows[1] || [];
    const sites = [];
    for (let i = 1; i < header.length; i++) {
      const [name, url] = parseHyperlinkCell(header[i]);
      if (!url) continue;
      sites.push({ id: SITE_ID_MAP[name] || String(name).toLowerCase().replace(/\W+/g, "-"), name, uwant_url: url });
    }
    const siteNames = sites.map((s) => s.name);

    const [cSite, cModel, , cPrice, cRegular, cClub, cStock, cKind, cUrl, cVerify, cNote, cChecked] = headerIndex(
      catRows[0] || [],
      [["אתר"], ["דגם ממופה"], ["שם מוצר"], ["מחיר נוכחי"], ["מחיר רגיל"],
       ["מחיר מועדון"], ["מלאי"], ["סוג"], ["קישור"], ["אימות"], ["הערה"], ["מועד בדיקה"]]
    );
    const cName = headerIndex(catRows[0] || [], [["שם מוצר"]])[0];

    const catalog = [];
    for (let r = 1; r < catRows.length; r++) {
      const row = catRows[r];
      if (!row || get(row, cSite) == null) continue;
      const [, url] = parseHyperlinkCell(get(row, cUrl));
      const modelN = normModel(get(row, cModel));
      if (EXCLUDED_MODELS.has(modelN)) continue;
      catalog.push({
        site: get(row, cSite), model: modelN, name: get(row, cName),
        price: toNumber(get(row, cPrice)), regular: toNumber(get(row, cRegular)),
        club: toNumber(get(row, cClub)), stock: get(row, cStock), kind: get(row, cKind),
        url, verify: get(row, cVerify), note: get(row, cNote), checked: get(row, cChecked),
      });
    }

    const matrix = {};
    ACTIVE_MODELS.forEach((m) => {
      matrix[m] = {};
      siteNames.forEach((s) => {
        const candidates = catalog.filter((c) => c.model === m && c.site === s && c.kind === "מכשיר" && c.price != null);
        if (!candidates.length) { matrix[m][s] = null; return; }
        const best = candidates.reduce((a, b) => (b.price < a.price ? b : a));
        matrix[m][s] = { price: best.price, regular: best.regular, club: best.club, stock: best.stock, url: best.url, note: best.note };
      });
    });

    const stats = siteNames.map((s) => {
      const siteCatalog = catalog.filter((c) => c.site === s);
      const variants = siteCatalog.filter((c) => c.kind === "מכשיר").length;
      const totalRecords = siteCatalog.length;
      const priceCoverage = new Set(
        siteCatalog.filter((c) => c.kind === "מכשיר" && ACTIVE_MODELS.includes(c.model) && c.price != null).map((c) => c.model)
      ).size;
      return {
        site: s, variants, total_records: totalRecords, price_coverage: priceCoverage,
        coverage_pct: Math.round((100 * priceCoverage) / ACTIVE_MODELS.length), note: null,
      };
    });

    const [eSite, eModel, eName, ePrice, eStock, eKind, eUrl] = headerIndex(
      extrasRows[0] || [], [["אתר"], ["דגם"], ["שם מדויק"], ["מחיר"], ["מלאי"], ["סוג"], ["קישור"]]
    );
    const extras = [];
    const seenUrls = new Set();
    for (let r = 1; r < extrasRows.length; r++) {
      const row = extrasRows[r];
      if (!row || get(row, eSite) == null) continue;
      const [, url] = parseHyperlinkCell(get(row, eUrl));
      const modelN = normModel(get(row, eModel));
      if (EXCLUDED_MODELS.has(modelN)) continue;
      extras.push({ site: get(row, eSite), model: modelN, name: get(row, eName), price: toNumber(get(row, ePrice)), stock: get(row, eStock), kind: get(row, eKind), url });
      if (url) seenUrls.add(url);
    }
    catalog.forEach((c) => {
      if ((!ACTIVE_MODELS.includes(c.model) || c.kind === "אביזר") && c.url && !seenUrls.has(c.url)) {
        extras.push({ site: c.site, model: c.model, name: c.name, price: c.price, stock: c.stock, kind: c.kind, url: c.url });
        seenUrls.add(c.url);
      }
    });

    return { title, checked_at: title, report_date: reportDate, sites, models: ACTIVE_MODELS, matrix, stats, extras };
  }

  function mergeHistory(existingHistory, snapshot) {
    const history = existingHistory || {};
    Object.entries(snapshot.matrix).forEach(([model, perSite]) => {
      Object.entries(perSite).forEach(([site, cell]) => {
        if (!cell || cell.price == null) return;
        if (!history[model]) history[model] = {};
        let series = history[model][site] || [];
        series = series.filter((p) => p.date !== snapshot.report_date);
        series.push({ date: snapshot.report_date, price: cell.price });
        series.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
        history[model][site] = series;
      });
    });
    return history;
  }

  function arrayBufferToBase64(buf) {
    let binary = "";
    const bytes = new Uint8Array(buf);
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }

  function utf8ToBase64(str) {
    return btoa(unescape(encodeURIComponent(str)));
  }

  async function githubRequest(path, token, options = {}) {
    const res = await fetch(`https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        ...(options.headers || {}),
      },
    });
    return res;
  }

  async function getFileSha(path, token) {
    const res = await githubRequest(`${path}?ref=${GITHUB_BRANCH}`, token);
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`GitHub read failed for ${path}: ${res.status}`);
    const json = await res.json();
    return json.sha;
  }

  async function putFile(path, base64Content, message, token) {
    const sha = await getFileSha(path, token);
    const body = { message, content: base64Content, branch: GITHUB_BRANCH };
    if (sha) body.sha = sha;
    const res = await githubRequest(path, token, { method: "PUT", body: JSON.stringify(body) });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`GitHub write failed for ${path}: ${res.status} ${text}`);
    }
  }

  function getToken() {
    let token = localStorage.getItem(TOKEN_KEY);
    if (!token) {
      token = prompt(
        "כדי לעדכן את הדוח נדרש טוקן גישה אישי לגיטהאב (Fine-grained Personal Access Token) עם הרשאת Contents: Read and write על ה-repo הזה בלבד.\n\n" +
        "ליצירה: github.com/settings/personal-access-tokens/new → Repository access: Only select repositories → " +
        `${GITHUB_OWNER}/${GITHUB_REPO} → Repository permissions → Contents: Read and write.\n\n` +
        "הטוקן נשמר רק בדפדפן שלך (לא נשלח לשום מקום מלבד GitHub).\n\nהדבק כאן:"
      );
      if (token) localStorage.setItem(TOKEN_KEY, token.trim());
    }
    return token ? token.trim() : null;
  }

  async function handleFile(file, statusEl) {
    const setStatus = (msg, isError) => {
      statusEl.textContent = msg;
      statusEl.className = "update-status" + (isError ? " error" : "");
    };

    try {
      const token = getToken();
      if (!token) { setStatus("בוטל — לא הוזן טוקן.", true); return; }

      setStatus("קורא את הקובץ…");
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      setStatus("מעבד נתונים…");
      const snapshot = buildSnapshot(wb, new Date(file.lastModified));

      setStatus("טוען את הנתונים הקיימים…");
      const existingRes = await githubRequest(`data/data.json?ref=${GITHUB_BRANCH}`, token);
      let existing = {};
      if (existingRes.ok) {
        const json = await existingRes.json();
        existing = JSON.parse(decodeURIComponent(escape(atob(json.content.replace(/\n/g, "")))));
      }

      const history = mergeHistory(existing.history, snapshot);
      const out = {
        generated_at: new Date().toISOString(),
        source_file: file.name,
        ...snapshot,
        history,
      };

      setStatus("שומר את הדוח המעודכן…");
      await putFile("data/data.json", utf8ToBase64(JSON.stringify(out, null, 1)), `עדכון דוח: ${snapshot.report_date}`, token);

      setStatus("שומר עותק של קובץ ה-XLSX…");
      await putFile("data/UWANT_price_report_latest.xlsx", arrayBufferToBase64(buf), `עדכון קובץ דוח: ${snapshot.report_date}`, token);

      setStatus("✓ עודכן בהצלחה! האתר יתעדכן עבור כולם תוך כדקה (GitHub Pages בונה מחדש).");
    } catch (err) {
      console.error(err);
      setStatus("שגיאה: " + err.message, true);
    }
  }

  function init() {
    const btn = document.getElementById("update-report");
    const input = document.getElementById("update-file-input");
    const statusEl = document.getElementById("update-status");
    if (!btn || !input) return;

    btn.addEventListener("click", () => input.click());
    input.addEventListener("change", () => {
      if (input.files && input.files[0]) {
        handleFile(input.files[0], statusEl);
        input.value = "";
      }
    });
  }

  document.addEventListener("DOMContentLoaded", init);

  window.UWANTUpdate = { buildSnapshot, mergeHistory };
})();
