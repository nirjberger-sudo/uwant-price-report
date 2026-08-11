(() => {
  "use strict";

  const UWANT_SITE = "UWANT IL";
  const IN_STOCK_WORDS = ["במלאי", "זמין"];
  const RISK_UNDERCUT_THRESHOLD = 3; // flag a site once it holds the cheapest-in-row (red) price on MORE than this many models

  // Product-line grouping for the default view, with a thick divider between groups.
  const CATEGORY_GROUPS = [
    ["U400", "T300", "U300 V25"],
    ["DX800", "D700", "D600", "D500"],
    ["V800", "V600", "V500"],
    ["M600", "M500", "M400"],
    ["Y200 Steam", "Y100 Steam", "Y200", "Y100"],
  ];
  const MODEL_TO_GROUP = {};
  CATEGORY_GROUPS.forEach((group, i) => group.forEach((m) => (MODEL_TO_GROUP[m] = i)));
  const CATEGORY_ORDER = CATEGORY_GROUPS.flat();
  const SITE_COLORS = {
    "KSP": "#e63946", "Walla Shops": "#f4a261", "NetoNeto": "#2a9d8f", "CWC": "#457b9d",
    "LastPrice": "#8338ec", "ALM": "#ff006e", "Traklin": "#3a86ff", "City Deal": "#ffbe0b",
    "Multi Store": "#06a77d", "UWANT IL": "#153b5b",
  };

  const state = { data: null, search: "", sites: new Set(), stock: "", sort: "default", riskSiteNames: new Set() };

  function getVisibleSites(data) {
    if (state.sites.size === 0) return data.sites;
    return data.sites.filter((s) => s.name === UWANT_SITE || state.sites.has(s.name));
  }

  const $ = (sel) => document.querySelector(sel);

  function isInStock(stock) {
    if (!stock) return false;
    return IN_STOCK_WORDS.some((w) => stock.includes(w));
  }

  function fmtPrice(p) {
    return p == null ? "" : "₪" + Math.round(p).toLocaleString("he-IL");
  }

  async function loadData() {
    if (window.__UWANT_DATA__) return window.__UWANT_DATA__;
    const res = await fetch("./data/data.json", { cache: "no-store" });
    if (!res.ok) throw new Error("data.json load failed: " + res.status);
    return res.json();
  }

  // ---------------- header + stats ----------------

  function renderHeader(data) {
    $("#updated-at").textContent = data.checked_at || data.report_date || "";
  }

  function renderStats(data, riskSiteNames) {
    const strip = $("#stats-strip");
    strip.innerHTML = "";
    data.stats.forEach((s) => {
      const site = data.sites.find((x) => x.name === s.site);
      const card = document.createElement("div");
      card.className = "stat-card" + (riskSiteNames.has(s.site) ? " is-risk" : "");
      card.innerHTML = `
        <div class="stat-site">
          <a href="${site ? site.uwant_url : "#"}" target="_blank" rel="noopener">${s.site}</a>
          <span>${s.coverage_pct}%</span>
        </div>
        <div class="stat-bar"><div class="stat-bar-fill" style="width:${s.coverage_pct}%"></div></div>
        <div class="stat-nums"><span>${s.price_coverage}/${data.models.length} דגמים</span><span>${s.total_records} רשומות</span></div>
      `;
      strip.appendChild(card);
    });
  }

  // ---------------- controls ----------------

  function syncSiteChips() {
    document.querySelectorAll(".site-chip").forEach((chip) => {
      const active = state.sites.has(chip.dataset.site);
      chip.classList.toggle("active", active);
      chip.setAttribute("aria-pressed", String(active));
    });
  }

  function renderControls(data) {
    const chipGroup = $("#site-chip-group");
    chipGroup.innerHTML = "";
    data.sites
      .filter((s) => s.name !== UWANT_SITE)
      .forEach((s) => {
        const chip = document.createElement("button");
        chip.type = "button";
        chip.className = "site-chip";
        chip.dataset.site = s.name;
        chip.textContent = s.name;
        chip.setAttribute("aria-pressed", "false");
        chip.addEventListener("click", () => {
          if (state.sites.has(s.name)) state.sites.delete(s.name);
          else state.sites.add(s.name);
          syncSiteChips();
          renderTable();
        });
        chipGroup.appendChild(chip);
      });

    $("#search-input").addEventListener("input", (e) => {
      state.search = e.target.value.trim().toLowerCase();
      renderTable();
    });
    $("#stock-filter").addEventListener("change", (e) => {
      state.stock = e.target.value;
      renderTable();
    });
    $("#sort-select").addEventListener("change", (e) => {
      state.sort = e.target.value;
      renderTable();
    });
    $("#download-xlsx").addEventListener("click", () => {
      const a = document.createElement("a");
      a.href = "./data/UWANT_price_report_latest.xlsx";
      a.download = "UWANT_price_report.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
    });
  }

  // ---------------- table ----------------

  function computeRow(data, model) {
    const perSite = data.sites.map((s) => ({ site: s, cell: data.matrix[model][s.name] || null }));
    const prices = perSite.filter((x) => x.cell && x.cell.price != null).map((x) => x.cell.price);
    const min = prices.length ? Math.min(...prices) : null;
    const max = prices.length ? Math.max(...prices) : null;
    const allEqual = prices.length > 0 && prices.every((p) => p === prices[0]);
    const uwantCell = data.matrix[model][UWANT_SITE];
    const uwantIsCheapest = uwantCell && uwantCell.price != null && min != null && uwantCell.price === min;
    return { perSite, min, max, allEqual, uwantCell, uwantIsCheapest, hasUwantPrice: !!(uwantCell && uwantCell.price != null) };
  }

  // A site is "at risk" when it holds the cheapest-in-row price (the red-highlighted
  // cell) on MORE than RISK_UNDERCUT_THRESHOLD active models — i.e. it isn't one or
  // two isolated cheap listings, but a pattern of consistently being the single
  // cheapest place to buy UWANT, which undermines UWANT's own pricing.
  function computeRiskSites(data) {
    const counts = {};
    data.sites.forEach((s) => (counts[s.name] = 0));
    data.models.forEach((model) => {
      const row = computeRow(data, model);
      if (row.allEqual || row.min == null) return;
      row.perSite.forEach(({ site, cell }) => {
        if (site.name === UWANT_SITE) return;
        if (cell && cell.price === row.min) counts[site.name]++;
      });
    });
    return data.sites
      .map((s) => ({ site: s.name, count: counts[s.name] || 0 }))
      .filter((x) => x.count > RISK_UNDERCUT_THRESHOLD)
      .sort((a, b) => b.count - a.count);
  }

  function renderDashboard(data) {
    const riskSites = computeRiskSites(data);
    const uwantCheapestCount = data.models.filter((m) => computeRow(data, m).uwantIsCheapest).length;
    const uwantCheapestPct = Math.round((100 * uwantCheapestCount) / data.models.length);

    $("#dashboard-kpis").innerHTML = `
      <div class="kpi-card">
        <div class="kpi-value">${data.models.length}</div>
        <div class="kpi-label">דגמים במעקב</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-value">${data.sites.length}</div>
        <div class="kpi-label">אתרים מושווים</div>
      </div>
      <div class="kpi-card kpi-good">
        <div class="kpi-value">${uwantCheapestPct}%</div>
        <div class="kpi-label">דגמים בהם UWANT הזול ביותר (${uwantCheapestCount}/${data.models.length})</div>
      </div>
      <div class="kpi-card ${riskSites.length ? "kpi-risk" : ""}">
        <div class="kpi-value">${riskSites.length}</div>
        <div class="kpi-label">אתרים לתשומת לב (מוכרים יותר מדי דגמים במחיר הזול ביותר)</div>
      </div>
    `;

    const panel = $("#risk-panel");
    if (!riskSites.length) {
      panel.innerHTML = `<p class="risk-empty">✓ אין כרגע אתר שמחזיק את המחיר הזול ביותר במספר חריג של דגמים.</p>`;
    } else {
      const chips = riskSites
        .map((r) => `<button type="button" class="risk-chip" data-site="${r.site}">⚠️ ${r.site} — הזול ביותר ב-${r.count} דגמים</button>`)
        .join("");
      panel.innerHTML = `
        <div class="risk-panel">
          <span class="risk-title">⚠️ אזהרה — האתרים הבאים מוכרים יותר מדי דגמי UWANT במחיר הזול ביותר בהשוואה:</span>
          ${chips}
        </div>
      `;
      panel.querySelectorAll(".risk-chip").forEach((btn) => {
        btn.addEventListener("click", () => {
          state.sites = new Set([btn.dataset.site]);
          syncSiteChips();
          renderTable();
          $("#compare-table").scrollIntoView({ behavior: "smooth", block: "start" });
        });
      });
    }

    return riskSites;
  }

  function renderThead(data, visibleSites) {
    const thead = $("#compare-thead");
    const th = document.createElement("tr");
    th.innerHTML =
      `<th>דגם</th>` +
      visibleSites
        .map((s) => {
          const isRisk = state.riskSiteNames.has(s.name);
          const warn = isRisk ? `⚠️ ` : "";
          const title = isRisk ? `⚠️ מוכר יותר מדי דגמי UWANT במחיר הזול ביותר בהשוואה` : `עמוד UWANT באתר ${s.name}`;
          return `<th class="${isRisk ? "is-risk" : ""}" title="${title}"><a href="${s.uwant_url}" target="_blank" rel="noopener">${warn}${s.name}</a></th>`;
        })
        .join("");
    thead.innerHTML = "";
    thead.appendChild(th);
  }

  function filterAndSortModels(data) {
    let models = CATEGORY_ORDER.filter((m) => data.models.includes(m));

    if (state.search) {
      models = models.filter((m) => m.toLowerCase().includes(state.search));
    }
    if (state.stock === "instock") {
      models = models.filter((m) => {
        const c = data.matrix[m][UWANT_SITE];
        return c && isInStock(c.stock);
      });
    } else if (state.stock === "warn") {
      models = models.filter((m) => !computeRow(data, m).uwantIsCheapest);
    }

    const rowOf = (m) => computeRow(data, m);
    if (state.sort === "name") {
      models.sort((a, b) => a.localeCompare(b, "he"));
    } else if (state.sort === "cheapest") {
      models.sort((a, b) => {
        const ra = rowOf(a).min, rb = rowOf(b).min;
        if (ra == null) return 1;
        if (rb == null) return -1;
        return ra - rb;
      });
    } else if (state.sort === "priciest") {
      models.sort((a, b) => {
        const ra = rowOf(a).max, rb = rowOf(b).max;
        if (ra == null) return 1;
        if (rb == null) return -1;
        return rb - ra;
      });
    } else if (state.sort === "warn-first") {
      models.sort((a, b) => Number(!rowOf(a).uwantIsCheapest) - Number(!rowOf(b).uwantIsCheapest) ? -1 : 0);
      models.sort((a, b) => {
        const wa = !rowOf(a).uwantIsCheapest, wb = !rowOf(b).uwantIsCheapest;
        return wa === wb ? 0 : wa ? -1 : 1;
      });
    }
    return models;
  }

  function cellHTML(cell, min, max, allEqual) {
    if (!cell || cell.price == null) {
      return `<span class="cell-na">—</span>`;
    }
    let priceClass = "";
    if (!allEqual) {
      if (cell.price === min) priceClass = "min";
      else if (cell.price === max) priceClass = "max";
    }
    const dotClass = isInStock(cell.stock) ? "in" : "out";
    const dot = `<i class="stock-dot ${dotClass}" title="${cell.stock || ""}"></i>`;
    const price = `<span class="cell-price ${priceClass}">${fmtPrice(cell.price)}</span>`;
    if (cell.url) {
      return `<a class="cell-link" href="${cell.url}" target="_blank" rel="noopener" title="${cell.stock || ""} · לצפייה במוצר">${price}${dot}</a>`;
    }
    return `<span class="cell-link">${price}${dot}</span>`;
  }

  function renderTable() {
    const data = state.data;
    const models = filterAndSortModels(data);
    const visibleSites = getVisibleSites(data);
    renderThead(data, visibleSites);

    const tbody = $("#compare-tbody");
    tbody.innerHTML = "";

    models.forEach((model, i) => {
      // Colors always reflect the full market (all 10 sites), never just the
      // currently visible columns — filtering changes what you see, not the prices.
      const fullRow = computeRow(data, model);
      const { min, max, allEqual } = fullRow;
      const visibleCells = visibleSites.map((s) => ({ site: s, cell: data.matrix[model][s.name] || null }));

      const tr = document.createElement("tr");

      if (state.sort === "default") {
        const nextModel = models[i + 1];
        const isGroupEnd = nextModel !== undefined && MODEL_TO_GROUP[nextModel] !== MODEL_TO_GROUP[model];
        if (isGroupEnd) tr.classList.add("group-end");
      }

      const nameTd = document.createElement("td");
      nameTd.innerHTML = `<span class="model-name">UWANT ${model}</span>`;
      tr.appendChild(nameTd);

      visibleCells.forEach(({ cell }) => {
        const td = document.createElement("td");
        if (!cell || cell.price == null) td.classList.add("is-na");
        td.innerHTML = `<span class="cell">${cellHTML(cell, min, max, allEqual)}</span>`;
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    $("#result-count").textContent = `${models.length} מתוך ${data.models.length} דגמים`;
    $("#empty-state").hidden = models.length !== 0;
    $("#compare-table").style.display = models.length === 0 ? "none" : "";
  }

  // ---------------- extras ----------------

  function renderExtras(data) {
    const tbody = $("#extras-tbody");
    tbody.innerHTML = "";
    data.extras.forEach((x) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${x.site}</td>
        <td>${x.model || ""}</td>
        <td style="text-align:start;max-width:420px;white-space:normal">${x.name || ""}</td>
        <td>${fmtPrice(x.price)}</td>
        <td>${x.stock || ""}</td>
        <td>${x.url ? `<a class="cell-link" href="${x.url}" target="_blank" rel="noopener">למוצר ↗</a>` : ""}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  // ---------------- history chart ----------------

  function renderHistoryControls(data) {
    const select = $("#history-model-select");
    const modelsWithHistory = data.models.filter((m) => data.history[m] && Object.keys(data.history[m]).length);
    modelsWithHistory.forEach((m) => {
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = "UWANT " + m;
      select.appendChild(opt);
    });
    select.addEventListener("change", () => renderHistoryChart(data, select.value));
    if (modelsWithHistory.length) {
      select.value = modelsWithHistory.includes("U300 V25") ? "U300 V25" : modelsWithHistory[0];
      renderHistoryChart(data, select.value);
    }
  }

  function renderHistoryChart(data, model) {
    const wrap = $("#history-chart-wrap");
    const seriesBySite = data.history[model] || {};
    const siteNames = Object.keys(seriesBySite).filter((s) => seriesBySite[s].length > 0);
    if (!siteNames.length) {
      wrap.innerHTML = `<p class="hint-text">אין נתוני היסטוריה זמינים לדגם זה עדיין.</p>`;
      return;
    }

    const allDates = Array.from(new Set(siteNames.flatMap((s) => seriesBySite[s].map((p) => p.date)))).sort();
    const allPrices = siteNames.flatMap((s) => seriesBySite[s].map((p) => p.price));
    const minP = Math.min(...allPrices), maxP = Math.max(...allPrices);
    const padP = Math.max((maxP - minP) * 0.15, 10);
    const yMin = Math.max(0, minP - padP), yMax = maxP + padP;

    const W = 960, H = 340, ML = 60, MR = 20, MT = 20, MB = 40;
    const plotW = W - ML - MR, plotH = H - MT - MB;
    const xFor = (date) => ML + (allDates.length <= 1 ? plotW / 2 : (allDates.indexOf(date) / (allDates.length - 1)) * plotW);
    const yFor = (price) => MT + plotH - ((price - yMin) / (yMax - yMin || 1)) * plotH;

    let svg = `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="גרף היסטוריית מחירים עבור ${model}">`;

    // gridlines + y labels
    const gridSteps = 4;
    for (let i = 0; i <= gridSteps; i++) {
      const y = MT + (plotH / gridSteps) * i;
      const val = Math.round(yMax - ((yMax - yMin) / gridSteps) * i);
      svg += `<line x1="${ML}" y1="${y}" x2="${W - MR}" y2="${y}" stroke="currentColor" stroke-opacity="0.12" />`;
      svg += `<text x="${ML - 8}" y="${y + 4}" text-anchor="end" font-size="11" fill="currentColor" opacity="0.6">₪${val}</text>`;
    }
    // x labels
    allDates.forEach((d) => {
      const x = xFor(d);
      svg += `<text x="${x}" y="${H - 12}" text-anchor="middle" font-size="11" fill="currentColor" opacity="0.6">${d}</text>`;
    });

    siteNames.forEach((site) => {
      const color = SITE_COLORS[site] || "#888";
      const pts = seriesBySite[site].filter((p) => allDates.includes(p.date));
      if (!pts.length) return;
      const isUwant = site === UWANT_SITE;
      const path = pts.map((p, i) => `${i === 0 ? "M" : "L"}${xFor(p.date).toFixed(1)},${yFor(p.price).toFixed(1)}`).join(" ");
      svg += `<path d="${path}" fill="none" stroke="${color}" stroke-width="${isUwant ? 3 : 2}" stroke-linecap="round" stroke-linejoin="round" opacity="${isUwant ? 1 : 0.85}"/>`;
      pts.forEach((p) => {
        svg += `<circle cx="${xFor(p.date).toFixed(1)}" cy="${yFor(p.price).toFixed(1)}" r="${isUwant ? 4.5 : 3.5}" fill="${color}"><title>${site} · ${p.date} · ${fmtPrice(p.price)}</title></circle>`;
      });
    });

    svg += `</svg>`;

    const legend = siteNames
      .map((s) => `<span class="legend-item"><span class="swatch" style="background:${SITE_COLORS[s] || "#888"}"></span>${s}</span>`)
      .join("");

    wrap.innerHTML = svg + `<div class="chart-legend">${legend}</div>`;
  }

  // ---------------- init ----------------

  async function init() {
    try {
      const data = await loadData();
      state.data = data;
      renderHeader(data);
      const riskSites = renderDashboard(data);
      state.riskSiteNames = new Set(riskSites.map((r) => r.site));
      renderStats(data, state.riskSiteNames);
      renderControls(data);
      renderTable();
      renderExtras(data);
      renderHistoryControls(data);
    } catch (err) {
      console.error(err);
      $("#compare-tbody").innerHTML = `<tr><td colspan="12">שגיאה בטעינת הנתונים. נסו לרענן את הדף.</td></tr>`;
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
