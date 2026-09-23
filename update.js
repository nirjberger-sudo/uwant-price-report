(() => {
  "use strict";

  // The button doesn't scrape anything itself — a browser can't reach 10
  // other domains (CORS), and a naive scrape got several sites wrong the
  // first time (see scripts/live_update.py's comments). Instead it triggers
  // scripts/live_update.py running inside GitHub Actions, which does the
  // real work server-side and pushes data/data.json when done.

  const GITHUB_OWNER = "nirjberger-sudo";
  const GITHUB_REPO = "uwant-price-report";
  const WORKFLOW_FILE = "update-prices.yml";
  const TOKEN_KEY = "uwant_gh_actions_token";

  function getToken() {
    let token = localStorage.getItem(TOKEN_KEY);
    if (!token) {
      token = prompt(
        "בפעם הראשונה בלבד: כדי להפעיל את העדכון נדרש טוקן גישה מוגבל לגיטהאב " +
        "(מותר לו רק להריץ את תהליך העדכון — הוא לא יכול לקרוא או לשנות שום קובץ).\n\n" +
        "ליצירה: github.com/settings/personal-access-tokens/new →\n" +
        "Repository access: Only select repositories → " +
        `${GITHUB_OWNER}/${GITHUB_REPO} →\n` +
        "Repository permissions → Actions: Read and write (זה הכל, לא לסמן שום דבר אחר).\n\n" +
        "הטוקן נשמר רק בדפדפן שלך.\n\nהדבק כאן:"
      );
      if (token) localStorage.setItem(TOKEN_KEY, token.trim());
    }
    return token ? token.trim() : null;
  }

  function setStatus(el, msg, cls) {
    el.textContent = msg;
    el.className = "update-status" + (cls ? " " + cls : "");
  }

  async function triggerUpdate(statusEl) {
    const token = getToken();
    if (!token) {
      setStatus(statusEl, "בוטל — לא הוזן טוקן.", "error");
      return;
    }
    setStatus(statusEl, "מפעיל בדיקה חיה בענן...");
    try {
      const res = await fetch(
        `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/actions/workflows/${WORKFLOW_FILE}/dispatches`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
          },
          body: JSON.stringify({ ref: "main" }),
        }
      );
      if (res.status === 204) {
        setStatus(
          statusEl,
          "✓ העדכון החל! בודק כרגע את כל האתרים בענן — האתר יתעדכן לבד תוך כ-3–5 דקות (אין צורך לחכות כאן, אפשר לסגור את הדף).",
          "success"
        );
      } else if (res.status === 401 || res.status === 403) {
        localStorage.removeItem(TOKEN_KEY);
        setStatus(statusEl, "הטוקן לא תקין או שאין לו הרשאה. נסו שוב ותידרשו להזין טוקן חדש.", "error");
      } else {
        const text = await res.text();
        setStatus(statusEl, `שגיאה (${res.status}): ${text}`, "error");
      }
    } catch (err) {
      setStatus(statusEl, "שגיאת רשת: " + err.message, "error");
    }
  }

  function init() {
    const btn = document.getElementById("update-report");
    const statusEl = document.getElementById("update-status");
    if (!btn || !statusEl) return;
    btn.addEventListener("click", () => triggerUpdate(statusEl));
  }

  document.addEventListener("DOMContentLoaded", init);
})();
