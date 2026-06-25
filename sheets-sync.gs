/**
 * Strategy stress-test data sync
 * Sheets → strategy.json → GitHub Pages
 *
 * Sheets structure required:
 *   Sheet "balance_sheet"      — A=key, B=value (see KEYS below)
 *   Sheet "convertible_ladder" — headers in row 1, one tranche per row
 *   Sheet "preferred_stack"    — headers in row 1, one series per row
 *
 * Script Properties (File → Project properties → Script properties):
 *   GITHUB_TOKEN   — Personal Access Token, scope: repo
 *   GITHUB_OWNER   — e.g. "pate"
 *   GITHUB_REPO    — e.g. "portfolio-dashboard"
 *   GITHUB_PATH    — e.g. "data/strategy.json"
 *   GITHUB_BRANCH  — e.g. "main"
 */

// ── Config ──────────────────────────────────────────────────────────────────

const CONFIG = {
  schemaVersion: "1.0",
  baseDate:      "2026-06-01",
  source:        "Strategy 10-K/10-Q + 8-K. Verify ladder from latest filings.",
};

// Keys expected in balance_sheet tab (col A → col B)
const BS_KEYS = [
  "btc_holdings",
  "avg_cost_basis_usd",
  "cash_reserve_usd_m",
  "shares_diluted_m",
  "annual_interest_m",
];
const TOP_KEYS = [
  "preferred_annual_div_m",
  "total_debt_m",
  "total_preferred_m",
];

// ── Main entry point ─────────────────────────────────────────────────────────

function syncToGitHub() {
  const json = buildJSON();
  pushToGitHub(JSON.stringify(json, null, 2));
  Logger.log("✅ Pushed strategy.json — last_updated: " + json.meta.last_updated);
}

// ── Build JSON ───────────────────────────────────────────────────────────────

function buildJSON() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- balance_sheet tab ---
  const bsSheet = ss.getSheetByName("balance_sheet");
  if (!bsSheet) throw new Error("Sheet 'balance_sheet' not found");

  const bsData = bsSheet.getDataRange().getValues(); // [[key, value], ...]
  const bsMap  = {};
  bsData.forEach(([k, v]) => { if (k) bsMap[String(k).trim()] = v; });

  const balance_sheet = {};
  BS_KEYS.forEach(k => {
    balance_sheet[k] = typeof bsMap[k] === "number" ? bsMap[k] : parseFloat(bsMap[k]) || 0;
  });

  const topLevel = {};
  TOP_KEYS.forEach(k => {
    topLevel[k] = typeof bsMap[k] === "number" ? bsMap[k] : parseFloat(bsMap[k]) || 0;
  });

  // --- convertible_ladder tab ---
  const ladSheet = ss.getSheetByName("convertible_ladder");
  if (!ladSheet) throw new Error("Sheet 'convertible_ladder' not found");
  const ladRows = sheetToObjects(ladSheet);
  const convertible_ladder = ladRows.map(r => ({
    id:               String(r.id || ""),
    label:            String(r.label || ""),
    months_from_base: Number(r.months_from_base) || 0,
    principal_m:      Number(r.principal_m)      || 0,
    coupon_pct:       Number(r.coupon_pct)        || 0,
    conversion_price: Number(r.conversion_price)  || 0,
    put_date_label:   String(r.put_date_label || ""),
    verified:         r.verified === true || r.verified === "TRUE",
  }));

  // --- preferred_stack tab ---
  const prefSheet = ss.getSheetByName("preferred_stack");
  if (!prefSheet) throw new Error("Sheet 'preferred_stack' not found");
  const prefRows = sheetToObjects(prefSheet);
  const preferred_stack = prefRows.map(r => ({
    id:                 String(r.id || ""),
    ticker:             String(r.ticker || ""),
    nominal_m:          Number(r.nominal_m)          || 0,
    dividend_rate_pct:  Number(r.dividend_rate_pct)   || 0,
    cumulative:         r.cumulative === true || r.cumulative === "TRUE",
    senior:             Number(r.senior) || 0,
  }));

  return {
    meta: {
      schema_version: CONFIG.schemaVersion,
      last_updated:   Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd"),
      source:         CONFIG.source,
      base_date:      CONFIG.baseDate,
      maintainer:     "Apps Script auto-push",
    },
    balance_sheet,
    ...topLevel,
    convertible_ladder,
    preferred_stack,
  };
}

// ── GitHub push ──────────────────────────────────────────────────────────────

function pushToGitHub(content) {
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty("GITHUB_TOKEN");
  const owner  = props.getProperty("GITHUB_OWNER");
  const repo   = props.getProperty("GITHUB_REPO");
  const path   = props.getProperty("GITHUB_PATH")  || "data/strategy.json";
  const branch = props.getProperty("GITHUB_BRANCH") || "main";

  if (!token || !owner || !repo) throw new Error("Missing GITHUB_TOKEN / GITHUB_OWNER / GITHUB_REPO in Script Properties");

  const apiBase = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
  const headers = {
    "Authorization": "token " + token,
    "Accept":        "application/vnd.github.v3+json",
    "User-Agent":    "Apps-Script-Strategy-Sync",
  };

  // 1. Get current SHA (needed for update)
  let sha = null;
  try {
    const getRes = UrlFetchApp.fetch(apiBase + "?ref=" + branch, { headers, muteHttpExceptions: true });
    if (getRes.getResponseCode() === 200) {
      sha = JSON.parse(getRes.getContentText()).sha;
    }
  } catch(e) { /* new file — no SHA needed */ }

  // 2. PUT new content
  const body = { message: "chore: update strategy.json " + new Date().toISOString(), content: Utilities.base64Encode(content), branch };
  if (sha) body.sha = sha;

  const putRes = UrlFetchApp.fetch(apiBase, {
    method:  "PUT",
    headers: { ...headers, "Content-Type": "application/json" },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });

  const code = putRes.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error("GitHub API error " + code + ": " + putRes.getContentText());
  }
}

// ── Utility ──────────────────────────────────────────────────────────────────

/** Converts a sheet with header row to array of plain objects */
function sheetToObjects(sheet) {
  const [headers, ...rows] = sheet.getDataRange().getValues();
  const keys = headers.map(h => String(h).trim());
  return rows
    .filter(r => r.some(v => v !== "" && v !== null))
    .map(r => {
      const obj = {};
      keys.forEach((k, i) => { obj[k] = r[i]; });
      return obj;
    });
}

// ── Optional: manual trigger via menu ────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Strategy Sync")
    .addItem("Push strategy.json → GitHub", "syncToGitHub")
    .addToUi();
}
