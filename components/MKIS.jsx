import React, { useState, useMemo, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import MKIS_SOURCE from "./MKIS.jsx?raw";
import { supabase } from "@/integrations/supabase/client";
// NOTE: Excel export uses the "xlsx" package (already available in Claude.ai artifacts;
// for a standalone project run: npm install xlsx).
// Word export below is dependency-free (HTML-to-.doc), so it works everywhere, including artifacts.
// ─── CONSTANTS ───────────────────────────────────────────────────────────────
const ALL_CLASSES = ["P1","P2","P3","P4","P5","P6","P7"];
const LOWER_CLASSES = ["P1","P2","P3"];
const UPPER_CLASSES = ["P4","P5","P6","P7"];
const LOWER_SUBJECTS = ["MATHS","LIT I","LIT II","ENG","RE","READING","WRITING"];
const LOWER_MONTHLY_SUBJECTS = ["MATHS","LIT I","LIT II","ENG","RE","READING","WRITING"];
const LOWER_SUBJECT_MAX = { READING: 50, WRITING: 50 };
const lowerSubjectMax = (sub) => LOWER_SUBJECT_MAX[sub] || 100;
const UPPER_SUBJECTS = ["ENG","MATH","SST","SCI"];
const MONTHLY_SUBJECTS = ["ENG","MATH","SST","SCI"];
const TERMS = ["Term I","Term II","Term III"];
const TERM_MONTHS = {
  "Term I": ["FEB","MAR","APR"],
  "Term II": ["MAY","JUN","JUL"],
  "Term III": ["SEP","OCT","NOV"],
};
const DEFAULT_BANDS = [
  { min:90, max:100, grade:"D1", label:"Excellent" },
  { min:80, max:89,  grade:"D2", label:"Very Good" },
  { min:75, max:79,  grade:"C3", label:"Good" },
  { min:65, max:74,  grade:"C4", label:"Good" },
  { min:55, max:64,  grade:"C5", label:"Quite Good" },
  { min:45, max:54,  grade:"C6", label:"Fair" },
  { min:35, max:44,  grade:"P7", label:"Pass" },
  { min:25, max:34,  grade:"P8", label:"Weak" },
  { min:0,  max:24,  grade:"F9", label:"Fail" },
];
const DEFAULT_DIVISIONS = [
  { name:"I",   min:4,  max:12 },
  { name:"II",  min:13, max:24 },
  { name:"III", min:25, max:36 },
  { name:"IV",  min:37, max:48 },
  { name:"U",   min:49, max:999 },
];
const DEFAULT_SCHOOL = {
  name: "ST. KIZITO'S PRIMARY SCHOOL",
  motto: "Knowledge is Power",
  poBox: "P.O. Box 172, Tororo",
  tel: "",
  email: "stkizitosprimaryschool@gmail.com",
  district: "Tororo",
  headTeacher: "",
  nextOpens: "",
  nextEnds: "",
  requirements: "",
  year: String(new Date().getFullYear()),
};
// ─── STORAGE (shared across all devices via window.storage) ───────────────────
// All MKIS data is saved with shared:true, so every device/browser that opens
// this app reads and writes the SAME records instead of separate, per-device
// localStorage copies. This is what makes data appear identically everywhere.
const STORAGE_KEYS = [
  "mkis_students","mkis_termmarks","mkis_monthlymarks","mkis_initials",
  "mkis_bands","mkis_divisions","mkis_school","mkis_accounts","mkis_changerequests",
  "mkis_locked_term","mkis_locked_monthly",
];
// One-time migration: if a browser still has old localStorage data and the
// shared store is empty, lift it into shared storage so nothing is lost.
async function migrateLocalStorageOnce() {
  try {
    if (localStorage.getItem("mkis_migrated_v1")) return;
    for (const key of STORAGE_KEYS) {
      const raw = localStorage.getItem(key);
      if (raw == null) continue;
      const existing = await window.storage.get(key, true).catch(() => null);
      if (existing) continue; // shared copy already exists, don't overwrite
      await window.storage.set(key, raw, true);
    }
    localStorage.setItem("mkis_migrated_v1", "1");
  } catch {}
}
async function loadShared(key, def) {
  try {
    const res = await window.storage.get(key, true);
    return res && res.value != null ? JSON.parse(res.value) : def;
  } catch { return def; }
}
async function saveShared(key, val) {
  try { await window.storage.set(key, JSON.stringify(val), true); } catch {}
}
// ─── NO-DATA-LOSS WRITE LAYER ───────────────────────────────────────────────
// Problem this solves: two devices can each hold a slightly different local
// snapshot of the same shared key. If a save simply pushes "whatever my
// local snapshot is right now", whichever device saves last wins -- and any
// change the OTHER device made (that this device never saw) is silently
// erased. mergeShared fixes this by re-reading the CURRENT value straight
// from storage immediately before writing, then combining it with the
// caller's intended change using a per-key merge strategy, so a write never
// overwrites information neither this update was about NOR the device has
// seen. updateShared() is the function every mutation should funnel through.
function mergeArrayById(remote, local, deletedIds) {
  const remoteArr = Array.isArray(remote) ? remote : [];
  const localArr = Array.isArray(local) ? local : [];
  const byId = new Map();
  remoteArr.forEach(r => r && r.id != null && byId.set(r.id, r));
  // local entries win for ids it touched (it's the freshest intent for those ids)
  localArr.forEach(r => r && r.id != null && byId.set(r.id, r));
  // honour deletions this device explicitly asked for, even if remote still has them
  if (deletedIds) deletedIds.forEach(id => byId.delete(id));
  return Array.from(byId.values());
}
function deepMergeObjects(remote, local) {
  // Recursively merges two plain-object trees (used for termMarks/monthlyMarks,
  // which are nested {studentId: {term: {subject: {field: value}}}} maps).
  // Any branch present on only one side is kept; branches present on both
  // sides are merged key-by-key rather than one side replacing the other,
  // so an edit to Student A's marks on one device can never wipe out a
  // concurrent edit to Student B's marks made on another device.
  if (remote == null) return local;
  if (local == null) return remote;
  const remoteIsObj = typeof remote === "object" && !Array.isArray(remote);
  const localIsObj = typeof local === "object" && !Array.isArray(local);
  if (!remoteIsObj || !localIsObj) return local; // leaf value: local edit wins
  const out = { ...remote };
  for (const k of Object.keys(local)) out[k] = deepMergeObjects(remote[k], local[k]);
  return out;
}
// Merge strategy per storage key. "leaf" keys (school/bands/divisions/password)
// are small, whole-object settings a person edits deliberately on one screen,
// so local intent simply wins there -- but we still merge them via this same
// read-then-write path so a save never clobbers a *different* settings key.
const MERGE_STRATEGIES = {
  mkis_students: (remote, local, ctx) => mergeArrayById(remote, local, ctx?.deletedStudentIds),
  mkis_termmarks: (remote, local) => deepMergeObjects(remote, local),
  mkis_monthlymarks: (remote, local) => deepMergeObjects(remote, local),
  mkis_bands: (remote, local) => local,
  mkis_divisions: (remote, local) => local,
  mkis_school: (remote, local) => local,
  mkis_accounts: (remote, local) => deepMergeObjects(remote, local),
  mkis_changerequests: (remote, local, ctx) => mergeArrayById(remote, local, ctx?.deletedRequestIds),
  mkis_initials: (remote, local) => local,
  // Locked-entry maps are { "CLASS__TERM__YEAR" (or "...__MONTH" for monthly): true|false }.
  // Deep-merged key-by-key so a Save/Unlock on one device never clobbers a
  // different class/term/month another device locked or unlocked.
  mkis_locked_term: (remote, local) => deepMergeObjects(remote, local),
  mkis_locked_monthly: (remote, local) => deepMergeObjects(remote, local),
};
// Reads the freshest shared value, merges in the local change using the
// strategy for that key, writes the merged result back, and returns it so
// the caller can sync local React state to exactly what was actually saved.
async function updateShared(key, localVal, ctx) {
  const remoteVal = await loadShared(key, undefined);
  const merge = MERGE_STRATEGIES[key] || ((r, l) => (r === undefined ? l : l));
  const merged = remoteVal === undefined ? localVal : merge(remoteVal, localVal, ctx);
  await saveShared(key, merged);
  return merged;
}

// ─── PASSWORD HASHING (Web Crypto API — no dependencies) ─────────────────────
// Passwords are stored as "sha256:<hex>" so we can detect plaintext legacy
// values and upgrade them on first load.
async function hashPassword(plain) {
  const buf = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(plain)
  );
  const hex = Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2,"0")).join("");
  return "sha256:" + hex;
}
function isHashed(val) { return typeof val === "string" && val.startsWith("sha256:"); }
async function verifyPassword(plain, stored) {
  if (!isHashed(stored)) return plain === stored; // legacy plaintext fallback
  return (await hashPassword(plain)) === stored;
}
// ─── ACCOUNTS (two roles: teacher + admin) ────────────────────────────────────
// mkis_accounts is a shared object keyed by lowercase username:
//   { [usernameLower]: { username, passwordHash, role: "teacher"|"admin" } }
// The teacher account is the everyday login everyone already knows; the admin
// account is a separate, more privileged login. Defaults match the school's
// requested credentials and are only used the very first time the app runs
// (i.e. when no accounts have been saved to shared storage yet).
const TEACHER_USERNAMES = [
  "OMIITA01","EMURON10","IJANG11","ODOI31","IMMACULATE71","AKOTH33",
  "MARYG0001","ANYANGO45","IGARI89","JAKISA23","NYACHWO",
];
const DEFAULT_ACCOUNTS_PLAIN = (() => {
  const out = {
    "gerald": { username: "Gerald", password: "GOODTOGO11", role: "admin" },
  };
  for (const u of TEACHER_USERNAMES) {
    out[u.toLowerCase()] = { username: u, password: `${u}172`, role: "teacher" };
  }
  return out;
})();
async function buildDefaultAccounts() {
  const out = {};
  for (const key of Object.keys(DEFAULT_ACCOUNTS_PLAIN)) {
    const a = DEFAULT_ACCOUNTS_PLAIN[key];
    out[key] = { username: a.username, passwordHash: await hashPassword(a.password), role: a.role };
  }
  return out;
}
function findAccount(accounts, usernameInput) {
  if (!accounts || !usernameInput) return null;
  return accounts[usernameInput.trim().toLowerCase()] || null;
}
async function verifyAccountLogin(accounts, usernameInput, passwordInput) {
  const acct = findAccount(accounts, usernameInput);
  if (!acct) return null;
  const ok = await verifyPassword(passwordInput, acct.passwordHash);
  return ok ? acct : null;
}
// ─── HELPERS ─────────────────────────────────────────────────────────────────
const toUpper = (s) => (s || "").toUpperCase();
function gradeFor(score, bands) {
  if (score === undefined || score === null || isNaN(score)) return null;
  return bands.find(b => score >= b.min && score <= b.max) || null;
}
function aggOf(score, bands) {
  const g = gradeFor(score, bands);
  if (!g) return 9;
  return parseInt(g.grade.replace(/\D/g,"")) || 9;
}
function gradeLabel(score, bands) {
  const g = gradeFor(score, bands);
  return g ? g.grade : "F9";
}
function divisionOf(totalAgg, numSubjects, divisions) {
  if (!totalAgg || numSubjects === 0) return "U";
  const d = divisions.find(d => totalAgg >= d.min && totalAgg <= d.max);
  return d ? d.name : "U";
}
function remarkFor(score) {
  if (score >= 90) return "Excellent";
  if (score >= 80) return "Very Good";
  if (score >= 75) return "Good";
  if (score >= 65) return "Good";
  if (score >= 55) return "Quite Good";
  if (score >= 45) return "Fair";
  if (score >= 35) return "Pass";
  if (score >= 25) return "Weak";
  return "Fail";
}
// Rank with ties. When `aggs` is supplied, ties on total marks are broken by
// lower aggregate (better aggregate = better position). Pupils only share a
// position when BOTH their total marks AND their aggregate match exactly.
function rankWithTies(totals, aggs) {
  return totals.map((t, i) => {
    if (!t || t === 0) return "-";
    const a = aggs ? aggs[i] : null;
    const aIsNum = typeof a === "number";
    let better = 0;
    for (let j = 0; j < totals.length; j++) {
      if (j === i) continue;
      const tj = totals[j];
      if (!tj || tj === 0) continue;
      if (tj > t) { better++; continue; }
      if (tj === t && aggs) {
        const aj = aggs[j];
        const ajIsNum = typeof aj === "number";
        // lower agg ranks higher; non-numeric (e.g. "X") is worst
        if (ajIsNum && (!aIsNum || aj < a)) better++;
      }
    }
    return better + 1;
  });
}
function ordinal(n) {
  if (n === "-" || !n) return "-";
  const s = ["th","st","nd","rd"];
  const v = n % 100;
  return n + (s[(v-20)%10] || s[v] || s[0]);
}
function ordinalSuffix(n) {
  if (!n || n === "-") return "";
  const s = ["TH","ST","ND","RD"];
  const v = n % 100;
  return (s[(v-20)%10] || s[v] || s[0]);
}
function clampMark(val, max = 100) {
  if (val === undefined || val === null || val === "") return undefined;
  const n = Number(val);
  if (isNaN(n)) return undefined;
  return Math.max(0, Math.min(max, n));
}
// ─── EXPORT HELPERS (Word + Excel downloads) ───────────────────────────────────
function safeFileName(s) {
  return String(s).replace(/[^a-z0-9]+/gi, "_").replace(/^_+|_+$/g, "");
}
function triggerBlobDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function escapeHtml(s) {
  return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
// Builds a real, openable Word file with zero extra dependencies: Word natively
// opens well-formed HTML saved with a .doc extension / application/msword mime type.
function downloadWordHtml(title, bodyHtml, filename, opts = {}) {
  // Default stays landscape A4 (297mm x 210mm) -- unchanged for the wide,
  // many-column Result Sheets / Mark Sheets exports. Pass
  // { pageSize: "210mm 297mm" } for narrow, vertical layouts like Report
  // Cards so the downloaded file's orientation matches the on-screen
  // preview instead of defaulting to landscape.
  const pageSize = opts.pageSize || "297mm 210mm";
  const pageMargin = opts.margin || "14mm";
  const html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8">
<title>${escapeHtml(title)}</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>90</w:Zoom></w:WordDocument></xml><![endif]-->
<style>
  @page { size: ${pageSize}; margin: ${pageMargin}; }
  body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color:#111; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 12px; }
  th, td { border: 1px solid #999; padding: 4px 6px; font-size: 9.5pt; text-align: center; }
  th { background:#1e3a6e; color:#fff; font-weight:bold; }
  .title { text-align:center; font-size:16pt; font-weight:bold; }
  .motto { text-align:center; font-style:italic; font-size:10pt; }
  .addr { text-align:center; font-size:10pt; margin-bottom:4px; }
  .subtitle { text-align:center; font-weight:bold; font-size:12pt; margin:6px 0 10px; }
  .section-title { font-weight:bold; font-size:11pt; margin:14px 0 6px; }
  .name-cell { text-align:left; font-weight:600; }
  tr:nth-child(even) td { background:#f8fafc; }
  /* Keep each pupil's full report card together as one block in the
     downloaded file -- never split a table/section across two pages. */
  .report-card-block, .report-card-block table, .report-card-block tr {
    page-break-inside: avoid;
    mso-pagination: none;
  }
</style>
</head>
<body>
${bodyHtml}
</body>
</html>`;
  const blob = new Blob(["\ufeff", html], { type: "application/msword" });
  triggerBlobDownload(blob, filename);
}
function htmlTable(headerRow, dataRows) {
  const head = `<tr>${headerRow.map(h => `<th>${escapeHtml(h)}</th>`).join("")}</tr>`;
  const body = dataRows.map(row => `<tr>${row.map((v, i) => `<td${i === 1 ? ' class="name-cell"' : ""}>${escapeHtml(v === undefined || v === null || v === "" ? "-" : v)}</td>`).join("")}</tr>`).join("");
  return `<table>${head}${body}</table>`;
}
function titleBlockHtml(school, subtitle) {
  let html = `<div class="title">${escapeHtml(school.name || "")}</div>`;
  if (school.motto) html += `<div class="motto">"${escapeHtml(school.motto)}"</div>`;
  const addr = [school.poBox, school.email].filter(Boolean).map(escapeHtml).join(" &bull; ");
  if (addr) html += `<div class="addr">${addr}</div>`;
  html += `<div class="subtitle">${escapeHtml(subtitle)}</div>`;
  return html;
}
// ── End-of-term Result Sheet: row builders shared by Excel + Word ──
function resultSheetHeaderRow(isLower, subjects) {
  const head = ["S/N", "NAME OF PUPIL"];
  if (isLower) subjects.forEach(s => head.push(s));
  else subjects.forEach(s => head.push(`${s} CA`, `${s} EXAM`, `${s} AVG`, `${s} AGG`));
  head.push("TOT MK");
  if (!isLower) head.push("TOT AGG", "DIV");
  head.push("POS");
  return head;
}
function resultSheetDataRow(r, i, isLower) {
  const row = [i + 1, r.s.name];
  r.perSub.forEach(p => {
    if (isLower) {
      row.push(p.isX ? "X" : (p.av ?? "-"));
    } else {
      row.push(p.isX ? "X" : (p.ca ?? "-"));
      row.push(p.isX ? "X" : (p.exam ?? "-"));
      row.push(p.isX ? "X" : (p.av ?? "-"));
      row.push(p.isX ? "X" : (p.av !== undefined ? p.agg : "-"));
    }
  });
  row.push(r.totMk || "-");
  if (!isLower) {
    row.push(r.hasX ? "X" : (r.totAgg || "-"));
    row.push(r.hasX ? "X" : (r.totMk ? r.div : "-"));
  }
  row.push(r.pos !== "-" ? r.pos : "-");
  return row;
}
function exportResultSheetExcel({ school, cls, term, year, isLower, subjects, sortedRows, best, worst, avg, subjectAnalysis, gradeKeys, divCounts, classCount }) {
  const headerRow = resultSheetHeaderRow(isLower, subjects);
  const numCols = headerRow.length;
  const dataRows = sortedRows.map((r, i) => resultSheetDataRow(r, i, isLower));
  const aoa = [
    [school.name],
    [school.poBox || ""],
    [`END OF ${term.toUpperCase()} ${year} - ${cls} RESULT SHEET`],
    [],
    headerRow,
    ...dataRows,
    [],
    [`Highest: ${best || "-"}    Lowest: ${worst || "-"}    Class Average: ${avg || "-"}    Best Pupil: ${sortedRows[0]?.s.name || "-"}`],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: numCols - 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: numCols - 1 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: numCols - 1 } },
  ];
  ws["!cols"] = headerRow.map((h, i) => (i === 1 ? { wch: 26 } : { wch: 11 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Result Sheet");
  if (!isLower && subjectAnalysis?.length) {
    const aHead = ["SUBJECT", ...gradeKeys, "X", "TOTAL"];
    const aRows = subjectAnalysis.map(sa => [sa.sub, ...gradeKeys.map(g => sa.gradeCounts[g] || 0), sa.xCount || 0, sa.total]);
    const gHead = ["NO. OF PUPILS", "DIV I", "DIV II", "DIV III", "DIV IV", "U", "X"];
    const gRow = [classCount, divCounts.I, divCounts.II, divCounts.III, divCounts.IV, divCounts.U, divCounts.X];
    const aoa2 = [["A. SUBJECT PERFORMANCE ANALYSIS"], aHead, ...aRows, [], ["B. GENERAL PERFORMANCE ANALYSIS"], gHead, gRow];
    const ws2 = XLSX.utils.aoa_to_sheet(aoa2);
    XLSX.utils.book_append_sheet(wb, ws2, "Analysis");
  }
  XLSX.writeFile(wb, `${safeFileName(cls)}_${safeFileName(term)}_${year}_Result_Sheet.xlsx`);
}
function exportResultSheetWord({ school, cls, term, year, isLower, subjects, sortedRows, best, worst, avg, subjectAnalysis, gradeKeys, divCounts, classCount }) {
  const headerRow = resultSheetHeaderRow(isLower, subjects);
  const dataRows = sortedRows.map((r, i) => resultSheetDataRow(r, i, isLower));
  let body = titleBlockHtml(school, `END OF ${term.toUpperCase()} ${year} - ${cls} RESULT SHEET`);
  body += htmlTable(headerRow, dataRows);
  body += `<p><b>Highest:</b> ${escapeHtml(best || "-")} &nbsp; <b>Lowest:</b> ${escapeHtml(worst || "-")} &nbsp; <b>Class Average:</b> ${escapeHtml(avg || "-")} &nbsp; <b>Best Pupil:</b> ${escapeHtml(sortedRows[0]?.s.name || "-")}</p>`;
  if (!isLower && subjectAnalysis?.length) {
    const aHead = ["SUBJECT", ...gradeKeys, "X", "TOTAL"];
    const aRows = subjectAnalysis.map(sa => [sa.sub, ...gradeKeys.map(g => sa.gradeCounts[g] || 0), sa.xCount || 0, sa.total]);
    const gHead = ["NO. OF PUPILS", "DIV I", "DIV II", "DIV III", "DIV IV", "U", "X"];
    const gRow = [classCount, divCounts.I, divCounts.II, divCounts.III, divCounts.IV, divCounts.U, divCounts.X];
    body += `<div class="section-title">A. Subject Performance Analysis</div>${htmlTable(aHead, aRows)}`;
    body += `<div class="section-title">B. General Performance Analysis</div>${htmlTable(gHead, [gRow])}`;
  }
  body += `<p>Class Teacher's Comment: .............................................................................. Sign: ......................</p>`;
  body += `<p>Head Teacher's Comment: .............................................................................. Sign: ......................</p>`;
  downloadWordHtml(`${cls} ${term} ${year} Result Sheet`, body, `${safeFileName(cls)}_${safeFileName(term)}_${year}_Result_Sheet.doc`);
}
// ── Monthly Mark Sheet: row builders + recomputation for export ──
function monthlySheetHeaderRow(isLower, subjects) {
  const head = ["S/N", "NAME OF PUPIL"];
  if (isLower) subjects.forEach(s => head.push(s));
  else subjects.forEach(s => head.push(`${s} MK`, `${s} AGG`));
  head.push("TOT MK");
  if (!isLower) head.push("TOT AGG", "DIV");
  head.push("POS");
  return head;
}
function monthlySheetDataRow(r, i, isLower) {
  const row = [i + 1, r.s.name];
  r.perSub.forEach(p => {
    if (isLower) row.push(p.mk ?? "-");
    else { row.push(p.mk ?? "-"); row.push(p.isX ? "X" : (p.mk !== undefined ? p.agg : "-")); }
  });
  // Total marks are unaffected by X — only sum subjects actually attempted.
  row.push(r.totMk || "-");
  if (!isLower) {
    row.push(r.hasX ? "X" : (r.totAgg || "-"));
    row.push(r.hasX ? "X" : (r.totMk ? r.div : "-"));
  }
  row.push(r.pos !== "-" ? r.pos : "-");
  return row;
}
// Mirrors MonthBlock's internal calculation so exported figures match what's on screen.
// X (Missing / Did not complete) rule: any required paper that is blank ("-")
// marks that subject as X and forces the overall aggregate + division to X,
// but never affects the total marks or the position ranking.
function computeMonthRows({ month, classStudents, monthlyMarks, tk, subjects, isLower, bands, divisions }) {
  const rows = classStudents.map(s => {
    const m = monthlyMarks[s.id]?.[tk]?.[month] || {};
    const perSub = subjects.map(sub => {
      const mk = m[sub]?.mk;
      const isX = (mk === undefined || mk === null);
      const agg = (!isLower && !isX) ? aggOf(mk, bands) : undefined;
      return { sub, mk, agg, isX };
    });
    const hasX = !isLower && perSub.some(p => p.isX);
    const totMk = perSub.reduce((a, p) => a + (p.mk ?? 0), 0);
    const totAgg = isLower ? null : (hasX ? "X" : perSub.reduce((a, p) => a + (p.agg ?? 0), 0));
    const div = isLower ? null : (hasX ? "X" : divisionOf(totAgg, 4, divisions));
    return { s, perSub, totMk, totAgg, div, hasX };
  });
  const positions = rankWithTies(rows.map(r => (r.totMk > 0 ? r.totMk : null)), rows.map(r => typeof r.totAgg === "number" ? r.totAgg : null));
  const indexed = rows.map((r, i) => ({ ...r, pos: positions[i] }));
  return [...indexed].sort((a, b) => { if (a.pos === "-") return 1; if (b.pos === "-") return -1; return a.pos - b.pos; });
}
function exportMonthlyExcel({ school, cls, term, year, isLower, subjects, monthsData }) {
  const wb = XLSX.utils.book_new();
  monthsData.forEach(({ month, sortedRows, divCounts }) => {
    const headerRow = monthlySheetHeaderRow(isLower, subjects);
    const numCols = headerRow.length;
    const dataRows = sortedRows.map((r, i) => monthlySheetDataRow(r, i, isLower));
    const aoa = [[school.name], [`${month} - ${term} ${year} - ${cls} MONTHLY MARK SHEET`], [], headerRow, ...dataRows];
    if (!isLower && divCounts) {
      const gHead = ["NO. OF PUPILS", "DIV I", "DIV II", "DIV III", "DIV IV", "U", "X"];
      const gRow = [sortedRows.length, divCounts.I, divCounts.II, divCounts.III, divCounts.IV, divCounts.U, divCounts.X];
      aoa.push([], ["GENERAL PERFORMANCE ANALYSIS"], gHead, gRow);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: numCols - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: numCols - 1 } },
    ];
    ws["!cols"] = headerRow.map((h, i) => (i === 1 ? { wch: 26 } : { wch: 11 }));
    XLSX.utils.book_append_sheet(wb, ws, month.slice(0, 31));
  });
  XLSX.writeFile(wb, `${safeFileName(cls)}_${safeFileName(term)}_${year}_Monthly_Mark_Sheets.xlsx`);
}
function exportMonthlyWord({ school, cls, term, year, isLower, subjects, monthsData }) {
  const headerRow = monthlySheetHeaderRow(isLower, subjects);
  let body = titleBlockHtml(school, `MONTHLY MARK SHEETS - ${term.toUpperCase()} ${year} - ${cls}`);
  monthsData.forEach(({ month, sortedRows, divCounts }) => {
    const dataRows = sortedRows.map((r, i) => monthlySheetDataRow(r, i, isLower));
    body += `<div class="section-title">${escapeHtml(month)} - ${escapeHtml(term)} ${escapeHtml(year)}</div>`;
    body += htmlTable(headerRow, dataRows);
    if (!isLower && divCounts) {
      const gHead = ["NO. OF PUPILS", "DIV I", "DIV II", "DIV III", "DIV IV", "U", "X"];
      const gRow = [sortedRows.length, divCounts.I, divCounts.II, divCounts.III, divCounts.IV, divCounts.U, divCounts.X];
      body += `<div class="section-title">General Performance Analysis</div>${htmlTable(gHead, [gRow])}`;
    }
  });
  downloadWordHtml(`${cls} ${term} ${year} Monthly Mark Sheets`, body, `${safeFileName(cls)}_${safeFileName(term)}_${year}_Monthly_Mark_Sheets.doc`);
}
// ── Monthly Report Card Word export ──
function exportMonthlyCardsWord({ school, cls, term, year, isLower, subjects, cardData, totalInClass }) {
  let body = titleBlockHtml(school, `MONTHLY TESTS REPORT CARDS - ${term.toUpperCase()} ${year} - ${cls}`);
  cardData.forEach(({ s, monthData }) => {
    // Student header
    body += `<div style="page-break-before:always;margin-top:20px;">`;
    body += `<table style="width:100%;border-collapse:collapse;margin-bottom:6px;">
      <tr>
        <td style="padding:4px 8px;font-size:12pt;"><b>NAME:</b> ${escapeHtml(s.name)}</td>
        <td style="padding:4px 8px;font-size:12pt;"><b>CLASS:</b> ${escapeHtml(cls)}</td>
        <td style="padding:4px 8px;font-size:12pt;"><b>YEAR:</b> ${escapeHtml(year)}</td>
      </tr>
    </table>`;
    // Marks table header
    const subHeaders = subjects.map(sub =>
      isLower
        ? `<th style="border:1px solid #999;padding:4px;background:#1e3a6e;color:white;font-size:9pt;">${escapeHtml(sub)}${lowerSubjectMax(sub)!==100?` /${lowerSubjectMax(sub)}`:""}</th>`
        : `<th colspan="2" style="border:1px solid #999;padding:4px;background:#1e3a6e;color:white;font-size:9pt;">${escapeHtml(sub)}</th>`
    ).join("");
    const subSubHeaders = isLower ? "" : `<tr>${subjects.map(()=>`<th style="border:1px solid #999;padding:3px;background:#3b82f6;color:white;font-size:8pt;">MK</th><th style="border:1px solid #999;padding:3px;background:#60a5fa;color:white;font-size:8pt;">AGG</th>`).join("")}</tr>`;
    const totHeaders = isLower
      ? `<th style="border:1px solid #999;padding:4px;background:#1e3a6e;color:white;font-size:9pt;">TOT MK</th><th style="border:1px solid #999;padding:4px;background:#1e3a6e;color:white;font-size:9pt;">POS</th>`
      : `<th style="border:1px solid #999;padding:4px;background:#1e3a6e;color:white;font-size:9pt;">TOT MK</th><th style="border:1px solid #999;padding:4px;background:#1e3a6e;color:white;font-size:9pt;">TOT AGG</th><th style="border:1px solid #999;padding:4px;background:#1e3a6e;color:white;font-size:9pt;">DIV</th><th style="border:1px solid #999;padding:4px;background:#1e3a6e;color:white;font-size:9pt;">POS</th>`;
    body += `<table style="width:100%;border-collapse:collapse;font-size:10pt;margin-bottom:6px;">
      <thead>
        <tr>
          <th rowspan="2" style="border:1px solid #999;padding:4px;background:#1e3a6e;color:white;font-size:9pt;text-align:left;min-width:70px;">MONTH</th>
          ${subHeaders}
          ${totHeaders}
        </tr>
        ${subSubHeaders ? `<tr>${subSubHeaders.replace(/<tr>|<\/tr>/g,"")}</tr>` : ""}
      </thead>
      <tbody>`;
    monthData.forEach(({ month, perSub, totMk, totAgg, div, pos }, mIdx) => {
      const bg = mIdx % 2 === 0 ? "#ffffff" : "#f8fafc";
      const subCells = perSub.map(p =>
        isLower
          ? `<td style="border:1px solid #999;padding:4px;text-align:center;background:${bg};">${p.mk !== undefined ? p.mk : "-"}</td>`
          : `<td style="border:1px solid #999;padding:4px;text-align:center;background:${bg};">${p.mk !== undefined ? p.mk : "-"}</td><td style="border:1px solid #999;padding:4px;text-align:center;background:#fff7ed;">${p.agg !== undefined ? p.agg : "-"}</td>`
      ).join("");
      const totCells = isLower
        ? `<td style="border:1px solid #999;padding:4px;text-align:center;font-weight:bold;background:#ede9fe;">${totMk > 0 ? totMk : "-"}</td><td style="border:1px solid #999;padding:4px;text-align:center;">${pos && pos !== "-" ? ordinal(pos) : "-"}</td>`
        : `<td style="border:1px solid #999;padding:4px;text-align:center;font-weight:bold;background:#ede9fe;">${totMk > 0 ? totMk : "-"}</td><td style="border:1px solid #999;padding:4px;text-align:center;background:#ede9fe;">${totAgg > 0 ? totAgg : "-"}</td><td style="border:1px solid #999;padding:4px;text-align:center;font-weight:bold;">${totMk > 0 ? div : "-"}</td><td style="border:1px solid #999;padding:4px;text-align:center;">${pos && pos !== "-" ? ordinal(pos) : "-"}</td>`;
      body += `<tr><td style="border:1px solid #999;padding:4px;font-weight:bold;background:#dbeafe;text-align:left;">${escapeHtml(month)}</td>${subCells}${totCells}</tr>`;
    });
    body += `<tr><td colspan="100" style="border:1px solid #999;padding:4px;font-size:9pt;color:#6b7280;">Total pupils in class: <b>${totalInClass}</b></td></tr>`;
    body += `</tbody></table>`;
    body += `<p style="font-size:11pt;line-height:2;">
      <b>Class Teacher's Comment:</b> .............................................................................. <b>Sign:</b> ......................</p>
      <p style="font-size:11pt;line-height:2;"><b>Head Teacher's Comment:</b> .............................................................................. <b>Sign:</b> ......................</p>`;
    body += `</div>`;
  });
  downloadWordHtml(`${cls} ${term} ${year} Monthly Report Cards`, body, `${safeFileName(cls)}_${safeFileName(term)}_${year}_Monthly_Report_Cards.doc`);
}
// ── Termly Report Card Word export ──
function exportReportCardsWord({ school, cls, term, year, isLower, rows, allPositions, totalInClass, bands, initials }) {
  let body = "";
  rows.forEach((r, idx) => {
    const { s, perSub, totMk, totAgg, div, hasX } = r;
    const position = allPositions[s.id];
    // Each pupil's page carries its own school-details header (matching the
    // on-screen preview) plus their pupil info, marks table, and comments,
    // all inside one report-card-block so nothing is split apart. Only
    // pages after the first force a break, so there's no leading blank page.
    body += `<div class="report-card-block" style="page-break-inside:avoid;${idx > 0 ? "page-break-before:always;" : ""}margin-top:20px;">`;
    body += titleBlockHtml(school, `PUPIL'S ACADEMIC REPORT CARD - END OF ${term.toUpperCase()} ${year} - ${cls}`);
    body += `<table style="width:100%;border-collapse:collapse;margin-bottom:6px;">
      <tr>
        <td style="padding:4px 8px;font-size:12pt;"><b>NAME:</b> ${escapeHtml(s.name)}</td>
        <td style="padding:4px 8px;font-size:12pt;"><b>CLASS:</b> ${escapeHtml(cls)}</td>
        <td style="padding:4px 8px;font-size:12pt;"><b>TERM:</b> ${escapeHtml(term)}</td>
      </tr>
      <tr>
        <td style="padding:4px 8px;font-size:12pt;"><b>YEAR:</b> ${escapeHtml(year)}</td>
        <td colspan="2" style="padding:4px 8px;font-size:12pt;"><b>POSITION:</b> ${position && position !== "-" ? ordinal(position) : "-"} out of ${totalInClass}</td>
      </tr>
    </table>`;
    const subHead = isLower
      ? `<th style="border:1px solid #999;padding:5px;background:#1e3a6e;color:white;">AVERAGE</th>`
      : `<th style="border:1px solid #999;padding:5px;background:#1e3a6e;color:white;">CA</th><th style="border:1px solid #999;padding:5px;background:#1e3a6e;color:white;">EXAM</th><th style="border:1px solid #999;padding:5px;background:#1e3a6e;color:white;">AVERAGE</th><th style="border:1px solid #999;padding:5px;background:#1e3a6e;color:white;">AGG</th>`;
    body += `<table style="width:100%;border-collapse:collapse;font-size:10pt;margin-bottom:6px;">
      <thead>
        <tr>
          <th style="border:1px solid #999;padding:5px;background:#1e3a6e;color:white;text-align:left;">SUBJECT</th>
          ${subHead}
          <th style="border:1px solid #999;padding:5px;background:#1e3a6e;color:white;">REMARKS</th>
          <th style="border:1px solid #999;padding:5px;background:#1e3a6e;color:white;">INITIALS</th>
        </tr>
      </thead>
      <tbody>`;
    perSub.forEach((p, i) => {
      const isUnscored = isLower && lowerSubjectMax(p.sub) !== 100;
      const bg = i % 2 === 0 ? "#ffffff" : "#f8fafc";
      const subName = `${escapeHtml(p.sub)}${isUnscored ? ` (/${lowerSubjectMax(p.sub)})` : ""}`;
      const subCells = isLower
        ? `<td style="border:1px solid #999;padding:5px;text-align:center;font-weight:bold;font-size:12pt;background:${bg};">${p.isX ? "X" : p.av ?? "-"}</td>`
        : `<td style="border:1px solid #999;padding:5px;text-align:center;background:#fefce8;">${p.isX ? "X" : p.ca ?? "-"}</td><td style="border:1px solid #999;padding:5px;text-align:center;background:#f0fdf4;">${p.isX ? "X" : p.exam ?? "-"}</td><td style="border:1px solid #999;padding:5px;text-align:center;font-weight:bold;font-size:12pt;background:${bg};">${p.isX ? "X" : p.av ?? "-"}</td><td style="border:1px solid #999;padding:5px;text-align:center;">${isUnscored ? "-" : p.isX ? "X" : p.av !== undefined ? p.agg : "-"}</td>`;
      const remark = isUnscored ? "-" : p.isX ? "Absent" : p.av !== undefined ? remarkFor(p.av) : "-";
      body += `<tr style="background:${bg};">
        <td style="border:1px solid #999;padding:5px;font-weight:600;text-align:left;">${subName}</td>
        ${subCells}
        <td style="border:1px solid #999;padding:5px;text-align:center;font-weight:bold;color:#1e40af;">${remark}</td>
        <td style="border:1px solid #999;padding:5px;text-align:center;font-weight:bold;color:#7c3aed;">${((initials||{})[cls]||{})[p.sub]||""}</td>
      </tr>`;
    });
    const totColspan = isLower ? 1 : 3;
    body += `<tr style="background:#dbeafe;font-weight:bold;">
      <td style="border:1px solid #999;padding:5px;text-align:left;" colspan="${totColspan + 1}">TOTAL</td>
      <td style="border:1px solid #999;padding:5px;text-align:center;font-size:12pt;">${totMk || "-"}</td>
      ${!isLower ? `<td style="border:1px solid #999;padding:5px;text-align:center;">${hasX ? "X" : totAgg || "-"}</td>` : ""}
      <td style="border:1px solid #999;padding:5px;"></td><td style="border:1px solid #999;padding:5px;"></td>
    </tr>`;
    body += `</tbody></table>`;
    body += `<p style="font-size:11pt;line-height:2.2;margin-top:6px;">`;
    if (!isLower) body += `<b>DIVISION:</b> ${hasX ? "X" : totMk ? div : "-"}&nbsp;&nbsp;&nbsp;`;
    body += `</p>`;
    body += `<p style="font-size:11pt;line-height:2;"><b>CONDUCT:</b> ...........................................................................................</p>
      <p style="font-size:11pt;line-height:2;"><b>Class Teacher's Comment:</b> .............................................................................. <b>Sign:</b> ......................</p>
      <p style="font-size:11pt;line-height:2;"><b>Head Teacher's Comment:</b> .............................................................................. <b>Sign:</b> ......................</p>
      <p style="font-size:11pt;line-height:2;"><b>Next Term begins on</b> ....................... <b>Ends on</b> .......................</p>
      <p style="font-size:11pt;line-height:2;"><b>Requirements:</b> ${escapeHtml(school.requirements || "...........................................................................................")}</p>
      <p style="font-size:11pt;line-height:2;"><b>Parent's Signature after reading:</b> ...................................................................</p>`;
    body += `</div>`;
  });
  downloadWordHtml(`${cls} ${term} ${year} Report Cards`, body, `${safeFileName(cls)}_${safeFileName(term)}_${year}_Report_Cards.doc`, { pageSize: "210mm 297mm", margin: "12mm" });
}
// ─── SHARED STYLES ───────────────────────────────────────────────────────────
const lbl = { display:"block", fontSize:11, fontWeight:700, color:"#374151", marginBottom:4, textTransform:"uppercase", letterSpacing:0.5 };
const inp = { padding:"8px 10px", border:"1.5px solid #d1d5db", borderRadius:7, fontSize:13, outline:"none", background:"white", minWidth:80 };
const th  = { padding:"8px 10px", fontWeight:700, textAlign:"center", fontSize:12, whiteSpace:"nowrap", border:"1px solid #d1d5db" };
const td  = { padding:"6px 8px", textAlign:"center", fontSize:12, border:"1px solid #d1d5db" };
const markInput = { width:52, padding:"3px 4px", border:"1px solid #d1d5db", borderRadius:4, fontSize:12, textAlign:"center", background:"white" };
// A mark-entry cell that lets the teacher type freely without interruption.
// Every keystroke only updates local draft text (instant, no save, no
// approval check) so entry feels smooth. The actual save -- and the
// decision on whether a change needs admin approval -- only happens once
// the teacher is done with the cell: on blur (tabbing/clicking away) or
// pressing Enter. That's the moment "existingVal" (the last saved value) is
// compared against the new value, never mid-keystroke.
function MarkInput({ value, existingVal, max, onCommit, style, locked }) {
  const [draft, setDraft] = useState(value ?? "");
  const [pendingFlash, setPendingFlash] = useState(false);
  // Keep the draft in sync whenever the saved value changes from elsewhere
  // (another device, an admin approval landing, switching class/term, etc.)
  // but never while the teacher is actively mid-keystroke in this cell.
  const focusedRef = useRef(false);
  useEffect(() => { if (!focusedRef.current) setDraft(value ?? ""); }, [value]);
  const commit = () => {
    const raw = draft;
    const newVal = raw === "" ? undefined : clampMark(raw, max);
    // Nothing actually changed (e.g. teacher tabbed through without typing) -
    // don't fire a save or a request for a no-op.
    const same = (newVal===undefined && (existingVal===undefined||existingVal===null)) || newVal===existingVal;
    if (!same) {
      const wasFirstEntry = existingVal===undefined || existingVal===null || existingVal==="";
      onCommit(newVal, existingVal);
      if (!wasFirstEntry) { setPendingFlash(true); setTimeout(()=>setPendingFlash(false), 1200); }
    }
  };
  // Once saved & locked, the cell is a genuinely disabled input -- not
  // focusable, not typable, no edit affordance of any kind. The only way
  // back in is the admin's "Unlock" button on the entry screen.
  return (
    <input type="number" min={0} max={max} value={locked ? (value ?? "") : draft} disabled={locked}
      onFocus={()=>{ focusedRef.current = true; }}
      onChange={e=>{
        const raw = e.target.value;
        if (raw === "") { setDraft(""); return; }
        const n = Number(raw);
        // Clamp on every keystroke, not just on blur, so the box can never
        // visually hold (or briefly show) a number outside 0..max.
        if (!isNaN(n)) setDraft(String(Math.max(0, Math.min(max, n))));
      }}
      onBlur={()=>{ focusedRef.current = false; commit(); }}
      onKeyDown={e=>{ if(e.key==="Enter"){ e.target.blur(); } }}
      style={{...style,
        ...(locked ? {background:"#f3f4f6", color:"#6b7280", cursor:"not-allowed"} : {}),
        ...(pendingFlash?{background:"#fef9c3",borderColor:"#f59e0b"}:{})}} />
  );
}
const btnPrimary = { padding:"8px 16px", background:"linear-gradient(135deg,#1e3a6e,#2563eb)", color:"white", border:"none", borderRadius:8, fontWeight:700, fontSize:13, cursor:"pointer" };
const btnWarning = { padding:"8px 16px", background:"linear-gradient(135deg,#d97706,#f59e0b)", color:"white", border:"none", borderRadius:8, fontWeight:700, fontSize:13, cursor:"pointer" };
const btnDanger  = { padding:"8px 16px", background:"linear-gradient(135deg,#dc2626,#ef4444)", color:"white", border:"none", borderRadius:8, fontWeight:700, fontSize:13, cursor:"pointer" };
const btnGhost   = { padding:"8px 16px", background:"white", color:"#374151", border:"1.5px solid #d1d5db", borderRadius:8, fontWeight:600, fontSize:13, cursor:"pointer" };
const btnSuccess = { padding:"8px 16px", background:"linear-gradient(135deg,#059669,#10b981)", color:"white", border:"none", borderRadius:8, fontWeight:700, fontSize:13, cursor:"pointer" };
const btnExcel   = { padding:"8px 16px", background:"linear-gradient(135deg,#0f7b3f,#16a34a)", color:"white", border:"none", borderRadius:8, fontWeight:700, fontSize:13, cursor:"pointer" };
const btnWord    = { padding:"8px 16px", background:"linear-gradient(135deg,#1e3a6e,#2b579a)", color:"white", border:"none", borderRadius:8, fontWeight:700, fontSize:13, cursor:"pointer" };
function Sel({ label, value, onChange, opts }) {
  return (
    <div>
      <label style={lbl}>{label}</label>
      <select value={value} onChange={e=>onChange(e.target.value)} style={inp}>
        {opts.map(o=><option key={o}>{o}</option>)}
      </select>
    </div>
  );
}
// ── Fixed PositionBadge: suffix sits inline to the RIGHT of the number ──
function PositionBadge({ pos, color = "#dc2626", size = 15 }) {
  if (pos === "-" || !pos) return <span style={{color, fontWeight:700, fontSize:size}}>-</span>;
  const suffix = ordinalSuffix(pos);
  return (
    <span style={{display:"inline-flex", alignItems:"flex-start", lineHeight:1, color, fontWeight:700}}>
      <span style={{fontSize:size}}>{pos}</span>
      <span style={{fontSize:Math.max(8, Math.round(size * 0.58)), marginTop:1, fontWeight:700}}>{suffix}</span>
    </span>
  );
}
// ── School crest: St. Kizito's P.S Tororo badge, reproduced as inline SVG ──
function SchoolCrest({ size = 64, ink = "#0f1115", paper = "#ffffff" }) {
  return (
    <svg viewBox="0 0 200 230" width={size} height={size * 1.15} role="img" aria-label="St. Kizito's school crest">
      <path d="M100,4 L116,16 C130,16 145,14 158,10 C160,40 160,70 162,95 C164,130 150,165 124,182 C114,189 106,193 100,196 C94,193 86,189 76,182 C50,165 36,130 38,95 C40,70 40,40 42,10 C55,14 70,16 84,16 Z"
        fill={paper} stroke={ink} strokeWidth="7" strokeLinejoin="round"/>
      <rect x="56" y="28" width="88" height="20" fill={ink}/>
      <text x="100" y="43" textAnchor="middle" fontFamily="Arial Black, Arial, sans-serif" fontSize="13" fontWeight="900" fill={paper} letterSpacing="0.5">ST. KIZITO'S</text>
      <text x="100" y="62" textAnchor="middle" fontFamily="Arial Black, Arial, sans-serif" fontSize="11.5" fontWeight="900" fill={ink} letterSpacing="0.5">P.S TORORO</text>
      <circle cx="100" cy="112" r="38" fill="none" stroke={ink} strokeWidth="5.5"/>
      <g stroke={ink} strokeWidth="2.4" fill="none" strokeLinecap="round" strokeLinejoin="round">
        <ellipse cx="100" cy="92" rx="10" ry="11" fill={ink} stroke="none"/>
        <path d="M91,101 C91,98 95,96 100,96 C105,96 109,98 109,101" fill="none"/>
        <path d="M89,102 C82,106 78,112 76,119 L120,108" fill="none"/>
        <path d="M111,102 C118,106 122,112 124,119 L80,108" fill="none"/>
        <path d="M72,122 L100,114 L128,122" fill="none"/>
        <path d="M72,126 L100,118 L128,126" fill="none"/>
        <path d="M66,125 C82,119 118,119 134,125 L134,130 C118,124 82,124 66,130 Z" fill={paper}/>
        <path d="M100,118 L100,130"/>
        <path d="M58,140 L142,140"/>
        <path d="M58,140 C70,150 88,150 100,142 C112,150 130,150 142,140" fill="none"/>
      </g>
      <path d="M14,200 L40,196 C70,206 130,206 160,196 L186,200 L182,212 L160,209 C130,217 70,217 40,209 L18,212 Z"
        fill={paper} stroke={ink} strokeWidth="4" strokeLinejoin="round"/>
      <path d="M14,200 L8,206 L18,212" fill="none" stroke={ink} strokeWidth="4"/>
      <path d="M186,200 L192,206 L182,212" fill="none" stroke={ink} strokeWidth="4"/>
      <text x="100" y="210" textAnchor="middle" fontFamily="Arial Black, Arial, sans-serif" fontSize="10.5" fontWeight="900" fill={ink} letterSpacing="0.3">BUILD FOR THE FUTURE</text>
    </svg>
  );
}
const PAGES = ["Dashboard","Mark Entry","Monthly Exams","Monthly Cards","Result Sheets","Report Cards","Students","Manage Requests","Settings","Audit Log","Download Centre"];
// Pages only the admin account can see/use. Teachers never see these in the sidebar.
const ADMIN_ONLY_PAGES = ["Manage Requests", "Settings", "Audit Log"];
// ─── APP ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage] = useState("Dashboard");
  const [students, setStudents] = useState([]);
  const [termMarks, setTermMarks] = useState({});
  const [monthlyMarks, setMonthlyMarks] = useState({});
  const [bands, setBands] = useState(DEFAULT_BANDS);
  const [divisions, setDivisions] = useState(DEFAULT_DIVISIONS);
  const [school, setSchool] = useState(DEFAULT_SCHOOL);
  const [accounts, setAccounts] = useState({});
  const [changeRequests, setChangeRequests] = useState([]);
  const [initials, setInitials] = useState({});
  // Locked entry-sessions: { "CLASS__TERM__YEAR": true } for term mark entry,
  // { "CLASS__TERM__YEAR__MONTH": true } for monthly exam entry. Set by the
  // Save button on those screens; once locked, MarkInput cells render
  // read-only and any further change must go through the existing
  // change-request / admin-approval flow.
  const [lockedTerm, setLockedTerm] = useState({});
  const [lockedMonthly, setLockedMonthly] = useState({});
  const [authed, setAuthed] = useState(false);
  const [role, setRole] = useState(null); // "teacher" | "admin"
  const [currentUser, setCurrentUser] = useState(null);
  const [loginUser, setLoginUser] = useState("");
  const [loginPw, setLoginPw] = useState("");
  const [loginErr, setLoginErr] = useState("");
  const [sideOpen, setSideOpen] = useState(true);
  const [dataReady, setDataReady] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [lastSyncedAt, setLastSyncedAt] = useState(null);
  // setter map + a ref mirror of the latest state, used by both the save
  // effects and the poll loop so neither has to be re-created on every render
  const setters = { mkis_students: setStudents, mkis_termmarks: setTermMarks, mkis_monthlymarks: setMonthlyMarks,
    mkis_bands: setBands, mkis_divisions: setDivisions, mkis_school: setSchool, mkis_accounts: setAccounts, mkis_changerequests: setChangeRequests, mkis_initials: setInitials,
    mkis_locked_term: setLockedTerm, mkis_locked_monthly: setLockedMonthly };
  const stateRef = useRef({});
  stateRef.current = { mkis_students: students, mkis_termmarks: termMarks, mkis_monthlymarks: monthlyMarks,
    mkis_bands: bands, mkis_divisions: divisions, mkis_school: school, mkis_accounts: accounts, mkis_changerequests: changeRequests, mkis_initials: initials,
    mkis_locked_term: lockedTerm, mkis_locked_monthly: lockedMonthly };
  // last value WE wrote (or loaded) per key, serialized -- used to tell "a
  // remote device changed this" apart from "this is just our own save echoing back"
  const lastSeenRef = useRef({});
  // brief guard so an incoming auto-refresh never overwrites a field the
  // person is actively typing into right this second
  const editingUntilRef = useRef(0);
  const markEditing = useCallback(() => { editingUntilRef.current = Date.now() + 2500; }, []);
  // Student ids this device has explicitly deleted. Needed because the merge
  // logic otherwise treats "missing from my local array" as "I haven't seen
  // it yet" and would resurrect it from the remote copy -- this set tells the
  // merge "no, this one was removed on purpose."
  const deletedStudentIdsRef = useRef(new Set());
  // Same idea as deletedStudentIdsRef, but for change-request records that
  // have been resolved (approved/rejected) and removed from the list.
  const deletedRequestIdsRef = useRef(new Set());
  // When a key is in this set, the NEXT save effect for that key writes the
  // local value as-is, bypassing the merge strategy entirely. Used only for
  // deliberate, user-confirmed full-backup restores, where overwriting
  // everything is the explicit point of the action. Consumed (removed) the
  // moment it's used so normal merge behavior resumes afterward.
  const forceWriteRef = useRef(new Set());
  // ── Initial load: pull everything from shared storage once on mount ──
  useEffect(() => {
    let mounted = true;
    (async () => {
      await migrateLocalStorageOnce();
      const [s, tm, mm, b, d, sc, acc, reqs, ini, lt, lm] = await Promise.all([
        loadShared("mkis_students", []),
        loadShared("mkis_termmarks", {}),
        loadShared("mkis_monthlymarks", {}),
        loadShared("mkis_bands", DEFAULT_BANDS),
        loadShared("mkis_divisions", DEFAULT_DIVISIONS),
        loadShared("mkis_school", DEFAULT_SCHOOL),
        loadShared("mkis_accounts", null),
        loadShared("mkis_changerequests", []),
        loadShared("mkis_initials", {}),
        loadShared("mkis_locked_term", {}),
        loadShared("mkis_locked_monthly", {}),
      ]);
      if (!mounted) return;
      setStudents(s); setTermMarks(tm); setMonthlyMarks(mm);
      // First run ever: no accounts saved yet, so seed the two default
      // logins (Teacher / admin) and persist them.
      let finalAccounts = acc;
      if (!finalAccounts || Object.keys(finalAccounts).length === 0) {
        finalAccounts = await buildDefaultAccounts();
        await saveShared("mkis_accounts", finalAccounts);
      }
      setBands(b); setDivisions(d); setSchool(sc); setAccounts(finalAccounts); setChangeRequests(reqs || []); setInitials(ini);
      setLockedTerm(lt || {}); setLockedMonthly(lm || {});
      lastSeenRef.current = { mkis_students: JSON.stringify(s), mkis_termmarks: JSON.stringify(tm),
        mkis_monthlymarks: JSON.stringify(mm), mkis_bands: JSON.stringify(b), mkis_divisions: JSON.stringify(d),
        mkis_school: JSON.stringify(sc), mkis_accounts: JSON.stringify(finalAccounts), mkis_changerequests: JSON.stringify(reqs || []), mkis_initials: JSON.stringify(ini),
        mkis_locked_term: JSON.stringify(lt || {}), mkis_locked_monthly: JSON.stringify(lm || {}) };
      setDataReady(true);
      setLastSyncedAt(new Date());
    })();
    return () => { mounted = false; };
  }, []);
  // ── Save effects: whenever local state changes, merge it into shared
  // storage instead of blindly overwriting it. Each effect re-reads the
  // CURRENT shared value, combines it with this device's local change via
  // the merge strategy for that key (see updateShared), writes the merged
  // result back, and then reconciles local state to match exactly what got
  // saved. That last step matters: if another device's change was folded in
  // by the merge, this device immediately sees it too, instead of silently
  // diverging from what's actually on disk. Skipped until the initial load
  // finishes, so we never merge shared data with empty defaults on first
  // render. ──
  useEffect(() => { if (dataReady) { (async () => {
    if (forceWriteRef.current.has("mkis_students")) {
      forceWriteRef.current.delete("mkis_students");
      await saveShared("mkis_students", students);
      lastSeenRef.current.mkis_students = JSON.stringify(students);
      return;
    }
    const merged = await updateShared("mkis_students", students, { deletedStudentIds: deletedStudentIdsRef.current });
    lastSeenRef.current.mkis_students = JSON.stringify(merged);
    if (JSON.stringify(merged) !== JSON.stringify(students)) setStudents(merged);
  })(); } }, [students, dataReady]);
  useEffect(() => { if (dataReady) { (async () => {
    if (forceWriteRef.current.has("mkis_termmarks")) {
      forceWriteRef.current.delete("mkis_termmarks");
      await saveShared("mkis_termmarks", termMarks);
      lastSeenRef.current.mkis_termmarks = JSON.stringify(termMarks);
      return;
    }
    const merged = await updateShared("mkis_termmarks", termMarks);
    lastSeenRef.current.mkis_termmarks = JSON.stringify(merged);
    if (JSON.stringify(merged) !== JSON.stringify(termMarks)) setTermMarks(merged);
  })(); } }, [termMarks, dataReady]);
  useEffect(() => { if (dataReady) { (async () => {
    if (forceWriteRef.current.has("mkis_monthlymarks")) {
      forceWriteRef.current.delete("mkis_monthlymarks");
      await saveShared("mkis_monthlymarks", monthlyMarks);
      lastSeenRef.current.mkis_monthlymarks = JSON.stringify(monthlyMarks);
      return;
    }
    const merged = await updateShared("mkis_monthlymarks", monthlyMarks);
    lastSeenRef.current.mkis_monthlymarks = JSON.stringify(merged);
    if (JSON.stringify(merged) !== JSON.stringify(monthlyMarks)) setMonthlyMarks(merged);
  })(); } }, [monthlyMarks, dataReady]);
  useEffect(() => { if (dataReady) { lastSeenRef.current.mkis_bands = JSON.stringify(bands); saveShared("mkis_bands", bands); } }, [bands, dataReady]);
  useEffect(() => { if (dataReady) { lastSeenRef.current.mkis_divisions = JSON.stringify(divisions); saveShared("mkis_divisions", divisions); } }, [divisions, dataReady]);
  useEffect(() => { if (dataReady) { lastSeenRef.current.mkis_school = JSON.stringify(school); saveShared("mkis_school", school); } }, [school, dataReady]);
  useEffect(() => { if (dataReady) { (async () => {
    const merged = await updateShared("mkis_accounts", accounts);
    lastSeenRef.current.mkis_accounts = JSON.stringify(merged);
    if (JSON.stringify(merged) !== JSON.stringify(accounts)) setAccounts(merged);
  })(); } }, [accounts, dataReady]);
  useEffect(() => { if (dataReady) { (async () => {
    const merged = await updateShared("mkis_changerequests", changeRequests, { deletedRequestIds: deletedRequestIdsRef.current });
    lastSeenRef.current.mkis_changerequests = JSON.stringify(merged);
    if (JSON.stringify(merged) !== JSON.stringify(changeRequests)) setChangeRequests(merged);
  })(); } }, [changeRequests, dataReady]);
  useEffect(() => { if (dataReady) { lastSeenRef.current.mkis_initials = JSON.stringify(initials); saveShared("mkis_initials", initials); } }, [initials, dataReady]);
  useEffect(() => { if (dataReady) { (async () => {
    const merged = await updateShared("mkis_locked_term", lockedTerm);
    lastSeenRef.current.mkis_locked_term = JSON.stringify(merged);
    if (JSON.stringify(merged) !== JSON.stringify(lockedTerm)) setLockedTerm(merged);
  })(); } }, [lockedTerm, dataReady]);
  useEffect(() => { if (dataReady) { (async () => {
    const merged = await updateShared("mkis_locked_monthly", lockedMonthly);
    lastSeenRef.current.mkis_locked_monthly = JSON.stringify(merged);
    if (JSON.stringify(merged) !== JSON.stringify(lockedMonthly)) setLockedMonthly(merged);
  })(); } }, [lockedMonthly, dataReady]);
  // ── Real-time auto-refresh: poll shared storage so that when another
  // device saves new data, this device picks it up automatically -- no
  // manual page reload needed. We only touch a key whose serialized value
  // actually differs from what we last saw, and we hold off on a key while
  // the person is mid-edit (within the editingUntil guard window) so we
  // never yank a field out from under someone who's actively typing. ──
  useEffect(() => {
    if (!dataReady) return;
    const POLL_MS = 4000;
    let stopped = false;
    const tick = async () => {
      if (stopped) return;
      const isEditing = Date.now() < editingUntilRef.current;
      try {
        setSyncing(true);
        const results = await Promise.all(STORAGE_KEYS.map(key => loadShared(key, undefined)));
        if (!stopped && !isEditing) {
          STORAGE_KEYS.forEach((key, i) => {
            const remoteVal = results[i];
            if (remoteVal === undefined) return;
            const remoteStr = JSON.stringify(remoteVal);
            if (remoteStr !== lastSeenRef.current[key]) {
              lastSeenRef.current[key] = remoteStr;
              setters[key](remoteVal);
            }
          });
          setLastSyncedAt(new Date());
        }
      } catch {}
      finally { if (!stopped) setSyncing(false); }
    };
    const id = setInterval(tick, POLL_MS);
    return () => { stopped = true; clearInterval(id); };
  }, [dataReady]);
  const updateTermMark = useCallback((sid, tk, sub, field, val) => {
    markEditing();
    setTermMarks(prev => ({
      ...prev,
      [sid]: { ...prev[sid], [tk]: { ...prev[sid]?.[tk], [sub]: { ...prev[sid]?.[tk]?.[sub], [field]: val } } }
    }));
  }, []);
  const updateMonthlyMark = useCallback((sid, tk, month, sub, field, val) => {
    markEditing();
    setMonthlyMarks(prev => ({
      ...prev,
      [sid]: { ...prev[sid], [tk]: { ...prev[sid]?.[tk], [month]: { ...prev[sid]?.[tk]?.[month], [sub]: { ...prev[sid]?.[tk]?.[month]?.[sub], [field]: val } } } }
    }));
  }, []);
  // ── Save & lock for Mark Entry / Monthly Exams ───────────────────────────
  // Clicking "Save" on an entry screen locks that class+term+year (or, for
  // monthly, that class+term+year+month) so the cells render read-only.
  // Only an admin can unlock it again; anyone else who needs to change a
  // locked value has to use the per-cell "request change" pencil, which
  // still funnels through the existing change-request/approval flow above.
  const lockTermEntry = useCallback((cls, tk) => {
    markEditing();
    setLockedTerm(prev => ({ ...prev, [`${cls}__${tk}`]: true }));
  }, []);
  const unlockTermEntry = useCallback((cls, tk) => {
    markEditing();
    setLockedTerm(prev => ({ ...prev, [`${cls}__${tk}`]: false }));
  }, []);
  const lockMonthlyEntry = useCallback((cls, tk, month) => {
    markEditing();
    setLockedMonthly(prev => ({ ...prev, [`${cls}__${tk}__${month}`]: true }));
  }, []);
  const unlockMonthlyEntry = useCallback((cls, tk, month) => {
    markEditing();
    setLockedMonthly(prev => ({ ...prev, [`${cls}__${tk}__${month}`]: false }));
  }, []);
  // ── Change-request workflow ──────────────────────────────────────────────
  // Anyone can enter a mark for the first time without approval. But once a
  // mark already has a value, *changing* it needs admin sign-off: instead of
  // writing straight to termMarks/monthlyMarks, we file a pending request
  // and the admin applies it (or rejects it) from the Manage Requests page.
  const submitChangeRequest = useCallback((req) => {
    markEditing();
    const full = {
      id: `${Date.now()}_${Math.random().toString(36).slice(2,8)}`,
      status: "pending",
      requestedAt: new Date().toISOString(),
      requestedBy: currentUser || "Unknown",
      ...req, // { kind: "term"|"monthly"|"unlock_term"|"unlock_monthly", ... }
    };
    setChangeRequests(prev => [...prev, full]);
  }, [currentUser]);
  // A teacher who hits a locked sheet can ask the admin to reopen it,
  // instead of being stuck with no recourse. This reuses the same
  // change-request queue/Manage Requests page as mark-change requests.
  const requestUnlockTerm = useCallback((cls, tk) => {
    submitChangeRequest({ kind: "unlock_term", cls, tk });
  }, [submitChangeRequest]);
  const requestUnlockMonthly = useCallback((cls, tk, month) => {
    submitChangeRequest({ kind: "unlock_monthly", cls, tk, month });
  }, [submitChangeRequest]);
  const approveChangeRequest = useCallback((id) => {
    setChangeRequests(prev => {
      const req = prev.find(r => r.id === id);
      if (!req) return prev;
      if (req.kind === "term") {
        updateTermMark(req.studentId, req.tk, req.sub, req.field, req.newVal);
      } else if (req.kind === "monthly") {
        updateMonthlyMark(req.studentId, req.tk, req.month, req.sub, req.field, req.newVal);
      } else if (req.kind === "unlock_term") {
        unlockTermEntry(req.cls, req.tk);
      } else if (req.kind === "unlock_monthly") {
        unlockMonthlyEntry(req.cls, req.tk, req.month);
      }
      deletedRequestIdsRef.current.add(id);
      return prev.filter(r => r.id !== id);
    });
  }, [updateTermMark, updateMonthlyMark, unlockTermEntry, unlockMonthlyEntry]);
  const rejectChangeRequest = useCallback((id) => {
    deletedRequestIdsRef.current.add(id);
    setChangeRequests(prev => prev.filter(r => r.id !== id));
  }, []);
  // Wrapper the Mark Entry / Monthly Exams screens call: admins always write
  // straight through (and can correct anything instantly); teachers write
  // straight through ONLY when the cell is currently empty (a first-time
  // entry, not a change), otherwise the edit is filed as a request.
  const requestOrApplyTermMark = useCallback((sid, studentName, tk, sub, field, val, existingVal) => {
    const isFirstEntry = existingVal === undefined || existingVal === null || existingVal === "";
    if (role === "admin" || isFirstEntry) {
      updateTermMark(sid, tk, sub, field, val);
    } else {
      submitChangeRequest({ kind: "term", studentId: sid, studentName, tk, sub, field, oldVal: existingVal, newVal: val });
    }
  }, [role, updateTermMark, submitChangeRequest]);
  const requestOrApplyMonthlyMark = useCallback((sid, studentName, tk, month, sub, field, val, existingVal) => {
    const isFirstEntry = existingVal === undefined || existingVal === null || existingVal === "";
    if (role === "admin" || isFirstEntry) {
      updateMonthlyMark(sid, tk, month, sub, field, val);
    } else {
      submitChangeRequest({ kind: "monthly", studentId: sid, studentName, tk, month, sub, field, oldVal: existingVal, newVal: val });
    }
  }, [role, updateMonthlyMark, submitChangeRequest]);
  const addStudent = useCallback((name, className, gender) => {
    markEditing();
    const newS = { id: Date.now().toString(), name: toUpper(name.trim()), className, gender };
    setStudents(prev => [...prev, newS].sort((a,b) => a.name.localeCompare(b.name)));
  }, []);
  const deleteStudent = useCallback((id) => {
    markEditing();
    deletedStudentIdsRef.current.add(id);
    setStudents(prev => prev.filter(s => s.id !== id));
  }, []);
  // Applies a full-backup restore. This is the one place a whole-object
  // overwrite is correct: the person explicitly chose a backup file and
  // confirmed they want it to replace current data. forceWriteRef tells the
  // save effects to skip merging for this write only, so the restored
  // snapshot lands exactly as-is instead of being blended with whatever was
  // previously in shared storage.
  const forceRestoreData = useCallback((d) => {
    if (d.students)     { forceWriteRef.current.add("mkis_students"); setStudents(d.students); }
    if (d.termMarks)    { forceWriteRef.current.add("mkis_termmarks"); setTermMarks(d.termMarks); }
    if (d.monthlyMarks) { forceWriteRef.current.add("mkis_monthlymarks"); setMonthlyMarks(d.monthlyMarks); }
    if (d.bands)         setBands(d.bands);
    if (d.divisions)     setDivisions(d.divisions);
    if (d.school)        setSchool(d.school);
    if (d.accounts) {
      setAccounts(d.accounts);
    }
    deletedStudentIdsRef.current = new Set();
  }, []);
  const promoteStudents = useCallback((fromClass) => {
    markEditing();
    const classMap = {"P1":"P2","P2":"P3","P3":"P4","P4":"P5","P5":"P6","P6":"P7","P7":"Completed"};
    setStudents(prev => prev.map(s => s.className === fromClass ? {...s, className: classMap[fromClass] || s.className} : s));
  }, []);
  if (!dataReady) {
    return (
      <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"#f1f5f9",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
        <div style={{textAlign:"center",color:"#1e3a6e"}}>
          <div style={{width:48,height:48,border:"4px solid #cbd5e1",borderTopColor:"#2563eb",borderRadius:"50%",margin:"0 auto 14px",animation:"mkis-spin 0.8s linear infinite"}} />
          <div style={{fontWeight:700,fontSize:14}}>Loading shared data…</div>
          <style>{`@keyframes mkis-spin { to { transform: rotate(360deg); } }`}</style>
        </div>
      </div>
    );
  }
  if (!authed) {
    const doLogin = () => {
      verifyAccountLogin(accounts, loginUser, loginPw).then(acct => {
        if (acct) {
          setAuthed(true); setRole(acct.role); setCurrentUser(acct.username); setLoginErr(""); setLoginPw("");
        } else {
          setLoginErr("Incorrect username or password");
        }
      });
    };
    return (
      <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#1e3a6e 0%,#2563eb 100%)",display:"flex",alignItems:"center",justifyContent:"center"}}>
        <div style={{background:"white",borderRadius:16,padding:40,width:360,boxShadow:"0 20px 60px rgba(0,0,0,0.3)"}}>
          <div style={{textAlign:"center",marginBottom:24}}>
            <div style={{margin:"0 auto 12px",display:"flex",justifyContent:"center"}}><SchoolCrest size={64}/></div>
            <div style={{fontWeight:800,fontSize:18,color:"#1e3a6e"}}>{school.name}</div>
            <div style={{fontSize:12,color:"#666",marginTop:4}}>Results Management System v2</div>
          </div>
          <div style={{marginBottom:14}}>
            <label style={lbl}>USERNAME</label>
            <input type="text" value={loginUser} onChange={e=>setLoginUser(e.target.value)}
              onKeyDown={e=>{ if(e.key==="Enter") doLogin(); }}
              style={{width:"100%",padding:"10px 12px",border:"2px solid #e5e7eb",borderRadius:8,fontSize:14,boxSizing:"border-box"}}
              placeholder="Enter username" autoFocus />
          </div>
          <div style={{marginBottom:16}}>
            <label style={lbl}>PASSWORD</label>
            <input type="password" value={loginPw} onChange={e=>setLoginPw(e.target.value)}
              onKeyDown={e=>{ if(e.key==="Enter") doLogin(); }}
              style={{width:"100%",padding:"10px 12px",border:"2px solid #e5e7eb",borderRadius:8,fontSize:14,boxSizing:"border-box"}}
              placeholder="Enter password" />
          </div>
          {loginErr && <div style={{color:"#dc2626",fontSize:13,marginBottom:12}}>{loginErr}</div>}
          <button onClick={doLogin}
            style={{width:"100%",padding:"11px",background:"linear-gradient(135deg,#1e3a6e,#2563eb)",color:"white",border:"none",borderRadius:8,fontWeight:700,fontSize:15,cursor:"pointer"}}>
            Login
          </button>
          <div style={{textAlign:"center",fontSize:11,color:"#9ca3af",marginTop:16}}>Contact your administrator if you have forgotten the password.</div>
        </div>
      </div>
    );
  }
  const props = { students, setStudents, termMarks, setTermMarks, monthlyMarks, setMonthlyMarks, bands, setBands, divisions, setDivisions, school, setSchool, accounts, setAccounts, initials, setInitials, updateTermMark, updateMonthlyMark, requestOrApplyTermMark, requestOrApplyMonthlyMark, addStudent, deleteStudent, forceRestoreData, promoteStudents, role, currentUser, changeRequests, submitChangeRequest, approveChangeRequest, rejectChangeRequest, lockedTerm, lockTermEntry, unlockTermEntry, lockedMonthly, lockMonthlyEntry, unlockMonthlyEntry, requestUnlockTerm, requestUnlockMonthly };
  return (
    <div style={{display:"flex",minHeight:"100vh",fontFamily:"'Segoe UI',system-ui,sans-serif",background:"#f1f5f9"}}>
      {/* SIDEBAR */}
      <div style={{width:sideOpen?230:60,background:"linear-gradient(180deg,#1e3a6e 0%,#1e40af 100%)",color:"white",transition:"width 0.2s",overflow:"hidden",flexShrink:0,display:"flex",flexDirection:"column"}}>
        <div style={{padding:"16px 12px",borderBottom:"1px solid rgba(255,255,255,0.1)",display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:36,height:36,borderRadius:"50%",background:"rgba(255,255,255,0.92)",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,padding:3,boxSizing:"border-box"}}>
            <SchoolCrest size={30} ink="#1e3a6e" paper="#ffffff"/>
          </div>
          {sideOpen && <div style={{fontSize:11,fontWeight:700,lineHeight:1.3}}>ST. KIZITO'S<br/><span style={{fontWeight:400,opacity:0.8}}>Results MIS v2</span></div>}
        </div>
        <nav style={{flex:1,padding:"8px 0"}}>
          {PAGES.filter(p => !ADMIN_ONLY_PAGES.includes(p) || role==="admin").map(p => {
            const icons = {"Dashboard":"📊","Mark Entry":"📝","Monthly Exams":"📅","Monthly Cards":"🗂️","Result Sheets":"📋","Report Cards":"🎓","Students":"👥","Manage Requests":"🛂","Settings":"⚙️","Audit Log":"🕓","Download Centre":"📥"};
            const pendingCount = p==="Manage Requests" ? changeRequests.filter(r=>r.status==="pending").length : 0;
            return (
              <button key={p} onClick={()=>setPage(p)}
                style={{width:"100%",padding:"10px 14px",background:page===p?"rgba(255,255,255,0.2)":"transparent",border:"none",color:"white",textAlign:"left",cursor:"pointer",display:"flex",alignItems:"center",gap:10,fontSize:13,fontWeight:page===p?700:400,borderLeft:page===p?"3px solid #60a5fa":"3px solid transparent"}}>
                <span style={{fontSize:16,flexShrink:0}}>{icons[p]}</span>
                {sideOpen && <span style={{flex:1,display:"flex",alignItems:"center",justifyContent:"space-between"}}>{p}{pendingCount>0 && <span style={{background:"#dc2626",color:"white",borderRadius:10,fontSize:10,fontWeight:800,padding:"1px 7px"}}>{pendingCount}</span>}</span>}
              </button>
            );
          })}
        </nav>
        {sideOpen && (
          <div style={{padding:"10px 14px",fontSize:11,color:"rgba(255,255,255,0.65)",borderTop:"1px solid rgba(255,255,255,0.1)"}}>
            Signed in as <b style={{color:"white"}}>{currentUser}</b> ({role === "admin" ? "Admin" : "Teacher"})
          </div>
        )}
        <button onClick={()=>{ setAuthed(false); setRole(null); setCurrentUser(null); setLoginUser(""); setPage("Dashboard"); }} style={{padding:"12px 14px",background:"transparent",border:"none",color:"rgba(255,255,255,0.6)",textAlign:"left",cursor:"pointer",fontSize:13,display:"flex",alignItems:"center",gap:10}}>
          <span>🚪</span>{sideOpen && "Logout"}
        </button>
      </div>
      {/* MAIN */}
      <div style={{flex:1,overflow:"auto"}}>
        <div style={{background:"white",padding:"12px 20px",borderBottom:"1px solid #e5e7eb",display:"flex",alignItems:"center",gap:12,position:"sticky",top:0,zIndex:10}}>
          <button onClick={()=>setSideOpen(v=>!v)} style={{background:"none",border:"none",cursor:"pointer",fontSize:20,color:"#374151"}}>☰</button>
          <h1 style={{fontSize:18,fontWeight:700,color:"#1e3a6e",margin:0}}>{page}</h1>
          <div style={{marginLeft:"auto",fontSize:12,color:"#6b7280"}}>{school.name} - {school.year}</div>
        </div>
        <div style={{padding:20}}>
          {page==="Dashboard" && <Dashboard {...props} />}
          {page==="Mark Entry" && <MarkEntry {...props} />}
          {page==="Monthly Exams" && <MonthlyExams {...props} />}
          {page==="Monthly Cards" && <MonthlyCards {...props} />}
          {page==="Result Sheets" && <ResultSheets {...props} />}
          {page==="Report Cards" && <ReportCards {...props} />}
          {page==="Students" && <Students {...props} />}
          {page==="Manage Requests" && role==="admin" && <ManageRequests {...props} />}
          {page==="Settings" && role==="admin" && <Settings {...props} />}
          {page==="Audit Log" && role==="admin" && <AuditLog />}
          {page==="Download Centre" && <DownloadCentre {...props} />}
        </div>
      </div>
      <style>{`
        @media print {
          @page { size: A4 portrait; margin: 8mm; }
          .no-print { display: none !important; }
          .page-break { page-break-after: always; }
          /* Report Card: every pupil's full card (school details, pupil
             info, subjects, grades, comments, signatures) stays together
             as one unbroken block, and each card starts a fresh page --
             except the last one, which should not leave a trailing blank
             page after it. */
          .report-card-sheet {
            page-break-inside: avoid;
            break-inside: avoid;
            page-break-after: always;
            break-after: page;
          }
          .report-card-sheet:last-child {
            page-break-after: auto;
            break-after: auto;
          }
          .report-card-sheet table,
          .report-card-sheet tr,
          .report-card-sheet thead,
          .report-card-sheet tbody {
            page-break-inside: avoid;
            break-inside: avoid;
          }
        }
        /* Screen preview: stack report cards as plain blocks (not flex) so
           the same page-break rules above apply predictably when printed. */
        .report-card-list { display: block; }
        .report-card-sheet { margin: 0 auto 24px; }
        @media print {
          .report-card-sheet { margin: 0 auto; }
        }
        input[type=number]::-webkit-inner-spin-button { opacity:1; }
        * { box-sizing: border-box; }
        table { border-collapse: collapse; }
        th, td { border: 1px solid #d1d5db; }
      `}</style>
    </div>
  );
}
// ─── DASHBOARD ───────────────────────────────────────────────────────────────
function Dashboard({ students, school }) {
  const active = students.filter(s=>s.className!=="Completed");
  const total = active.length;
  const boys = active.filter(s=>s.gender==="M").length;
  const girls = active.filter(s=>s.gender==="F").length;
  const classCounts = ALL_CLASSES.map(c => ({ cls:c, count: students.filter(s=>s.className===c).length }));
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:16,marginBottom:24}}>
        {[
          { label:"Total Pupils", value:total, color:"#1e40af", bg:"#dbeafe" },
          { label:"Boys", value:boys, color:"#065f46", bg:"#d1fae5" },
          { label:"Girls", value:girls, color:"#7c2d12", bg:"#fee2e2" },
          { label:"Classes", value:ALL_CLASSES.length, color:"#6b21a8", bg:"#f3e8ff" },
        ].map(c => (
          <div key={c.label} style={{background:c.bg,borderRadius:12,padding:20,borderLeft:`4px solid ${c.color}`}}>
            <div style={{fontSize:32,fontWeight:800,color:c.color}}>{c.value}</div>
            <div style={{fontSize:13,color:c.color,fontWeight:600}}>{c.label}</div>
          </div>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <div style={{background:"white",borderRadius:12,padding:20,border:"1px solid #e5e7eb"}}>
          <h3 style={{margin:"0 0 16px",color:"#1e3a6e",fontSize:15}}>Enrolment by Class</h3>
          {classCounts.map(({cls,count}) => (
            <div key={cls} style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
              <div style={{width:40,fontWeight:700,color:"#1e3a6e",fontSize:13}}>{cls}</div>
              <div style={{flex:1,background:"#f1f5f9",borderRadius:4,height:20,overflow:"hidden"}}>
                <div style={{width:`${total?Math.round(count/total*100):0}%`,background:"linear-gradient(90deg,#1e40af,#3b82f6)",height:"100%",borderRadius:4}}/>
              </div>
              <div style={{width:30,textAlign:"right",fontSize:13,fontWeight:600}}>{count}</div>
            </div>
          ))}
        </div>
        <div style={{background:"white",borderRadius:12,padding:20,border:"1px solid #e5e7eb"}}>
          <h3 style={{margin:"0 0 16px",color:"#1e3a6e",fontSize:15}}>School Info</h3>
          {[["School",school.name],["Motto",school.motto],["P.O. Box",school.poBox],["District",school.district],["Head Teacher",school.headTeacher],["Year",school.year]].map(([k,v])=>
            v ? <div key={k} style={{fontSize:13,marginBottom:8,display:"flex",gap:8}}><span style={{fontWeight:600,color:"#374151",minWidth:90}}>{k}:</span><span style={{color:"#6b7280"}}>{v}</span></div> : null
          )}
        </div>
      </div>
    </div>
  );
}
// ─── STUDENTS ────────────────────────────────────────────────────────────────
function Students({ students, setStudents, addStudent, deleteStudent, promoteStudents }) {
  const [name, setName] = useState("");
  const [cls, setCls] = useState("P1");
  const [gender, setGender] = useState("M");
  const [search, setSearch] = useState("");
  const [filterCls, setFilterCls] = useState("All");
  const [filterGender, setFilterGender] = useState("All");
  const [promoteClass, setPromoteClass] = useState("P1");
  const [showPromote, setShowPromote] = useState(false);
  const [editId, setEditId] = useState(null);
  const [editName, setEditName] = useState("");
  const [editCls, setEditCls] = useState("P1");
  const [editGender, setEditGender] = useState("M");
  const [bulkPreview, setBulkPreview] = useState(null);
  const [bulkCls, setBulkCls] = useState("P1");
  const [bulkGender, setBulkGender] = useState("M");
  const [showBulk, setShowBulk] = useState(false);
  const [bulkText, setBulkText] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const confirmDeleteStudent = students.find(s => s.id === confirmDeleteId);
  const filtered = useMemo(() =>
    students.filter(s =>
      (filterCls==="All"||s.className===filterCls) &&
      (s.className!=="Completed") &&
      (filterGender==="All"||s.gender===filterGender) &&
      s.name.toLowerCase().includes(search.toLowerCase())
    ).sort((a,b)=>a.name.localeCompare(b.name)),
    [students, filterCls, filterGender, search]
  );
  const handleAdd = () => {
    if (!name.trim()) return;
    addStudent(name, cls, gender);
    setName("");
  };
  const handleDelete = (id) => {
    setConfirmDeleteId(id);
  };
  const handleEdit = (s) => { setEditId(s.id); setEditName(s.name); setEditCls(s.className); setEditGender(s.gender); };
  const handleSaveEdit = (id) => {
    setStudents(prev => prev.map(s=>s.id===id?{...s,name:toUpper(editName.trim()),className:editCls,gender:editGender}:s).sort((a,b)=>a.name.localeCompare(b.name)));
    setEditId(null);
  };
  const handleBulkTextPreview = () => {
    if (!bulkText.trim()) return;
    const names = bulkText.split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
    const preview = names.map((n,i)=>({ id:`bulk_${Date.now()}_${i}`, name:toUpper(n), className:bulkCls, gender:bulkGender, keep:true }));
    setBulkPreview(preview);
  };
  const confirmBulk = () => {
    if (!bulkPreview) return;
    const toAdd = bulkPreview.filter(r=>r.keep&&r.name.trim());
    setStudents(prev => [...prev, ...toAdd.map(r=>({id:Date.now().toString()+Math.random(),name:r.name,className:r.className,gender:r.gender}))].sort((a,b)=>a.name.localeCompare(b.name)));
    setBulkPreview(null); setShowBulk(false); setBulkText("");
  };
  return (
    <div>
      <div style={{background:"white",borderRadius:12,padding:20,border:"1px solid #e5e7eb",marginBottom:16}}>
        <h3 style={{margin:"0 0 16px",color:"#1e3a6e",fontSize:15,fontWeight:700}}>Add New Student</h3>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end"}}>
          <div><label style={lbl}>Full Name</label><input value={name} onChange={e=>setName(toUpper(e.target.value))} onKeyDown={e=>e.key==="Enter"&&handleAdd()} style={{...inp,textTransform:"uppercase"}} placeholder="e.g. AKELLO TIM" /></div>
          <div><label style={lbl}>Class</label><select value={cls} onChange={e=>setCls(e.target.value)} style={inp}>{ALL_CLASSES.map(c=><option key={c}>{c}</option>)}</select></div>
          <div><label style={lbl}>Gender</label><select value={gender} onChange={e=>setGender(e.target.value)} style={inp}><option value="M">Male</option><option value="F">Female</option></select></div>
          <button onClick={handleAdd} style={btnPrimary}>+ Add Student</button>
          <button onClick={()=>setShowBulk(v=>!v)} style={btnWarning}>📋 Bulk Import</button>
        </div>
      </div>
      {showBulk && (
        <div style={{background:"#f0fdf4",border:"2px solid #22c55e",borderRadius:12,padding:20,marginBottom:16}}>
          <h3 style={{margin:"0 0 12px",color:"#15803d",fontSize:14}}>📋 Bulk Import Pupils</h3>
          <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap",marginBottom:12}}>
            <div><label style={lbl}>Assign to Class</label><select value={bulkCls} onChange={e=>setBulkCls(e.target.value)} style={inp}>{ALL_CLASSES.map(c=><option key={c}>{c}</option>)}</select></div>
            <div><label style={lbl}>Default Gender</label><select value={bulkGender} onChange={e=>setBulkGender(e.target.value)} style={inp}><option value="M">Male</option><option value="F">Female</option></select></div>
            <button onClick={()=>{setShowBulk(false);setBulkPreview(null);setBulkText("");}} style={btnGhost}>Cancel</button>
          </div>
          <div style={{marginBottom:10}}>
            <label style={lbl}>Names (one per line)</label>
            <textarea
              value={bulkText}
              onChange={e=>setBulkText(e.target.value)}
              placeholder={"AKELLO SARAH\nOKELLO TIM\nNAMUKASA GRACE"}
              style={{width:"100%",minHeight:120,padding:"8px 10px",border:"1.5px solid #d1d5db",borderRadius:7,fontSize:13,fontFamily:"inherit",resize:"vertical",textTransform:"uppercase",boxSizing:"border-box"}}
            />
            <div style={{fontSize:11,color:"#6b7280",marginTop:3}}>Type or paste pupil names - one per line. You can adjust gender and class individually in the preview.</div>
          </div>
          <button onClick={handleBulkTextPreview} disabled={!bulkText.trim()} style={{...btnPrimary,opacity:bulkText.trim()?1:0.5}}>🔍 Preview ({bulkText.split(/\r?\n/).filter(l=>l.trim()).length} names)</button>
          {bulkPreview && (
            <>
              <div style={{maxHeight:300,overflowY:"auto",marginTop:12,marginBottom:12,border:"1px solid #d1fae5",borderRadius:8}}>
                <table style={{width:"100%",fontSize:12}}>
                  <thead><tr style={{background:"#dcfce7"}}><th style={th}>Include</th><th style={{...th,textAlign:"left"}}>S/N</th><th style={{...th,textAlign:"left"}}>Name</th><th style={th}>Class</th><th style={th}>Gender</th></tr></thead>
                  <tbody>
                    {bulkPreview.map((r,i)=>(
                      <tr key={r.id} style={{background:i%2===0?"white":"#f0fdf4"}}>
                        <td style={{...td,textAlign:"center"}}><input type="checkbox" checked={r.keep} onChange={()=>setBulkPreview(prev=>prev.map((x,j)=>j===i?{...x,keep:!x.keep}:x))}/></td>
                        <td style={td}>{i+1}</td>
                        <td style={td}><input value={r.name} onChange={e=>setBulkPreview(prev=>prev.map((x,j)=>j===i?{...x,name:toUpper(e.target.value)}:x))} style={{...inp,padding:"2px 6px",fontSize:12,width:"100%",textTransform:"uppercase"}}/></td>
                        <td style={td}><select value={r.className} onChange={e=>setBulkPreview(prev=>prev.map((x,j)=>j===i?{...x,className:e.target.value}:x))} style={{...inp,padding:"2px 6px",fontSize:12}}>{ALL_CLASSES.map(c=><option key={c}>{c}</option>)}</select></td>
                        <td style={td}><select value={r.gender} onChange={e=>setBulkPreview(prev=>prev.map((x,j)=>j===i?{...x,gender:e.target.value}:x))} style={{...inp,padding:"2px 6px",fontSize:12}}><option value="M">M</option><option value="F">F</option></select></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{display:"flex",gap:10}}>
                <button onClick={confirmBulk} style={btnPrimary}>💾 Save {bulkPreview.filter(r=>r.keep).length} Pupils</button>
                <button onClick={()=>setBulkPreview(null)} style={btnGhost}>Clear Preview</button>
              </div>
            </>
          )}
        </div>
      )}
      <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} style={{...inp,width:200}} placeholder="🔍 Search student..." />
        <select value={filterCls} onChange={e=>setFilterCls(e.target.value)} style={inp}><option value="All">All Classes</option>{ALL_CLASSES.map(c=><option key={c}>{c}</option>)}</select>
        <select value={filterGender} onChange={e=>setFilterGender(e.target.value)} style={inp}><option value="All">All Genders</option><option value="M">👦 Males only</option><option value="F">👧 Females only</option></select>
        <div style={{marginLeft:"auto"}}><button onClick={()=>setShowPromote(v=>!v)} style={btnWarning}>🎓 Promote Students</button></div>
      </div>
      {showPromote && (
        <div style={{background:"#fefce8",border:"2px solid #f59e0b",borderRadius:12,padding:16,marginBottom:16,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
          <span style={{fontWeight:700,color:"#92400e"}}>Promote class:</span>
          <select value={promoteClass} onChange={e=>setPromoteClass(e.target.value)} style={inp}>{ALL_CLASSES.map(c=><option key={c}>{c}</option>)}</select>
          <span style={{color:"#92400e",fontSize:13}}>➡️ {{"P1":"P2","P2":"P3","P3":"P4","P4":"P5","P5":"P6","P6":"P7","P7":"Completed"}[promoteClass]}</span>
          <button onClick={()=>{ if(window.confirm(`Promote all ${promoteClass} students?`)){promoteStudents(promoteClass);setShowPromote(false);} }} style={btnPrimary}>Confirm Promote</button>
          <button onClick={()=>setShowPromote(false)} style={btnGhost}>Cancel</button>
        </div>
      )}
      <div style={{background:"white",borderRadius:12,border:"1px solid #e5e7eb",overflow:"hidden"}}>
        <div style={{padding:"12px 16px",background:"#1e3a6e",color:"white",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontWeight:700}}>Students ({filtered.length})</span>
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",fontSize:13}}>
            <thead>
              <tr style={{background:"#dbeafe"}}>
                {["#","Name","Class","Gender","Actions"].map(h=>(
                  <th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,color:"#1e3a6e",whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((s,i)=>(
                <tr key={s.id} style={{background:i%2===0?"white":"#f8fafc"}}>
                  <td style={{padding:"9px 12px",color:"#6b7280"}}>{i+1}</td>
                  <td style={{padding:"9px 12px",fontWeight:600}}>
                    {editId===s.id ? <input value={editName} onChange={e=>setEditName(toUpper(e.target.value))} style={{...inp,padding:"4px 8px",width:200,textTransform:"uppercase"}} /> : s.name}
                  </td>
                  <td style={{padding:"9px 12px"}}>
                    {editId===s.id
                      ? <select value={editCls} onChange={e=>setEditCls(e.target.value)} style={{...inp,padding:"4px 8px",fontSize:12}}>{ALL_CLASSES.map(c=><option key={c}>{c}</option>)}</select>
                      : <span style={{background:"#dbeafe",color:"#1e40af",borderRadius:6,padding:"2px 8px",fontWeight:700,fontSize:12}}>{s.className}</span>}
                  </td>
                  <td style={{padding:"9px 12px"}}>
                    {editId===s.id
                      ? <select value={editGender} onChange={e=>setEditGender(e.target.value)} style={{...inp,padding:"4px 8px",fontSize:12}}><option value="M">Male</option><option value="F">Female</option></select>
                      : (s.gender==="M"?"👦 Male":"👧 Female")}
                  </td>
                  <td style={{padding:"9px 12px",display:"flex",gap:6}}>
                    {editId===s.id
                      ? <><button onClick={()=>handleSaveEdit(s.id)} style={{...btnPrimary,padding:"4px 10px",fontSize:12}}>Save</button><button onClick={()=>setEditId(null)} style={{...btnGhost,padding:"4px 10px",fontSize:12}}>Cancel</button></>
                      : <><button onClick={()=>handleEdit(s)} style={{...btnGhost,padding:"4px 10px",fontSize:12}}>Edit</button><button onClick={()=>handleDelete(s.id)} style={{...btnDanger,padding:"4px 10px",fontSize:12}}>Delete</button></>
                    }
                  </td>
                </tr>
              ))}
              {filtered.length===0 && <tr><td colSpan={5} style={{padding:24,textAlign:"center",color:"#9ca3af"}}>No students found.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
      {confirmDeleteStudent && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000,padding:16}}>
          <div style={{background:"white",borderRadius:16,padding:28,width:"100%",maxWidth:380,boxShadow:"0 24px 80px rgba(0,0,0,0.4)"}}>
            <div style={{fontSize:36,textAlign:"center",marginBottom:8}}>🗑️</div>
            <h3 style={{margin:"0 0 16px",color:"#991b1b",fontSize:16,fontWeight:800,textAlign:"center"}}>Delete Student</h3>
            <p style={{margin:"0 0 20px",fontSize:14,color:"#374151",textAlign:"center",lineHeight:1.6}}>Are you sure you want to delete <b>{confirmDeleteStudent.name}</b>?</p>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>{deleteStudent(confirmDeleteId);setConfirmDeleteId(null);}} style={{flex:1,padding:"11px",background:"#dc2626",color:"white",border:"none",borderRadius:8,fontWeight:700,fontSize:14,cursor:"pointer"}}>Delete</button>
              <button onClick={()=>setConfirmDeleteId(null)} style={{...btnGhost,padding:"11px 20px"}}>Cancel</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
// ─── MARK ENTRY ──────────────────────────────────────────────────────────────
function MarkEntry({ students, termMarks, setTermMarks, updateTermMark, requestOrApplyTermMark, role, bands, divisions, school, lockedTerm, lockTermEntry, unlockTermEntry, changeRequests, requestUnlockTerm }) {
  const [cls, setCls] = useState("P1");
  const [term, setTerm] = useState("Term I");
  const [year, setYear] = useState(school.year||String(new Date().getFullYear()));
  const [showBulkMark, setShowBulkMark] = useState(false);
  const [bulkMarkPreview, setBulkMarkPreview] = useState(null);
  const [bulkMarkError, setBulkMarkError] = useState("");
  const [pendingToast, setPendingToast] = useState("");
  const [sortByPos, setSortByPos] = useState(false);
  const bulkMarkFileRef = useRef();
  // Wraps requestOrApplyTermMark to surface a brief toast whenever an edit
  // was filed for admin approval rather than saved immediately, so teachers
  // aren't left wondering why a number they typed doesn't show up yet.
  const handleMarkChange = useCallback((sid, name, tk, sub, field, val, existingVal) => {
    requestOrApplyTermMark(sid, name, tk, sub, field, val, existingVal);
    const isFirstEntry = existingVal === undefined || existingVal === null || existingVal === "";
    if (role !== "admin" && !isFirstEntry) {
      setPendingToast(`Change to ${name}'s ${sub} mark sent to admin for approval.`);
      setTimeout(()=>setPendingToast(""), 3500);
    }
  }, [requestOrApplyTermMark, role]);
  const isLower = LOWER_CLASSES.includes(cls);
  const subjects = isLower ? LOWER_SUBJECTS : UPPER_SUBJECTS;
  const tk = `${term}__${year}`;
  // Save & lock: once clicked, marks for this class+term+year render
  // read-only (see MarkInput) until an admin unlocks them again.
  const lockKey = `${cls}__${tk}`;
  const isLocked = !!lockedTerm?.[lockKey];
  // Has a teacher already asked the admin to reopen this exact sheet?
  // Checked against the live change-request queue so the button can't be
  // double-clicked into duplicate requests.
  const unlockRequestPending = useMemo(
    () => (changeRequests||[]).some(r => r.kind==="unlock_term" && r.cls===cls && r.tk===tk),
    [changeRequests, cls, tk]
  );
  const handleSave = useCallback(() => {
    lockTermEntry(cls, tk);
    setPendingToast(`${cls} ${term} ${year} marks saved and locked.`);
    setTimeout(()=>setPendingToast(""), 3500);
  }, [lockTermEntry, cls, tk, term, year]);
  const handleUnlock = useCallback(() => {
    unlockTermEntry(cls, tk);
    setPendingToast(`${cls} ${term} ${year} marks unlocked for editing.`);
    setTimeout(()=>setPendingToast(""), 3500);
  }, [unlockTermEntry, cls, tk, term, year]);
  const handleRequestUnlock = useCallback(() => {
    requestUnlockTerm(cls, tk);
    setPendingToast(`Unlock request for ${cls} ${term} ${year} sent to admin.`);
    setTimeout(()=>setPendingToast(""), 3500);
  }, [requestUnlockTerm, cls, tk, term, year]);
  const classStudents = useMemo(()=>
    students.filter(s=>s.className===cls).sort((a,b)=>a.name.localeCompare(b.name)),
    [students, cls]
  );
  const rows = useMemo(()=> classStudents.map(s => {
    const m = termMarks[s.id]?.[tk] || {};
    const perSub = subjects.map(sub => {
      const ca = m[sub]?.ca;
      const exam = m[sub]?.exam;
      const isX = isLower ? (exam===undefined||exam===null) : (ca===undefined||ca===null) && (exam===undefined||exam===null);
      if (isX) return { sub, ca, exam, av: undefined, agg: undefined, isX: true };
      const hasBoth = typeof ca==="number" && typeof exam==="number";
      const av = hasBoth ? Math.round((ca+exam)/2) : (typeof exam==="number"?exam:typeof ca==="number"?ca:undefined);
      const agg = av !== undefined ? aggOf(av, bands) : undefined;
      return { sub, ca, exam, av, agg, isX: false };
    });
    const hasX = perSub.some(p => p.isX);
    const totMk = perSub.reduce((a,p)=>a+(p.av??0),0);
    const totAgg = hasX ? "X" : perSub.reduce((a,p)=>a+(p.agg||0),0);
    const div = hasX ? "X" : (typeof totAgg==="number" ? divisionOf(totAgg, isLower?5:4, divisions) : "X");
    return { s, perSub, totMk, totAgg, div, hasX };
  }), [classStudents, termMarks, tk, subjects, bands, divisions, isLower]);
  const positions = useMemo(()=> rankWithTies(rows.map(r=>r.totMk>0?r.totMk:null), rows.map(r=>typeof r.totAgg==="number"?r.totAgg:null)), [rows]);
  const indexedRows = useMemo(()=> rows.map((r,i)=>({...r, pos:positions[i]})), [rows, positions]);
  const sortedRows = useMemo(()=>{
    return [...indexedRows].sort((a,b)=>{ if(a.pos==="-") return 1; if(b.pos==="-") return -1; return a.pos - b.pos; });
  }, [indexedRows]);
  const handleBulkMarkFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setBulkMarkError("");
    try {
      const text = await file.text();
      const lines = text.split(/\r?\n/).filter(l=>l.trim());
      if (lines.length < 2) { setBulkMarkError("File must have a header row and at least one data row."); return; }
      const headers = lines[0].split(",").map(h=>h.trim().toUpperCase());
      const nameIdx = headers.findIndex(h=>h==="NAME"||h==="PUPIL"||h==="STUDENT");
      if (nameIdx === -1) { setBulkMarkError("CSV must have a column named NAME, PUPIL, or STUDENT."); return; }
      const curSubjects = isLower ? LOWER_SUBJECTS : UPPER_SUBJECTS;
      // Build column map: subject -> { caIdx, examIdx } or { markIdx }
      const subjectCols = {};
      curSubjects.forEach(sub => {
        if (isLower) {
          const idx = headers.findIndex(h=>h===sub||h===sub.replace(" ","_"));
          if (idx !== -1) subjectCols[sub] = { markIdx: idx };
        } else {
          const caIdx = headers.findIndex(h=>h===`${sub}_CA`||h===`${sub} CA`||h===`${sub}CA`);
          const exIdx = headers.findIndex(h=>h===`${sub}_EXAM`||h===`${sub} EXAM`||h===`${sub}EXAM`||h===`${sub}_EX`||h===`${sub} EX`);
          if (caIdx !== -1 || exIdx !== -1) subjectCols[sub] = { caIdx, examIdx: exIdx };
        }
      });
      const rows = lines.slice(1).map((line, i) => {
        const cells = line.split(",").map(c=>c.trim());
        const rawName = toUpper(cells[nameIdx] || "");
        const matched = students.find(s =>
          s.className === cls && (
            s.name === rawName ||
            s.name.toLowerCase().includes(rawName.toLowerCase().split(" ")[0]) ||
            rawName.toLowerCase().includes(s.name.toLowerCase().split(" ")[0])
          )
        );
        const marks = {};
        curSubjects.forEach(sub => {
          const col = subjectCols[sub];
          if (!col) return;
          if (isLower) {
            const v = cells[col.markIdx];
            marks[sub] = { mk: v !== undefined && v !== "" ? Number(v) : null };
          } else {
            const ca = col.caIdx !== undefined && col.caIdx !== -1 && cells[col.caIdx] !== "" ? Number(cells[col.caIdx]) : null;
            const exam = col.examIdx !== undefined && col.examIdx !== -1 && cells[col.examIdx] !== "" ? Number(cells[col.examIdx]) : null;
            marks[sub] = { ca, exam };
          }
        });
        return { id:`bmr_${i}`, rawName, studentId: matched?.id || null, include: true, marks };
      }).filter(r => r.rawName);
      if (rows.length === 0) { setBulkMarkError("No data rows found in the file."); return; }
      setBulkMarkPreview({ rows, subjectCols, headers });
    } catch(err) {
      setBulkMarkError("Could not read file: " + err.message);
    }
    if (bulkMarkFileRef.current) bulkMarkFileRef.current.value = "";
  };
  const confirmBulkMarks = () => {
    if (!bulkMarkPreview) return;
    const curSubjects = isLower ? LOWER_SUBJECTS : UPPER_SUBJECTS;
    bulkMarkPreview.rows.forEach(row => {
      if (!row.include || !row.studentId) return;
      const studentName = students.find(s=>s.id===row.studentId)?.name || row.rawName;
      curSubjects.forEach(sub => {
        const m = row.marks[sub];
        if (!m) return;
        const existing = termMarks[row.studentId]?.[tk]?.[sub] || {};
        if (isLower) {
          if (m.mk !== null && m.mk !== undefined) requestOrApplyTermMark(row.studentId, studentName, tk, sub, "exam", clampMark(m.mk, lowerSubjectMax(sub)), existing.exam);
        } else {
          if (m.ca !== null && m.ca !== undefined) requestOrApplyTermMark(row.studentId, studentName, tk, sub, "ca", clampMark(m.ca), existing.ca);
          if (m.exam !== null && m.exam !== undefined) requestOrApplyTermMark(row.studentId, studentName, tk, sub, "exam", clampMark(m.exam), existing.exam);
        }
      });
    });
    setBulkMarkPreview(null);
    setShowBulkMark(false);
  };
  // Generate a template CSV download
  const downloadMarkTemplate = () => {
    const curSubjects = isLower ? LOWER_SUBJECTS : UPPER_SUBJECTS;
    let header = "NAME";
    curSubjects.forEach(sub => {
      if (isLower) header += `,${sub}`;
      else header += `,${sub}_CA,${sub}_EXAM`;
    });
    const rows = classStudents.map(s => {
      let row = s.name;
      curSubjects.forEach(() => { if (isLower) row += ","; else row += ",,"; });
      return row;
    });
    const csv = [header, ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    triggerBlobDownload(blob, `${cls}_${term.replace(" ","_")}_${year}_mark_template.csv`);
  };
  return (
    <div>
      {pendingToast && (
        <div style={{position:"fixed",top:16,right:16,zIndex:3000,background:"#1e3a6e",color:"white",padding:"12px 18px",borderRadius:10,fontSize:13,fontWeight:600,boxShadow:"0 8px 24px rgba(0,0,
