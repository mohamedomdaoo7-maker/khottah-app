import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";

const INK = "#F5F6FA";
const MUTED = "#9AA3C7";
const FAINT = "#6B7398";
const INDIGO = "#5B6EF5";
const INDIGO_SOFT = "#8A97F8";
const GOLD = "#E3B65C";
const BG = "#07080F";
const GLASS = "rgba(255,255,255,0.045)";
const GLASS_BORDER = "rgba(255,255,255,0.09)";
const GLASS_HI = "rgba(255,255,255,0.08)";

const CHANNEL_COLORS = {
  snapchat: "#F5D90A",
  tiktok: "#2FD9E0",
  meta: "#4C8DFF",
  instagram: "#E1548D",
  facebook: "#4C8DFF",
  "google search": "#39C982",
  google: "#39C982",
  youtube: "#FF5C5C",
  x: "#B8BEDB",
  twitter: "#B8BEDB",
  linkedin: "#3EA6FF",
  programmatic: "#B18CF5",
  display: "#B18CF5",
};

function channelColor(name) {
  const key = Object.keys(CHANNEL_COLORS).find((k) =>
    name.toLowerCase().includes(k)
  );
  return key ? CHANNEL_COLORS[key] : "#8890B5";
}

const fmtSAR = (n) =>
  "SAR " + Number(n).toLocaleString("en-US", { maximumFractionDigits: 0 });

const OBJECTIVES = [
  "Brand awareness",
  "Traffic",
  "Lead generation",
  "Conversions / registrations",
  "App installs",
];

const MARKETS = ["Saudi Arabia — national", "Riyadh only", "GCC-wide"];

const SCENARIOS = [
  { label: "70% budget", instruction: "Regenerate the plan with only 70% of the original budget. State what was cut and why." },
  { label: "Shift to awareness", instruction: "Regenerate the plan optimized for brand awareness instead of the original objective." },
  { label: "Performance-max", instruction: "Regenerate the plan weighted aggressively toward lower-funnel performance channels." },
];

async function callClaude(promptText) {
  const response = await fetch("/api/generate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt: promptText }),
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || "Request failed");
  }
  return data;
}

function buildAudiencePrompt(brand, industry, objective, market) {
  return `You are Khottah, an expert Saudi/GCC digital media planner and audience strategist.

Given only a product/brand and industry, propose the ideal target audience segments for a paid social media campaign in the GCC.

PRODUCT: ${brand}
INDUSTRY: ${industry}
OBJECTIVE: ${objective || "not specified"}
MARKET: ${market || "Saudi Arabia"}

RULES
- Propose exactly 3 distinct audience personas, ordered by priority (primary first).
- Each persona must be realistic for the Saudi/GCC social media landscape (Snapchat, TikTok, Instagram, X, LinkedIn usage patterns by segment).
- Keep every string SHORT (under 18 words per field).
- All output in English only. Never mention Arabic or translation.
- Respond with ONLY valid JSON, no markdown, no preamble, exactly this schema:
{"personas":[{"name":"","age_range":"","demographics":"","interests":"","platform_behavior":"","primary_channel":"","rationale":""}]}`;
}

function buildPrompt(brief, scenario, prevPlan) {
  const base = `You are Khottah, an expert Saudi/GCC digital media planner. Build a paid digital media plan.

BRIEF
Brand: ${brief.brand}
Industry: ${brief.industry}
Objective: ${brief.objective}
Net budget: SAR ${brief.budget}
Market: ${brief.market}
Audience: ${brief.audience}
Duration: ${brief.weeks} weeks
Notes: ${brief.notes || "none"}

RULES
- Use GCC platform logic: Snapchat and TikTok are dominant reach channels for Saudi consumers; LinkedIn only for B2B; Google Search for intent; YouTube for video reach; Meta and X as support.
- Choose 3-6 channels. Budgets must sum exactly to the net budget.
- For each channel, name the single BEST-PERFORMING ad format for this objective on this platform right now (e.g. "Snapchat: Story Ads with Snap Lead Gen form" not just "Story Ads"), and give a concrete creative recommendation covering: creative concept/angle, visual style, and a hook or CTA suited to the audience and objective — specific enough that a designer could brief off it.
- All output in English. Never mention Arabic, translation, or non-English creative. Benchmarks are indicative estimates.
- Keep every string SHORT (under 20 words), except creative_recommendation which may be up to 35 words. Respond with ONLY valid JSON, no markdown, no preamble, exactly this schema:
{"campaign_name":"","strategy_summary":"","channels":[{"name":"","budget_sar":0,"share_pct":0,"role":"","formats":"","best_format":"","creative_recommendation":"","kpi":"","est_result":""}],"flighting":[{"phase":"","weeks":"","focus":""}],"kpis":["",""],"rationale":"","assumptions":""}
- flighting: 2-4 phases covering all ${brief.weeks} weeks.`;

  if (scenario && prevPlan) {
    return (
      base +
      `\n\nPREVIOUS PLAN (for reference): ${JSON.stringify({
        channels: prevPlan.channels,
        campaign_name: prevPlan.campaign_name,
      })}\n\nSCENARIO CHANGE REQUESTED: ${scenario}\nApply this change and regenerate the full JSON plan.`
    );
  }
  return base;
}

function buildReportPrompt(rb, lang) {
  const langRule = lang === "ar"
    ? "- Write EVERY string field (executive_summary, theme names, descriptions, snippets, recommendations) entirely in formal Modern Standard Arabic. Do not mix English words in unless it's a brand name."
    : "- All output in English only. Never mention Arabic or translation.";
  return `You are Khottah, an expert Saudi/GCC social media analyst. Analyze raw social listening data and produce a thematic performance report.

BRAND: ${rb.brand}
PLATFORMS: ${rb.platforms}
DATE RANGE: ${rb.dateRange}
CONTEXT: ${rb.context || "none"}

RAW SOCIAL LISTENING DATA (mentions, comments, post captions, exported text):
"""
${rb.rawData}
"""

TASK
- Read the raw data above and identify the 4-6 most significant recurring themes.
- Estimate overall sentiment split (positive/neutral/negative, must sum to 100).
- For each theme, estimate its share of volume (rough %, all themes need not sum to 100), assign a sentiment label, and include one short representative example drawn from or closely reflecting the raw data (paraphrase, do not copy long verbatim passages).
- Write a 2-sentence executive summary of what's happening overall.
- Give 3-5 concrete recommended actions for the social/marketing team.

RULES
${langRule}
- Keep every string SHORT (under 25 words per field, executive_summary under 45 words).
- The "sentiment" field inside each theme object must always be one of exactly: Positive, Neutral, Negative, Mixed (in English, regardless of output language) — this is used for UI color coding.
- If the raw data is sparse, still produce a best-effort structured analysis; do not refuse.
- Respond with ONLY valid JSON, no markdown, no preamble, exactly this schema:
{"executive_summary":"","sentiment":{"positive":0,"neutral":0,"negative":0},"total_mentions_estimate":"","themes":[{"name":"","sentiment":"Positive|Neutral|Negative|Mixed","volume_pct":0,"description":"","example_snippet":"","recommendation":""}],"overall_recommendations":["",""]}`;
}

const TEXT_COL_HINTS = /message|text|content|comment|post|snippet|mention|body|caption|title|description|review|tweet/i;
const MAX_CHARS = 150000; // generous cap — Claude's context window comfortably fits this; just a safety ceiling for pathological files

function extractTextFromSheetRows(rows) {
  if (!rows || rows.length === 0) return { lines: [], columnsUsed: [] };

  const headers = Object.keys(rows[0]);
  let textCols = headers.filter((h) => TEXT_COL_HINTS.test(h));

  if (textCols.length === 0) {
    // Fallback: pick columns with the longest average string length (likely free-text)
    const avgLen = headers.map((h) => {
      const vals = rows.slice(0, 30).map((r) => String(r[h] ?? ""));
      const avg = vals.reduce((a, v) => a + v.length, 0) / (vals.length || 1);
      return { h, avg };
    });
    avgLen.sort((a, b) => b.avg - a.avg);
    textCols = avgLen.slice(0, 2).filter((c) => c.avg > 15).map((c) => c.h);
    if (textCols.length === 0) textCols = headers;
  }

  const lines = [];
  for (const row of rows) {
    const parts = textCols.map((c) => row[c]).filter((v) => v !== undefined && v !== null && String(v).trim() !== "");
    if (parts.length) lines.push(parts.join(" — "));
  }
  return { lines, columnsUsed: textCols };
}

function extractTextFromWorkbook(workbook) {
  const sheetNames = workbook.SheetNames || [];
  let allLines = [];
  let totalRows = 0;
  const sheetsUsed = [];
  const columnsBySheet = {};

  for (const name of sheetNames) {
    const sheet = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    if (!rows.length) continue;
    const { lines, columnsUsed } = extractTextFromSheetRows(rows);
    if (lines.length === 0) continue;

    totalRows += rows.length;
    sheetsUsed.push(name);
    columnsBySheet[name] = columnsUsed;

    const prefix = sheetNames.length > 1 ? [`--- Sheet: ${name} ---`] : [];
    allLines = allLines.concat(prefix, lines);
  }

  let text = allLines.join("\n");
  let truncated = false;
  let includedChars = text.length;
  if (text.length > MAX_CHARS) {
    text = text.slice(0, MAX_CHARS);
    truncated = true;
    includedChars = MAX_CHARS;
  }

  return { text, rowCount: totalRows, truncated, sheetsUsed, columnsBySheet, includedChars };
}

// ---------- Small animated primitives ----------

function CountUp({ value, prefix = "", duration = 900 }) {
  const [display, setDisplay] = useState(0);
  const raf = useRef(null);
  const start = useRef(null);
  const from = useRef(0);

  useEffect(() => {
    from.current = display;
    start.current = null;
    const target = Number(value) || 0;
    function tick(ts) {
      if (!start.current) start.current = ts;
      const p = Math.min(1, (ts - start.current) / duration);
      const eased = 1 - Math.pow(1 - p, 3);
      setDisplay(Math.round(from.current + (target - from.current) * eased));
      if (p < 1) raf.current = requestAnimationFrame(tick);
    }
    raf.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);

  return <>{prefix}{display.toLocaleString("en-US")}</>;
}

function Reveal({ children, delay = 0, y = 14 }) {
  const [shown, setShown] = useState(false);
  useEffect(() => {
    const t = setTimeout(() => setShown(true), delay);
    return () => clearTimeout(t);
  }, [delay]);
  return (
    <div
      style={{
        opacity: shown ? 1 : 0,
        transform: shown ? "translateY(0)" : `translateY(${y}px)`,
        transition: "opacity 0.6s cubic-bezier(.16,1,.3,1), transform 0.6s cubic-bezier(.16,1,.3,1)",
      }}
    >
      {children}
    </div>
  );
}

export default function Khottah() {
  const [mode, setMode] = useState("plan"); // "plan" | "report"

  const [brief, setBrief] = useState({
    brand: "",
    industry: "",
    objective: OBJECTIVES[0],
    budget: 150000,
    market: MARKETS[0],
    audience: "",
    weeks: 6,
    notes: "",
  });
  const [plan, setPlan] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [lastScenario, setLastScenario] = useState(null);
  const [customScenario, setCustomScenario] = useState("");
  const [personas, setPersonas] = useState(null);
  const [personasLoading, setPersonasLoading] = useState(false);
  const [personasError, setPersonasError] = useState(null);
  const [selectedPersona, setSelectedPersona] = useState(null);
  const [barsIn, setBarsIn] = useState(false);
  const [pressed, setPressed] = useState(null);

  const [reportBrief, setReportBrief] = useState({
    brand: "",
    platforms: "",
    dateRange: "",
    context: "",
    rawData: "",
  });
  const [report, setReport] = useState(null);
  const [reportLoading, setReportLoading] = useState(false);
  const [reportError, setReportError] = useState(null);
  const [reportBarsIn, setReportBarsIn] = useState(false);
  const [fileInfo, setFileInfo] = useState(null);
  const [fileParsing, setFileParsing] = useState(false);
  const [fileError, setFileError] = useState(null);
  const [reportLang, setReportLang] = useState("en");
  const [exportMsg, setExportMsg] = useState(null);
  const [exporting, setExporting] = useState(null);

  const setR = (k) => (e) => setReportBrief({ ...reportBrief, [k]: e.target.value });

  const handleFileUpload = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setFileError(null);
    setFileParsing(true);
    setFileInfo(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const { text, rowCount, truncated, sheetsUsed, columnsBySheet, includedChars } = extractTextFromWorkbook(workbook);

        if (!text || text.trim().length < 20) {
          setFileError("Couldn't find readable text columns in this file. Try pasting the data manually instead.");
          setFileParsing(false);
          return;
        }

        setReportBrief((prev) => ({ ...prev, rawData: text }));
        setFileInfo({
          name: file.name,
          rowCount,
          truncated,
          sheetsUsed,
          columnsBySheet,
          totalSheets: workbook.SheetNames.length,
          includedChars,
        });
      } catch (err) {
        setFileError("Couldn't read this file. Make sure it's a valid .xlsx, .xls, or .csv export.");
      }
      setFileParsing(false);
    };
    reader.onerror = () => {
      setFileError("Couldn't read this file. Try again or paste the data manually.");
      setFileParsing(false);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const generateReport = async () => {
    setReportLoading(true);
    setReportError(null);
    setReportBarsIn(false);
    try {
      const result = await callClaude(buildReportPrompt(reportBrief, reportLang));
      setReport(result);
      setTimeout(() => setReportBarsIn(true), 80);
    } catch (err) {
      setReportError((err && err.message) || "The report couldn't be generated. Check the raw data field isn't empty, then try again.");
    }
    setReportLoading(false);
  };

  const reportReady = reportBrief.brand && reportBrief.rawData && reportBrief.rawData.trim().length > 40;

  const printHTML = (title, bodyHTML, rtl) => `<!doctype html>
<html lang="${rtl ? "ar" : "en"}" dir="${rtl ? "rtl" : "ltr"}">
<head>
<meta charset="utf-8" />
<title>${title}</title>
<style>
  body { font-family: ${rtl ? "'Tajawal', sans-serif" : "-apple-system, Arial, sans-serif"}; color: #1a1a2e; padding: 32px 40px; max-width: 900px; margin: 0 auto; }
  h1 { font-size: 24px; margin: 0 0 4px; }
  h2 { font-size: 15px; color: #3B4CCA; margin: 28px 0 8px; border-bottom: 1px solid #ddd; padding-bottom: 4px; }
  .meta { font-family: 'Courier New', monospace; font-size: 11px; color: #C9A24B; letter-spacing: 0.05em; text-transform: uppercase; margin-bottom: 6px; }
  .summary { font-size: 13.5px; color: #444; margin: 10px 0 4px; line-height: 1.5; }
  table { width: 100%; border-collapse: collapse; margin-top: 6px; font-size: 12.5px; }
  th { text-align: ${rtl ? "right" : "left"}; background: #1F2A56; color: #fff; padding: 8px 10px; font-size: 11px; }
  td { padding: 8px 10px; border-bottom: 1px solid #eee; vertical-align: top; }
  tr:nth-child(even) td { background: #f7f8fc; }
  .budget { text-align: ${rtl ? "left" : "right"}; font-weight: 700; color: #8a6d1f; font-family: 'Courier New', monospace; }
  .card { border: 1px solid #ddd; border-radius: 8px; padding: 12px 14px; margin-bottom: 10px; }
  .card b { color: #1F2A56; }
  ul { padding-${rtl ? "right" : "left"}: 20px; font-size: 13px; line-height: 1.7; }
  .footer { margin-top: 30px; font-size: 10.5px; color: #888; border-top: 1px solid #ddd; padding-top: 10px; }
  @media print { body { padding: 10px 20px; } }
</style>
</head>
<body>
${bodyHTML}
<div class="footer">Khottah — decision support only; human approval required before any spend. © ${new Date().getFullYear()} Mohamed Emad</div>
<script>window.onload = () => { window.print(); };</script>
</body>
</html>`;

  const escapeHTML = (s) => String(s ?? "").replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));

  const exportPlanPDF = () => {
    if (!plan) return;
    const total = plan.channels.reduce((a, c) => a + Number(c.budget_sar), 0);
    const body = `
      <div class="meta">MEDIA PLAN · ${escapeHTML(brief.market.toUpperCase())} · ${brief.weeks} WEEKS</div>
      <h1>${escapeHTML(plan.campaign_name)}</h1>
      <p class="summary">${escapeHTML(plan.strategy_summary)}</p>
      <p><b>Net budget:</b> SAR ${total.toLocaleString("en-US")}</p>
      <h2>Channel Mix & Budget</h2>
      <table><thead><tr><th>Channel</th><th>Role</th><th>KPI</th><th>Est. result</th><th>Budget</th></tr></thead><tbody>
        ${plan.channels.map((c) => `<tr><td><b>${escapeHTML(c.name)}</b></td><td>${escapeHTML(c.role)}</td><td>${escapeHTML(c.kpi)}</td><td>${escapeHTML(c.est_result)}</td><td class="budget">${fmtSAR(c.budget_sar)}</td></tr>`).join("")}
      </tbody></table>
      <h2>Creative Recommendations</h2>
      ${plan.channels.map((c) => `<div class="card"><b>${escapeHTML(c.name)}</b> — Best format: ${escapeHTML(c.best_format || c.formats)}<br/>${escapeHTML(c.creative_recommendation)}</div>`).join("")}
      <h2>Flighting</h2>
      ${plan.flighting.map((f) => `<div class="card"><b>${escapeHTML(f.weeks)} — ${escapeHTML(f.phase)}</b><br/>${escapeHTML(f.focus)}</div>`).join("")}
      <h2>KPIs</h2>
      <ul>${plan.kpis.map((k) => `<li>${escapeHTML(k)}</li>`).join("")}</ul>
      <h2>Rationale</h2>
      <p class="summary">${escapeHTML(plan.rationale)}</p>
    `;
    const w = window.open("", "_blank");
    if (!w) { setExportMsg("Your browser blocked the print tab. Allow pop-ups for this site and try again."); return; }
    w.document.write(printHTML(plan.campaign_name, body, false));
    w.document.close();
  };

  const exportReportPDF = () => {
    if (!report) return;
    const rtl = reportLang === "ar";
    const body = `
      <div class="meta">SOCIAL LISTENING · ${escapeHTML((reportBrief.platforms || "ALL PLATFORMS").toUpperCase())} · ${escapeHTML(reportBrief.dateRange || "PERIOD N/A")}</div>
      <h1>${escapeHTML(reportBrief.brand)} — ${rtl ? "ملخص المحاور" : "Theme Summary"}</h1>
      <p class="summary">${escapeHTML(report.executive_summary)}</p>
      <p><b>${rtl ? "الحجم التقديري" : "Estimated volume"}:</b> ${escapeHTML(report.total_mentions_estimate)}</p>
      <h2>${rtl ? "المشاعر العامة" : "Overall sentiment"}</h2>
      <p>${rtl ? "إيجابي" : "Positive"}: ${report.sentiment.positive}% · ${rtl ? "محايد" : "Neutral"}: ${report.sentiment.neutral}% · ${rtl ? "سلبي" : "Negative"}: ${report.sentiment.negative}%</p>
      <h2>${rtl ? "المحاور" : "Themes"}</h2>
      <table><thead><tr><th>${rtl ? "المحور" : "Theme"}</th><th>${rtl ? "المشاعر" : "Sentiment"}</th><th>${rtl ? "الحجم" : "Volume"}</th><th>${rtl ? "الوصف" : "Description"}</th></tr></thead><tbody>
        ${report.themes.map((t) => `<tr><td><b>${escapeHTML(t.name)}</b></td><td>${escapeHTML(t.sentiment)}</td><td>~${t.volume_pct}%</td><td>${escapeHTML(t.description)}<br/><i>"${escapeHTML(t.example_snippet)}"</i><br/>→ ${escapeHTML(t.recommendation)}</td></tr>`).join("")}
      </tbody></table>
      <h2>${rtl ? "الإجراءات الموصى بها" : "Recommended actions"}</h2>
      <ul>${report.overall_recommendations.map((r) => `<li>${escapeHTML(r)}</li>`).join("")}</ul>
    `;
    const w = window.open("", "_blank");
    if (!w) { setExportMsg("Your browser blocked the print tab. Allow pop-ups for this site and try again."); return; }
    w.document.write(printHTML(`${reportBrief.brand} — Theme Summary`, body, rtl));
    w.document.close();
  };

  const exportPPT = async (type) => {
    setExporting(type);
    setExportMsg(null);
    try {
      const payload = type === "plan"
        ? { type: "plan", data: plan, meta: brief }
        : { type: "report", data: report, meta: reportBrief };
      const res = await fetch("/api/export-pptx", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!res.ok) throw new Error("not deployed");
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `khottah-${type}.pptx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      setExportMsg("PPT export needs the deployed version of Khottah (it runs on the server). Use Download PDF here, or deploy to Netlify/Vercel for PPT.");
    }
    setExporting(null);
  };


  const set = (k) => (e) => setBrief({ ...brief, [k]: e.target.value });

  const suggestAudience = async () => {
    setPersonasLoading(true);
    setPersonasError(null);
    setPersonas(null);
    setSelectedPersona(null);
    try {
      const result = await callClaude(
        buildAudiencePrompt(brief.brand, brief.industry, brief.objective, brief.market)
      );
      setPersonas(result.personas || []);
    } catch (err) {
      setPersonasError("Couldn't generate audience suggestions. Try again in a moment.");
    }
    setPersonasLoading(false);
  };

  const applyPersona = (p) => {
    const desc = `${p.name} (${p.age_range}): ${p.demographics}. Interests: ${p.interests}. Primarily active on ${p.primary_channel}.`;
    setBrief({ ...brief, audience: desc });
    setSelectedPersona(p.name);
  };

  const generate = async (scenario) => {
    setLoading(true);
    setError(null);
    setBarsIn(false);
    try {
      const result = await callClaude(buildPrompt(brief, scenario, plan));
      setPlan(result);
      setLastScenario(scenario || null);
      setTimeout(() => setBarsIn(true), 80);
    } catch (err) {
      setError((err && err.message) || "The plan couldn't be generated. Check the brief is complete, then try again.");
    }
    setLoading(false);
  };

  const audienceReady = brief.brand && brief.industry;
  const ready = brief.brand && brief.audience && brief.budget > 0;
  const total = plan ? plan.channels.reduce((a, c) => a + Number(c.budget_sar), 0) : 0;

  const pressStyle = (id) => ({
    transform: pressed === id ? "scale(0.97)" : "scale(1)",
    transition: "transform 0.15s cubic-bezier(.16,1,.3,1)",
  });

  return (
    <div style={{ background: BG, minHeight: "100vh", color: INK, position: "relative", overflow: "hidden" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600&family=Tajawal:wght@400;500;700;800&display=swap');
        * { font-family: -apple-system, BlinkMacSystemFont, 'Inter', sans-serif; box-sizing: border-box; }
        .mono { font-family: 'IBM Plex Mono', ui-monospace, monospace; }
        html { scroll-behavior: smooth; }
        input, select, textarea { outline: none; color: ${INK}; }
        input::placeholder, textarea::placeholder { color: #5A6289; }
        input:focus, select:focus, textarea:focus {
          border-color: ${INDIGO} !important;
          box-shadow: 0 0 0 4px rgba(91,110,245,0.18);
        }
        select option { background: #14162A; color: ${INK}; }
        button { font-family: inherit; }
        button:focus-visible { outline: 2px solid ${INDIGO}; outline-offset: 2px; }
        ::-webkit-scrollbar { width: 10px; height: 10px; }
        ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.12); border-radius: 8px; }
        ::-webkit-scrollbar-track { background: transparent; }

        @keyframes floatA { 0%,100% { transform: translate(0,0) scale(1); } 50% { transform: translate(40px,-30px) scale(1.08); } }
        @keyframes floatB { 0%,100% { transform: translate(0,0) scale(1); } 50% { transform: translate(-50px,40px) scale(1.1); } }
        @keyframes floatC { 0%,100% { transform: translate(0,0) scale(1); } 50% { transform: translate(30px,50px) scale(0.95); } }
        @keyframes shimmer { 0% { background-position: -400px 0; } 100% { background-position: 400px 0; } }
        @keyframes pulseDot { 0%,100% { opacity: 0.3; transform: scale(0.85); } 50% { opacity: 1; transform: scale(1); } }
        @keyframes barGrow { from { width: 0; } }
        @keyframes spin { to { transform: rotate(360deg); } }

        .orb { position: absolute; border-radius: 50%; filter: blur(80px); pointer-events: none; }
        .shimmer-block {
          background: linear-gradient(90deg, rgba(255,255,255,0.03) 0%, rgba(255,255,255,0.09) 50%, rgba(255,255,255,0.03) 100%);
          background-size: 800px 100%;
          animation: shimmer 1.6s infinite linear;
        }
        .card-hover { transition: transform 0.35s cubic-bezier(.16,1,.3,1), box-shadow 0.35s cubic-bezier(.16,1,.3,1), border-color 0.35s ease; }
        .card-hover:hover { transform: translateY(-3px); border-color: rgba(255,255,255,0.18) !important; box-shadow: 0 20px 50px -20px rgba(0,0,0,0.6); }
        .btn-primary { transition: transform 0.15s cubic-bezier(.16,1,.3,1), box-shadow 0.25s ease, filter 0.25s ease; }
        .btn-primary:hover:not(:disabled) { filter: brightness(1.08); box-shadow: 0 10px 30px -8px rgba(91,110,245,0.55); }
        .btn-primary:active:not(:disabled) { transform: scale(0.97); }
        .chip { transition: all 0.2s cubic-bezier(.16,1,.3,1); }
        .chip:hover:not(:disabled) { background: rgba(91,110,245,0.16) !important; transform: translateY(-1px); }
        .chip:active:not(:disabled) { transform: scale(0.96); }
        @media (prefers-reduced-motion: reduce) {
          * { animation-duration: 0.01ms !important; animation-iteration-count: 1 !important; transition-duration: 0.01ms !important; }
        }
        @media print {
          .print-hide { display: none !important; }
          body, .print-area { background: #ffffff !important; }
          .print-area, .print-area * {
            color: #111111 !important;
            background: transparent !important;
            box-shadow: none !important;
            backdrop-filter: none !important;
            border-color: #ddd !important;
          }
          .print-area { border: none !important; }
        }
      `}</style>

      {/* Animated gradient mesh background */}
      <div className="orb print-hide" style={{ width: 520, height: 520, top: -160, left: -120, background: "radial-gradient(circle, rgba(91,110,245,0.35), transparent 70%)", animation: "floatA 22s ease-in-out infinite" }} />
      <div className="orb print-hide" style={{ width: 460, height: 460, top: 240, right: -140, background: "radial-gradient(circle, rgba(227,182,92,0.22), transparent 70%)", animation: "floatB 26s ease-in-out infinite" }} />
      <div className="orb print-hide" style={{ width: 380, height: 380, bottom: -140, left: "30%", background: "radial-gradient(circle, rgba(138,151,248,0.25), transparent 70%)", animation: "floatC 20s ease-in-out infinite" }} />

      <div style={{ position: "relative", zIndex: 1 }}>
        {/* Header */}
        <header
          className="px-6 py-5 flex items-center justify-between flex-wrap gap-3 print-hide"
          style={{ borderBottom: `1px solid ${GLASS_BORDER}`, backdropFilter: "blur(20px)", background: "rgba(7,8,15,0.6)", position: "sticky", top: 0, zIndex: 10 }}
        >
          <div className="flex items-center gap-4 flex-wrap">
            <div className="flex items-center gap-3">
              <div
                style={{
                  width: 34, height: 34, borderRadius: 10,
                  background: `linear-gradient(135deg, ${INDIGO}, ${INDIGO_SOFT})`,
                  display: "flex", alignItems: "center", justifyContent: "center",
                  boxShadow: "0 6px 20px -6px rgba(91,110,245,0.6)",
                }}
              >
                <span style={{ fontWeight: 800, fontSize: 16, color: "#fff" }}>K</span>
              </div>
              <div>
                <div style={{ fontWeight: 700, fontSize: 17, letterSpacing: "-0.02em" }}>Khottah</div>
                <div className="mono" style={{ fontSize: 10, color: FAINT, letterSpacing: "0.08em" }}>
                  AI MEDIA INTELLIGENCE · KSA / GCC
                </div>
              </div>
            </div>

            <div className="flex rounded-full p-1" style={{ background: "rgba(255,255,255,0.05)", border: `1px solid ${GLASS_BORDER}` }}>
              {[
                { id: "plan", label: "Media Plan" },
                { id: "report", label: "Social Report" },
              ].map((m) => (
                <button
                  key={m.id}
                  onClick={() => setMode(m.id)}
                  className="chip rounded-full px-4 py-1.5 text-xs font-semibold"
                  style={{
                    background: mode === m.id ? `linear-gradient(135deg, ${INDIGO}, #4655D9)` : "transparent",
                    color: mode === m.id ? "#fff" : MUTED,
                    cursor: "pointer",
                  }}
                >
                  {m.label}
                </button>
              ))}
            </div>
          </div>
          <div className="text-right">
            <div style={{ fontSize: 12, color: MUTED }}>
              {mode === "plan" ? "Plans in minutes. Judgment stays human." : "Themes from the noise, in minutes."}
            </div>
            <div style={{ fontSize: 11, fontWeight: 600, color: GOLD }}>Developed by Mohamed Emad</div>
          </div>
        </header>

        {mode === "plan" ? (
        <main className="max-w-6xl mx-auto px-5 py-8 grid gap-6 lg:grid-cols-3">
          {/* Brief form */}
          <Reveal delay={40}>
            <section
              className="rounded-2xl p-6 h-fit lg:col-span-1 print-hide"
              style={{ background: GLASS, border: `1px solid ${GLASS_BORDER}`, backdropFilter: "blur(24px)" }}
            >
              <h2 style={{ fontSize: 19, fontWeight: 700, letterSpacing: "-0.01em" }}>Campaign brief</h2>
              <p style={{ fontSize: 12.5, color: FAINT, marginTop: 4, marginBottom: 20 }}>
                Fill the brief. Khottah drafts the plan; you approve the spend.
              </p>

              {[
                { k: "brand", label: "Brand / product", ph: "e.g. Training Portal" },
                { k: "industry", label: "Industry", ph: "e.g. Government services" },
              ].map((f) => (
                <label key={f.k} className="block mb-3.5">
                  <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>{f.label}</span>
                  <input
                    value={brief[f.k]}
                    onChange={set(f.k)}
                    placeholder={f.ph}
                    className="mt-1.5 w-full rounded-lg px-3.5 py-2.5 text-sm"
                    style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)" }}
                  />
                </label>
              ))}

              <label className="block mb-3.5">
                <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>Objective</span>
                <select
                  value={brief.objective}
                  onChange={set("objective")}
                  className="mt-1.5 w-full rounded-lg px-3.5 py-2.5 text-sm"
                  style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)" }}
                >
                  {OBJECTIVES.map((o) => <option key={o}>{o}</option>)}
                </select>
              </label>

              <div className="grid grid-cols-2 gap-3 mb-3.5">
                <label className="block">
                  <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>Net budget (SAR)</span>
                  <input
                    type="number" min="1000" value={brief.budget} onChange={set("budget")}
                    className="mono mt-1.5 w-full rounded-lg px-3.5 py-2.5 text-sm"
                    style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)" }}
                  />
                </label>
                <label className="block">
                  <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>Duration (wks)</span>
                  <input
                    type="number" min="1" max="26" value={brief.weeks} onChange={set("weeks")}
                    className="mono mt-1.5 w-full rounded-lg px-3.5 py-2.5 text-sm"
                    style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)" }}
                  />
                </label>
              </div>

              <label className="block mb-3.5">
                <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>Market</span>
                <select
                  value={brief.market}
                  onChange={set("market")}
                  className="mt-1.5 w-full rounded-lg px-3.5 py-2.5 text-sm"
                  style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)" }}
                >
                  {MARKETS.map((m) => <option key={m}>{m}</option>)}
                </select>
              </label>

              <div className="mb-3.5">
                <div className="flex items-center justify-between mb-1.5">
                  <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>Target audience</span>
                  <button
                    onClick={suggestAudience}
                    disabled={!audienceReady || personasLoading}
                    className="chip text-xs font-semibold rounded-full px-3 py-1"
                    style={{
                      color: audienceReady ? INDIGO_SOFT : "#4A5178",
                      border: `1px solid ${audienceReady ? "rgba(91,110,245,0.4)" : GLASS_BORDER}`,
                      background: "rgba(91,110,245,0.08)",
                      cursor: audienceReady && !personasLoading ? "pointer" : "not-allowed",
                    }}
                  >
                    {personasLoading ? (
                      <span className="flex items-center gap-1.5">
                        <span style={{ width: 6, height: 6, borderRadius: "50%", background: INDIGO_SOFT, display: "inline-block", animation: "pulseDot 1s infinite" }} />
                        Thinking
                      </span>
                    ) : "✨ Suggest audience"}
                  </button>
                </div>
                <textarea
                  value={brief.audience}
                  onChange={(e) => { setBrief({ ...brief, audience: e.target.value }); setSelectedPersona(null); }}
                  placeholder="e.g. Saudi jobseekers 20-35, mobile-first — or click Suggest audience"
                  rows={2}
                  className="w-full rounded-lg px-3.5 py-2.5 text-sm"
                  style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)" }}
                />
                {!audienceReady && (
                  <p style={{ fontSize: 11, color: FAINT, marginTop: 6 }}>
                    Add brand and industry above to enable AI audience suggestions.
                  </p>
                )}
                {personasError && <p style={{ fontSize: 11, color: "#F27272", marginTop: 6 }}>{personasError}</p>}

                {personasLoading && (
                  <div className="mt-3 space-y-2">
                    {[0, 1, 2].map((i) => (
                      <div key={i} className="shimmer-block rounded-lg" style={{ height: 62, opacity: 1 - i * 0.15 }} />
                    ))}
                  </div>
                )}

                {personas && personas.length > 0 && (
                  <div className="mt-3 space-y-2">
                    <p style={{ fontSize: 11, fontWeight: 600, color: MUTED }}>Suggested personas — tap one to use it</p>
                    {personas.map((p, i) => {
                      const active = selectedPersona === p.name;
                      return (
                        <Reveal key={p.name} delay={i * 90}>
                          <button
                            onClick={() => applyPersona(p)}
                            className="card-hover w-full text-left rounded-xl p-3"
                            style={{
                              border: `1px solid ${active ? "rgba(91,110,245,0.55)" : GLASS_BORDER}`,
                              background: active ? "rgba(91,110,245,0.14)" : "rgba(255,255,255,0.025)",
                              cursor: "pointer",
                            }}
                          >
                            <div className="flex justify-between items-baseline">
                              <span style={{ fontSize: 13.5, fontWeight: 700 }}>{p.name}</span>
                              <span className="mono" style={{ fontSize: 11, color: FAINT }}>{p.age_range}</span>
                            </div>
                            <p style={{ fontSize: 11.5, color: "#B7BEDD", marginTop: 4 }}>{p.demographics}</p>
                            <p style={{ fontSize: 11.5, color: FAINT, marginTop: 3 }}>
                              <span style={{ fontWeight: 600 }}>Interests:</span> {p.interests}
                            </p>
                            <p style={{ fontSize: 11.5, color: INDIGO_SOFT, marginTop: 3 }}>
                              <span style={{ fontWeight: 600 }}>Primary:</span> {p.primary_channel} — {p.platform_behavior}
                            </p>
                            <p style={{ fontSize: 11, color: "#7C84A8", marginTop: 3, fontStyle: "italic" }}>{p.rationale}</p>
                          </button>
                        </Reveal>
                      );
                    })}
                  </div>
                )}
              </div>

              <label className="block mb-5">
                <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>Notes (optional)</span>
                <textarea
                  value={brief.notes}
                  onChange={set("notes")}
                  placeholder="B2B focus, exclusions, seasonal timing…"
                  rows={2}
                  className="mt-1.5 w-full rounded-lg px-3.5 py-2.5 text-sm"
                  style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)" }}
                />
              </label>

              <button
                onClick={() => generate(null)}
                disabled={!ready || loading}
                onMouseDown={() => setPressed("gen")}
                onMouseUp={() => setPressed(null)}
                onMouseLeave={() => setPressed(null)}
                className="btn-primary w-full rounded-xl py-3 text-sm font-semibold"
                style={{
                  background: ready && !loading ? `linear-gradient(135deg, ${INDIGO}, #4655D9)` : "rgba(255,255,255,0.06)",
                  color: ready && !loading ? "#fff" : "#5A6289",
                  cursor: ready && !loading ? "pointer" : "not-allowed",
                  ...pressStyle("gen"),
                }}
              >
                {loading ? (
                  <span className="flex items-center justify-center gap-2">
                    <span style={{
                      width: 14, height: 14, borderRadius: "50%",
                      border: "2px solid rgba(255,255,255,0.35)", borderTopColor: "#fff",
                      animation: "spin 0.7s linear infinite", display: "inline-block",
                    }} />
                    Building the plan…
                  </span>
                ) : plan ? "Regenerate plan" : "Generate media plan"}
              </button>
              {!ready && <p style={{ fontSize: 11, color: FAINT, marginTop: 8 }}>Brand, audience, and budget are required.</p>}
            </section>
          </Reveal>

          {/* Plan sheet */}
          <section className="lg:col-span-2">
            {!plan && !loading && (
              <Reveal delay={120}>
                <div
                  className="rounded-2xl p-16 text-center"
                  style={{ border: `1px dashed ${GLASS_BORDER}`, background: "rgba(255,255,255,0.015)" }}
                >
                  <div style={{
                    width: 52, height: 52, borderRadius: 16, margin: "0 auto 18px",
                    background: `linear-gradient(135deg, rgba(91,110,245,0.25), rgba(227,182,92,0.2))`,
                    display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22,
                  }}>📋</div>
                  <p style={{ fontSize: 19, fontWeight: 700, letterSpacing: "-0.01em" }}>Your plan sheet renders here</p>
                  <p style={{ fontSize: 13.5, color: FAINT, marginTop: 8, maxWidth: 380, marginLeft: "auto", marginRight: "auto" }}>
                    Complete the brief and generate. You'll get channel mix, budget split, flighting, and rationale — ready for expert review.
                  </p>
                </div>
              </Reveal>
            )}

            {loading && (
              <div className="rounded-2xl p-6" style={{ border: `1px solid ${GLASS_BORDER}`, background: GLASS }}>
                <div className="shimmer-block rounded-lg mb-4" style={{ height: 90 }} />
                <div className="shimmer-block rounded-lg mb-4" style={{ height: 36 }} />
                {[0, 1, 2].map((i) => (
                  <div key={i} className="shimmer-block rounded-lg mb-3" style={{ height: 56, opacity: 1 - i * 0.12 }} />
                ))}
                <p style={{ textAlign: "center", fontSize: 13, color: MUTED, marginTop: 10 }}>
                  Allocating budget across GCC platform logic…
                </p>
              </div>
            )}

            {error && !loading && (
              <div className="rounded-xl p-4 mb-4 text-sm" style={{ background: "rgba(242,114,114,0.1)", border: "1px solid rgba(242,114,114,0.35)", color: "#F5A3A3" }}>
                {error}
              </div>
            )}

            {plan && !loading && (
              <Reveal delay={20}>
                <div className="rounded-2xl overflow-hidden print-area" style={{ border: `1px solid ${GLASS_BORDER}`, background: GLASS, backdropFilter: "blur(24px)" }}>
                  {/* Plan header */}
                  <div className="px-7 py-6" style={{ background: "linear-gradient(135deg, rgba(91,110,245,0.16), rgba(227,182,92,0.08))", borderBottom: `1px solid ${GLASS_BORDER}` }}>
                    <div className="flex justify-between items-start flex-wrap gap-3">
                      <div>
                        <p className="mono" style={{ fontSize: 11, color: GOLD, letterSpacing: "0.06em", marginBottom: 6 }}>
                          MEDIA PLAN · {brief.market.toUpperCase()} · {brief.weeks} WEEKS
                          {lastScenario ? " · SCENARIO REVISION" : ""}
                        </p>
                        <h3 style={{ fontSize: 24, fontWeight: 800, letterSpacing: "-0.02em" }}>{plan.campaign_name}</h3>
                      </div>
                      <div className="text-right">
                        <p style={{ fontSize: 11, color: MUTED }}>Net budget</p>
                        <p className="mono" style={{ fontSize: 22, fontWeight: 700, color: GOLD }}>
                          SAR <CountUp value={total} />
                        </p>
                      </div>
                    </div>
                    <p style={{ fontSize: 13.5, color: "#C7CDE6", marginTop: 12, lineHeight: 1.5 }}>{plan.strategy_summary}</p>

                    <div className="flex gap-2 mt-4">
                      <button
                        onClick={exportPlanPDF}
                        className="chip rounded-full px-3.5 py-1.5 text-xs font-semibold flex items-center gap-1.5"
                        style={{ border: `1px solid ${GLASS_BORDER}`, color: "#C7CDE6", background: "rgba(255,255,255,0.05)", cursor: "pointer" }}
                      >
                        ⬇ Download PDF
                      </button>
                      <button
                        onClick={() => exportPPT("plan")}
                        disabled={exporting === "plan"}
                        className="chip rounded-full px-3.5 py-1.5 text-xs font-semibold flex items-center gap-1.5"
                        style={{ border: `1px solid ${GLASS_BORDER}`, color: "#C7CDE6", background: "rgba(255,255,255,0.05)", cursor: "pointer" }}
                      >
                        {exporting === "plan" ? "Preparing…" : "⬇ Download PPT"}
                      </button>
                    </div>
                    {exportMsg && exporting === null && (
                      <p className="print-hide" style={{ fontSize: 11, color: GOLD, marginTop: 8, maxWidth: 480 }}>{exportMsg}</p>
                    )}
                  </div>

                  {/* Allocation strip + donut */}
                  <div className="px-7 pt-6 grid gap-6 md:grid-cols-[1fr_180px] items-center">
                  <div>
                    <p style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 10 }}>Budget allocation</p>
                    <div className="flex w-full h-3 rounded-full overflow-hidden" style={{ background: "rgba(255,255,255,0.06)" }}>
                      {plan.channels.map((c, i) => (
                        <div
                          key={c.name}
                          title={`${c.name} — ${c.share_pct}%`}
                          style={{
                            width: barsIn ? `${c.share_pct}%` : "0%",
                            background: channelColor(c.name),
                            transition: `width 0.9s cubic-bezier(.16,1,.3,1) ${i * 0.08}s`,
                          }}
                        />
                      ))}
                    </div>
                    <div className="flex flex-wrap gap-x-4 gap-y-1.5 mt-3">
                      {plan.channels.map((c) => (
                        <span key={c.name} className="flex items-center gap-1.5" style={{ fontSize: 12 }}>
                          <span className="inline-block w-2.5 h-2.5 rounded-full" style={{ background: channelColor(c.name) }} />
                          {c.name}
                          <span className="mono" style={{ color: FAINT }}>{c.share_pct}%</span>
                        </span>
                      ))}
                    </div>
                  </div>
                  <div style={{ width: 160, height: 160, margin: "0 auto" }} className="print-hide">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie data={plan.channels} dataKey="share_pct" nameKey="name" innerRadius={45} outerRadius={70} paddingAngle={2} stroke="none">
                          {plan.channels.map((c) => <Cell key={c.name} fill={channelColor(c.name)} />)}
                        </Pie>
                        <Tooltip contentStyle={{ background: "#14162A", border: `1px solid ${GLASS_BORDER}`, borderRadius: 8, fontSize: 12 }} />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                  </div>

                  {/* Channel table */}
                  <div className="px-7 pt-6 overflow-x-auto">
                    <table className="w-full" style={{ minWidth: 580, fontSize: 13.5 }}>
                      <thead>
                        <tr style={{ borderBottom: `1px solid ${GLASS_BORDER}` }}>
                          {["Channel", "Role", "Formats", "Primary KPI", "Est. result", ""].map((h, i) => (
                            <th key={h} style={{ textAlign: i === 5 ? "right" : "left", padding: "0 12px 10px 0", fontSize: 11, color: FAINT, fontWeight: 600 }}>{h || "Budget"}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {plan.channels.map((c, i) => (
                          <tr key={c.name} style={{ borderBottom: i < plan.channels.length - 1 ? `1px solid rgba(255,255,255,0.05)` : "none" }}>
                            <td style={{ padding: "12px 12px 12px 0", fontWeight: 700 }}>
                              <span className="flex items-center gap-2">
                                <span className="inline-block w-2 h-6 rounded-full" style={{ background: channelColor(c.name) }} />
                                {c.name}
                              </span>
                            </td>
                            <td style={{ padding: "12px 12px 12px 0", color: "#C7CDE6" }}>{c.role}</td>
                            <td style={{ padding: "12px 12px 12px 0", color: "#C7CDE6" }}>{c.formats}</td>
                            <td style={{ padding: "12px 12px 12px 0", color: "#C7CDE6" }}>{c.kpi}</td>
                            <td style={{ padding: "12px 12px 12px 0", color: FAINT }}>{c.est_result}</td>
                            <td className="mono" style={{ padding: "12px 0", textAlign: "right", fontWeight: 700, color: GOLD }}>{fmtSAR(c.budget_sar)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  {/* Creative recommendations */}
                  <div className="px-7 pt-6">
                    <p style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 10 }}>Creative recommendations, per channel</p>
                    <div className="grid gap-2.5 sm:grid-cols-2">
                      {plan.channels.map((c, i) => (
                        <Reveal key={c.name} delay={i * 70}>
                          <div className="card-hover rounded-xl p-4" style={{ background: "rgba(255,255,255,0.03)", border: `1px solid ${GLASS_BORDER}` }}>
                            <div className="flex items-center gap-2 mb-2">
                              <span className="inline-block w-2.5 h-2.5 rounded-full" style={{ background: channelColor(c.name) }} />
                              <span style={{ fontSize: 13.5, fontWeight: 700 }}>{c.name}</span>
                            </div>
                            <p style={{ fontSize: 11.5, color: GOLD, fontWeight: 600 }}>
                              Best format: <span style={{ color: "#E8CE8F", fontWeight: 500 }}>{c.best_format || c.formats}</span>
                            </p>
                            <p style={{ fontSize: 12, color: "#C7CDE6", marginTop: 6, lineHeight: 1.5 }}>{c.creative_recommendation}</p>
                          </div>
                        </Reveal>
                      ))}
                    </div>
                  </div>

                  {/* Flighting */}
                  <div className="px-7 pt-6">
                    <p style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 10 }}>Flighting</p>
                    <div className="grid gap-2.5 sm:grid-cols-2 lg:grid-cols-3">
                      {plan.flighting.map((f, i) => (
                        <Reveal key={f.phase} delay={i * 80}>
                          <div className="card-hover rounded-xl p-3.5" style={{ background: "rgba(255,255,255,0.03)", border: `1px solid ${GLASS_BORDER}` }}>
                            <p className="mono" style={{ fontSize: 11, color: INDIGO_SOFT }}>{f.weeks}</p>
                            <p style={{ fontSize: 13.5, fontWeight: 700, marginTop: 2 }}>{f.phase}</p>
                            <p style={{ fontSize: 11.5, color: FAINT, marginTop: 4 }}>{f.focus}</p>
                          </div>
                        </Reveal>
                      ))}
                    </div>
                  </div>

                  {/* KPIs + rationale */}
                  <div className="px-7 pt-6 grid gap-5 md:grid-cols-2">
                    <div>
                      <p style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 8 }}>Campaign KPIs</p>
                      <ul style={{ fontSize: 13, lineHeight: 1.7 }}>
                        {plan.kpis.map((k) => (
                          <li key={k} className="flex gap-2"><span style={{ color: GOLD }}>—</span>{k}</li>
                        ))}
                      </ul>
                    </div>
                    <div>
                      <p style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 8 }}>Planner's rationale</p>
                      <p style={{ fontSize: 13, color: "#C7CDE6", lineHeight: 1.6 }}>{plan.rationale}</p>
                    </div>
                  </div>

                  {/* Scenario lab */}
                  <div className="px-7 py-6 mt-6" style={{ background: "rgba(255,255,255,0.02)", borderTop: `1px solid ${GLASS_BORDER}` }}>
                    <p style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 10 }}>Scenario lab — regenerate under new assumptions</p>
                    <div className="flex flex-wrap gap-2 mb-3">
                      {SCENARIOS.map((s) => (
                        <button
                          key={s.label}
                          onClick={() => generate(s.instruction)}
                          disabled={loading}
                          className="chip rounded-full px-3.5 py-1.5 text-xs font-semibold"
                          style={{ border: `1px solid rgba(91,110,245,0.4)`, color: INDIGO_SOFT, background: "rgba(91,110,245,0.08)", cursor: "pointer" }}
                        >
                          {s.label}
                        </button>
                      ))}
                    </div>
                    <div className="flex gap-2">
                      <input
                        value={customScenario}
                        onChange={(e) => setCustomScenario(e.target.value)}
                        placeholder="Custom scenario, e.g. remove X, add LinkedIn for B2B decision-makers"
                        className="flex-1 rounded-lg px-3.5 py-2.5 text-sm"
                        style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)" }}
                      />
                      <button
                        onClick={() => customScenario && generate(customScenario)}
                        disabled={loading || !customScenario}
                        className="btn-primary rounded-lg px-5 py-2.5 text-sm font-semibold"
                        style={{
                          background: customScenario ? `linear-gradient(135deg, ${INDIGO}, #4655D9)` : "rgba(255,255,255,0.06)",
                          color: customScenario ? "#fff" : "#5A6289",
                          cursor: customScenario ? "pointer" : "not-allowed",
                        }}
                      >
                        Apply
                      </button>
                    </div>
                    <p style={{ fontSize: 11.5, color: FAINT, marginTop: 12, lineHeight: 1.5 }}>
                      Assumptions: {plan.assumptions} · All figures are indicative estimates. A qualified planner must validate benchmarks before any spend is approved.
                    </p>
                  </div>
                </div>
              </Reveal>
            )}
          </section>
        </main>
        ) : (
        <main className="max-w-6xl mx-auto px-5 py-8 grid gap-6 lg:grid-cols-3">
          {/* Report brief form */}
          <Reveal delay={40}>
            <section
              className="rounded-2xl p-6 h-fit lg:col-span-1 print-hide"
              style={{ background: GLASS, border: `1px solid ${GLASS_BORDER}`, backdropFilter: "blur(24px)" }}
            >
              <div className="flex items-center justify-between mb-1">
                <h2 style={{ fontSize: 19, fontWeight: 700, letterSpacing: "-0.01em" }}>Social listening report</h2>
                <div className="flex rounded-full p-0.5" style={{ background: "rgba(255,255,255,0.05)", border: `1px solid ${GLASS_BORDER}` }}>
                  {[{ id: "en", label: "EN" }, { id: "ar", label: "AR" }].map((l) => (
                    <button
                      key={l.id}
                      onClick={() => setReportLang(l.id)}
                      className="chip rounded-full px-2.5 py-1 text-xs font-bold"
                      style={{
                        background: reportLang === l.id ? `linear-gradient(135deg, ${INDIGO}, #4655D9)` : "transparent",
                        color: reportLang === l.id ? "#fff" : MUTED,
                        cursor: "pointer",
                      }}
                    >
                      {l.label}
                    </button>
                  ))}
                </div>
              </div>
              <p style={{ fontSize: 12.5, color: FAINT, marginTop: 4, marginBottom: 20 }}>
                Paste raw mentions, comments, or a Sprinklr/GA4 export. Khottah finds the themes.
              </p>

              {[
                { k: "brand", label: "Brand", ph: "e.g. Musaned" },
                { k: "platforms", label: "Platforms", ph: "e.g. X, Instagram, TikTok" },
                { k: "dateRange", label: "Date range", ph: "e.g. Jan – Jun 2026" },
              ].map((f) => (
                <label key={f.k} className="block mb-3.5">
                  <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>{f.label}</span>
                  <input
                    value={reportBrief[f.k]}
                    onChange={setR(f.k)}
                    placeholder={f.ph}
                    className="mt-1.5 w-full rounded-lg px-3.5 py-2.5 text-sm"
                    style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)" }}
                  />
                </label>
              ))}

              <label className="block mb-3.5">
                <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>Context (optional)</span>
                <input
                  value={reportBrief.context}
                  onChange={setR("context")}
                  placeholder="e.g. post-campaign review, crisis check"
                  className="mt-1.5 w-full rounded-lg px-3.5 py-2.5 text-sm"
                  style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)" }}
                />
              </label>

              <div className="mb-4">
                <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>Upload export (optional)</span>
                <label
                  className="chip mt-1.5 flex flex-col items-center justify-center rounded-xl px-4 py-5 text-center cursor-pointer"
                  style={{ border: `1.5px dashed ${GLASS_BORDER}`, background: "rgba(255,255,255,0.02)" }}
                >
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={handleFileUpload}
                    style={{ display: "none" }}
                  />
                  {fileParsing ? (
                    <span className="flex items-center gap-2" style={{ fontSize: 12.5, color: INDIGO_SOFT }}>
                      <span style={{
                        width: 13, height: 13, borderRadius: "50%",
                        border: "2px solid rgba(138,151,248,0.35)", borderTopColor: INDIGO_SOFT,
                        animation: "spin 0.7s linear infinite", display: "inline-block",
                      }} />
                      Reading file…
                    </span>
                  ) : (
                    <>
                      <span style={{ fontSize: 13, fontWeight: 600, color: "#C7CDE6" }}>
                        📎 Drop or click to upload .xlsx / .xls / .csv
                      </span>
                      <span style={{ fontSize: 11, color: FAINT, marginTop: 3 }}>
                        A Sprinklr, Excel, or CSV mentions export
                      </span>
                    </>
                  )}
                </label>

                {fileError && <p style={{ fontSize: 11, color: "#F5A3A3", marginTop: 6 }}>{fileError}</p>}

                {fileInfo && !fileParsing && (
                  <div className="rounded-lg mt-2 px-3 py-2" style={{ background: "rgba(57,201,130,0.08)", border: "1px solid rgba(57,201,130,0.3)" }}>
                    <p style={{ fontSize: 11.5, color: "#6FE3A8", fontWeight: 600 }}>
                      ✓ Loaded {fileInfo.rowCount} rows across {fileInfo.sheetsUsed.length} sheet{fileInfo.sheetsUsed.length > 1 ? "s" : ""} from {fileInfo.name}
                    </p>
                    <p style={{ fontSize: 10.5, color: FAINT, marginTop: 3 }}>
                      {fileInfo.sheetsUsed.map((s) => `${s} (${fileInfo.columnsBySheet[s].join(", ")})`).join("  ·  ")}
                    </p>
                    {fileInfo.truncated && (
                      <p style={{ fontSize: 10.5, color: GOLD, marginTop: 3 }}>
                        ⚠ File is large — using the first ~{Math.round(fileInfo.includedChars / 1000)}K characters. Consider splitting very large exports for a fully complete read.
                      </p>
                    )}
                  </div>
                )}
              </div>

              <label className="block mb-5">
                <span style={{ fontSize: 12, fontWeight: 600, color: MUTED }}>Raw listening data</span>
                <textarea
                  value={reportBrief.rawData}
                  onChange={setR("rawData")}
                  placeholder="Paste mentions, comments, or post captions here — one per line — or upload a file above…"
                  rows={9}
                  className="mt-1.5 w-full rounded-lg px-3.5 py-2.5 text-sm"
                  style={{ border: `1px solid ${GLASS_BORDER}`, background: "rgba(255,255,255,0.03)", resize: "vertical" }}
                />
                <p style={{ fontSize: 11, color: FAINT, marginTop: 6 }}>
                  {reportBrief.rawData.trim().length} characters — needs at least ~40 to analyze. Editable after upload.
                </p>
              </label>

              <button
                onClick={generateReport}
                disabled={!reportReady || reportLoading}
                onMouseDown={() => setPressed("rep")}
                onMouseUp={() => setPressed(null)}
                onMouseLeave={() => setPressed(null)}
                className="btn-primary w-full rounded-xl py-3 text-sm font-semibold"
                style={{
                  background: reportReady && !reportLoading ? `linear-gradient(135deg, ${INDIGO}, #4655D9)` : "rgba(255,255,255,0.06)",
                  color: reportReady && !reportLoading ? "#fff" : "#5A6289",
                  cursor: reportReady && !reportLoading ? "pointer" : "not-allowed",
                  ...pressStyle("rep"),
                }}
              >
                {reportLoading ? (
                  <span className="flex items-center justify-center gap-2">
                    <span style={{
                      width: 14, height: 14, borderRadius: "50%",
                      border: "2px solid rgba(255,255,255,0.35)", borderTopColor: "#fff",
                      animation: "spin 0.7s linear infinite", display: "inline-block",
                    }} />
                    Reading the noise…
                  </span>
                ) : report ? "Regenerate report" : "Generate theme report"}
              </button>
              {!reportReady && <p style={{ fontSize: 11, color: FAINT, marginTop: 8 }}>Brand and raw listening data are required.</p>}
            </section>
          </Reveal>

          {/* Report output */}
          <section className="lg:col-span-2">
            {!report && !reportLoading && (
              <Reveal delay={120}>
                <div className="rounded-2xl p-16 text-center" style={{ border: `1px dashed ${GLASS_BORDER}`, background: "rgba(255,255,255,0.015)" }}>
                  <div style={{
                    width: 52, height: 52, borderRadius: 16, margin: "0 auto 18px",
                    background: `linear-gradient(135deg, rgba(91,110,245,0.25), rgba(227,182,92,0.2))`,
                    display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22,
                  }}>💬</div>
                  <p style={{ fontSize: 19, fontWeight: 700, letterSpacing: "-0.01em" }}>Your theme report renders here</p>
                  <p style={{ fontSize: 13.5, color: FAINT, marginTop: 8, maxWidth: 380, marginLeft: "auto", marginRight: "auto" }}>
                    Paste social listening data and generate. You'll get sentiment split, ranked themes, and recommended actions.
                  </p>
                </div>
              </Reveal>
            )}

            {reportLoading && (
              <div className="rounded-2xl p-6" style={{ border: `1px solid ${GLASS_BORDER}`, background: GLASS }}>
                <div className="shimmer-block rounded-lg mb-4" style={{ height: 90 }} />
                <div className="shimmer-block rounded-lg mb-4" style={{ height: 36 }} />
                {[0, 1, 2].map((i) => (
                  <div key={i} className="shimmer-block rounded-lg mb-3" style={{ height: 70, opacity: 1 - i * 0.12 }} />
                ))}
                <p style={{ textAlign: "center", fontSize: 13, color: MUTED, marginTop: 10 }}>Clustering mentions into themes…</p>
              </div>
            )}

            {reportError && !reportLoading && (
              <div className="rounded-xl p-4 mb-4 text-sm" style={{ background: "rgba(242,114,114,0.1)", border: "1px solid rgba(242,114,114,0.35)", color: "#F5A3A3" }}>
                {reportError}
              </div>
            )}

            {report && !reportLoading && (
              <Reveal delay={20}>
                <div
                  className="rounded-2xl overflow-hidden print-area"
                  dir={reportLang === "ar" ? "rtl" : "ltr"}
                  style={{
                    border: `1px solid ${GLASS_BORDER}`, background: GLASS, backdropFilter: "blur(24px)",
                    fontFamily: reportLang === "ar" ? "'Tajawal', sans-serif" : undefined,
                  }}
                >
                  {/* Report header */}
                  <div className="px-7 py-6" style={{ background: "linear-gradient(135deg, rgba(91,110,245,0.16), rgba(227,182,92,0.08))", borderBottom: `1px solid ${GLASS_BORDER}` }}>
                    <p className="mono" style={{ fontSize: 11, color: GOLD, letterSpacing: "0.06em", marginBottom: 6, direction: "ltr" }}>
                      SOCIAL LISTENING · {(reportBrief.platforms || "ALL PLATFORMS").toUpperCase()} · {reportBrief.dateRange || "PERIOD N/A"}
                    </p>
                    <h3 style={{ fontSize: 24, fontWeight: 800, letterSpacing: "-0.02em" }}>
                      {reportBrief.brand} — {reportLang === "ar" ? "ملخص المحاور" : "Theme Summary"}
                    </h3>
                    <p style={{ fontSize: 13.5, color: "#C7CDE6", marginTop: 12, lineHeight: 1.5 }}>{report.executive_summary}</p>
                    <p style={{ fontSize: 11.5, color: MUTED, marginTop: 8 }}>
                      {reportLang === "ar" ? "الحجم التقديري" : "Estimated volume"}: {report.total_mentions_estimate}
                    </p>

                    <div className="flex gap-2 mt-4">
                      <button
                        onClick={exportReportPDF}
                        className="chip rounded-full px-3.5 py-1.5 text-xs font-semibold flex items-center gap-1.5"
                        style={{ border: `1px solid ${GLASS_BORDER}`, color: "#C7CDE6", background: "rgba(255,255,255,0.05)", cursor: "pointer" }}
                      >
                        ⬇ Download PDF
                      </button>
                      <button
                        onClick={() => exportPPT("report")}
                        disabled={exporting === "report"}
                        className="chip rounded-full px-3.5 py-1.5 text-xs font-semibold flex items-center gap-1.5"
                        style={{ border: `1px solid ${GLASS_BORDER}`, color: "#C7CDE6", background: "rgba(255,255,255,0.05)", cursor: "pointer" }}
                      >
                        {exporting === "report" ? "Preparing…" : "⬇ Download PPT"}
                      </button>
                    </div>
                    {exportMsg && exporting === null && (
                      <p className="print-hide" style={{ fontSize: 11, color: GOLD, marginTop: 8, maxWidth: 480 }}>{exportMsg}</p>
                    )}
                  </div>

                  {/* Sentiment split */}
                  <div className="px-7 pt-6 grid gap-6 md:grid-cols-[1fr_160px] items-center">
                  <div>
                    <p style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 10 }}>
                      {reportLang === "ar" ? "المشاعر العامة" : "Overall sentiment"}
                    </p>
                    <div className="flex w-full h-3 rounded-full overflow-hidden" style={{ background: "rgba(255,255,255,0.06)" }}>
                      <div style={{ width: reportBarsIn ? `${report.sentiment.positive}%` : "0%", background: "#39C982", transition: "width 0.9s cubic-bezier(.16,1,.3,1)" }} />
                      <div style={{ width: reportBarsIn ? `${report.sentiment.neutral}%` : "0%", background: "#8890B5", transition: "width 0.9s cubic-bezier(.16,1,.3,1) 0.08s" }} />
                      <div style={{ width: reportBarsIn ? `${report.sentiment.negative}%` : "0%", background: "#F27272", transition: "width 0.9s cubic-bezier(.16,1,.3,1) 0.16s" }} />
                    </div>
                    <div className="flex gap-5 mt-3 flex-wrap">
                      <span className="flex items-center gap-1.5" style={{ fontSize: 12 }}><span className="inline-block w-2.5 h-2.5 rounded-full" style={{ background: "#39C982" }} />{reportLang === "ar" ? "إيجابي" : "Positive"} <span className="mono" style={{ color: FAINT }}>{report.sentiment.positive}%</span></span>
                      <span className="flex items-center gap-1.5" style={{ fontSize: 12 }}><span className="inline-block w-2.5 h-2.5 rounded-full" style={{ background: "#8890B5" }} />{reportLang === "ar" ? "محايد" : "Neutral"} <span className="mono" style={{ color: FAINT }}>{report.sentiment.neutral}%</span></span>
                      <span className="flex items-center gap-1.5" style={{ fontSize: 12 }}><span className="inline-block w-2.5 h-2.5 rounded-full" style={{ background: "#F27272" }} />{reportLang === "ar" ? "سلبي" : "Negative"} <span className="mono" style={{ color: FAINT }}>{report.sentiment.negative}%</span></span>
                    </div>
                  </div>
                  <div style={{ width: 140, height: 140, margin: "0 auto" }} className="print-hide">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={[
                            { name: "Positive", value: report.sentiment.positive },
                            { name: "Neutral", value: report.sentiment.neutral },
                            { name: "Negative", value: report.sentiment.negative },
                          ]}
                          dataKey="value" nameKey="name" innerRadius={38} outerRadius={62} paddingAngle={2} stroke="none"
                        >
                          <Cell fill="#39C982" /><Cell fill="#8890B5" /><Cell fill="#F27272" />
                        </Pie>
                        <Tooltip contentStyle={{ background: "#14162A", border: `1px solid ${GLASS_BORDER}`, borderRadius: 8, fontSize: 12 }} />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                  </div>

                  {/* Theme volume bar chart */}
                  <div className="px-7 pt-6 print-hide" style={{ height: 40 + report.themes.length * 34 }}>
                    <p style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 6 }}>
                      {reportLang === "ar" ? "حجم المحاور" : "Theme volume"}
                    </p>
                    <ResponsiveContainer width="100%" height="90%">
                      <BarChart data={report.themes} layout="vertical" margin={{ left: 10, right: 20 }}>
                        <XAxis type="number" hide domain={[0, "dataMax + 5"]} />
                        <YAxis type="category" dataKey="name" width={140} tick={{ fill: MUTED, fontSize: 11 }} axisLine={false} tickLine={false} />
                        <Tooltip contentStyle={{ background: "#14162A", border: `1px solid ${GLASS_BORDER}`, borderRadius: 8, fontSize: 12 }} />
                        <Bar dataKey="volume_pct" radius={[0, 6, 6, 0]}>
                          {report.themes.map((t) => (
                            <Cell key={t.name} fill={t.sentiment === "Positive" ? "#39C982" : t.sentiment === "Negative" ? "#F27272" : t.sentiment === "Mixed" ? GOLD : "#8890B5"} />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>

                  {/* Themes */}
                  <div className="px-7 pt-6">
                    <p style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 10 }}>
                      {reportLang === "ar" ? "المحاور، مرتبة حسب الحجم" : "Themes, ranked by volume"}
                    </p>
                    <div className="space-y-3">
                      {report.themes.map((t, i) => {
                        const sentColor = t.sentiment === "Positive" ? "#39C982" : t.sentiment === "Negative" ? "#F27272" : t.sentiment === "Mixed" ? GOLD : "#8890B5";
                        return (
                          <Reveal key={t.name} delay={i * 80}>
                            <div className="card-hover rounded-xl p-4" style={{ background: "rgba(255,255,255,0.03)", border: `1px solid ${GLASS_BORDER}` }}>
                              <div className="flex justify-between items-start flex-wrap gap-2">
                                <div className="flex items-center gap-2">
                                  <span className="inline-block w-2.5 h-2.5 rounded-full" style={{ background: sentColor }} />
                                  <span style={{ fontSize: 15, fontWeight: 700 }}>{t.name}</span>
                                  <span className="mono" style={{ fontSize: 10.5, color: sentColor, border: `1px solid ${sentColor}55`, borderRadius: 999, padding: "1px 8px" }}>{t.sentiment}</span>
                                </div>
                                <span className="mono" style={{ fontSize: 12, color: GOLD, fontWeight: 700 }}>~{t.volume_pct}% of volume</span>
                              </div>
                              <p style={{ fontSize: 13, color: "#C7CDE6", marginTop: 8 }}>{t.description}</p>
                              <p style={{ fontSize: 12, color: FAINT, marginTop: 6, fontStyle: "italic" }}>"{t.example_snippet}"</p>
                              <p style={{ fontSize: 12, color: INDIGO_SOFT, marginTop: 6 }}>
                                <span style={{ fontWeight: 600 }}>Recommendation:</span> {t.recommendation}
                              </p>
                            </div>
                          </Reveal>
                        );
                      })}
                    </div>
                  </div>

                  {/* Overall recommendations */}
                  <div className="px-7 py-6 mt-6" style={{ background: "rgba(255,255,255,0.02)", borderTop: `1px solid ${GLASS_BORDER}` }}>
                    <p style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 10 }}>
                      {reportLang === "ar" ? "الإجراءات الموصى بها" : "Recommended actions"}
                    </p>
                    <ul style={{ fontSize: 13, lineHeight: 1.8 }}>
                      {report.overall_recommendations.map((r) => (
                        <li key={r} className="flex gap-2"><span style={{ color: GOLD }}>—</span>{r}</li>
                      ))}
                    </ul>
                    <p style={{ fontSize: 11.5, color: FAINT, marginTop: 14, lineHeight: 1.5 }}>
                      {reportLang === "ar"
                        ? "المحاور والمشاعر هي تقديرات مبنية على الذكاء الاصطناعي من العينة الملصقة. يجب على مسؤول التواصل الاجتماعي التحقق قبل اعتمادها في الاستراتيجية أو مشاركتها خارجياً."
                        : "Themes and sentiment are AI-generated estimates from the pasted sample. A social lead should validate before this informs strategy or is shared externally."}
                    </p>
                  </div>
                </div>
              </Reveal>
            )}
          </section>
        </main>
        )}

        <footer className="px-6 py-6 text-center print-hide" style={{ borderTop: `1px solid ${GLASS_BORDER}` }}>
          <p style={{ fontSize: 11.5, color: FAINT }}>Khottah — decision support only; human approval required before any spend.</p>
          <p style={{ fontSize: 11.5, fontWeight: 600, color: GOLD, marginTop: 4 }}>
            © {new Date().getFullYear()} Mohamed Emad — Developer &amp; Owner
          </p>
        </footer>
      </div>
    </div>
  );
}
