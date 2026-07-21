// Server-side PPTX generation using pptxgenjs. This runs on the deployed
// backend only — it cannot run inside a sandboxed browser preview, which is
// why "Download PPT" only works once this app is actually deployed.

const NAVY = "1F2A56";
const INDIGO = "3B4CCA";
const GOLD = "C9A24B";
const PAPER = "F7F8FC";
const INK = "232A45";
const GREY = "6B7398";

function buildPlanDeck(pptxgen, data, meta) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";

  const total = (data.channels || []).reduce((a, c) => a + Number(c.budget_sar || 0), 0);

  // Title slide
  let s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText("KHOTTAH", { x: 0.7, y: 2.3, w: 8, h: 0.8, fontFace: "Calibri", fontSize: 40, bold: true, color: "FFFFFF" });
  s.addText(data.campaign_name || "Media Plan", { x: 0.7, y: 3.1, w: 10, h: 0.6, fontFace: "Calibri", fontSize: 20, color: "C7CDE6" });
  s.addText(`Net budget: SAR ${total.toLocaleString("en-US")}`, { x: 0.7, y: 3.7, w: 8, h: 0.4, fontFace: "Courier New", fontSize: 13, color: GOLD });
  s.addText(data.strategy_summary || "", { x: 0.7, y: 4.3, w: 10.5, h: 1.2, fontFace: "Calibri", fontSize: 13, color: "AEB6D9" });

  // Channel table slide
  s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  s.addText("Channel Mix & Budget", { x: 0.5, y: 0.35, w: 10, h: 0.5, fontFace: "Calibri", fontSize: 24, bold: true, color: NAVY });
  const rows = [[
    { text: "Channel", options: { bold: true, color: "FFFFFF", fill: NAVY } },
    { text: "Role", options: { bold: true, color: "FFFFFF", fill: NAVY } },
    { text: "KPI", options: { bold: true, color: "FFFFFF", fill: NAVY } },
    { text: "Budget (SAR)", options: { bold: true, color: "FFFFFF", fill: NAVY } },
  ]];
  (data.channels || []).forEach((c) => {
    rows.push([
      { text: c.name || "" },
      { text: c.role || "" },
      { text: c.kpi || "" },
      { text: Number(c.budget_sar || 0).toLocaleString("en-US"), options: { align: "right" } },
    ]);
  });
  s.addTable(rows, { x: 0.5, y: 1.0, w: 12.3, colW: [2.6, 4.5, 3.2, 2.0], fontSize: 11, border: { type: "solid", color: "D9DEEF", pt: 0.5 } });

  // Creative recommendations slide
  s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  s.addText("Creative Recommendations", { x: 0.5, y: 0.35, w: 10, h: 0.5, fontFace: "Calibri", fontSize: 24, bold: true, color: NAVY });
  const creativeRows = [[
    { text: "Channel", options: { bold: true, color: "FFFFFF", fill: NAVY } },
    { text: "Best Format", options: { bold: true, color: "FFFFFF", fill: NAVY } },
    { text: "Creative Recommendation", options: { bold: true, color: "FFFFFF", fill: NAVY } },
  ]];
  (data.channels || []).forEach((c) => {
    creativeRows.push([
      { text: c.name || "" },
      { text: c.best_format || c.formats || "" },
      { text: c.creative_recommendation || "" },
    ]);
  });
  s.addTable(creativeRows, { x: 0.5, y: 1.0, w: 12.3, colW: [2.2, 3.3, 6.8], fontSize: 10.5, border: { type: "solid", color: "D9DEEF", pt: 0.5 }, valign: "top" });

  // Flighting + KPIs + rationale slide
  s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  s.addText("Flighting, KPIs & Rationale", { x: 0.5, y: 0.35, w: 10, h: 0.5, fontFace: "Calibri", fontSize: 24, bold: true, color: NAVY });
  let fy = 1.1;
  (data.flighting || []).forEach((f) => {
    s.addText(`${f.weeks}  —  ${f.phase}: ${f.focus}`, { x: 0.5, y: fy, w: 12, h: 0.4, fontFace: "Calibri", fontSize: 12.5, color: INK });
    fy += 0.45;
  });
  s.addText("KPIs", { x: 0.5, y: fy + 0.2, w: 5, h: 0.35, fontFace: "Calibri", fontSize: 14, bold: true, color: INDIGO });
  s.addText((data.kpis || []).map((k) => `• ${k}`).join("\n"), { x: 0.5, y: fy + 0.6, w: 6, h: 1.2, fontFace: "Calibri", fontSize: 12, color: INK });
  s.addText("Rationale", { x: 6.8, y: fy + 0.2, w: 5, h: 0.35, fontFace: "Calibri", fontSize: 14, bold: true, color: INDIGO });
  s.addText(data.rationale || "", { x: 6.8, y: fy + 0.6, w: 6, h: 1.5, fontFace: "Calibri", fontSize: 12, color: INK });

  return pres;
}

function buildReportDeck(pptxgen, data, meta) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";

  // Title slide
  let s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText("KHOTTAH", { x: 0.7, y: 2.1, w: 8, h: 0.8, fontFace: "Calibri", fontSize: 40, bold: true, color: "FFFFFF" });
  s.addText(`${meta.brand || ""} — Social Listening Report`, { x: 0.7, y: 2.95, w: 11, h: 0.6, fontFace: "Calibri", fontSize: 20, color: "C7CDE6" });
  s.addText(`${meta.platforms || ""}  ·  ${meta.dateRange || ""}`, { x: 0.7, y: 3.5, w: 8, h: 0.4, fontFace: "Courier New", fontSize: 12, color: GOLD });
  s.addText(data.executive_summary || "", { x: 0.7, y: 4.1, w: 10.8, h: 1.3, fontFace: "Calibri", fontSize: 13, color: "AEB6D9" });

  const sent = data.sentiment || { positive: 0, neutral: 0, negative: 0 };
  s.addChart(pptxgen.ChartType.doughnut,
    [{ name: "Sentiment", labels: ["Positive", "Neutral", "Negative"], values: [sent.positive, sent.neutral, sent.negative] }],
    { x: 9.5, y: 4.0, w: 3.0, h: 2.4, chartColors: ["39C982", "8890B5", "F27272"], showLegend: true, legendPos: "b", legendFontSize: 9, dataLabelColor: "FFFFFF" }
  );

  // Themes slide
  s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  s.addText("Themes", { x: 0.5, y: 0.35, w: 10, h: 0.5, fontFace: "Calibri", fontSize: 24, bold: true, color: NAVY });
  const rows = [[
    { text: "Theme", options: { bold: true, color: "FFFFFF", fill: NAVY } },
    { text: "Sentiment", options: { bold: true, color: "FFFFFF", fill: NAVY } },
    { text: "Volume", options: { bold: true, color: "FFFFFF", fill: NAVY } },
    { text: "Description", options: { bold: true, color: "FFFFFF", fill: NAVY } },
  ]];
  (data.themes || []).forEach((t) => {
    rows.push([
      { text: t.name || "" },
      { text: t.sentiment || "" },
      { text: `~${t.volume_pct || 0}%`, options: { align: "right" } },
      { text: t.description || "" },
    ]);
  });
  s.addTable(rows, { x: 0.5, y: 1.0, w: 12.3, colW: [2.6, 1.8, 1.4, 6.5], fontSize: 10.5, border: { type: "solid", color: "D9DEEF", pt: 0.5 } });

  // Recommendations slide
  s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  s.addText("Recommended Actions", { x: 0.5, y: 0.35, w: 10, h: 0.5, fontFace: "Calibri", fontSize: 24, bold: true, color: NAVY });
  s.addText((data.overall_recommendations || []).map((r) => `•  ${r}`).join("\n\n"), { x: 0.5, y: 1.1, w: 12, h: 5, fontFace: "Calibri", fontSize: 14, color: INK, lineSpacingMultiple: 1.4 });

  return pres;
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }
  const { type, data, meta } = req.body || {};
  if (!type || !data) {
    return res.status(400).json({ error: "Missing 'type' or 'data' in request body" });
  }

  try {
    const pptxgen = (await import("pptxgenjs")).default;
    const pres = type === "plan" ? buildPlanDeck(pptxgen, data, meta || {}) : buildReportDeck(pptxgen, data, meta || {});
    const buffer = await pres.write({ outputType: "nodebuffer" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="khottah-${type}.pptx"`);
    return res.status(200).send(buffer);
  } catch (err) {
    return res.status(500).json({ error: "PPTX generation failed", detail: String(err) });
  }
}
