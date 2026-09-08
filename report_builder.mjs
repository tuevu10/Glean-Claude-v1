import fs from "node:fs/promises";
import { pathToFileURL } from "node:url";

const [inputPath, outputPath, logoPath] = process.argv.slice(2);
if (!inputPath || !outputPath || !logoPath || !process.env.ARTIFACT_TOOL_MODULE) {
  throw new Error("Report input, output, logo, and ARTIFACT_TOOL_MODULE are required");
}
const { Presentation, PresentationFile } = await import(
  pathToFileURL(process.env.ARTIFACT_TOOL_MODULE).href
);
const data = JSON.parse(await fs.readFile(inputPath, "utf8"));
const logo = new Uint8Array(await fs.readFile(logoPath));

const W = 1280;
const H = 720;
const C = {
  ink: "#141414", paper: "#FCFBF8", cream: "#F1EEE8", stone: "#DAD6CD",
  slate: "#6F6C66", indigo: "#343CED", violet: "#B3A6F5", magenta: "#C049C4",
  grid: "#E7E4DE", usage: "#6F52D9", projection: "#9B95A4",
};
const FONT = "Arial";
const presentation = Presentation.create({ slideSize: { width: W, height: H } });

function addText(slide, text, position, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox", position, fill: "none", line: { fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    typeface: FONT, fontSize: style.fontSize ?? 20, bold: style.bold ?? false,
    italic: style.italic ?? false, color: style.color ?? C.ink,
    alignment: style.alignment ?? "left", verticalAlignment: style.verticalAlignment ?? "top",
    autoFit: style.autoFit ?? "shrinkText", insets: style.insets ?? { left: 0, right: 0, top: 0, bottom: 0 },
  };
  return shape;
}

function addRect(slide, position, fill, line = "none", radius = "rect") {
  return slide.shapes.add({
    geometry: radius, position, fill,
    line: line === "none" ? { fill: "none", width: 0 } : line,
  });
}

function addLogo(slide, large = false) {
  slide.images.add({
    blob: logo, contentType: "image/svg+xml", alt: "Glean logo", fit: "contain",
    position: large ? { left: 1090, top: 53, width: 132, height: 54 }
                    : { left: 1155, top: 55, width: 78, height: 32 },
  });
}

function addChrome(slide, slideNumber) {
  slide.background.fill = C.paper;
  addRect(slide, { left: 0, top: 0, width: W, height: 8 }, C.indigo);
  addLogo(slide);
  addText(slide, String(slideNumber), { left: 1190, top: 676, width: 28, height: 18 },
          { fontSize: 11, color: C.slate, alignment: "right" });
}

function usd(value) {
  if (value === null || value === undefined) return "N/A";
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(value);
}
function usdPrecise(value) {
  if (value === null || value === undefined) return "N/A";
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
}
function credits(value) {
  if (value === null || value === undefined) return "N/A";
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(value);
}
function pct(value) {
  if (value === null || value === undefined) return "N/A";
  return `${(value * 100).toFixed(1)}%`;
}

// Cover
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  addRect(slide, { left: 0, top: 0, width: W, height: 8 }, C.indigo);
  addLogo(slide, true);
  addText(slide, data.title, { left: 96, top: 300, width: 850, height: 120 },
          { fontSize: 58, bold: true, autoFit: "none", verticalAlignment: "middle" });
  addText(slide, data.month, { left: 98, top: 518, width: 400, height: 42 },
          { fontSize: 26, italic: true, color: C.slate, autoFit: "none" });
}

// Executive summary
{
  const slide = presentation.slides.add();
  addChrome(slide, 2);
  addText(slide, "Executive Summary", { left: 56, top: 50, width: 860, height: 54 },
          { fontSize: 34, bold: true, autoFit: "none" });

  const p = data.portfolio;
  const projectedShare = p.year_end_consumption != null && p.year_end_contracts
    ? pct(p.year_end_consumption / p.year_end_contracts) : "unavailable";
  const outlook = p.year_end_consumption == null
    ? `Customers consumed ${credits(p.ytd_usage)} credits through ${data.as_of}. The current data does not support a complete year-end forecast.`
    : `Customers consumed ${credits(p.ytd_usage)} credits through ${data.as_of}. At validated recent usage rates, calendar-year consumption projects to ${credits(p.year_end_consumption)} credits against ${credits(p.year_end_contracts)} projected contracted credits at year end.`;
  addRect(slide, { left: 54, top: 126, width: 1172, height: 96 }, "#EAE7FF",
          { fill: C.violet, width: 1 }, "roundRect");
  addText(slide, outlook, { left: 72, top: 146, width: 1135, height: 58 },
          { fontSize: 20, bold: true, verticalAlignment: "middle" });

  addText(slide, "1  Current usage and outlook", { left: 58, top: 258, width: 360, height: 34 },
          { fontSize: 20, bold: true, color: C.indigo, autoFit: "none" });
  addText(slide,
    `Current contract utilization is ${pct(p.utilization)} across ${credits(p.entitlement)} credits of entitlement, representing ${usd(p.contract_value)} in annual contract value. The year-end usage outlook equals ${projectedShare} of projected year-end contracted credits.`,
    { left: 58, top: 300, width: 1165, height: 74 }, { fontSize: 18, autoFit: "shrinkText" });

  addText(slide, "2  Top customers by contract value", { left: 58, top: 392, width: 430, height: 34 },
          { fontSize: 20, bold: true, color: C.indigo, autoFit: "none" });
  const topText = data.top_three.map((item) =>
    `${item.tracking_sentence} Its annual contract value is ${usd(item.contract_value)}.`).join(" ");
  addText(slide, topText, { left: 58, top: 433, width: 1165, height: 92 },
          { fontSize: 18, autoFit: "shrinkText" });

  addText(slide, "3  Issues requiring attention", { left: 58, top: 542, width: 420, height: 34 },
          { fontSize: 20, bold: true, color: C.indigo, autoFit: "none" });
  const issueText = data.issues.length
    ? data.issues.map((item) =>
        `${item.customer_id} has ${usdPrecise(item.amount)} of ${item.status === "OVER ENTITLEMENT" ? "current" : "projected"} overage exposure. The item is ${item.resolution.toLowerCase()} and assigned to ${item.owner}.`).join(" ")
    : "No customers currently meet the over-entitlement or early-exhaustion rules.";
  addText(slide, issueText, { left: 58, top: 583, width: 1165, height: 64 },
          { fontSize: 18, autoFit: "shrinkText" });
  addText(slide, `As of ${data.as_of}. Forecasts use validated trailing 30-day rates.`,
          { left: 58, top: 680, width: 900, height: 18 }, { fontSize: 10, color: C.slate });
}

// Annual chart and metrics
{
  const slide = presentation.slides.add();
  addChrome(slide, 3);
  addText(slide, "Annual Usage Overview", { left: 56, top: 50, width: 820, height: 54 },
          { fontSize: 34, bold: true, autoFit: "none" });
  addText(slide, `Actual usage through ${data.as_of}. Remaining months show the current outlook`,
          { left: 58, top: 106, width: 900, height: 30 }, { fontSize: 18, italic: true, color: C.slate });

  const allValues = Object.values(data.chart).flat().filter((v) => typeof v === "number");
  const maxValue = Math.max(...allValues, 1);
  const axisMax = Math.ceil(maxValue / 1000000) * 1000000;
  const majorUnit = axisMax / 4;
  const contractedSeries = data.chart.actual_contracts.map((value, index) =>
    value ?? data.chart.projected_contracts[index]);
  const usageSeries = data.chart.actual_usage.map((value, index) =>
    value ?? data.chart.projected_usage[index]);
  const plotLeft = 108;
  const plotWidth = 1090;
  const projectionLeft = plotLeft + plotWidth * (data.chart.snapshot_month / 12);
  const chart = slide.charts.add("line", {
    position: { left: 58, top: 162, width: 1164, height: 360 },
    categories: data.chart.categories,
    series: [
      { name: "Total contracted credits", values: contractedSeries,
        line: { style: "solid", fill: "#AAA6B0", width: 3 }, marker: { symbol: "none" } },
      { name: "Cumulative consumption", values: usageSeries,
        line: { style: "solid", fill: C.usage, width: 3 }, marker: { symbol: "none" } },
    ],
    lineOptions: { grouping: "standard", smooth: false }, hasLegend: false,
    chartFill: "transparent",
    plotAreaFill: "transparent",
    chartLine: { fill: "none", width: 0 }, plotAreaLine: { fill: "none", width: 0 },
    xAxis: { majorUnit: 1, textStyle: { typeface: FONT, fill: C.slate, fontSize: 13 },
             line: { style: "solid", fill: C.stone, width: 1 }, majorGridlines: null },
    yAxis: { min: 0, max: axisMax, majorUnit, numberFormatCode: "0.0,,\\M",
             textStyle: { typeface: FONT, fill: C.slate, fontSize: 13 },
             line: { fill: "none", width: 0 },
             majorGridlines: { style: "solid", fill: C.grid, width: 1 } },
  });
  chart.title = "";
  addRect(slide, { left: projectionLeft, top: 190, width: plotLeft + plotWidth - projectionLeft, height: 304 },
          `${C.cream}/55`);
  addText(slide, "Projection", { left: 934, top: 144, width: 150, height: 24 },
          { fontSize: 14, italic: true, color: C.slate, alignment: "center" });
  addText(slide, "●  Total contracted credits", { left: 350, top: 522, width: 255, height: 24 },
          { fontSize: 14, color: "#8E8A96", autoFit: "none" });
  addText(slide, "●  Cumulative consumption", { left: 640, top: 522, width: 260, height: 24 },
          { fontSize: 14, color: C.usage, autoFit: "none" });

  const metrics = [
    ["Contracted credits", credits(data.portfolio.entitlement)],
    ["Credits consumed YTD", credits(data.portfolio.ytd_usage)],
    ["Contract utilization", pct(data.portfolio.utilization)],
    ["Over entitlement", String(data.portfolio.over_count)],
    ["Early exhaustion risk", String(data.portfolio.risk_count)],
    ["Current overage value", usdPrecise(data.portfolio.current_overage)],
  ];
  const left = 58;
  const width = 1164 / metrics.length;
  metrics.forEach(([label, value], index) => {
    const x = left + index * width;
    if (index > 0) addRect(slide, { left: x, top: 578, width: 1, height: 75 }, C.stone);
    addText(slide, value, { left: x + 12, top: 580, width: width - 24, height: 35 },
            { fontSize: 24, bold: true, color: index >= 3 ? C.indigo : C.ink, alignment: "center" });
    addText(slide, label, { left: x + 10, top: 620, width: width - 20, height: 36 },
            { fontSize: 13, color: C.slate, alignment: "center" });
  });
  addText(slide,
    "Projected contracted credits use the average monthly additions observed through the last completed month. Projected consumption uses each account’s validated trailing 30-day daily rate through contract end.",
    { left: 58, top: 678, width: 1080, height: 24 }, { fontSize: 10, color: C.slate });
  slide.speakerNotes.textFrame.setText(
    `Source workbook: ${data.source_name}. Source SHA-256: ${data.source_hash}. ` +
    "All figures were calculated in Python before presentation generation."
  );
}

await (await PresentationFile.exportPptx(presentation)).save(outputPath);
