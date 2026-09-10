import fs from "node:fs/promises";

const [inputPath, outputPath, previewPath] = process.argv.slice(2);
const artifactModule = process.env.ARTIFACT_TOOL_MODULE;
if (!artifactModule) throw new Error("ARTIFACT_TOOL_MODULE is required");
const { SpreadsheetFile, Workbook } = await import(`file:///${artifactModule.replaceAll("\\", "/")}`);
const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
const wb = Workbook.create();
const queue = wb.worksheets.add("Action Queue");
const metrics = wb.worksheets.add("Account Metrics");
const assumptions = wb.worksheets.add("Assumptions");
const contracts = wb.worksheets.add("Contracts");
const usage = wb.worksheets.add("Usage Events");

const purple = "#6E35C8", violetWash = "#F3EEFF", ink = "#24343D";
const stone = "#DAD6CD", cream = "#F1EEE8", gray = "#F5F8F9", green = "#008A52";
const border = { preset: "all", style: "thin", color: stone };
const excelDate = (iso) => iso ? new Date(`${iso}T00:00:00Z`) : null;
const colName = (n) => {
  let s = "";
  while (n > 0) { n--; s = String.fromCharCode(65 + n % 26) + s; n = Math.floor(n / 26); }
  return s;
};
const title = (sheet, lastCol, text, subtitle) => {
  sheet.getRange(`A1:${lastCol}1`).merge();
  sheet.getRange("A1").values = [[text]];
  sheet.getRange("A1").format = { fill: purple, font: { color: "#FFFFFF", bold: true, size: 18 }, rowHeight: 30 };
  sheet.getRange(`A2:${lastCol}2`).merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2").format = { fill: violetWash, font: { color: "#62537C", italic: true, size: 10 }, rowHeight: 22 };
};
const header = (range) => {
  range.format = { fill: cream, font: { color: ink, bold: true }, borders: border, wrapText: true, verticalAlignment: "center" };
  range.format.rowHeight = 34;
};

// Assumptions are explicit inputs used by every rule formula.
title(assumptions, "D", "Monitoring Assumptions", "Editable thresholds and immutable source metadata");
const assumptionRows = [
  ["Parameter", "Value", "Unit", "Purpose"],
  ["Analysis as of", excelDate(payload.as_of), "date", "Latest usage-event date; reproducible reporting boundary"],
  ["Cutoff", null, "date", "Exclusive day boundary after the as-of date"],
  ["Projected utilization threshold", payload.rules.projected_utilization_threshold, "%", "Early exhaustion risk when strictly exceeded"],
  ["Early exhaustion buffer", payload.rules.early_exhaustion_days, "days", "Risk when exhaustion is earlier than contract end minus buffer"],
  ["Underutilization threshold", payload.rules.underutilization_threshold, "%", "Underutilizing when projection is strictly below"],
  ["Minimum contract elapsed", payload.rules.underutilization_min_elapsed, "%", "Minimum elapsed time for underutilization flag"],
  ["Usage acceleration threshold", payload.rules.usage_acceleration_threshold, "%", "Recent growth must strictly exceed this threshold"],
  ["Acceleration minimum increase", payload.rules.acceleration_min_increase_credits, "credits", "Absolute increase required across the two windows"],
  ["Acceleration history", payload.rules.acceleration_min_history_days, "days", "Complete comparable history required"],
  ["Over entitlement boundary", payload.rules.over_entitlement_threshold, "%", "Fixed classification boundary"],
  ["Source workbook", payload.source_name, "", "Workbook used to create this export"],
  ["Source SHA-256", payload.source_hash, "", "Source integrity fingerprint"],
];
assumptions.getRange(`A3:D${2 + assumptionRows.length}`).values = assumptionRows;
assumptions.getRange("B5").formulas = [["=B4+1"]];
header(assumptions.getRange("A3:D3"));
assumptions.getRange("B4:B5").format.numberFormat = "yyyy-mm-dd";
assumptions.getRange("B6:B6").format.numberFormat = "0.0%";
assumptions.getRange("B8:B10").format.numberFormat = "0.0%";
assumptions.getRange("B13:B13").format.numberFormat = "0.0%";
assumptions.getRange("A4:D15").format.borders = border;
assumptions.getRange("A1:D15").format.autofitColumns();
assumptions.getRange("A:A").format.columnWidth = 31;
assumptions.getRange("D:D").format.columnWidth = 62;
assumptions.showGridLines = false;
assumptions.freezePanes.freezeRows(3);

// Normalized source tables.
title(contracts, "F", "Contracts", "Normalized source fields; one row per customer");
contracts.getRange("A3:F3").values = [["Customer ID", "Contract Start", "Term Months", "Annual Entitlement Credits", "Annual Contract Value USD", "Source Excel Row"]];
header(contracts.getRange("A3:F3"));
const contractRows = payload.contracts.map(r => [r.customer_id, excelDate(r.contract_start), r.term_months, r.annual_entitlement_credits, r.annual_contract_value_usd, r.source_excel_row]);
if (contractRows.length) contracts.getRange(`A4:F${3 + contractRows.length}`).values = contractRows;
contracts.getRange(`B4:B${3 + contractRows.length}`).format.numberFormat = "yyyy-mm-dd";
contracts.getRange(`D4:D${3 + contractRows.length}`).format.numberFormat = "#,##0.0";
contracts.getRange(`E4:E${3 + contractRows.length}`).format.numberFormat = "$#,##0.00";
contracts.getRange(`A3:F${3 + contractRows.length}`).format.borders = border;
contracts.getRange(`A1:F${3 + contractRows.length}`).format.autofitColumns();
contracts.getRange("A:A").format.columnWidth = 18;
contracts.getRange("D:E").format.columnWidth = 24;
contracts.showGridLines = false;
contracts.freezePanes.freezeRows(3);

title(usage, "E", "Usage Events", "Normalized daily source rows with a formula-based cumulative audit column");
usage.getRange("A3:E3").values = [["Date", "Customer ID", "Credits Used", "Cumulative Customer Credits", "Source Excel Row"]];
header(usage.getRange("A3:E3"));
const usageRows = payload.usage.map(r => [excelDate(r.date), r.customer_id, r.credits_used, null, r.source_excel_row]);
if (usageRows.length) usage.getRange(`A4:E${3 + usageRows.length}`).values = usageRows;
if (usageRows.length) {
  usage.getRange("D4").formulas = [["=SUMIFS($C$4:C4,$B$4:B4,B4)"]];
  usage.getRange(`D4:D${3 + usageRows.length}`).fillDown();
}
usage.getRange(`A4:A${3 + usageRows.length}`).format.numberFormat = "yyyy-mm-dd";
usage.getRange(`C4:D${3 + usageRows.length}`).format.numberFormat = "#,##0.0";
usage.getRange(`D4:D${3 + usageRows.length}`).format.font = { color: green };
usage.getRange(`A3:E${3 + usageRows.length}`).format.borders = border;
usage.getRange(`A1:E${3 + usageRows.length}`).format.autofitColumns();
usage.getRange("B:B").format.columnWidth = 18;
usage.getRange("D:D").format.columnWidth = 27;
usage.showGridLines = false;
usage.freezePanes.freezeRows(3);

const metricHeaders = ["Customer ID","Contract Start","Term Months","Contract End","Entitlement Credits","Contract Value USD","Value per Credit","As Of","Cutoff","Lifecycle","Credits Used","Remaining Credits","Utilization %","Contract Days","Days Elapsed","Days Remaining","Contract Elapsed %","Missing Usage Days","Forecast Available","Pacing Index","Trailing Window Days","Trailing 30-Day Credits","Trailing Daily Average","Prior Window Days","Prior 30-Day Credits","Prior Daily Average","30-Day Growth %","Usage Accelerating","Forecast Remaining Usage","Projected Total Usage","Projected Utilization %","Days to Exhaustion","Estimated Exhaustion Date","Overage Credits","Estimated Overage Value","Projected Overage Credits","Projected Overage Value","Projected Unused Credits","Projected Unused Value","Status","Priority","Recommended Owner","Recommended Action","Status Reason","Data Notes"];
const metricLast = colName(metricHeaders.length);
title(metrics, metricLast, "Account Metrics", "Green formula text indicates a calculated or linked value; source data remains on Contracts and Usage Events");
metrics.getRange(`A3:${metricLast}3`).values = [metricHeaders];
header(metrics.getRange(`A3:${metricLast}3`));
const uEnd = 3 + usageRows.length;
const cStart = 4;
for (let i = 0; i < contractRows.length; i++) {
  const r = 4 + i, cr = cStart + i;
  const f = [];
  f[0] = `='Contracts'!A${cr}`;
  f[1] = `='Contracts'!B${cr}`;
  f[2] = `='Contracts'!C${cr}`;
  f[3] = `=EDATE(B${r},C${r})`;
  f[4] = `='Contracts'!D${cr}`;
  f[5] = `='Contracts'!E${cr}`;
  f[6] = `=F${r}/E${r}`;
  f[7] = "='Assumptions'!B4";
  f[8] = "='Assumptions'!B5";
  f[9] = `=IF(I${r}<=B${r},"NOT STARTED",IF(I${r}>=D${r},"EXPIRED","ACTIVE"))`;
  f[10] = `=SUMIFS('Usage Events'!$C$4:$C$${uEnd},'Usage Events'!$B$4:$B$${uEnd},A${r})`;
  f[11] = `=E${r}-K${r}`;
  f[12] = `=K${r}/E${r}`;
  f[13] = `=D${r}-B${r}`;
  f[14] = `=MAX(0,MIN(I${r}-B${r},N${r}))`;
  f[15] = `=MAX(0,N${r}-O${r})`;
  f[16] = `=O${r}/N${r}`;
  f[17] = `=MAX(0,MAX(0,MIN(I${r},D${r})-B${r})-COUNTIFS('Usage Events'!$B$4:$B$${uEnd},A${r},'Usage Events'!$A$4:$A$${uEnd},">="&B${r},'Usage Events'!$A$4:$A$${uEnd},"<"&MIN(I${r},D${r})))`;
  f[18] = `=AND(J${r}<>"NOT STARTED",R${r}=0)`;
  f[19] = `=IF(OR(R${r}>0,O${r}=0),"",M${r}/Q${r})`;
  f[20] = `=MAX(0,MIN(I${r},D${r})-MAX(I${r}-30,B${r}))`;
  f[21] = `=SUMIFS('Usage Events'!$C$4:$C$${uEnd},'Usage Events'!$B$4:$B$${uEnd},A${r},'Usage Events'!$A$4:$A$${uEnd},">="&I${r}-30,'Usage Events'!$A$4:$A$${uEnd},"<"&I${r})`;
  f[22] = `=IF(R${r}>0,"",IF(U${r}>0,V${r}/U${r},0))`;
  f[23] = `=MAX(0,MIN(I${r}-30,D${r})-MAX(I${r}-60,B${r}))`;
  f[24] = `=SUMIFS('Usage Events'!$C$4:$C$${uEnd},'Usage Events'!$B$4:$B$${uEnd},A${r},'Usage Events'!$A$4:$A$${uEnd},">="&I${r}-60,'Usage Events'!$A$4:$A$${uEnd},"<"&I${r}-30)`;
  f[25] = `=IF(R${r}>0,"",IF(X${r}>0,Y${r}/X${r},0))`;
  f[26] = `=IF(R${r}>0,"",IF(Y${r}>0,V${r}/Y${r}-1,IF(V${r}=0,0,"")))`;
  f[27] = `=AND(J${r}="ACTIVE",R${r}=0,O${r}>='Assumptions'!B12,U${r}=30,X${r}=30,V${r}-Y${r}>='Assumptions'!B11,OR(AND(Y${r}>0,AA${r}>'Assumptions'!B10),AND(Y${r}=0,V${r}>0)))`;
  f[28] = `=IF(S${r},IF(J${r}="ACTIVE",W${r}*P${r},0),"")`;
  f[29] = `=IF(S${r},K${r}+AC${r},"")`;
  f[30] = `=IF(S${r},AD${r}/E${r},"")`;
  f[31] = `=IF(K${r}>=E${r},0,IF(AND(S${r},J${r}="ACTIVE",W${r}>0),L${r}/W${r},""))`;
  f[32] = `=IF(K${r}>=E${r},IFERROR(_xlfn.MINIFS('Usage Events'!$A$4:$A$${uEnd},'Usage Events'!$B$4:$B$${uEnd},A${r},'Usage Events'!$D$4:$D$${uEnd},">="&E${r}),""),IF(AND(S${r},J${r}="ACTIVE",W${r}>0),H${r}+ROUNDUP(AF${r},0),""))`;
  f[33] = `=MAX(K${r}-E${r},0)`;
  f[34] = `=AH${r}*G${r}`;
  f[35] = `=IF(S${r},MAX(AD${r}-E${r},0),"")`;
  f[36] = `=IF(S${r},AJ${r}*G${r},"")`;
  f[37] = `=IF(S${r},MAX(E${r}-AD${r},0),"")`;
  f[38] = `=IF(S${r},AL${r}*G${r},"")`;
  f[39] = `=IF(M${r}>='Assumptions'!B13,"OVER ENTITLEMENT",IF(R${r}>0,"DATA REVIEW",IF(J${r}="NOT STARTED","NOT STARTED",IF(AND(J${r}="ACTIVE",OR(AE${r}>'Assumptions'!B6,AND(AG${r}<>"",AG${r}<D${r}-'Assumptions'!B7))),"EARLY EXHAUSTION RISK",IF(AND(Q${r}>='Assumptions'!B9,AE${r}<'Assumptions'!B8),"UNDERUTILIZING","ON TRACK")))))`;
  f[40] = `=IF(AN${r}="OVER ENTITLEMENT","P1 · Urgent",IF(OR(AN${r}="EARLY EXHAUSTION RISK",AN${r}="DATA REVIEW"),"P2 · High",IF(OR(AN${r}="UNDERUTILIZING",AND(AN${r}="ON TRACK",OR(J${r}="EXPIRED",AB${r}))),"P3 · Medium","P4 · Routine")))`;
  f[41] = `=IF(OR(AN${r}="OVER ENTITLEMENT",AN${r}="DATA REVIEW"),"Revenue Accounting / Billing",IF(AN${r}="EARLY EXHAUSTION RISK","Sales / Account Management",IF(AN${r}="UNDERUTILIZING","Customer Success",IF(OR(J${r}="EXPIRED",AB${r}),"Sales / Account Management","No Action"))))`;
  const baseAction = `IF(AN${r}="OVER ENTITLEMENT","Reconcile credit ledger and verify overage terms; coordinate a top-up or contract amendment.",IF(AN${r}="DATA REVIEW","Verify usage-feed completeness and supply explicit zero-usage days before acting on forecasts.",IF(AN${r}="NOT STARTED","Wait for contract activation; no usage forecast is available.",IF(AN${r}="EARLY EXHAUSTION RISK","Review forecast with the account team and discuss additional credits before exhaustion.",IF(AN${r}="UNDERUTILIZING","Review adoption barriers, agree a usage plan, and assess renewal implications.","Continue routine monitoring.")))))`;
  f[42] = `=${baseAction}&IF(AND(J${r}="EXPIRED",AN${r}<>"NOT STARTED")," Contract expired: confirm final reconciliation and renewal disposition.","")&IF(AB${r}," Investigate the recent usage increase and confirm whether it will persist.","")`;
  f[43] = `=IF(AN${r}="OVER ENTITLEMENT","Consumed credits have reached or exceeded entitlement.",IF(AN${r}="DATA REVIEW","Missing daily records prevent a reliable pacing or forecast classification.",IF(AN${r}="NOT STARTED","The contract has not started; forecast-dependent values are unavailable.",IF(AN${r}="EARLY EXHAUSTION RISK",IF(AE${r}>'Assumptions'!B6,"Projected utilization exceeds the configured threshold. ","")&IF(AND(AG${r}<>"",AG${r}<D${r}-'Assumptions'!B7),"Estimated exhaustion is earlier than the configured contract-end buffer.",""),IF(AN${r}="UNDERUTILIZING","Enough contract time has elapsed and projected utilization is below the configured threshold.","No primary exception rule is triggered.")))))`;
  f[44] = `=IF(R${r}>0,R${r}&" missing daily records; observed usage is incomplete and forecasts are withheld",IF(AND(U${r}>0,U${r}<30),"Short trailing window: rate uses eligible contract days",IF(AND(W${r}=0,K${r}<E${r}),"No recent rate; exhaustion date unavailable",IF(J${r}="NOT STARTED","Forecast unavailable before contract start",""))))`;
  metrics.getRange(`A${r}:${metricLast}${r}`).formulas = [f];
}
const mEnd = 3 + contractRows.length;
if (contractRows.length) {
  metrics.getRange(`A4:${metricLast}${mEnd}`).format.font = { color: green };
  metrics.getRange(`B4:B${mEnd}`).format.numberFormat = "yyyy-mm-dd";
  metrics.getRange(`D4:D${mEnd}`).format.numberFormat = "yyyy-mm-dd";
  metrics.getRange(`H4:I${mEnd}`).format.numberFormat = "yyyy-mm-dd";
  metrics.getRange(`AG4:AG${mEnd}`).format.numberFormat = "yyyy-mm-dd";
  for (const col of ["E","K","L","V","W","Y","Z","AC","AD","AH","AJ","AL"]) metrics.getRange(`${col}4:${col}${mEnd}`).format.numberFormat = "#,##0.0";
  for (const col of ["F","G","AI","AK","AM"]) metrics.getRange(`${col}4:${col}${mEnd}`).format.numberFormat = "$#,##0.00";
  for (const col of ["M","Q","AA","AE"]) metrics.getRange(`${col}4:${col}${mEnd}`).format.numberFormat = "0.0%";
  metrics.getRange(`A3:${metricLast}${mEnd}`).format.borders = border;
}
metrics.getRange(`A1:${metricLast}${Math.max(mEnd,4)}`).format.autofitColumns();
for (const col of ["AQ","AR","AS"]) metrics.getRange(`${col}:${col}`).format.columnWidth = 60;
metrics.getRange("A:A").format.columnWidth = 18;
metrics.showGridLines = false;
metrics.freezePanes.freezeRows(3);
metrics.freezePanes.freezeColumns(1);

const queueHeaders = ["Priority","Customer","Status","Credits Used","Entitlement","Utilization %","Contract Elapsed %","Pacing Index","Projected Utilization %","Days to Exhaustion","Estimated Overage Value","Recommended Owner","Recommended Action"];
title(queue, "M", "Finance Action Queue", "Filtered dashboard selection, linked to formula-based Account Metrics");
queue.getRange("A3:M3").values = [queueHeaders];
header(queue.getRange("A3:M3"));
for (let i = 0; i < payload.queue_customer_ids.length; i++) {
  const qr = 4 + i;
  const index = payload.contracts.findIndex(c => c.customer_id === payload.queue_customer_ids[i]);
  if (index < 0) continue;
  const mr = 4 + index;
  const refs = ["AO","A","AN","K","E","M","Q","T","AE","AF","AI","AP","AQ"];
  queue.getRange(`A${qr}:M${qr}`).formulas = [refs.map(c => `='Account Metrics'!${c}${mr}`)];
}
const qEnd = 3 + payload.queue_customer_ids.length;
if (payload.queue_customer_ids.length) {
  queue.getRange(`A4:M${qEnd}`).format.font = { color: green };
  queue.getRange(`D4:E${qEnd}`).format.numberFormat = "#,##0.0";
  queue.getRange(`F4:I${qEnd}`).format.numberFormat = "0.0%";
  queue.getRange(`J4:J${qEnd}`).format.numberFormat = "0.0";
  queue.getRange(`K4:K${qEnd}`).format.numberFormat = "$#,##0.00";
  queue.getRange(`A3:M${qEnd}`).format.borders = border;
  queue.getRange(`A4:M${qEnd}`).conditionalFormats.add("expression", { formula: `=$C4="OVER ENTITLEMENT"`, format: { fill: "#EEE7FF", font: { color: "#56309A" } } });
  queue.getRange(`A4:M${qEnd}`).conditionalFormats.add("expression", { formula: `=$C4="EARLY EXHAUSTION RISK"`, format: { fill: "#F6F1FF", font: { color: "#725A9C" } } });
}
queue.getRange(`A1:M${Math.max(qEnd,4)}`).format.autofitColumns();
queue.getRange("A:A").format.columnWidth = 18;
queue.getRange("B:B").format.columnWidth = 16;
queue.getRange("C:C").format.columnWidth = 26;
queue.getRange("L:L").format.columnWidth = 30;
queue.getRange("M:M").format.columnWidth = 68;
queue.getRange(`A4:M${Math.max(qEnd,4)}`).format.wrapText = true;
queue.showGridLines = false;
queue.freezePanes.freezeRows(3);

wb.recalculate();
if (process.env.VERIFY_XLSX === "1") {
  const formulaCheck = await wb.inspect({ kind: "formula", sheetId: "Account Metrics", range: `A3:${metricLast}${Math.min(mEnd,7)}`, maxChars: 8000, options: { maxResults: 200 } });
  const queueCheck = await wb.inspect({ kind: "region", sheetId: "Action Queue", range: `A3:M${Math.max(qEnd,4)}`, include: "values,formulas", maxChars: 8000, tableMaxRows: 15, tableMaxCols: 13 });
  const errorCheck = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, maxChars: 6000 });
  console.log(formulaCheck.ndjson || formulaCheck);
  console.log(queueCheck.ndjson || queueCheck);
  console.log(errorCheck.ndjson || errorCheck);
  const preview = await wb.render({ sheetName: "Assumptions", range: "A1:D15", scale: 1, format: "png" });
  await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
