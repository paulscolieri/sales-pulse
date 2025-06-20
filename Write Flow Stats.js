function writeFlowEmailStatsToSheet() {
  writeKlaviyoFlowEmailDataToSheet(); // Refresh flow metadata

  const sheetName = "Flow Email Stats";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Date", "Flow ID", "Flow Name", "Flow Message ID", "Message Name", "Subject Line",
      "Recipients", "Opens", "Opens Unique", "Open Rate", "Clicks", "Clicks Unique",
      "Click Rate", "Click-to-Open Rate", "Conversions", "Conversion Uniques", "Conversion Rate",
      "Conversion Value", "Revenue per Recipient", "Unsubscribes", "Unsubscribe Rate",
      "Spam Complaints", "Spam Complaint Rate"
    ]);
  }

  const flowMessageMetaMap = buildFlowMessageReferenceMap();
  const results = getYesterdayFlowEmailStats();
  if (!results || results.length === 0) return;

  const today = new Date();
  const dateStr = today.toISOString().split("T")[0];

  // 🧠 Build key map: "YYYY-MM-DD|message_id" → row number
  const data = sheet.getDataRange().getValues();
  const rowIndexMap = {};
  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][0];
    const messageId = data[i][3];
    let normalizedDate;

    if (rawDate instanceof Date) {
      normalizedDate = rawDate.toISOString().split("T")[0];
    } else if (typeof rawDate === "string" && rawDate.includes("/")) {
      const [mm, dd, yyyy] = rawDate.split("/").map(s => parseInt(s, 10));
      const d = new Date(yyyy, mm - 1, dd);
      normalizedDate = d.toISOString().split("T")[0];
    } else {
      normalizedDate = String(rawDate);
    }

    const key = `${normalizedDate}|${messageId}`;
    rowIndexMap[key] = i + 1; // 1-based row index
  }

  // 🧠 Summary tracking
  let totalRevenue = 0;
  let totalRecipients = 0;
  const revenueByFlow = {};

  results.forEach(({ flow_id, flow_message_id, stats }) => {
    const meta = flowMessageMetaMap[flow_message_id] || {};
    const flowName = meta.flowName || "Unknown Flow";
    const revenue = parseFloat(stats.conversion_value || 0);
    const recipients = parseInt(stats.recipients || 0);
    const key = `${dateStr}|${flow_message_id}`;

    const newRow = [
      dateStr, flow_id, flowName, flow_message_id,
      meta.messageName || "", meta.subjectLine || "",
      recipients,
      stats.opens || 0,
      stats.opens_unique || 0,
      stats.open_rate || 0,
      stats.clicks || 0,
      stats.clicks_unique || 0,
      stats.click_rate || 0,
      stats.click_to_open_rate || 0,
      stats.conversions || 0,
      stats.conversion_uniques || 0,
      stats.conversion_rate || 0,
      revenue,
      stats.revenue_per_recipient || 0,
      stats.unsubscribes || 0,
      stats.unsubscribe_rate || 0,
      stats.spam_complaints || 0,
      stats.spam_complaint_rate || 0
    ];

    if (rowIndexMap[key]) {
      const existingRow = data[rowIndexMap[key] - 1];
      const existingRevenue = parseFloat(existingRow[17]) || 0;

      if (existingRevenue === 0) {
        Logger.log(`🔁 Overwriting row ${rowIndexMap[key]} for ${key} (revenue was 0)`);
        sheet.getRange(rowIndexMap[key], 1, 1, newRow.length).setValues([newRow]);
      } else {
        Logger.log(`⏭ Skipping ${key} (existing revenue: ${existingRevenue})`);
        return;
      }
    } else {
      sheet.appendRow(newRow);
      const newRowNum = sheet.getLastRow();
      rowIndexMap[key] = newRowNum; // 💡 update map
    }

    // Track summary
    totalRevenue += revenue;
    totalRecipients += recipients;
    revenueByFlow[flowName] = (revenueByFlow[flowName] || 0) + revenue;
  });

  const topFlow = Object.entries(revenueByFlow).sort((a, b) => b[1] - a[1])[0] || ["None", 0];
  const summary = {
    source: "klaviyo_flows",
    date: dateStr,
    totalRevenue: parseFloat(totalRevenue.toFixed(2)),
    totalRecipients,
    topFlow: topFlow[0],
    topFlowRevenue: parseFloat(topFlow[1].toFixed(2))
  };

  Logger.log("✅ Flow Email Stats updated with duplicate protection.");
  Logger.log(JSON.stringify(summary, null, 2));
  const top7 = getTopConvertingFlowLast7Days();
  Logger.log("🔥 Top converting flow over last 7 days:");
  Logger.log(JSON.stringify(top7, null, 2));

  if (top7) {
  summary.topFlow7d = top7;
  }

  return summary;
}
