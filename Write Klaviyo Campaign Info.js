function writeCampaignDailyStatsToLog(daysBack = 16) {
  const results = getCampaignStatsWindow(daysBack);
  if (!results || results.length === 0) {
    Logger.log("No campaign stats found.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refSheet = ss.getSheetByName("Campaign Reference");
  const logSheetName = "Email Log";
  const logSheet = ss.getSheetByName(logSheetName) || ss.insertSheet(logSheetName);

  // Add headers if sheet is empty
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow([
      "Date", "Campaign Name", "Campaign ID", "Subject Line", "Preview Text", "Send Time",
      "Opens", "Unique Opens", "Open Rate",
      "Clicks", "Unique Clicks", "Click Rate", "Click-to-Open Rate",
      "Conversions", "Revenue", "Revenue per Recipient",
      "Unsubscribes", "Spam Complaints"
    ]);
  }

  // Load campaign reference data
  const refData = refSheet.getDataRange().getValues();
  const refMap = {};
  for (let i = 1; i < refData.length; i++) {
    const row = refData[i];
    refMap[row[1]] = {
      name: row[0],
      subject_line: row[2],
      preview_text: row[3],
      send_time: row[4]
    };
  }

  // Load existing rows for deduplication (date + campaign_id)
  const existingRows = {};
  const lastRow = logSheet.getLastRow();

  if (lastRow > 1) {
    const existingData = logSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < existingData.length; i++) {
      const [rowDate, , campaignId] = existingData[i];
      const key = `${rowDate}|${campaignId}`;
      existingRows[key] = i + 2; // Actual sheet row
    }
  }


  let writes = 0;
  let updates = 0;

  const today = new Date();
  const tz = Session.getScriptTimeZone();

  results.forEach(r => {
    const meta = refMap[r.campaign_id] || {};
    const stats = r.stats || {};

    // Use the Klaviyo-reported send_time (if available) or fallback to yesterday
    const sendDate = meta.send_time
      ? new Date(meta.send_time)
      : new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);

    const dateStr = Utilities.formatDate(sendDate, tz, "yyyy-MM-dd");
    const key = `${dateStr}|${r.campaign_id}`;

    const newRow = [
      dateStr,
      meta.name || "",
      r.campaign_id || "",
      meta.subject_line || "",
      meta.preview_text || "",
      meta.send_time || "",
      stats.opens || 0,
      stats.opens_unique || 0,
      stats.open_rate || 0,
      stats.clicks || 0,
      stats.clicks_unique || 0,
      stats.click_rate || 0,
      stats.click_to_open_rate || 0,
      stats.conversions || 0,
      stats.conversion_value || 0,
      stats.revenue_per_recipient || 0,
      stats.unsubscribes || 0,
      stats.spam_complaints || 0
    ];

    if (existingRows[key]) {
      logSheet.getRange(existingRows[key], 1, 1, newRow.length).setValues([newRow]);
      updates++;
    } else {
      logSheet.appendRow(newRow);
      writes++;
    }
  });

  Logger.log(`Wrote ${writes} new rows to Email Log.`);
  Logger.log(`Updated ${updates} existing rows in Email Log.`);
}
