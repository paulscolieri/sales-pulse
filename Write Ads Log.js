function writeMetaAdInsightsToSheet() {
  const results = getMetaAdInsightsWindow(7); // pull last 7 days
  if (!results.length) return;

  const sheetName = "Ads Log";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`📄 Created new sheet: ${sheetName}`);
  }

  // Add headers if empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Date", "Campaign", "Ad Set", "Ad", "Ad ID",
      "Impressions", "Clicks", "Spend", "CPM", "CPC", "CTR",
      "Adds to Cart", "Checkouts Initiated", "Purchases",
      "Revenue", "ROAS", "Cost Per Conversion"
    ]);
  }

  // Helper to normalize date format
  function normalizeDate(date) {
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  // Map existing "ad_id|date" to row number
  const existingMap = new Map();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const existingData = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // A2:E
    existingData.forEach((row, i) => {
      const rawDate = row[0];
      const adId = row[4];
      if (rawDate && adId) {
        const key = `${adId}|${normalizeDate(rawDate)}`;
        existingMap.set(key, i + 2); // actual row number
      }
    });
  }

  results.forEach(r => {
    const rowDate = r.date;
    const uniqueKey = `${r.ad_id}|${rowDate}`;
    const rowData = [
      rowDate, r.campaign_name, r.adset_name, r.ad_name, r.ad_id,
      r.impressions, r.clicks, r.spend, r.cpm, r.cpc, r.ctr,
      r.add_to_cart, r.checkout, r.purchases,
      r.revenue, r.roas, r.cost_per_conversion
    ];

    if (existingMap.has(uniqueKey)) {
      const rowNum = existingMap.get(uniqueKey);
      sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
      Logger.log(`🔁 Updated row for ${r.ad_id} on ${rowDate}`);
    } else {
      sheet.appendRow(rowData);
      Logger.log(`➕ Appended new row for ${r.ad_id} on ${rowDate}`);
    }
  });

  Logger.log("✅ Ads Log synced.");
}
