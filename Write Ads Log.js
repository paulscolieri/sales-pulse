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

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Date", "Campaign", "Ad Set", "Ad", "Ad ID",
      "Impressions", "Clicks", "Spend", "CPM", "CPC", "CTR",
      "Adds to Cart", "Checkouts Initiated", "Purchases",
      "Revenue", "ROAS", "Cost Per Conversion"
    ]);
  }

  function normalizeDate(date) {
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  const existingMap = new Map();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const existingData = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // A2:E
    existingData.forEach((row, i) => {
      const rawDate = row[0];
      const adId = row[4];
      if (rawDate && adId) {
        const key = `${adId}|${normalizeDate(rawDate)}`;
        existingMap.set(key, i + 2);
      }
    });
  }

  // 🧠 Summary tracking
  let totalSpend = 0;
  let totalRevenue = 0;
  let totalPurchases = 0;
  let bestROAS = 0;
  let topAd = null;
  const adStats = {};

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

    // Update summary tracking
    const spend = parseFloat(r.spend || 0);
    const revenue = parseFloat(r.revenue || 0);
    const purchases = parseInt(r.purchases || 0);
    


    totalSpend += spend;
    totalRevenue += revenue;
    totalPurchases += purchases;

    const adId = r.ad_id;
    if (!adStats[adId]) {
      adStats[adId] = {
        adName: r.ad_name || "Unnamed Ad",
        adId,
        revenue: 0,
        spend: 0,
        purchases: 0
      };
    }
    adStats[adId].revenue += revenue;
    adStats[adId].spend += spend;
    adStats[adId].purchases += purchases;

  });

  Object.values(adStats).forEach(ad => {
  const roas = ad.spend > 0 ? ad.revenue / ad.spend : 0;
  ad.roas = parseFloat(roas.toFixed(2));

  if (ad.roas > bestROAS) {
    bestROAS = ad.roas;
    topAd = ad;
  }
  });


  const avgROAS = totalSpend > 0 ? parseFloat((totalRevenue / totalSpend).toFixed(2)) : 0;

  const summary = {
    source: "meta_ads",
    date: normalizeDate(new Date()),
    totalSpend: parseFloat(totalSpend.toFixed(2)),
    totalRevenue: parseFloat(totalRevenue.toFixed(2)),
    totalPurchases,
    avgROAS,
    topAd
  };

  Logger.log("✅ Meta Ads Summary:");
  Logger.log(JSON.stringify(summary, null, 2));

  return summary;
}
