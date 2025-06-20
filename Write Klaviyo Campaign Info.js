function writeCampaignDailyStatsToLog(daysBack = 7) {
  const results = getCampaignStatsWindow(daysBack);
  if (!results || results.length === 0) {
    Logger.log("No campaign stats found.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refSheet = ss.getSheetByName("Campaign Reference");
  const logSheetName = "Email Log";
  const logSheet = ss.getSheetByName(logSheetName) || ss.insertSheet(logSheetName);

  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow([
      "Date", "Campaign Name", "Campaign ID", "Subject Line", "Preview Text", "Send Time",
      "Opens", "Unique Opens", "Open Rate",
      "Clicks", "Unique Clicks", "Click Rate", "Click-to-Open Rate",
      "Conversions", "Revenue", "Revenue per Recipient",
      "Unsubscribes", "Spam Complaints"
    ]);
  }

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

  const existingRows = {};
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1) {
    const existingData = logSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < existingData.length; i++) {
      const [rowDate, , campaignId] = existingData[i];
      const key = `${rowDate}|${campaignId}`;
      existingRows[key] = i + 2;
    }
  }

  let writes = 0;
  let updates = 0;

  // 🧠 Summary tracking
  let totalRevenue = 0;
  let totalRecipients = 0;
  let totalOpens = 0;
  let totalClicks = 0;
  let totalConversions = 0;
  const revenueByCampaign = {};

  const today = new Date();
  const tz = Session.getScriptTimeZone();
  let mostRecentDate = null;

  results.forEach(r => {
    const meta = refMap[r.campaign_id] || {};
    const stats = r.stats || {};

    const sendDate = meta.send_time
      ? new Date(meta.send_time)
      : new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
    const dateStr = Utilities.formatDate(sendDate, tz, "yyyy-MM-dd");
    if (!mostRecentDate || sendDate > mostRecentDate) mostRecentDate = sendDate;

    const key = `${dateStr}|${r.campaign_id}`;
    const revenue = parseFloat(stats.conversion_value || 0);
    const recipients = parseInt(stats.opens_unique || 0);

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
      revenue,
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

    // 🧠 Aggregate summary stats
    totalRevenue += revenue;
    totalRecipients += recipients;
    totalOpens += parseInt(stats.opens || 0);
    totalClicks += parseInt(stats.clicks || 0);
    totalConversions += parseInt(stats.conversions || 0);

    const campaignName = meta.name || "Unnamed Campaign";
    revenueByCampaign[campaignName] = (revenueByCampaign[campaignName] || 0) + revenue;
  });

  // 🥇 Top campaign today
  const topCampaign = Object.entries(revenueByCampaign)
    .sort((a, b) => b[1] - a[1])[0] || ["None", 0];

  const summary = {
    source: "klaviyo_campaigns",
    date: mostRecentDate ? Utilities.formatDate(mostRecentDate, tz, "yyyy-MM-dd") : null,
    totalRevenue: parseFloat(totalRevenue.toFixed(2)),
    totalRecipients,
    totalOpens,
    totalClicks,
    totalConversions,
    topCampaign: topCampaign[0],
    topCampaignRevenue: parseFloat(topCampaign[1].toFixed(2))
  };

  // 🧠 Add top campaign from past 7 days
  const top7 = getTopCampaignLast7Days(); // <- defined separately
  if (top7) {
    summary.topCampaign7d = top7;
  }

  Logger.log(`📬 Wrote ${writes} new rows to Email Log.`);
  Logger.log(`🔁 Updated ${updates} existing rows in Email Log.`);
  Logger.log("📊 Klaviyo Campaign Summary:");
  Logger.log(JSON.stringify(summary, null, 2));

  return summary;
}
