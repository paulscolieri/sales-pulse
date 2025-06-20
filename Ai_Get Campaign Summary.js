function getCampaignSummaryFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Log");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(today.getDate() - 7);

  const yStr = Utilities.formatDate(yesterday, tz, "yyyy-MM-dd");
  const d7Str = Utilities.formatDate(sevenDaysAgo, tz, "yyyy-MM-dd");

  const dateCol = 0;
  const campaignCol = 1;
  const recipientsCol = 6;
  const opensCol = 7;
  const clicksCol = 9;
  const conversionsCol = 13;
  const revenueCol = 14;

  let dailyRevenue = 0;
  let dailyRecipients = 0;
  let dailyOpens = 0;
  let dailyClicks = 0;
  let dailyConversions = 0;

  let revenue7d = 0;
  let recipients7d = 0;
  let opens7d = 0;
  let clicks7d = 0;
  let conversions7d = 0;

  const dailyRevenueByCampaign = {};
  const revenueByCampaign7d = {};

  rows.forEach(row => {
    const rowDateStr = Utilities.formatDate(new Date(row[dateCol]), tz, "yyyy-MM-dd");
    const campaign = row[campaignCol];
    const recipients = Number(row[recipientsCol]) || 0;
    const opens = Number(row[opensCol]) || 0;
    const clicks = Number(row[clicksCol]) || 0;
    const conversions = Number(row[conversionsCol]) || 0;
    const revenue = Number(row[revenueCol]) || 0;

    if (rowDateStr >= d7Str) {
      revenue7d += revenue;
      recipients7d += recipients;
      opens7d += opens;
      clicks7d += clicks;
      conversions7d += conversions;
      revenueByCampaign7d[campaign] = (revenueByCampaign7d[campaign] || 0) + revenue;
    }

    if (rowDateStr === yStr) {
      dailyRevenue += revenue;
      dailyRecipients += recipients;
      dailyOpens += opens;
      dailyClicks += clicks;
      dailyConversions += conversions;
      dailyRevenueByCampaign[campaign] = (dailyRevenueByCampaign[campaign] || 0) + revenue;
    }
  });

  const sortedDaily = Object.entries(dailyRevenueByCampaign).sort((a, b) => b[1] - a[1]);
  const topCampaignDaily = dailyRevenue > 0 ? sortedDaily[0] : ["None", 0];

  const sorted7d = Object.entries(revenueByCampaign7d).sort((a, b) => b[1] - a[1]);
  const topCampaign7d = revenue7d > 0 ? sorted7d[0] : ["None", 0];

  const summary = {
    source: "klaviyo_campaigns",
    date: yStr,

    dailyRevenue: parseFloat(dailyRevenue.toFixed(2)),
    dailyRecipients,
    dailyOpens,
    dailyClicks,
    dailyConversions,
    topCampaignDaily: topCampaignDaily[0],
    topCampaignRevenueDaily: parseFloat(topCampaignDaily[1].toFixed(2)),

    revenue7d: parseFloat(revenue7d.toFixed(2)),
    recipients7d,
    opens7d,
    clicks7d,
    conversions7d,
    topCampaign7d: topCampaign7d[0],
    topCampaignRevenue7d: parseFloat(topCampaign7d[1].toFixed(2))
  };

  Logger.log("📣 Klaviyo Campaign Summary:");
  Logger.log(JSON.stringify(summary, null, 2));

  return summary;
}
