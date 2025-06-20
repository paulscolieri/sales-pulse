function getMetaAdSummaryFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ads Log");
  const data = sheet.getDataRange().getValues();
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
  const adNameCol = 3;
  const adIdCol = 4;
  const spendCol = 7;
  const purchasesCol = 13;
  const revenueCol = 14;

  let dailySpend = 0;
  let dailyRevenue = 0;
  let dailyPurchases = 0;

  let spend7d = 0;
  let revenue7d = 0;
  let purchases7d = 0;

  const dailyAdStats = {};
  const adStats7d = {};

  rows.forEach(row => {
    const rowDateStr = Utilities.formatDate(new Date(row[dateCol]), tz, "yyyy-MM-dd");
    const adName = row[adNameCol];
    const adId = row[adIdCol];
    const spend = parseFloat(row[spendCol]) || 0;
    const revenue = parseFloat(row[revenueCol]) || 0;
    const purchases = parseInt(row[purchasesCol]) || 0;

    // 7-day totals
    if (rowDateStr >= d7Str) {
      spend7d += spend;
      revenue7d += revenue;
      purchases7d += purchases;

      if (!adStats7d[adId]) {
        adStats7d[adId] = { adName, adId, revenue: 0, spend: 0, purchases: 0 };
      }
      adStats7d[adId].revenue += revenue;
      adStats7d[adId].spend += spend;
      adStats7d[adId].purchases += purchases;
    }

    // Daily totals
    if (rowDateStr === yStr) {
      dailySpend += spend;
      dailyRevenue += revenue;
      dailyPurchases += purchases;

      if (!dailyAdStats[adId]) {
        dailyAdStats[adId] = { adName, adId, revenue: 0, spend: 0, purchases: 0 };
      }
      dailyAdStats[adId].revenue += revenue;
      dailyAdStats[adId].spend += spend;
      dailyAdStats[adId].purchases += purchases;
    }
  });

  // Add ROAS to ad objects
  const enrichAdStats = stats => {
    return Object.values(stats).map(ad => ({
      ...ad,
      roas: ad.spend > 0 ? parseFloat((ad.revenue / ad.spend).toFixed(2)) : 0
    }));
  };

  const enrichedDailyAds = enrichAdStats(dailyAdStats);
  const enriched7dAds = enrichAdStats(adStats7d);

  const topAdDaily = enrichedDailyAds.sort((a, b) => b.revenue - a.revenue)[0] || { adName: "None", revenue: 0, spend: 0, roas: 0 };
  const topAd7d = enriched7dAds.sort((a, b) => b.revenue - a.revenue)[0] || { adName: "None", revenue: 0, spend: 0, roas: 0 };

  const avgROAS = spend7d > 0 ? parseFloat((revenue7d / spend7d).toFixed(2)) : 0;

  const summary = {
    source: "meta_ads",
    date: yStr,

    dailySpend: parseFloat(dailySpend.toFixed(2)),
    dailyRevenue: parseFloat(dailyRevenue.toFixed(2)),
    dailyPurchases,
    topAdDaily,

    spend7d: parseFloat(spend7d.toFixed(2)),
    revenue7d: parseFloat(revenue7d.toFixed(2)),
    purchases7d,
    avgROAS,
    topAd7d
  };

  Logger.log("📊 Meta Ad Summary:");
  Logger.log(JSON.stringify(summary, null, 2));

  return summary;
}
