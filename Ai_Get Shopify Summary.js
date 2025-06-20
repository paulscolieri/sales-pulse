function getShopifySummaryFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sales_Log");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const dateCol = headers.indexOf("Order Date");
  const grossCol = headers.indexOf("Gross Profit");
  const totalCol = headers.indexOf("Total Paid");
  const titleCol = headers.indexOf("Product Title");
  const qtyCol = headers.indexOf("Quantity");

  const today = new Date();
  const tz = Session.getScriptTimeZone();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(yesterday, tz, "yyyy-MM-dd");
  const last7Days = new Date(today);
  last7Days.setDate(today.getDate() - 7);

  let dailyRevenue = 0;
  let dailyGrossProfit = 0;
  let dailyOrders = 0;
  let dailyUnitsSold = 0;
  let revenue7d = 0;
  let grossProfit7d = 0;
  const dailyProductStats = {};
  const productStats = {};

  data.slice(1).forEach(row => {
    const orderDate = new Date(row[dateCol]);
    const paid = parseFloat(row[totalCol]) || 0;
    const profit = parseFloat(row[grossCol]) || 0;
    const qty = parseInt(row[qtyCol]) || 0;
    const title = row[titleCol] || "Unknown Product";

    if (isNaN(orderDate)) return;

    const orderDateStr = Utilities.formatDate(orderDate, tz, "yyyy-MM-dd");

    if (orderDateStr === yesterdayStr) {
      dailyRevenue += paid;
      dailyGrossProfit += profit;
      dailyOrders += 1;
      dailyUnitsSold += qty;
      dailyProductStats[title] = (dailyProductStats[title] || 0) + qty;
    }

    if (orderDate >= last7Days) {
      revenue7d += paid;
      grossProfit7d += profit;
      productStats[title] = (productStats[title] || 0) + qty;
    }
  });

  const topProductEntry = Object.entries(productStats).sort((a, b) => b[1] - a[1])[0] || ["None", 0];
  const topProductDaily = Object.entries(dailyProductStats).sort((a, b) => b[1] - a[1])[0] || ["None", 0];

  // Optional: pull goals from AI tab if it exists
  let DailySalesGoalUSD = null;
  let grossMarginGoal = null;
  const aiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AI");
  if (aiSheet) {
    const aiData = aiSheet.getDataRange().getValues();
    aiData.forEach(row => {
      if (row[0] === "Sales Goal (Daily)") {
        DailySalesGoalUSD = parseFloat(String(row[1]).replace(/[^\d.]/g, "")) || null;
      }
      if (row[0] === "Gross Margin Target") {
        grossMarginGoal = parseFloat(String(row[1]).replace(/[^\d.]/g, "")) || null;
      }
    });
  }

  const grossMarginPercent = dailyRevenue > 0 ? parseFloat((dailyGrossProfit / dailyRevenue).toFixed(3)) : null;

  const summary = {
    source: "shopify",
    date: yesterdayStr,
    dailyRevenue: parseFloat(dailyRevenue.toFixed(2)),
    dailyOrders,
    dailyGrossProfit: parseFloat(dailyGrossProfit.toFixed(2)),
    grossMarginPercent,
    topProduct: topProductDaily[0],
    topProductUnits: topProductDaily[1],
    dailyUnitsSold,
    revenue7d: parseFloat(revenue7d.toFixed(2)),
    grossProfit7d: parseFloat(grossProfit7d.toFixed(2)),
    topProduct7d: topProductEntry[0],
    topProductUnits7d: topProductEntry[1],
    DailySalesGoalUSD,
    grossMarginGoal
  };

  Logger.log("🧾 Shopify Summary:");
  Logger.log(JSON.stringify(summary, null, 2));

  return summary;
}
