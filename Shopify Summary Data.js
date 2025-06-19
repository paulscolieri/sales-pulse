function getShopifyKpiTrends() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sales_Log");
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);

  const getDaysAgo = (days) => {
    const d = new Date(today);
    d.setDate(d.getDate() - days);
    return d;
  };

  // Helper
  function parseDate(str) {
    if (str instanceof Date) {
      return new Date(str.getFullYear(), str.getMonth(), str.getDate());
    }
    const [mm, dd, yyyy] = str.split("/").map(s => parseInt(s, 10));
    return new Date(yyyy, mm - 1, dd);
  }


  const summary = {
    yesterday: {
      totalRevenue: 0,
      totalOrders: new Set(),
      grossProfit: 0,
      productCounts: {}
    },
    last7Days: {
      totalRevenue: 0,
      totalOrders: new Set(),
      grossProfit: 0,
      productCounts: {}
    }
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const orderDate = parseDate(row[1]);
    const orderName = row[0];
    const totalPaid = parseFloat(row[6]) || 0;
    const grossProfit = parseFloat(row[13]) || 0;
    const productTitle = row[8];
    const quantity = parseFloat(row[9]) || 0;

    const daysAgo = Math.floor((today - orderDate) / (1000 * 60 * 60 * 24));

    if (daysAgo === 1) {
      summary.yesterday.totalRevenue += totalPaid;
      summary.yesterday.totalOrders.add(orderName);
      summary.yesterday.grossProfit += grossProfit;
      summary.yesterday.productCounts[productTitle] = (summary.yesterday.productCounts[productTitle] || 0) + quantity;
    }

    if (daysAgo <= 7) {
      summary.last7Days.totalRevenue += totalPaid;
      summary.last7Days.totalOrders.add(orderName);
      summary.last7Days.grossProfit += grossProfit;
      summary.last7Days.productCounts[productTitle] = (summary.last7Days.productCounts[productTitle] || 0) + quantity;

    }
  }

  // Reduce to final summary
  const topProduct = Object.entries(summary.yesterday.productCounts).sort((a, b) => b[1] - a[1])[0] || ["N/A", 0];
  const top7Product = Object.entries(summary.last7Days.productCounts).sort((a, b) => b[1] - a[1])[0] || ["N/A", 0];


  const result = {
    source: "shopify",
    yesterday: {
      revenue: parseFloat(summary.yesterday.totalRevenue.toFixed(2)),
      orders: summary.yesterday.totalOrders.size,
      grossProfit: parseFloat(summary.yesterday.grossProfit.toFixed(2)),
      topProduct: topProduct[0],
      topProductQty: topProduct[1]
    },
    last7Days: {
      avgRevenue: parseFloat((summary.last7Days.totalRevenue / 7).toFixed(2)),
      avgOrders: parseFloat((summary.last7Days.totalOrders.size / 7).toFixed(2)),
      avgProfit: parseFloat((summary.last7Days.grossProfit / 7).toFixed(2)),
      topProduct: top7Product[0],
      topProductQty: top7Product[1]
    }
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
