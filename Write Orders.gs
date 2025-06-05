function writeLineItemsShopifyStyle() {
  const orders = getYesterdayOrders();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sales_Log") ||
                SpreadsheetApp.getActiveSpreadsheet().insertSheet("Sales_Log");

  const headers = [
    "Order Name",
    "Order Date",
    "Financial Status",
    "Shipping Total",
    "Subtotal",
    "Total Discounts",
    "Total Paid",
    "SKU",
    "Product Title",
    "Quantity",
    "Unit Price",
    "Discounted Total",
    "Unit Cost",
    "Gross Profit"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  const rows = [];

  function formatDateMMDDYYYY(dateString) {
    const date = new Date(dateString);
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    const yyyy = date.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
  }

  orders.forEach(order => {
    const orderName = order.name;
    const createdAt = formatDateMMDDYYYY(order.createdAt);
    const financialStatus = order.displayFinancialStatus;
    const shippingTotal = parseFloat(order.totalShippingPriceSet?.shopMoney?.amount || 0);
    const subtotal = parseFloat(order.subtotalPriceSet?.shopMoney?.amount || 0);
    const totalDiscounts = parseFloat(order.totalDiscountsSet?.shopMoney?.amount || 0);
    const totalPaid = parseFloat(order.totalPriceSet?.shopMoney?.amount || 0);

    let isFirstLineItem = true;

    order.lineItems.edges.forEach(itemEdge => {
      const item = itemEdge.node;

      const sku = item.sku || "";
      const title = item.title;
      const quantity = item.quantity;
      const unitPrice = parseFloat(item.originalUnitPriceSet?.shopMoney?.amount || 0);
      const discountedTotal = parseFloat(item.discountedTotalSet?.shopMoney?.amount || 0);
      const unitCost = parseFloat(item.unitCostAmount || 0);
      const grossProfit = ((unitPrice - unitCost) * quantity).toFixed(2);

      rows.push([
        orderName,            // repeat every row
        createdAt,            // repeat every row
        isFirstLineItem ? financialStatus : "",
        isFirstLineItem ? shippingTotal : "",
        isFirstLineItem ? subtotal : "",
        isFirstLineItem ? totalDiscounts : "",
        isFirstLineItem ? totalPaid : "",
        sku,
        title,
        quantity,
        unitPrice,
        discountedTotal,
        unitCost || "",
        grossProfit
      ]);

      isFirstLineItem = false;
    });
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    Logger.log(`Wrote ${rows.length} line items in Shopify export format`);
  }
}
