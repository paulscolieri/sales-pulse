function getYesterdayOrders() {
  const SHOPIFY_ADMIN_TOKEN = PropertiesService.getScriptProperties().getProperty('SHOPIFY_ADMIN_TOKEN');
  const SHOPIFY_DOMAIN = PropertiesService.getScriptProperties().getProperty('SHOPIFY_DOMAIN');
  const url = `https://${SHOPIFY_DOMAIN}/admin/api/2025-04/graphql.json`;

  if (!SHOPIFY_ADMIN_TOKEN || !SHOPIFY_DOMAIN) {
    throw new Error("Missing Shopify credentials. Run setCredentials() first.");
  }

  // 👉 1. Fetch store IANA timezone (e.g. "America/Los_Angeles")
  const tzQuery = JSON.stringify({
    query: `{ shop { ianaTimezone } }`
  });

  const tzResponse = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'X-Shopify-Access-Token': SHOPIFY_ADMIN_TOKEN
    },
    payload: tzQuery
  });

  const storeTimezone = JSON.parse(tzResponse.getContentText()).data.shop.ianaTimezone;

  // 👉 2. Calculate yesterday's range in store's local time
  const now = new Date();
  const yesterday = new Date(now);
  yesterday.setDate(yesterday.getDate() - 1); // or -13 if you're testing historical

  const start = Utilities.formatDate(new Date(yesterday.setHours(0, 0, 0, 0)), storeTimezone, `yyyy-MM-dd'T'HH:mm:ssXXX`);
  const end = Utilities.formatDate(new Date(), storeTimezone, `yyyy-MM-dd'T'HH:mm:ssXXX`);
  const dateFilter = `created_at:>=${start} AND created_at:<=${end}`;

  Logger.log(`Using date filter: ${dateFilter}`);

  let hasNextPage = true;
  let endCursor = null;
  let allOrders = [];

  while (hasNextPage) {
    const query = `
      {
        orders(first: 100${endCursor ? `, after: "${endCursor}"` : ''}, query: "${dateFilter}") {
          edges {
            cursor
            node {
              id
              name
              createdAt
              displayFinancialStatus
              subtotalPriceSet {
                shopMoney { amount currencyCode }
              }
              totalPriceSet {
                shopMoney { amount currencyCode }
              }
              totalShippingPriceSet {
                shopMoney { amount currencyCode }
              }
              totalDiscountsSet {
                shopMoney { amount currencyCode }
              }
              lineItems(first: 100) {
                edges {
                  node {
                    title
                    sku
                    quantity
                    originalUnitPriceSet {
                      shopMoney { amount currencyCode }
                    }
                    discountedTotalSet {
                      shopMoney { amount currencyCode }
                    }
                    variant {
                      inventoryItem {
                        unitCost {
                          amount
                          currencyCode
                        }
                      }
                    }
                  }
                }
              }
            }
          }
          pageInfo {
            hasNextPage
          }
        }
      }
    `;

    const options = {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'X-Shopify-Access-Token': SHOPIFY_ADMIN_TOKEN
      },
      payload: JSON.stringify({ query })
    };

    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (!json.data || !json.data.orders) {
      throw new Error(`Shopify API error:\n${JSON.stringify(json, null, 2)}`);
    }

    const edges = json.data.orders.edges;

    edges.forEach(edge => {
      const order = edge.node;

      order.lineItems.edges.forEach(itemEdge => {
        const item = itemEdge.node;
        const cost = item.variant?.inventoryItem?.unitCost?.amount ?? null;
        item.unitCostAmount = cost;
      });

      allOrders.push(order);

      // Log timestamps for debugging
      const createdAtUTC = new Date(order.createdAt);
      const createdAtStoreTime = Utilities.formatDate(createdAtUTC, storeTimezone, "yyyy-MM-dd HH:mm:ss");
      order.createdAtStoreTime = createdAtStoreTime; // <- Add this line

      Logger.log(`Order: ${order.name}`);
      Logger.log(`  createdAt (UTC):        ${createdAtUTC.toISOString()}`);
      Logger.log(`  createdAt (Store Time): ${createdAtStoreTime}`);
    });

    hasNextPage = json.data.orders.pageInfo.hasNextPage;
    if (hasNextPage) {
      endCursor = edges[edges.length - 1].cursor;
    }
  }

  Logger.log(`Fetched ${allOrders.length} orders from ${start} to ${end}`);
  return allOrders;
}
