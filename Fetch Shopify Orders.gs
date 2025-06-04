function getYesterdayOrders() {
  const SHOPIFY_ADMIN_TOKEN = PropertiesService.getScriptProperties().getProperty('SHOPIFY_ADMIN_TOKEN');
  const SHOPIFY_DOMAIN = PropertiesService.getScriptProperties().getProperty('SHOPIFY_DOMAIN');
  const url = `https://${SHOPIFY_DOMAIN}/admin/api/2025-04/graphql.json`;

  if (!SHOPIFY_ADMIN_TOKEN || !SHOPIFY_DOMAIN) {
    throw new Error("Missing Shopify credentials. Run setCredentials() first.");
  }

  // 1. Get UTC date filter for yesterday
  const now = new Date();
  const yesterday = new Date(now);
  yesterday.setDate(yesterday.getDate() - 1);

  const start = new Date(yesterday.setUTCHours(0, 0, 0, 0)).toISOString();
  const end = new Date(yesterday.setUTCHours(23, 59, 59, 999)).toISOString();
  const dateFilter = `created_at:>=${start} AND created_at:<=${end}`;

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

    // Add a safety check for cost
    edges.forEach(edge => {
      const order = edge.node;

      order.lineItems.edges.forEach(itemEdge => {
        const item = itemEdge.node;

        const cost = item.variant?.inventoryItem?.unitCost?.amount ?? null;
        item.unitCostAmount = cost; // Injected for use downstream
      });

      allOrders.push(order);
    });

    hasNextPage = json.data.orders.pageInfo.hasNextPage;
    if (hasNextPage) {
      endCursor = edges[edges.length - 1].cursor;
    }
  }

  Logger.log(`Fetched ${allOrders.length} orders from ${start} to ${end}`);
  return allOrders;
}
