function getMetaAdInsightsWindow(daysBack = 1) {
  const scriptProps = PropertiesService.getScriptProperties();
  const accessToken = scriptProps.getProperty("META_ACCESS_TOKEN");
  const adAccountId = scriptProps.getProperty("META_AD_ACCOUNT_ID");
  const apiVersion = "v18.0";
  const url = `https://graph.facebook.com/${apiVersion}/${adAccountId}/insights`;

  function shiftMetaDateToLocal(utcDateStr) {
  // Just use the raw string — Meta already returns it in 'YYYY-MM-DD' format
  return utcDateStr;
  }


  const today = new Date();
  const start = new Date(today);
  start.setDate(today.getDate() - daysBack);

  const since = Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const until = Utilities.formatDate(new Date(today.setDate(today.getDate() - 1)), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const timeRange = JSON.stringify({ since, until });

  const fields = [
    "campaign_id", "campaign_name",
    "adset_id", "adset_name",
    "ad_id", "ad_name",
    "impressions", "clicks", "spend", "cpm", "cpc", "ctr",
    "actions", "action_values"
  ].join(",");

  const params = {
    method: "get",
    muteHttpExceptions: true,
    headers: { Authorization: `Bearer ${accessToken}` }
  };

  const query = [
    `fields=${encodeURIComponent(fields)}`,
    `time_range=${encodeURIComponent(timeRange)}`,
    `level=ad`,
    `time_increment=1`
  ].join("&");

  const fullUrl = `${url}?${query}`;
  const res = UrlFetchApp.fetch(fullUrl, params);
  const json = JSON.parse(res.getContentText());

  if (!json.data) {
    Logger.log("❌ Error fetching insights: " + res.getContentText());
    return [];
  }

  const cleanResults = json.data.map(row => {
    const actions = parseActions(row.actions || []);
    const values = parseActions(row.action_values || []);
    const revenue = values.purchase || 0;
    const spend = parseFloat(row.spend || 0);
    const roas = spend > 0 ? revenue / spend : 0;
    const cost_per_conversion = (actions.purchase && actions.purchase > 0)
      ? spend / actions.purchase
      : 0;

    const rowDate = shiftMetaDateToLocal(row.date_start);


    return {
      date: rowDate,
      campaign_name: row.campaign_name || "",
      adset_name: row.adset_name || "",
      ad_name: row.ad_name || "",
      ad_id: row.ad_id || "",
      impressions: parseInt(row.impressions || 0),
      clicks: parseInt(row.clicks || 0),
      spend: spend,
      cpm: parseFloat(row.cpm || 0),
      cpc: parseFloat(row.cpc || 0),
      ctr: parseFloat(row.ctr || 0),
      add_to_cart: actions["add_to_cart"] || 0,
      checkout: actions["initiate_checkout"] || 0,
      purchases: actions.purchase || 0,
      revenue: revenue,
      roas: roas,
      cost_per_conversion
    };
  });

  Logger.log(`✅ Pulled ${cleanResults.length} ad rows from ${since} to ${until}`);
  return cleanResults;

  function parseActions(arr) {
    return arr.reduce((map, obj) => {
      if (obj.action_type && obj.value) {
        map[obj.action_type] = parseFloat(obj.value);
      }
      return map;
    }, {});
  }


}
