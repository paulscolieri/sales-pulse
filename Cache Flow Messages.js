function getYesterdayFlowEmailStats() {
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty('KLAVIYO_API_KEY');
  const revision = '2025-04-15';

  const conversionMetricId = getConversionMetricId();

  const url = `https://a.klaviyo.com/api/flow-values-reports`;

  const payload = {
    data: {
      type: "flow-values-report",
      attributes: {
        timeframe: {
          key: "yesterday"
        },
        statistics: [
          "recipients",
          "opens",
          "opens_unique",
          "open_rate",
          "clicks",
          "clicks_unique",
          "click_rate",
          "click_to_open_rate",
          "conversions",
          "conversion_uniques",
          "conversion_value",
          "conversion_rate",
          "revenue_per_recipient",
          "unsubscribes",
          "unsubscribe_rate",
          "spam_complaints",
          "spam_complaint_rate"
        ],
        conversion_metric_id: conversionMetricId,
        filter: 'equals(send_channel,"email")'
      }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/vnd.api+json",
    headers: {
      "Authorization": `Klaviyo-API-Key ${apiKey}`,
      "Accept": "application/vnd.api+json",
      "revision": revision
    },
    payload: JSON.stringify(payload)
  });

  const data = JSON.parse(response.getContentText());
  const results = data?.data?.attributes?.results || [];

  if (results.length === 0) {
    Logger.log("No flow email stats for yesterday.");
    return;
  }

  Logger.log("Flow Email Stats:");
  results.forEach(stat => {
    const group = stat.groupings;
    const stats = stat.statistics;
    Logger.log({
      flow_id: group.flow_id,
      flow_message_id: group.flow_message_id,
      stats: stats
    });
  });

  return results.map(item => ({
  flow_id: item.groupings.flow_id,
  flow_message_id: item.groupings.flow_message_id,
  stats: item.statistics
  }));

}