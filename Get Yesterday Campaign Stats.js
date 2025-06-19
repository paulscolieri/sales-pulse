function getCampaignStatsWindow(daysBack = 7) {
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty('KLAVIYO_API_KEY');
  const revision = '2025-04-15';
  const conversionMetricId = getConversionMetricId();
  const url = 'https://a.klaviyo.com/api/campaign-values-reports';

  const startDate = new Date();
  startDate.setDate(startDate.getDate() - daysBack);
  startDate.setHours(0, 0, 0, 0);

  const endDate = new Date(); // today
  endDate.setHours(23, 59, 59, 999);


  const payload = {
    data: {
      type: "campaign-values-report",
      attributes: {
        timeframe: {
          start: startDate.toISOString(),
          end: endDate.toISOString()
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
    method: 'post',
    contentType: 'application/vnd.api+json',
    headers: {
      'Authorization': `Klaviyo-API-Key ${apiKey}`,
      'Accept': 'application/vnd.api+json',
      'revision': revision
    },
    payload: JSON.stringify(payload)
  });

  const data = JSON.parse(response.getContentText());
  const results = data?.data?.attributes?.results || [];

  if (results.length === 0) {
    Logger.log(`No campaign stats found from ${startDate.toISOString()}.`);
    return [];
  }

  return results.map(item => ({
    campaign_id: item.groupings.campaign_id,
    campaign_message_id: item.groupings.campaign_message_id,
    stats: item.statistics
  }));
}
