function getConversionMetricId() {
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty('KLAVIYO_API_KEY');
  const revision = '2025-04-15';

  const url = 'https://a.klaviyo.com/api/metrics/';

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': `Klaviyo-API-Key ${apiKey}`,
      'Accept': 'application/json',
      'revision': revision
    }
  });

  const data = JSON.parse(response.getContentText());
  const placedOrderMetric = data.data.find(metric => metric.attributes.name === 'Placed Order');

  if (!placedOrderMetric) {
    throw new Error('Placed Order metric not found.');
  }

  return placedOrderMetric.id;
}
