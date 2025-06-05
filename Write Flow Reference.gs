function writeKlaviyoFlowEmailDataToSheet() {
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty('KLAVIYO_API_KEY');
  const revision = '2025-04-15';

  const sheetName = 'Flow Reference';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear(); // clear old data
  }

  sheet.appendRow([
    'Flow Name',
    'Flow ID',
    'Action Index',
    'Action Type',
    'Action ID',
    'Message ID',
    'Message Name',
    'Subject Line'
  ]);

  let nextPage = 'https://a.klaviyo.com/api/flows?page[size]=50';

  while (nextPage) {
    const flowRes = UrlFetchApp.fetch(nextPage, {
      method: 'get',
      headers: {
        'Authorization': `Klaviyo-API-Key ${apiKey}`,
        'Accept': 'application/json',
        'revision': revision
      }
    });

    const flowData = JSON.parse(flowRes.getContentText());

    (flowData.data || []).forEach(flow => {
      const flowId = flow.id;
      const flowName = flow.attributes.name;
      const defUrl = `https://a.klaviyo.com/api/flows/${flowId}/?additional-fields[flow]=definition`;

      try {
        const defRes = UrlFetchApp.fetch(defUrl, {
          method: 'get',
          headers: {
            'Authorization': `Klaviyo-API-Key ${apiKey}`,
            'Accept': 'application/json',
            'revision': revision
          }
        });

        const defData = JSON.parse(defRes.getContentText());
        const actions = defData.data.attributes.definition.actions;

        (actions || []).forEach((action, index) => {
          let actionType = action.type || '';
          let actionId = action.id || '';
          let message = action.data?.message || null;

          // If it's an ab-test, override with main_action
          if (actionType === 'ab-test' && action.data?.main_action?.data?.message?.id) {
            const main = action.data.main_action;
            message = main.data.message;
            actionType = main.type;
            actionId = main.id;
          }

          if (message?.id) {
            const messageId = message.id;
            const messageName = message.name || '';
            const subjectLine = message.subject_line || '';

            sheet.appendRow([
              flowName,
              flowId,
              index,
              actionType,
              actionId,
              messageId,
              messageName,
              subjectLine
            ]);
          }
        });

      } catch (e) {
        Logger.log(`Error fetching flow definition for ${flowName}: ${e.message}`);
      }
    });

    nextPage = flowData.links?.next || null;
  }

  Logger.log('Flow reference data written to sheet.');
}
