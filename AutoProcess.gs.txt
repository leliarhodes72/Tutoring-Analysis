const TRACKING_HEADER = "__Processed__"; 

function setupTrackingColumn(sheet, headers) {
  if (!headers.includes(TRACKING_HEADER)) {
    sheet.insertColumnAfter(headers.length);
    sheet.getRange(1, headers.length + 1).setValue(TRACKING_HEADER);
  }
}

function getTrackingColumnIndex(headers) {
  return headers.indexOf(TRACKING_HEADER);
}

function processAllReflectionsScheduled() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const scriptProperties = PropertiesService.getScriptProperties();
  const selectedColumns = JSON.parse(scriptProperties.getProperty('selectedColumns'));

  if (!selectedColumns || selectedColumns.length === 0) return;

  setupTrackingColumn(sheet, headers);
  const trackingColIndex = getTrackingColumnIndex(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);

  let colShift = 0;
  for (let idx = 0; idx < selectedColumns.length; idx++) {
    const colName = selectedColumns[idx];
    const headersNow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const originalColIndex = headersNow.indexOf(colName);

    if (originalColIndex === -1) continue;

    const reflectionColIndex = originalColIndex + colShift;
    const scoreColIndex = reflectionColIndex + 1;

    const scoreHeader = "Sentiment Score for " + colName;
    if (headersNow[scoreColIndex] !== scoreHeader) {
      sheet.insertColumnAfter(reflectionColIndex + 1);
      sheet.getRange(1, scoreColIndex + 1).setValue(scoreHeader);
    }

    for (let row = 1; row < sheet.getLastRow(); row++) {
      const alreadyTagged = sheet.getRange(row + 1, trackingColIndex + 1).getValue();
      if (alreadyTagged) continue;

      const reflection = sheet.getRange(row + 1, reflectionColIndex + 1).getValue();
      if (!reflection) continue;

      const sentiment = analyzeSentiment(reflection);
      if (!sentiment) continue;

      let label = labelSentiment(sentiment.score);
      if (detectProgression(sentiment.sentences)) label += " ↗️ Positive Progress";
      if (detectConcernWords(reflection)) label += " 🚨 Concern Flag";
      if (isReflectionTooShort(reflection)) label += " ⚠️ Reflection Too Short";

      const keywords = extractEmotionalKeywords(sentiment.sentences);
      const updatedReflection = highlightKeywordsInTextSimple(reflection, keywords);

      sheet.getRange(row + 1, reflectionColIndex + 1).setValue(updatedReflection);
      sheet.getRange(row + 1, scoreColIndex + 1).setValue(label);
      sheet.getRange(row + 1, trackingColIndex + 1).setValue("✅"); // invisible tag
    }
    colShift++;
  }
}

function onInstall(e) {
  onOpen(e);
}
