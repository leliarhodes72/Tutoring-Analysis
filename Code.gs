function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem("Select Columns", "openColumnSelectionSidebar")
    .addItem("Run Analysis Now", "runAnalysisFromMenu") 
    .addToUi();
}

function openColumnSelectionSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('ColumnSelection')
      .setTitle('Select Columns for Analysis');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getSheetColumns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers;
}

function saveColumnSelection(selectedColumns) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('selectedColumns', JSON.stringify(selectedColumns));
  SpreadsheetApp.getUi().alert("Columns saved! Sentiment analysis will be run on: " + selectedColumns.join(", "));
}

function getSelectedColumnData() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var selectedColumns = JSON.parse(scriptProperties.getProperty('selectedColumns'));

  if (!selectedColumns || selectedColumns.length === 0) {
    SpreadsheetApp.getUi().alert("No columns selected. Please select columns first.");
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getDataRange().getValues();

  var columnIndices = selectedColumns.map(col => headers.indexOf(col));

  if (columnIndices.includes(-1)) {
    SpreadsheetApp.getUi().alert("Error: One or more selected columns were not found in the sheet.");
    return;
  }

  var extractedData = [];
  for (var i = 1; i < data.length; i++) {
    var rowData = {};
    var isEmpty = true;

    selectedColumns.forEach((col, index) => {
      var cellValue = data[i][columnIndices[index]];
      if (cellValue) isEmpty = false;
      rowData[col] = cellValue;
    });

    if (!isEmpty) {
      extractedData.push(rowData);
    }
  }

  Logger.log("Extracted Data: " + JSON.stringify(extractedData.slice(0, 3), null, 2));
  return extractedData;
}

function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GOOGLE_NLP_API_KEY');
}

function analyzeSentiment(reflectionText) {
  if (!reflectionText) return null;

  var url = "https://language.googleapis.com/v1/documents:analyzeSentiment?key=" + getApiKey();
  var document = {
    document: {
      type: "PLAIN_TEXT",
      content: reflectionText
    },
    encodingType: "UTF8"
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(document)
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (json.documentSentiment) {
    return {
      score: json.documentSentiment.score,
      magnitude: json.documentSentiment.magnitude,
      sentences: json.sentences || []
    };
  }
  return null;
}

function labelSentiment(score) {
  if (score >= 0.6) return "✅ Positive";
  if (score >= 0.25) return "➡️ Slightly Positive";
  if (score > -0.25 && score < 0.25) return "❕ Neutral";
  if (score <= -0.25 && score > -0.6) return "⬅️ Slightly Negative";
  return "❗ Negative";
}

function detectProgression(sentences) {
  if (!sentences || sentences.length < 2) return false;
  const first = sentences[0].sentiment?.score || 0;
  const last = sentences[sentences.length - 1].sentiment?.score || 0;
  return first <= -0.5 && last >= 0.5;
}

const POSITIVE_WORDS = [
  "apply", "applying", "applying strategies", "appreciate", "appreciated", "appreciates", "beneficial",
  "collaborate", "collaboration", "collaborative", "comfortable", "confidence", "confident", "cooperate",
  "cooperated", "cooperates", "cooperatively", "encourage", "encouragement", "encourages", "engaged",
  "enjoy", "enjoyed", "enjoyment", "enjoys", "flow", "flowed", "flows", "focus", "good", "great", "helpful",
  "helping", "improve", "improvement", "improving", "independant", "independence", "initiative", "like",
  "motivated", "motivation", "nice", "nicely", "on track", "organize", "organized", "pay attention",
  "paying attention", "plan", "prepare", "prepared", "productive", "progress", "reassure", "reinforcing",
  "student directed", "student led", "student-directed", "student-led", "successful", "successfully",
  "understand", "understanded", "understanding", "useful", "welcome", "well"
];

const NEGATIVE_WORDS = [
  "ai", "apprehensive", "artificial intelligence", "awful", "awkward", "bad", "behind", "challenge",
  "challenges", "chat", "chatgpt", "chatted", "chatting", "confused", "deepseek", "dependant",
  "difficulties", "disengaged", "distract", "distracted", "distracting", "forgot", "forgotten",
  "frustrated", "gave answers", "give answers", "gossip", "gossiped", "gossiping", "lack", "lacked",
  "lacking", "nervous", "nervousness", "overwhelmed", "overwhelming", "panicing", "panicked", "poor",
  "quiz", "reliance", "resist", "resistant", "resisted", "resisting", "rushing", "rusty", "stress",
  "stressed", "stressing", "struggle", "struggled ", "struggling", "stuck", "taught ", "teach",
  "teaching", "terrible", "test", "tired", "tough", "unclear", "uncomfortable", "unfocused",
  "uninspired", "uninterested ", "unprepared", "unsure", "weakness", "worried", "worst"
];

const CONCERN_WORDS = [
  "911", "abuse", "anxiety", "burnout", "crisis", "danger", "depressed", "depression", "emergency",
  "emergency room", "exhausted", "giving up", "grieving", "harm", "helpless", "hopeless", "isolation",
  "lonely", "mental health", "numb", "panic", "runaway", "self-harm", "self-injury", "self-isolate",
  "suicide", "threat", "trauma", "unsafe", "urgent", "withdrawn", "worthless"
];

const NEGATION_WORDS = ["not", "never", "didn't", "doesn't", "isn't", "wasn't", "aren't", "won't", "can't"];

const INTENSIFIERS = ["very", "really", "quite", "super", "truly", "completely", "totally"];

const EMOTION_THRESHOLD = 0.6; // Adjustable threshold for emotional sentence filtering

function isNegatedWithIntensifier(text, word) {
  const neg = NEGATION_WORDS.join('|');
  const intens = INTENSIFIERS.join('|');
  const regex = new RegExp(`\\b(?:${neg})(?:\\s+${intens})?\\s+${word}\\b`, 'i');
  return regex.test(text);
}

function extractEmotionalKeywords(sentences) {
  const keywords = [];

  sentences.forEach(sentence => {
    const score = sentence.sentiment?.score || 0;
    if (Math.abs(score) < EMOTION_THRESHOLD) return; // Configurable threshold

    const content = sentence.text.content;
    const words = content.toLowerCase().match(/\b[a-z]{4,}\b/g);

    if (words) {
      words.forEach(word => {
        if (POSITIVE_WORDS.includes(word)) {
          const emoji = isNegatedWithIntensifier(content, word) ? "❗" : "✅";
          if (!keywords.find(k => k.word === word)) {
            keywords.push({ word, emoji });
          }
        } else if (NEGATIVE_WORDS.includes(word)) {
          if (!keywords.find(k => k.word === word)) {
            keywords.push({ word, emoji: "❗" });
          }
        }
      });
    }
  });

  return keywords;
}


function processAllReflections() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const scriptProperties = PropertiesService.getScriptProperties();
  const selectedColumns = JSON.parse(scriptProperties.getProperty('selectedColumns'));

  if (!selectedColumns || selectedColumns.length === 0) {
    SpreadsheetApp.getUi().alert("No columns selected for sentiment analysis.");
    return;
  }

  let colShift = 0;

  for (let idx = 0; idx < selectedColumns.length; idx++) {
    const colName = selectedColumns[idx];
    const originalColIndex = headers.indexOf(colName);

    if (originalColIndex === -1) {
      SpreadsheetApp.getUi().alert(`Column "${colName}" not found.`);
      continue;
    }

    const reflectionColIndex = originalColIndex + colShift;
    const scoreColIndex = reflectionColIndex + 1;

    sheet.insertColumnAfter(reflectionColIndex + 1);
    sheet.getRange(1, scoreColIndex + 1).setValue("Sentiment Score for " + colName);

    for (let row = 1; row < sheet.getLastRow(); row++) {
      const reflection = sheet.getRange(row + 1, reflectionColIndex + 1).getValue();
      if (!reflection) continue;

      const sentiment = analyzeSentiment(reflection);
      if (!sentiment) continue;

      let label = labelSentiment(sentiment.score);

      if (detectProgression(sentiment.sentences)) {
        label += " ↗️ Positive Progress";
      }

      if (detectConcernWords(reflection)) {
        label += " 🚨 Concern Flag";
      }

      if (isReflectionTooShort(reflection)) {
        label += " ⚠️ Reflection Too Short";
      }

      const keywords = extractEmotionalKeywords(sentiment.sentences);
      const updatedReflection = highlightKeywordsInTextSimple(reflection, keywords);

      sheet.getRange(row + 1, reflectionColIndex + 1).setValue(updatedReflection);
      sheet.getRange(row + 1, scoreColIndex + 1).setValue(label);
    }

    colShift++;
  }

  SpreadsheetApp.getUi().alert("Sentiment analysis complete with emoji tagging, negation handling, and progression detection!");
}

function highlightKeywordsInTextSimple(text, keywordList) {
  if (!text) return text;

  let updatedText = text;

  // Add 🚨 for concern keywords first
  CONCERN_WORDS.forEach(word => {
    const regex = new RegExp(`\\b(${word})\\b`, "gi");
    updatedText = updatedText.replace(regex, "$1🚨");
  });

  // Add ✅ or ❗ for sentiment keywords
  if (keywordList && keywordList.length > 0) {
    keywordList.forEach(({ word, emoji }) => {
      const regex = new RegExp(`\\b(${word})\\b`, "gi");
      updatedText = updatedText.replace(regex, `$1${emoji}`);
    });
  }

  return updatedText;
}


function detectConcernWords(text) {
  if (!text) return false;

  const lowerText = text.toLowerCase();

  return CONCERN_WORDS.some(keyword => lowerText.includes(keyword));
}

function isReflectionTooShort(text, wordThreshold = 100) {
  if (!text) return false;

  const words = text.trim().split(/\s+/); // Split by spaces
  return words.length < wordThreshold;
}

function runAnalysisFromMenu() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const scriptProperties = PropertiesService.getScriptProperties();
  const selectedColumns = JSON.parse(scriptProperties.getProperty('selectedColumns'));

  if (!selectedColumns || selectedColumns.length === 0) {
    SpreadsheetApp.getUi().alert("No columns selected for sentiment analysis.");
    return;
  }

  let colShift = 0;

  for (let idx = 0; idx < selectedColumns.length; idx++) {
    const colName = selectedColumns[idx];
    const originalColIndex = headers.indexOf(colName);

    if (originalColIndex === -1) {
      SpreadsheetApp.getUi().alert(`Column "${colName}" not found.`);
      continue;
    }

    const reflectionColIndex = originalColIndex + colShift;
    const scoreColIndex = reflectionColIndex + 1;

    // Insert score column if it doesn't exist
    const scoreHeader = "Sentiment Score for " + colName;
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (currentHeaders[scoreColIndex] !== scoreHeader) {
      sheet.insertColumnAfter(reflectionColIndex + 1);
      sheet.getRange(1, scoreColIndex + 1).setValue(scoreHeader);
    }

    for (let row = 1; row < sheet.getLastRow(); row++) {
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
    }

    colShift++;
  }

  SpreadsheetApp.getUi().alert("Sentiment analysis complete.");
}
