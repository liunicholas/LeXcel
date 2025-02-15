function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Chatbot")
    .addItem("Open Chat", "showChatbotSidebar")
    .addToUi();
}

function showChatbotSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("ChatbotUI")
    .setTitle("Chatbot Assistant");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  return data;
}

function processMessage(message) {
  var data = getSheetData();
  console.log(data);
  
  message = message.toLowerCase().trim();
  
  if (message.includes("hello")) {
    return "Hello! How can I help you with your sheet data?";
  } else if (message.includes("summary")) {
    return "Your sheet has " + data.length + " rows and " + data[0].length + " columns.";
  } else if (message.includes("first row")) {
    return "The first row is: " + data[0].join(", ");
  } else {
    return "I can answer questions about your spreadsheet. Try asking about 'summary' or 'first row'.";
  }
}
