function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("LeXcel")
    .addItem("Open LeXcel Assistant", "showChatbotSidebar")
    .addToUi();
}

function showChatbotSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("ChatbotUI")
    .setTitle("LeXcel Assistant");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  return data;
}

function single_cell_formula(user_prompt, rangeNotation, cellValuesString) {
  var apiKey = ""
  var url = "https://api.openai.com/v1/chat/completions";

  var payload = {
    model: "gpt-4o-mini",
    messages: [
      {
        role: "system",
        content: "You are an excel specialist that only returns excel formulas. " +
                 "Write me an excel formula to do the following, only return the formula with the = sign."
      },
      { 
        role: "user", 
        content: "user prompt: " + user_prompt + " range notation: " + rangeNotation + " cell values string: " + cellValuesString 
      }
    ],
    max_tokens: 100
  };

  var options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (json.choices && json.choices.length > 0) {
    return json.choices[0].message.content.trim();
  } else {
    return "Error: " + JSON.stringify(json);
  }
}

function processMessage(message) {
  message = message.toLowerCase().trim();
  // Check if the message is a request to generate a formula.
  var userPrompt = message;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeRange = sheet.getActiveRange();
  var rangeNotation = activeRange.getA1Notation();
  var cellValues = activeRange.getValues();
  var cellValuesString = cellValues.map(row => row.join(", ")).join("; ");

  // Call the GPT API to generate a formula
  var formula = single_cell_formula(userPrompt, rangeNotation, cellValuesString);
  
  // Ensure the formula begins with "=" so Sheets treats it as a formula.
  if (!formula.startsWith('=')) {
    formula = '=' + formula;
  }
  
  // Ask the user for the target cell address using an input box.
  var outputCellAddress = Browser.inputBox('Enter the cell address (e.g., A5) where you want the formula to be placed:');
  
  // Write the formula to the specified cell.
  sheet.getRange(outputCellAddress).setFormula(formula);
  
  return "Formula applied to " + outputCellAddress + ": " + formula;

  // Other message processing logic can go here.
}

