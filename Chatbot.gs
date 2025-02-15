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
        content: "You are an excel specialist that only returns google sheet formulas as you would type them in google sheet, do not include extra syntax. Please only use formulas that apply to one single cell, which generally means do not use ARRAYFORMULA" +
                 "Write me an excel formula to do the following, only return the formula with the = sign."
      },
      { 
        role: "user", 
        content: user_prompt + " The output will fall in (range notation): " + rangeNotation
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

  
  // var outputCellAddress = Browser.inputBox('Enter the cell address (e.g., A5) or range (e.g. A5:A10) where you want the formula to be placed:');

  // // Validate range
  // if (!outputCellAddress || outputCellAddress.trim() === "") {
  //   return "Error: Please enter a valid cell range.";
  // }
  
  // var range;
  // try {
  //   range = sheet.getRange(outputCellAddress);
  // } catch (e) {
  //   return "Error: Invalid range format.";
  // }

  var numRows = activeRange.getNumRows();
  var numCols = activeRange.getNumColumns();

  if (numRows === 0 || numCols === 0) {
    return "Error: The selected range is empty.";
  }

  // Select the first cell in the range
  var firstCell = activeRange.getCell(1, 1);
  if (!firstCell) {
    return "Error: Could not determine the first cell in the range.";
  }

  var firstCellNotation = firstCell.getA1Notation();
  
  // Get data from surrounding cells
  var surroundingData = sheet.getDataRange().getValues();
  var cellValuesString = surroundingData.map(row => row.join(", ")).join("; ");
  
  // Generate the formula using AI
  var formula = single_cell_formula(userPrompt, firstCellNotation, cellValuesString);

  if (!formula.startsWith('=')) {
    formula = '=' + formula;
  }

  // Apply formula to the first cell
  firstCell.setFormula(formula);

  // Autofill the rest of the range
  if (numRows > 1 || numCols > 1) {
    var autofillRange = sheet.getRange(rangeNotation);
    firstCell.autoFill(autofillRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }

  return "Formula " + formula + " applied to " + firstCellNotation + " and autofilled to " + rangeNotation;
}

function applyFormulaWithAutofill(rangeNotation, userPrompt) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Validate range
  if (!rangeNotation || rangeNotation.trim() === "") {
    return "Error: Please enter a valid cell range.";
  }
  
  var range;
  try {
    range = sheet.getRange(rangeNotation);
  } catch (e) {
    return "Error: Invalid range format.";
  }

  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  if (numRows === 0 || numCols === 0) {
    return "Error: The selected range is empty.";
  }

  // Select the first cell in the range
  var firstCell = range.getCell(1, 1);
  if (!firstCell) {
    return "Error: Could not determine the first cell in the range.";
  }

  var firstCellNotation = firstCell.getA1Notation();
  
  // Get data from surrounding cells
  var surroundingData = sheet.getDataRange().getValues();
  var cellValuesString = surroundingData.map(row => row.join(", ")).join("; ");
  
  // Generate the formula using AI
  var formula = single_cell_formula(userPrompt, firstCellNotation, cellValuesString);

  if (!formula.startsWith('=')) {
    formula = '=' + formula;
  }

  // Apply formula to the first cell
  firstCell.setFormula(formula);

  // Autofill the rest of the range
  if (numRows > 1 || numCols > 1) {
    var autofillRange = sheet.getRange(firstCellNotation + ":" + rangeNotation);
    firstCell.autoFill(autofillRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }

  return "Formula applied to " + firstCellNotation + " and autofilled to " + rangeNotation;
}

function getSelectedRange() {
    return SpreadsheetApp.getActiveRange()?.getA1Notation() ?? null;
}


