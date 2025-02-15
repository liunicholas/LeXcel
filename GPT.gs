function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GPT Tools')
    .addItem('Get GPT Formula', 'writeGPTFormulaToCell')
    .addToUi();
}

function single_cell_formula(user_prompt, rangeNotation, cellValuesString) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!apiKey) {
    throw new Error("API key not set in script properties. Please add it in the Project Properties.");
  }
  var url = "https://api.openai.com/v1/chat/completions";

  var payload = {
    model: "gpt-4o-mini",
    messages: [
      {
        role: "system",
        content: "You are an excel specialist that only returns excel formulas. " +
                 "Write me an excel formula to do the following, only return the formula with the = sign."
      },
      { role: "user", content: "user prompt: " + user_prompt + " range notation: " + rangeNotation + " cell values string: " + cellValuesString}
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

function writeGPTFormulaToCell() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Get the currently selected range for context
  var activeRange = sheet.getActiveRange();
  var rangeNotation = activeRange.getA1Notation();
  var cellValues = activeRange.getValues();
  var cellValuesString = cellValues.map(row => row.join(", ")).join("; ");
  
  // Ask the user for the formula prompt (using the selected range for context)
  var userPrompt = Browser.inputBox(
    'Enter your formula request. ' +
    'Selected range ' + rangeNotation + ' has values: ' + cellValuesString + 
    '. What would you like to do?'
  );
  
  // Ask the user to specify the cell address where the formula should be placed
  var outputCellAddress = Browser.inputBox('Enter the cell address (e.g. A5) where you want the formula to be placed:');
  var outputCell = sheet.getRange(outputCellAddress);
  
  // Get the GPT generated formula (ensuring it starts with "=")
  var formula = GPT(userPrompt);
  if (!formula.startsWith('=')) {
    formula = '=' + formula;
  }
  
  // Set the formula into the chosen cell
  outputCell.setFormula(formula);
}
