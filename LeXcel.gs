function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("LeXcel")
    .addItem("Open LeXcel Assistant", "showChatbotSidebar")
    .addToUi();
}

function showChatbotSidebar() {
  var html =
    HtmlService.createHtmlOutputFromFile("LeXcelUI").setTitle(
      "LeXcel Assistant"
    );
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  return data;
}

function single_cell_formula(user_prompt, rangeNotation, cellValuesString) {
  var apiKey =
  var url = "https://api.openai.com/v1/chat/completions";

  var payload = {
    model: "gpt-4o-mini",
    messages: [
      {
        role: "system",
        content:
          "You are an google sheets specialist that only returns google sheet formulas as you would type them in google sheet, do not include extra syntax. Please only use formulas that apply to one single cell, which generally means do not use ARRAYFORMULA. Do NOT include any cell from the output range in the formula. " +
          "Write me an google sheets formula to do the following, only return the formula with the = sign.",
      },
      {
        role: "user",
        content:
          user_prompt +
          " The output will fall in (range notation): " +
          rangeNotation,
      },
    ],
    max_tokens: 100,
  };

  var options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + apiKey,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (json.choices && json.choices.length > 0) {
    return json.choices[0].message.content.trim();
  } else {
    return "Error: " + JSON.stringify(json);
  }
}

const ERROR_VALUES = [
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#ERROR!",
];

function checkFormulaForErrors(formula, range) {
  try {
    // First try to parse the formula without applying it
    var sheet = SpreadsheetApp.getActiveSheet();
    var dummyRange = sheet.getRange("ZZ100"); // Use a temporary cell
    var originalValue = dummyRange.getValue();
    var originalFormula = dummyRange.getFormula();

    try {
      dummyRange.setFormula(formula);
      var error = dummyRange.getValue();

      // Check if the result is an error
      if (typeof error === "string" && ERROR_VALUES.includes(error)) {
        return {
          hasError: true,
          errorType: error,
          details: `Formula resulted in ${error}`,
        };
      }

      return { hasError: false };
    } finally {
      // Restore original state of dummy cell
      if (originalFormula) {
        dummyRange.setFormula(originalFormula);
      } else {
        dummyRange.setValue(originalValue);
      }
    }
  } catch (e) {
    return {
      hasError: true,
      errorType: "#ERROR!",
      details: e.toString(),
    };
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
  var cellValuesString = cellValues.map((row) => row.join(", ")).join("; ");

  var firstCell = activeRange.getCell(1, 1);
  var numRows = activeRange.getNumRows();
  var numCols = activeRange.getNumColumns();
  var firstCellNotation = firstCell.getA1Notation();

  // Check if this is a request to apply a specific formula without error checking
  if (
    message
      .toLowerCase()
      .startsWith("apply this exact formula without any changes")
  ) {
    const formulaMatch = message.match(/: (=.+)$/);
    if (formulaMatch) {
      const formula = formulaMatch[1];

      // Apply formula directly without error checking
      firstCell.setFormula(formula);

      // Autofill if needed
      if (numRows > 1 || numCols > 1) {
        var autofillRange = sheet.getRange(rangeNotation);
        firstCell.autoFill(
          autofillRange,
          SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
        );
      }

      return JSON.stringify({
        error: false,
        formula: formula,
        range: rangeNotation,
        warning: "Formula applied as requested, ignoring potential errors",
      });
    }
  }

  // Get data from surrounding cells
  var surroundingData = sheet.getDataRange().getValues();
  var cellValuesString = surroundingData
    .map((row) => row.join(", "))
    .join("; ");

  // Generate the formula using AI
  var formula = single_cell_formula(
    userPrompt,
    firstCellNotation,
    cellValuesString
  );

  if (!formula.startsWith("=")) {
    formula = "=" + formula;
  }

  // Check formula for errors before applying
  var errorCheck = checkFormulaForErrors(formula, firstCell);
  if (errorCheck.hasError) {
    return JSON.stringify({
      error: true,
      formula: formula,
      errorType: errorCheck.errorType,
      errorDetails: errorCheck.details,
    });
  }

  // Apply formula to the first cell
  firstCell.setFormula(formula);

  // Autofill the rest of the range
  if (numRows > 1 || numCols > 1) {
    var autofillRange = sheet.getRange(rangeNotation);
    firstCell.autoFill(
      autofillRange,
      SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );
  }

  return JSON.stringify({
    error: false,
    formula: formula,
    range: rangeNotation,
  });
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
  var cellValuesString = surroundingData
    .map((row) => row.join(", "))
    .join("; ");

  // Generate the formula using AI
  var formula = single_cell_formula(
    userPrompt,
    firstCellNotation,
    cellValuesString
  );

  if (!formula.startsWith("=")) {
    formula = "=" + formula;
  }

  // Check formula for errors before applying
  var errorCheck = checkFormulaForErrors(formula, firstCell);
  if (errorCheck.hasError) {
    return JSON.stringify({
      error: true,
      formula: formula,
      errorType: errorCheck.errorType,
      errorDetails: errorCheck.details,
    });
  }

  // Apply formula to the first cell
  firstCell.setFormula(formula);

  // Autofill the rest of the range
  if (numRows > 1 || numCols > 1) {
    var autofillRange = sheet.getRange(firstCellNotation + ":" + rangeNotation);
    firstCell.autoFill(
      autofillRange,
      SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );
  }

  return JSON.stringify({
    error: false,
    formula: formula,
    range: rangeNotation,
  });
}

function getSelectedRange() {
  return SpreadsheetApp.getActiveRange()?.getA1Notation() ?? null;
}

function checkRangeIsEmpty() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();

  if (!range) {
    return { isEmpty: true };
  }

  var values = range.getValues();
  var hasContent = false;

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] !== "") {
        hasContent = true;
        break;
      }
    }
    if (hasContent) break;
  }

  return {
    isEmpty: !hasContent,
    rangeNotation: range.getA1Notation(),
  };
}
