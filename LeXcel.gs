var openai_key = ""

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 LeXcel")
    .addItem("Open LeXcel Assistant", "showChatbotSidebar")
    .addToUi();

  // Clear message history when spreadsheet is opened
  clearMessageHistory();
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

function single_cell_formula(user_prompt, rangeNotation, exampleCellFormat) {
  var url = "https://api.openai.com/v1/chat/completions";

  // const history = getMessageHistoryContext();
  const prompt = [
    "**OUTPUT LOCATION**:" + rangeNotation + "\n",
    exampleCellFormat ? "**EXAMPLE INPUT**:" + exampleCellFormat + "\n" : "",
    "**TASK**:" + user_prompt,
  ]
    .filter(Boolean)
    .join("");

  // Browser.msgBox(prompt);

  var payload = {
    model: "gpt-4o",
    messages: [
      {
        role: "system",
        content:
          "You are an google sheets specialist that only returns google sheet formulas as you would type them in google sheet, do not include extra syntax. Do NOT include any cell from the output range in the formula. Do NOT use any markdown notation like ```. Be careful about using fixed references vs. relative references (use $ or not). " +
          "Write me an google sheets formula to do the following, only return the formula with the = sign.",
      },
      {
        role: "user",
        content: prompt,
      },
    ],
    max_tokens: 100,
  };

  var options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + openai_key,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (json.choices && json.choices.length > 0) {
      const formula = json.choices[0].message.content.trim();
      // Check if formula starts with =
      if (!formula.startsWith("=")) {
        return JSON.stringify({
          error: true,
          type: "formula_format",
          formula: formula,
          details: "Generated formula does not start with =",
        });
      }
      return JSON.stringify({
        error: false,
        formula: formula,
        type: "formula",
        details: null,
      });
    } else {
      return JSON.stringify({
        error: true,
        type: "api_error",
        details: JSON.stringify(json),
      });
    }
  } catch (error) {
    return JSON.stringify({
      error: true,
      type: "runtime_error",
      details: error.toString(),
    });
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
    var dummyRange = sheet.getRange("Z100"); // Use a temporary cell
    var originalValue = dummyRange.getValue();
    var originalFormula = dummyRange.getFormula();

    try {
      dummyRange.setFormula(formula);
      var error = dummyRange.getValue();

      // Check if the result is an error
      if (typeof error === "string" && ERROR_VALUES.includes(error)) {
        return {
          error: true,
          errorType: error,
          details: `Formula resulted in ${error}`,
        };
      }

      return { error: false };
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
      error: true,
      errorType: "#ERROR!",
      details: e.toString(),
    };
  }
}

function classifyMessage(message) {
  const OPENAI_API_KEY = ""
  const url = "https://api.openai.com/v1/chat/completions";

  const prompt = [
    {
      role: "system",
      content:
        "You are a classifier that categorizes user requests into 'formula' or 'plot'. Only respond with the exact category name, nothing else.",
    },
    {
      role: "user",
      content:
        "Here are some examples:\n1. 'Create a line chart comparing column A and B' -> plot\n2. 'Find the sum of A1 and B1' -> formula\n3. 'Make a pie chart with values B2:B10 and labels A2:A10' -> plot\n4. 'Fill each box with the row that it's in' -> formula",
    },
    {
      role: "user",
      content: `Classify this request: "${message}"`,
    },
  ];

  const payload = {
    model: "gpt-4o",
    messages: prompt,
    temperature: 0,
    max_tokens: 10,
  };

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${OPENAI_API_KEY}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    return result.choices[0].message.content.trim();
  } catch (error) {
    console.error("Classification error:", error);
    return "none";
  }
}

function extractPlotInfo(message) {
  // Browser.msgBox("1");
  const OPENAI_API_KEY = openai_key;
  const url = "https://api.openai.com/v1/chat/completions";

  const prompt = [
    {
      role: "system",
      content: `You are a plot information extractor. Extract plot details from user requests and output them in a specific JSON format. 
      Pay special attention to which data should be on the X and Y axes. 
      For scatter plots, detect if the user wants a trend line (also known as best fit line or regression line).
      For line and bar charts without explicit X-axis data, use row numbers as X-axis.
      Only output valid JSON, nothing else.`,
    },
    {
      role: "user",
      content: `Here are examples:
Request: "Create a scatter plot with temperature in column B against time in column A with a trend line"
{
  "type": "scatter",
  "xAxisRange": "A:A",
  "yAxisRanges": ["B:B"],
  "title": "Temperature over Time",
  "xAxisTitle": "Time",
  "yAxisTitle": "Temperature",
  "trendLine": true,
  "trendLineType": "linear"
}

Request: "Make a scatter plot comparing height (C1:C20) on the y axis and weight (D1:D20) on the x axis and add a best fit line"
{
  "type": "scatter",
  "xAxisRange": "D1:D20",
  "yAxisRanges": ["C1:C20"],
  "title": "Height vs Weight",
  "xAxisTitle": "Height",
  "yAxisTitle": "Weight",
  "trendLine": true,
  "trendLineType": "linear"
}

Request: "Make a pie chart with sales [B3:B8] as values and product [C3:C8] as labels."
{
  "type": "pie",
  "xAxisRange": "B3:B8",
  "yAxisRanges": ["C3:C8"],
  "title": "Sales by product",
}

Request: "create a bar chart representing flavor preferences by gender with [B22:B24] as female data [C22:C24] as male data and [A22:A24] as labels. the data represents what percentage of each gender likes the given flavor."
{
  "type": "bar",
  "xAxisRange": "A22:A24",
  "yAxisRanges": ["B22:B24", "C22:C24"],
  "title": "Flavor Preferences by Gender",
  "xAxisTitle": "Flavor",
  "yAxisTitle": "Percent Preferred",
  "seriesNames": ["Female", "Male"],
  "legendPosition": "right"
}
`,
    },
    {
      role: "user",
      content: `Request: "${message}"`,
    },
  ];



  const payload = {
    model: "gpt-4o-mini",
    messages: prompt,
    temperature: 0,
  };

  // Browser.msgBox("2");

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${OPENAI_API_KEY}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  // Browser.msgBox("3");

  try {
    // Browser.msgBox("url" +  url);
    // Browser.msgBox("options" + options);
    const response = UrlFetchApp.fetch(url, options);
    // Browser.msgBox("response " + response.getContextText());
    const result = JSON.parse(response.getContentText());
    // Browser.msgBox("result, " + result.choices[0].message.content);
    const plotInfo = JSON.parse(result.choices[0].message.content);
    return plotInfo;
  } catch (error) {
    // Browser.msgBox(error);
    console.error("Error extracting plot info:", error);
    return null;
  }
}

function createPlot(plotInfo) {
  const sheet = SpreadsheetApp.getActiveSheet();

  // Helper function to get range
  const getRange = (rangeNotation) => {
    if (!rangeNotation) return null;
    if (rangeNotation.match(/^[A-Z]:[A-Z]$/)) {
      const col = rangeNotation.split(":")[0];
      const lastRow = sheet.getLastRow();
      return sheet.getRange(`${col}1:${col}${lastRow}`);
    }
    return sheet.getRange(rangeNotation);
  };

  // Get X and Y axis ranges
  const xAxisRange = plotInfo.xAxisRange ? getRange(plotInfo.xAxisRange) : null;
  const yAxisRanges = plotInfo.yAxisRanges.map(getRange);

  // Create the chart
  const chart = sheet.newChart();

  // Set chart type
  switch (plotInfo.type.toLowerCase()) {
    case "line":
      chart.setChartType(Charts.ChartType.LINE);
      break;
    case "bar":
      chart.setChartType(Charts.ChartType.COLUMN);
      break;
    case "scatter":
      chart.setChartType(Charts.ChartType.SCATTER);
      break;
    case "pie":
      chart.setChartType(Charts.ChartType.PIE);
      break;
    default:
      chart.setChartType(Charts.ChartType.LINE);
  }

  // Handle different chart types
  if (plotInfo.type.toLowerCase() === "scatter") {
    if (!xAxisRange || yAxisRanges.length === 0) {
      throw new Error("Scatter plots require both X and Y axis data");
    }
    chart.addRange(xAxisRange);
    yAxisRanges.forEach((range) => chart.addRange(range));

    // Add trend line if requested
    if (plotInfo.trendLine) {
      chart
        .setOption("trendlines.0.type", "linear")
        .setOption("trendlines.0.showR2", true)
        .setOption("trendlines.0.visibleInLegend", true)
        .setOption("trendlines.0.color", "#666666")
        .setOption("trendlines.0.lineWidth", 2)
        .setOption("trendlines.0.opacity", 0.8)
        .setOption("trendlines.0.labelInLegend", "Trend Line");
    }
  } else if (plotInfo.type.toLowerCase() === "pie") {
    if (yAxisRanges.length === 0) {
      throw new Error("Pie charts require at least one data range");
    }
    chart.addRange(yAxisRanges[0]);
    if (xAxisRange) {
      chart.addRange(xAxisRange);
    }
  } else {
    // For line and bar charts
    if (xAxisRange) {
      // Add X-axis range first
      chart.addRange(xAxisRange);

      // Add each Y series in order
      yAxisRanges.forEach((yRange, index) => {
        chart
          .addRange(yRange)
          .setOption(`series.${index}.labelInLegend`, yRange.getA1Notation());
      });
    } else {
      // Use row numbers as X-axis
      yAxisRanges.forEach((range) => chart.addRange(range));
    }
  }

  // Configure chart
  chart
    .setPosition(5, 5, 0, 0)
    .setOption("title", plotInfo.title)
    .setOption("width", 600)
    .setOption("height", 400)
    .setOption("hAxis.title", plotInfo.xAxisTitle)
    .setOption("vAxis.title", plotInfo.yAxisTitle);

  // Set series names if specified
  if (plotInfo.seriesNames) {
    plotInfo.seriesNames.forEach((name, index) => {
      chart.setOption(`series.${index}.labelInLegend`, name);
    });
  }

  // Set legend position if specified
  if (plotInfo.legendPosition) {
    chart.setOption("legend.position", plotInfo.legendPosition);
  }

  try {
    // Build and insert the chart
    sheet.insertChart(chart.build());

    return {
      type: "plot",
      plotType: plotInfo.type,
      xAxisRange: plotInfo.xAxisRange || "Row Numbers",
      yAxisRanges: plotInfo.yAxisRanges.join(", "),
      trendLine: plotInfo.trendLine ? "with trend line" : "without trend line",
      success: true,
    };
  } catch (error) {
    console.error("Error creating chart:", error);
    return {
      error: true,
      message: `Failed to create chart: ${error.toString()}`,
    };
  }
}

function processMessage(message) {
  // Add message to history before processing
  addToMessageHistory(message);

  message = message.trim();

  // First, classify the message
  const messageType = classifyMessage(message);

  if (messageType === "plot") {
    // Handle plot request
    // Browser.msgBox(message);
    const plotInfo = extractPlotInfo(message);
    if (plotInfo) {
      try {
        // Browser.msgBox(JSON.stringify(plotInfo));

        const result = createPlot(plotInfo);

        // Browser.msgBox(JSON.stringify(result));

        addToMessageHistory(
          `Created a ${
            plotInfo.type
          } plot with data from ${plotInfo.yAxisRanges.join(", ")}`,
          true
        );
        return JSON.stringify(result);
      } catch (error) {
        const errorMsg = `Error creating chart: ${error.toString()}`;
        addToMessageHistory(errorMsg, true);
        return JSON.stringify({
          error: true,
          message: errorMsg,
        });
      }
    } else {
      const errorMsg =
        "Could not understand plot requirements. Please specify what type of chart you want to create and what data to use.";
      addToMessageHistory(errorMsg, true);
      return JSON.stringify({
        error: true,
        message: errorMsg,
      });
    }
  } else if (messageType === "formula") {
    // Regular expression to match A1 notation (e.g., A1, B2, AA12, etc.) including ranges (e.g., A1:B2)
    var a1NotationRegex = /[A-Za-z]+[0-9]+(?::[A-Za-z]+[0-9]+)?/g;
    var exampleCellFormat = "";

    // Search for A1 notation in the message
    var matches = message.match(a1NotationRegex);
    if (matches && matches.length === 1) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var cellReference = matches[0];
      // Get the actual value from the referenced cell or first cell in range
      var range = cellReference.includes(":")
        ? cellReference.split(":")[0]
        : cellReference;
      exampleCellFormat = sheet.getRange(range).getValue();
    } else {
      // If no matches or multiple matches found, use empty string
      exampleCellFormat = "";
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var activeRange = sheet.getActiveRange();
    var rangeNotation = activeRange.getA1Notation();
    var cellValues = activeRange.getValues();
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
      // Browser.msgBox("applying formula");
      // Browser.msgBox(message);
      const formulaMatch = message.match(/: (=.+)$/);
      if (formulaMatch) {
        const formula = formulaMatch[1];
        // Browser.msgBox(formula);
        firstCell.setFormula(formula);
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
    var formulaResponse = single_cell_formula(
      message,
      firstCellNotation,
      exampleCellFormat
    );
    // Browser.msgBox(formulaResponse);

    var formulaResponseJson = JSON.parse(formulaResponse);


    if (formulaResponseJson.error) {
      const errorMsg = `Formula error: ${formulaResponseJson.details}`;
      addToMessageHistory(errorMsg, true);
      return JSON.stringify({
        error: true,
        type: formulaResponseJson.type,
        formula: formulaResponseJson.formula,
        errorType: formulaResponseJson.errorType,
        errorDetails: formulaResponseJson.details,
      });
    }

    var formula = formulaResponseJson.formula;

    // Browser.msgBox(formula);
    // Check formula for errors before applying
    var errorCheck = checkFormulaForErrors(formula, firstCell);

    // Browser.msgBox(JSON.stringify(errorCheck));
    if (errorCheck.error) {
      const errorMsg = `Formula validation error: ${errorCheck.message}`;
      addToMessageHistory(errorMsg, true);
      return JSON.stringify({
        error: true,
        type: errorCheck.errorType,
        formula: formula,
        errorType: errorCheck.errorType,
        errorDetails: errorCheck.details,
      });
    }

    // Apply the formula
    firstCell.setFormula(formula);

    // If range is more than one cell, autofill
    if (numRows > 1 || numCols > 1) {
      var autofillRange = sheet.getRange(rangeNotation);
      firstCell.autoFill(
        autofillRange,
        SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
      );
    }

    const successMsg = `Applied formula ${formula} to range ${rangeNotation}`;
    addToMessageHistory(successMsg, true);
    return JSON.stringify({
      error: false,
      formula: formula,
      range: rangeNotation,
    });
  } else {
    return JSON.stringify({
      error: true,
      message:
        "I can only help with formulas and plots. Please try rephrasing your request.",
    });
  }
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

/**
 * Custom function that can be used directly in cells with =LEX("query")
 * @param {string} query The user's text generation request
 * @return {Array<Array<string>>} Array of text responses
 * @customfunction
 */
function LEX(query) {
  var url = "https://api.openai.com/v1/chat/completions";

  var payload = {
    model: "gpt-4o",
    messages: [
      {
        role: "system",
        content:
          "You are a text processing assistant. When given a request, generate the appropriate text output. " +
          "If the request implies multiple cells or lines, return them as separate items. Keep responses concise and direct. " +
          "Do not include any formatting instructions like bold or italic. Just return the text.",
      },
      {
        role: "user",
        content: query,
      },
    ],
    max_tokens: 1000,
  };

  var options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + openai_key,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/chat/completions",
    options
  );
  var json = JSON.parse(response.getContentText());

  if (json.choices && json.choices.length > 0) {
    // Split the response into lines
    var lines = json.choices[0].message.content.trim().split("\n");

    // Check if the query contains keywords suggesting horizontal layout
    var wantsHorizontal =
      query.toLowerCase().includes("horizontal") ||
      query.toLowerCase().includes("in a row") ||
      query.toLowerCase().includes("across");

    if (wantsHorizontal) {
      // Return as a single row
      return [lines.map((line) => line.trim())];
    } else {
      // Return as a column (default behavior)
      return lines.map((line) => [line.trim()]);
    }
  } else {
    return [["Error: " + JSON.stringify(json)]];
  }
}

function showHelpDialog() {
  var html = HtmlService.createHtmlOutput(
    "<style>" +
      "body { font-family: 'Inter', sans-serif; color: #333; padding: 20px; }" +
      "h2 { color: #7C3AED; font-size: 20px; margin-bottom: 10px; margin-top: 0px; }" +
      "ol { padding-left: 20px; font-size: 14px; line-height: 1.6; }" +
      "li { margin-bottom: 8px; }" +
      "code { background-color: #EDE9FE; padding: 2px 5px; border-radius: 4px; font-weight: bold; color: #7C3AED; }" +
      ".tip { background-color: #F3F0FF; padding: 10px; border-radius: 6px; margin-top: 10px; font-size: 14px; }" +
      "</style>" +
      "<h2>Example Usage: </h2>" +
      "<ol>" +
      "<li><b>Type Your Message:</b> Type <i>“Sum all the elements in ”</i> in the text box.</li>" +
      "<li><b>Select Input Range:</b> Highlight <code>[A1:A9]</code> and click <b><big>@</big> Insert Selected Range</b>.</li>" +
      "<li><b>(Alternative Range Specification):</b> Instead of selecting <code>[A1:A9]</code> you can simply specify in text “column A”---Lex can infer from the context.</li>" +
      "<li><b>Select Output Range:</b> Click on <code>[D4]</code> (where the result should appear) and click <b>Send</b>.</li>" +
      "</ol>" +
      "<div class='tip'>💡 <b>Tip:</b> While Lex can infer missing details, selecting ranges ensures accuracy!</div>"
  )
    .setWidth(600)
    .setHeight(310);

  SpreadsheetApp.getUi().showModalDialog(html, "LeXcel Assistant Help");
}

function addToMessageHistory(message) {
  const userProperties = PropertiesService.getUserProperties();
  const history = JSON.parse(userProperties.getProperty('messageHistory') || '[]');
  history.push(message);
  // Keep last 50 messages
  if (history.length > 50) {
    history.shift();
  }
  userProperties.setProperty('messageHistory', JSON.stringify(history));
}

function getMessageHistory() {
  const userProperties = PropertiesService.getUserProperties();
  return JSON.parse(userProperties.getProperty('messageHistory') || '[]');
}

function getMessageHistoryContext() {
  const messageHistory = getMessageHistory();

  if (messageHistory.length <= 1) return "";

  // Get unique messages, keeping only the most recent occurrence of each
  const uniqueMessages = [...new Set(messageHistory.reverse())].reverse();
  
  // Take last 4 messages, excluding the most recent one
  const relevantHistory = uniqueMessages.slice(-5, -1);

  if (relevantHistory.length === 0) return "";

  const historyText = relevantHistory.join("\n");
  return "\n\nHere is some context from previous messages:\n" + historyText;
}

function clearMessageHistory() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('messageHistory');
}
