/** @format */

function createChart(row, chartNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const categoryCol = headerRow.indexOf("Category Totals") + 1;
  const actualCol = headerRow.indexOf("Actual") + 1;
  const remainingCol = headerRow.indexOf("Remaining") + 1;

  if (!categoryCol || !actualCol || !remainingCol) {
    throw new Error('Missing required summary headers: "Category Totals", "Actual", or "Remaining"');
  }

  const cat = sheet.getRange(row, categoryCol).getValue();
  const spent = sheet.getRange(row, actualCol).getValue();
  const remaining = sheet.getRange(row, remainingCol).getValue();
  const rowStart = chartNum * 3 - 2;

  const rowEnd = chartNum * 3;
  const startCol = remainingCol + 2 + (chartNum - 1) * 2;
  Logger.log(`row start: ${rowStart}, col: ${startCol}, num rows: ${2}, num cols: ${2}`);
  const hiddenDataRange = sheet.getRange(1, startCol, 2, 2);
  hiddenDataRange.clearContent();
  hiddenDataRange.setValues([
    ["Spent", spent],
    ["Remaining", remaining],
  ]);

  //   hiddenDataRange.setValues([
  //   ["Label", cat],
  //   ["Spent", spent],
  //   ["Remaining", remaining],
  // ]);

  // chart is 5 rows tall, each time, shift it down 5 rows
  const chartRow = chartNum === 1 ? 1 : 5 * (chartNum - 1);
  // build the chart
  const chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(hiddenDataRange)
    .setOption("title", `${cat}`)
    .setOption("titleTextStyle", {
      bold: true,
      fontSize: 10,
      color: "#333",
      alignment: "center",
    })
    .setOption("pieHole", 0.5)
    .setOption("pieSliceText", "value")
    .setOption("colors", ["#b32222", "#22b322"]) //spent, remaining
    .setOption("width", 100)
    .setOption("height", 100)
    .setPosition(chartRow, remainingCol, 0, 0)
    .setOption("legend", { position: "none" }) // hide callouts

    .build();

  sheet.insertChart(chart);
}
