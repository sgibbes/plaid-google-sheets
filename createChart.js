/** @format */

function createChart(row, chartNum, chartData) {
  const sheet = chartData
    ? chartData.sheet
    : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let categoryCol = chartData ? chartData.categoryCol : 0;
  let actualCol = chartData ? chartData.actualCol : 0;
  let remainingCol = chartData ? chartData.remainingCol : 0;

  if (!chartData) {
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    categoryCol = headerRow.indexOf("Category Totals") + 1;
    actualCol = headerRow.indexOf("Actual") + 1;
    remainingCol = headerRow.indexOf("Remaining") + 1;
  }

  if (!categoryCol || !actualCol || !remainingCol) {
    throw new Error('Missing required summary headers: "Category Totals", "Actual", or "Remaining"');
  }

  let cat = chartData ? chartData.category : null;
  let spent = chartData ? chartData.spent : null;
  let remaining = chartData ? chartData.remaining : null;
  if (!chartData) {
    const rowValues = sheet
      .getRange(row, categoryCol, 1, remainingCol - categoryCol + 1)
      .getValues()[0];
    cat = rowValues[0];
    spent = rowValues[actualCol - categoryCol];
    remaining = rowValues[remainingCol - categoryCol];
  }

  const startCol = remainingCol + 2 + (chartNum - 1) * 2;
  const hiddenDataRange = sheet.getRange(1, startCol, 2, 2);
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
