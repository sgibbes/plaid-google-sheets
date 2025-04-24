/** @format */

function createChart(row, chartNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cat = sheet.getRange(row, 9).getValue();
  const spent = sheet.getRange(row, 10).getValue();
  const remaining = sheet.getRange(row, 12).getValue();
  const rowStart = chartNum * 3 - 2;

  const rowEnd = chartNum * 3;
  const startCol = 15 + (chartNum - 1) * 2;
  Logger.log(`row start: ${rowStart}, col: ${15}, num rows: ${2}, num cols: ${2}`);
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
    .setPosition(chartRow, 13, 0, 0)
    .setOption("legend", { position: "none" }) // hide callouts

    .build();

  sheet.insertChart(chart);
}
