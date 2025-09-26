/* global Office, Excel */
Office.onReady(() => {
  document.getElementById("run").addEventListener("click", run);
});

async function run() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const a1 = sheet.getRange("A1");
    a1.values = [["Hello, Excel Online!"]];
    a1.format.autofitColumns();
    await context.sync();
  }).catch((err) => console.error(err));
}
