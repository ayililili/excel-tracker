/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      // 訂閱工作表的 onChanged 事件
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onChanged.add(onCellChange);
      console.log("已訂閱儲存格變動事件。");
    });
  } catch (error) {
    console.error(error);
  }
}

// 當儲存格內容變動時觸發此函數
async function onCellChange(event) {
  try {
    await Excel.run(async (context) => {
      // 獲取變動範圍
      const changedRange = event.address;
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(changedRange);

      // 更改變動儲存格的填充顏色為紅色
      range.format.fill.color = "red";

      await context.sync();
      console.log(`變動的儲存格範圍為 ${changedRange}，已將底色變為紅色。`);
    });
  } catch (error) {
    console.error("錯誤：", error);
  }
}
