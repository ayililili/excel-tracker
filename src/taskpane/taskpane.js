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
      // 訂閱當前工作表的 onChanged 事件
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onChanged.add(onCellChange);
      // 新增一個工作表來存儲變動的儲存格資訊
      const logSheet = context.workbook.worksheets.add("ChangeLog");
      logSheet.getRange("A1").values = [["Cell Address", "New Value"]]; // 設定標題行
      await context.sync();
      console.log("已訂閱儲存格變動事件，並建立 ChangeLog 表單。");
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

      // 載入變動儲存格的值
      range.load("values");
      await context.sync();

      // 獲取或建立變動紀錄的工作表
      const logSheet = context.workbook.worksheets.getItem("ChangeLog");

      // 獲取最後一個使用的行
      const lastCell = logSheet.getRange("A1").getEntireColumn().getLastCell();
      lastCell.load("row"); // 加載行號
      await context.sync();
      const nextRow = lastCell.row + 1;

      // 在 ChangeLog 表單中記錄變動的儲存格地址和新值
      logSheet.getRange(`A${nextRow}`).values = [[changedRange]];
      logSheet.getRange(`B${nextRow}`).values = [[range.values[0][0]]];

      await context.sync();
      console.log(`變動的儲存格範圍為 ${changedRange}，已將底色變為紅色，並將新值記錄到 ChangeLog 表單。`);
    });
  } catch (error) {
    console.error("錯誤：", error);
  }
}
