/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = copySingleCellToAnotherSheet;
  }
});

async function copySingleCellToAnotherSheet() {
  try {
    await Excel.run(async (context) => {
      // 獲取當前工作表的單一儲存格（例如 A1 儲存格）
      const sourceSheet = context.workbook.worksheets.getActiveWorksheet();
      const sourceCell = sourceSheet.getRange("A1");
      sourceCell.load("values"); // 載入 A1 儲存格的數據
      await context.sync();

      // 檢查是否已存在目標工作表 (B 工作表)
      let targetSheet;
      try {
        targetSheet = context.workbook.worksheets.getItem("TargetSheet");
      } catch (e) {
        targetSheet = context.workbook.worksheets.add("TargetSheet");
      }

      // 將 A1 儲存格的數據寫入到 TargetSheet 中的 B2 儲存格
      const targetCell = targetSheet.getRange("B2");
      targetCell.values = sourceCell.values;

      await context.sync();
      console.log("已成功將 A1 儲存格的內容複製到 TargetSheet 的 B2 儲存格。");
    });
  } catch (error) {
    console.error("錯誤：", error);
  }
}
