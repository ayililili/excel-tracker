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

      // 獲取 A1 儲存格的值
      const cellValue = sourceCell.values[0][0];

      // 準備 API 請求的數據
      const requestBody = { data: cellValue };

      // 使用 fetch 進行 POST 請求

      // eslint-disable-next-line no-undef
      const response = await fetch("http://192.168.50.56:3000/", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
      });

      if (response.ok) {
        console.log("數據已成功上傳到 API");
      } else {
        console.error("上傳失敗，狀態碼：", response.status);
      }
    });
  } catch (error) {
    console.error("錯誤：", error);
  }
}
