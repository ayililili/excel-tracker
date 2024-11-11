/* global console, document, Excel, Office */

// 用來儲存變更的儲存格紀錄
let changes = [];

// 當 Add-in 加載後啟動
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // 監聽儲存格變化
    monitorCellChanges();

    // 當用戶點擊 'run' 按鈕時發送 API 請求
    document.getElementById("run").onclick = sendChangesToApi;
  }
});

// 監聽儲存格變化
async function monitorCellChanges() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // 監聽儲存格變化事件
      sheet.onChanged.add((eventArgs) => {
        // 紀錄變更的儲存格資料
        const changedCell = eventArgs.address;
        const newValue = eventArgs.newValue;

        // 紀錄資料，包含儲存格座標和數值
        changes.push({ address: changedCell, value: newValue });

        console.log(`儲存格 ${changedCell} 改為：${newValue}`);
      });

      await context.sync();
    });
  } catch (error) {
    console.error("監聽儲存格變化錯誤：", error);
  }
}

// 當用戶點擊 'run' 時，將儲存格紀錄發送到 API
async function sendChangesToApi() {
  try {
    if (changes.length > 0) {
      // 準備 API 請求的數據
      const requestBody = { data: changes };

      // 使用 fetch 進行 POST 請求
      const response = await fetch("http://localhost:3001/", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
      });

      if (response.ok) {
        console.log("數據已成功上傳到 API");
        // 上傳後清空紀錄
        changes = [];
      } else {
        console.error("上傳失敗，狀態碼：", response.status);
      }
    } else {
      console.log("沒有儲存格變更記錄");
    }
  } catch (error) {
    console.error("錯誤：", error);
  }
}
