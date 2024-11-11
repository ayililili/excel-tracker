/* global console, document, Excel, Office */

// 用來儲存變更的儲存格紀錄，以欄位標頭作為索引
let changes = {};
let workbookName = "";

// 當 Add-in 加載後啟動
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // 獲取檔案名
    getWorkbookName();

    // 監聽儲存格變化
    monitorCellChanges();

    // 當用戶點擊 'run' 按鈕時發送 API 請求
    document.getElementById("run").onclick = sendChangesToApi;
  }
});

// 獲取當前檔案名
async function getWorkbookName() {
  try {
    await Excel.run(async (context) => {
      workbookName = context.workbook.name;
      await context.sync();
      console.log(`檔案名：${workbookName}`);
    });
  } catch (error) {
    console.error("無法獲取檔案名：", error);
  }
}

// 監聽儲存格變化
async function monitorCellChanges() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // 監聽儲存格變化事件
      sheet.onChanged.add(async (eventArgs) => {
        console.log(eventArgs);
        const changedCell = eventArgs.address;
        const newValue = eventArgs.details.valueAfter;

        const [column, row] = changedCell.match(/[A-Z]+|\d+/g);

        // 如果是標頭列（第一行），則不予記錄
        if (parseInt(row, 10) === 1) {
          return;
        }

        // 獲取標頭值
        const headerCell = `${column}1`;
        const headerRange = sheet.getRange(headerCell);
        headerRange.load("values");

        await context.sync();

        const headerValue = headerRange.values[0][0];

        // 使用標頭值作為索引，僅保留最後更動的值
        changes[headerValue] = newValue;

        console.log(`儲存格 ${changedCell}（${headerValue}）改為：${newValue}`);
      });

      await context.sync();
    });
  } catch (error) {
    console.error("監聽儲存格變化錯誤：", error);
  }
}

// 當用戶點擊 'run' 時，將儲存格紀錄和檔案名發送到 API
async function sendChangesToApi() {
  try {
    const changeEntries = Object.entries(changes);
    if (changeEntries.length > 0) {
      // 將物件轉換成數組，以便於發送
      const requestBody = {
        filename: workbookName,
        data: changeEntries.map(([header, value]) => ({ header, value })),
      };

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
        changes = {};
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
