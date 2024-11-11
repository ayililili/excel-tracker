/* global console, document, Excel, Office */

// 用來儲存變更的儲存格紀錄，以編號作為索引
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
      const workbook = context.workbook;
      workbook.load("name"); // Explicitly load the 'name' property

      await context.sync();
      workbookName = workbook.name;
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

        // 如果是第一排（編號列）或第一行（項目列），則不予記錄
        if (parseInt(row, 10) === 1 || parseInt(column, 10) === 1) {
          return;
        }

        // 獲取編號列（第一列）的值
        const idCell = `A${row}`;  // 取得第一列的編號
        const idRange = sheet.getRange(idCell);
        idRange.load("values");

        // 獲取項目列（第一行）的值
        const headerCell = `${column}1`; // 取得第一行的標題
        const headerRange = sheet.getRange(headerCell);
        headerRange.load("values");

        await context.sync();

        const idValue = idRange.values[0][0]; // 取得編號
        const headerValue = headerRange.values[0][0]; // 取得項目名稱

        // 確保編號在紀錄物件中
        if (!changes[idValue]) {
          changes[idValue] = {};
        }

        // 使用編號作為索引，並將項目名稱和值加入
        changes[idValue][headerValue] = newValue;

        console.log(`儲存格 ${changedCell}（編號: ${idValue}, 項目: ${headerValue}）改為：${newValue}`);
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
        id: workbookName,
        data: changeEntries.map(([id, items]) => ({
          id,
          items: Object.entries(items).map(([header, value]) => ({ header, value }))
        })),
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
