import { ExcelService } from "../services/excel.service";
import { ApiService } from "../services/api.service";
import { ChangesStore } from "../stores/changes.store";

class TaskPane {
  constructor() {
    this.excelService = new ExcelService();
    this.apiService = new ApiService();
    this.changesStore = new ChangesStore();
    this.isValidDocumentType = false;
  }

  async initialize() {
    if (Office.context.host === Office.HostType.Excel) {
      await this.setupWorkbook();
      this.setupEventListeners();
    }
  }

  async setupWorkbook() {
    try {
      const { type } = await this.excelService.determineDocumentType();

      if (type >= 1 && type <= 3) {
        this.isValidDocumentType = true;
        await this.excelService.captureSnapshot();
        console.log("檔案類型有效，已捕獲初始快照");
      } else {
        this.isValidDocumentType = false;
        this.showInvalidFileTypeBanner();
        console.log("檔案類型無效");
      }
    } catch (error) {
      console.error("設置工作簿時發生錯誤：", error);
      await this.showNotification("設置工作簿時發生錯誤", "error");
    }
  }

  showInvalidFileTypeBanner() {
    // 創建或更新無效檔案類型的橫幅
    let banner = document.getElementById("invalid-type-banner");
    if (!banner) {
      banner = document.createElement("div");
      banner.id = "invalid-type-banner";
      banner.className = "warning-banner";

      const content = document.createElement("div");
      content.innerHTML = `
      <strong>⚠️ 檔案名稱格式不符合要求</strong>
      <p>目前檔案: ${this.excelService.workbookName}</p>
      <p>變動將不予紀錄，請確認檔案名稱格式是否正確。</p>
    `;

      banner.appendChild(content);
      document.getElementById("app-body").insertBefore(banner, document.getElementById("app-body").firstChild);
    } else {
      banner.querySelector("p").textContent = `目前檔案: ${this.excelService.workbookName}`;
    }
  }

  setupEventListeners() {
    // 基本功能按鈕
    document.getElementById("push").onclick = () => this.sendChangesToApi();
    document.getElementById("pull").onclick = () => this.syncTableWithApi();
    document.getElementById("restart").onclick = () => this.handleFileNameChange();
  }

  async handleFileNameChange() {
    try {
      const { type } = await this.excelService.determineDocumentType();
      const newWorkbookName = await this.excelService.getWorkbookName();

      // 檢查檔名是否真的有變更
      if (newWorkbookName === this.excelService.workbookName) {
        await this.showNotification("檔案名稱沒有變更", "info");
        return;
      }

      if (type >= 1 && type <= 3) {
        this.isValidDocumentType = true;
        await this.excelService.captureSnapshot();
        await this.showNotification("檔案名稱已更新，已重新捕獲快照", "success");
        console.log("檔案名稱已更新，類型有效，已捕獲新快照");

        // 移除警告橫幅（如果存在）
        const banner = document.getElementById("invalid-type-banner");
        if (banner) banner.remove();

        // 如果是類型3，提醒用戶工作表已被保護
        if (type === 3) {
          await this.showNotification("工作表已啟用保護，僅允許必要的編輯操作", "info");
        }
      } else {
        this.isValidDocumentType = false;
        this.showInvalidFileTypeBanner();
        await this.showNotification("新檔案名稱格式不符合要求，變動將不予紀錄", "warning");
        console.log("新檔案名稱類型無效");
      }
    } catch (error) {
      console.error("處理檔名變更時發生錯誤：", error);
      await this.showNotification("處理檔名變更時發生錯誤", "error");
    }
  }

  async checkForChanges() {
    if (!this.isValidDocumentType) {
      console.log("檔案類型無效，不進行變更檢查");
      return null;
    }

    try {
      const changes = await this.excelService.compareWithSnapshot();
      if (changes && Object.keys(changes).length > 0) {
        this.changesStore.setChanges(changes);
        console.log("變更已記錄:", changes);
        return changes;
      }
      return null;
    } catch (error) {
      console.error("檢查變更時發生錯誤：", error);
      await this.showNotification("檢查變更時發生錯誤", "error");
      throw error;
    }
  }

  async sendChangesToApi() {
    if (!this.isValidDocumentType) {
      await this.showNotification("檔案類型無效，不進行資料上傳", "warning");
      return;
    }

    try {
      const changes = await this.checkForChanges();
      if (changes) {
        await this.apiService.sendChanges(this.excelService.workbookName, changes);
        this.changesStore.clear();
        await this.showNotification("數據已成功上傳", "success");
        console.log("數據已成功上傳到 API");

        await this.excelService.captureSnapshot();
        console.log("已捕獲新快照");
      }
    } catch (error) {
      console.error("錯誤：", error);
      await this.showNotification("上傳資料時發生錯誤", "error");
    }
  }

  async syncTableWithApi() {
    if (!this.isValidDocumentType) {
      await this.showNotification("檔案類型無效，不進行資料同步", "warning");
      return;
    }

    try {
      const data = await this.apiService.fetchData();
      const workbookData = data[this.excelService.workbookName];

      if (!workbookData) {
        await this.showNotification("沒有可同步的資料", "info");
        return;
      }

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // 清除格式
        const clearRange = sheet.getRange("A2:Z1000");
        clearRange.format.fill.clear();

        // 載入範圍
        const dataRange = sheet.getRange("A2:Z1000");
        const headerRange = sheet.getRange("A1:Z1");
        dataRange.load("values");
        headerRange.load("values");
        await context.sync();

        // 更新資料
        for (const item of workbookData) {
          // 尋找對應的 ID 行
          for (let row = 0; row < dataRange.values.length; row++) {
            if (dataRange.values[row][0] === item.id) {
              // 更新每個欄位
              for (const field of item.items) {
                const colIndex = headerRange.values[0].indexOf(field.header);
                if (colIndex !== -1) {
                  const range = sheet.getRange(row + 2, colIndex + 1);
                  range.values = [[field.value]];
                  range.format.fill.color = "yellow";
                }
              }
              break;
            }
          }
        }

        await context.sync();
        await this.excelService.captureSnapshot();
        await this.showNotification("資料同步完成", "success");
      });
    } catch (error) {
      console.error("同步表格資料時發生錯誤：", error);
      await this.showNotification("同步資料時發生錯誤", "error");
    }
  }

  async showNotification(message, type = "info") {
    try {
      // 創建通知元素
      const notification = document.createElement("div");
      notification.className = `notification ${type}`;
      notification.textContent = message;

      // 添加到文件中
      document.body.appendChild(notification);

      // 3秒後移除通知
      setTimeout(() => {
        notification.remove();
      }, 3000);
    } catch (error) {
      console.error("顯示通知時發生錯誤：", error);
    }
  }

  // 清理資源
  dispose() {
    if (this.fileNameCheckInterval) {
      clearInterval(this.fileNameCheckInterval);
    }
    if (this.changeTimeout) {
      clearTimeout(this.changeTimeout);
    }
  }
}

// 當 Office 準備就緒時初始化
Office.onReady(() => {
  const taskPane = new TaskPane();
  taskPane.initialize();

  // 在窗口關閉時清理資源
  window.addEventListener("unload", () => {
    taskPane.dispose();
  });
});
