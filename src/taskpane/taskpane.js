import { ExcelService } from "../services/excel.service";
import { ApiService } from "../services/api.service";
import { ChangesStore } from "../stores/changes.store";

class TaskPane {
  constructor() {
    this.excelService = new ExcelService();
    this.apiService = new ApiService();
    this.changesStore = new ChangesStore();
    this.workbookName = "";
    this.isValidDocumentType = false;
  }

  async initialize() {
    if (Office.context.host === Office.HostType.Excel) {
      await this.setupWorkbook();
      this.setupEventListeners();
      this.updateUIState();
    }
  }

  async setupWorkbook() {
    try {
      const { type } = await this.excelService.determineDocumentType();
      this.workbookName = await this.excelService.getWorkbookName();

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
    let banner = document.getElementById("invalid-type-banner");
    if (!banner) {
      banner = document.createElement("div");
      banner.id = "invalid-type-banner";
      banner.className = "warning-banner";

      const content = document.createElement("div");
      content.innerHTML = `
        <strong>⚠️ 檔案名稱格式不符合要求</strong>
        <p>目前檔案: ${this.workbookName}</p>
        <p>變動將不予紀錄，請確認檔案名稱格式是否正確。</p>
      `;

      banner.appendChild(content);
      document.getElementById("app-body").insertBefore(banner, document.getElementById("app-body").firstChild);
    } else {
      banner.querySelector("p").textContent = `目前檔案: ${this.workbookName}`;
    }
  }

  updateUIState() {
    const pushBtn = document.getElementById("push");
    const pullBtn = document.getElementById("pull");

    [pushBtn, pullBtn].forEach((btn) => {
      if (this.isValidDocumentType) {
        btn.disabled = false;
        btn.classList.remove("disabled");
      } else {
        btn.disabled = true;
        btn.classList.add("disabled");
      }
    });
  }

  setupEventListeners() {
    document.getElementById("push").onclick = () => this.sendChangesToApi();
    document.getElementById("pull").onclick = () => this.syncTableWithApi();
    document.getElementById("restart").onclick = () => this.handleFileNameChange();
  }

  async handleFileNameChange() {
    try {
      const { type } = await this.excelService.determineDocumentType();
      const newWorkbookName = await this.excelService.getWorkbookName();

      if (newWorkbookName === this.workbookName) {
        await this.showNotification("檔案名稱沒有變更", "info");
        return;
      }

      this.workbookName = newWorkbookName;

      if (type >= 1 && type <= 3) {
        this.isValidDocumentType = true;
        await this.excelService.captureSnapshot();
        await this.showNotification("檔案名稱已更新，已重新捕獲快照", "success");

        const banner = document.getElementById("invalid-type-banner");
        if (banner) banner.remove();
      } else {
        this.isValidDocumentType = false;
        this.showInvalidFileTypeBanner();
        await this.showNotification("新檔案名稱格式不符合要求，變動將不予紀錄", "warning");
      }

      this.updateUIState();
    } catch (error) {
      console.error("處理檔名變更時發生錯誤：", error);
      await this.showNotification("處理檔名變更時發生錯誤", "error");
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
        await this.apiService.sendChanges(this.workbookName, changes);
        this.changesStore.clear();
        await this.showNotification("數據已成功上傳", "success");

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
      const workbookData = data[this.workbookName];

      if (!workbookData) {
        await this.showNotification("沒有可同步的資料", "info");
        return;
      }

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        const clearRange = sheet.getRange("A2:Z1000");
        clearRange.format.fill.clear();

        const dataRange = sheet.getRange("A2:Z1000");
        const headerRange = sheet.getRange("A1:Z1");
        dataRange.load("values");
        headerRange.load("values");
        await context.sync();

        for (const item of workbookData) {
          for (let row = 0; row < dataRange.values.length; row++) {
            if (dataRange.values[row][0] === item.id) {
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
    const notification = document.createElement("div");
    notification.className = `notification ${type}`;
    notification.textContent = message;

    document.body.appendChild(notification);

    setTimeout(() => {
      notification.remove();
    }, 3000);
  }

  dispose() {
    if (this.fileNameCheckInterval) {
      clearInterval(this.fileNameCheckInterval);
    }
  }
}

Office.onReady(() => {
  const taskPane = new TaskPane();
  taskPane.initialize();

  window.addEventListener("unload", () => {
    taskPane.dispose();
  });
});
