import { ExcelService } from "../services/excel.service";
import { ApiService } from "../services/api.service";
import { ChangesStore } from "../stores/changes.store";

class TaskPane {
  constructor() {
    this.excelService = new ExcelService();
    this.apiService = new ApiService();
    this.changesStore = new ChangesStore();
    this.workbookName = "";
  }

  async initialize() {
    if (Office.HostType.Excel) {
      document.getElementById("app-body").style.display = "flex";
      await this.setupWorkbook();
      this.setupEventListeners();
    }
  }

  async setupWorkbook() {
    this.workbookName = await this.excelService.getWorkbookName();
    await this.excelService.captureSnapshot(); // 捕獲初始快照
  }

  setupEventListeners() {
    document.getElementById("push").onclick = () => this.sendChangesToApi();
    document.getElementById("pull").onclick = () => this.syncTableWithApi();
  }

  async checkForChanges() {
    try {
      const changes = await this.excelService.compareWithSnapshot();
      if (changes) {
        this.changesStore.setChanges(changes);
        console.log("變更已記錄:", changes);
      }
    } catch (error) {
      console.error("檢查變更時發生錯誤：", error);
    }
  }

  async sendChangesToApi() {
    try {
      await this.checkForChanges(); // 檢查變更
      await this.apiService.sendChanges(this.workbookName, this.changesStore.getChanges());
      this.changesStore.clear();
      console.log("數據已成功上傳到 API");
    } catch (error) {
      console.error("錯誤：", error);
    }
  }

  async syncTableWithApi() {
    try {
      const data = await this.apiService.fetchData();
      const workbookData = data["Project.xlsx"];
      if (!workbookData) {
        console.log("沒有匹配的資料來自 API");
        return;
      }
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const clearRange = sheet.getRange("A2:Z1000");
        clearRange.format.fill.clear();

        const colRange = sheet.getRange("A2:A1000");
        const rowRange = sheet.getRange("B1:Z1");
        colRange.load("values");
        rowRange.load("values");
        await context.sync();

        for (const item of workbookData) {
          const rowIndex = colRange.values.findIndex((row) => row[0] == item.id);
          if (rowIndex === -1) continue;

          const row = rowIndex + 2;
          for (const field of item.items) {
            const colIndex = rowRange.values[0].findIndex((col) => col == field.header);
            if (colIndex === -1) continue;

            const col = String.fromCharCode(66 + colIndex);
            const cell = sheet.getRange(`${col}${row}`);
            cell.values = [[field.value]];
            cell.format.fill.color = "yellow";
          }
        }
        await context.sync();
      });
    } catch (error) {
      console.error("同步表格資料時發生錯誤：", error);
    }
  }
}

// Initialize when Office is ready
Office.onReady(() => {
  const taskPane = new TaskPane();
  taskPane.initialize();
});
