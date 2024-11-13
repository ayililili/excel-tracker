import { ExcelService } from "../services/excel.service";
import { ApiService } from "../services/api.service";
import { ChangesStore } from "../stores/changes.store";
import { CellChangeHandler } from "../handlers/cell-change.handler";

class TaskPane {
  constructor() {
    this.excelService = new ExcelService();
    this.apiService = new ApiService();
    this.changesStore = new ChangesStore();
    this.cellChangeHandler = new CellChangeHandler(this.changesStore);
    this.workbookName = "";
  }

  async initialize() {
    if (Office.HostType.Excel) {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";

      await this.setupWorkbook();
      this.setupEventListeners();
    }
  }

  async setupWorkbook() {
    this.workbookName = await this.excelService.getWorkbookName();
    await this.monitorCellChanges();
  }

  setupEventListeners() {
    document.getElementById("run").onclick = () => this.sendChangesToApi();
    document.getElementById("sync").onclick = () => this.syncTableWithApi();
  }

  async monitorCellChanges() {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.onChanged.add((eventArgs) => this.cellChangeHandler.handleCellChange(eventArgs, sheet));
        await context.sync();
      });
    } catch (error) {
      console.error("監聽儲存格變化錯誤：", error);
    }
  }

  async sendChangesToApi() {
    try {
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
      const workbookData = data[this.workbookName];

      if (!workbookData) {
        console.log("沒有匹配的資料來自 API");
        return;
      }

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Clear existing highlights
        const clearRange = sheet.getRange("A2:Z1000");
        clearRange.format.fill.clear();

        // Load ranges
        const colRange = sheet.getRange("A2:A1000");
        const rowRange = sheet.getRange("B1:Z1");
        colRange.load("values");
        rowRange.load("values");
        await context.sync();

        // Update cells
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
