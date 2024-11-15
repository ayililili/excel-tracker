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
    if (Office.HostType.Excel) {
      document.getElementById("app-body").style.display = "flex";
      await this.setupWorkbook();
      this.setupEventListeners();
    }
  }

  async setupWorkbook() {
    try {
      // 檢測檔案類型
      const { type } = await this.excelService.determineDocumentType();
      this.workbookName = await this.excelService.getWorkbookName();

      if (type >= 1 && type <= 3) {
        this.isValidDocumentType = true;
        await this.excelService.captureSnapshot(); // 捕獲初始快照
        console.log("檔案類型有效，已捕獲初始快照");
      } else {
        this.isValidDocumentType = false;
        await this.showErrorMessage("檔案名稱格式不符合要求，變動將不予紀錄");
        console.log("檔案類型無效");
      }

      // 監聽檔案名稱變更
      await this.setupFileNameChangeListener();
    } catch (error) {
      console.error("設置工作簿時發生錯誤：", error);
      await this.showErrorMessage("設置工作簿時發生錯誤");
    }
  }

  async setupFileNameChangeListener() {
    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.onNameChanged.add(async () => {
          await this.handleFileNameChange();
        });
        await context.sync();
      });
    } catch (error) {
      console.error("設置檔名變更監聽器時發生錯誤：", error);
    }
  }

  async handleFileNameChange() {
    try {
      const { type } = await this.excelService.determineDocumentType();
      this.workbookName = await this.excelService.getWorkbookName();

      if (type >= 1 && type <= 3) {
        this.isValidDocumentType = true;
        await this.excelService.captureSnapshot();
        console.log("檔案名稱已更新，類型有效，已捕獲新快照");
      } else {
        this.isValidDocumentType = false;
        await this.showErrorMessage("新檔案名稱格式不符合要求，變動將不予紀錄");
        console.log("新檔案名稱類型無效");
      }
    } catch (error) {
      console.error("處理檔名變更時發生錯誤：", error);
    }
  }

  setupEventListeners() {
    // 監聽按鈕點擊
    document.getElementById("push").onclick = () => this.sendChangesToApi();
    document.getElementById("pull").onclick = () => this.syncTableWithApi();

    // 監聽檔案儲存
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
      this.handleDocumentSave();
    });
  }

  async handleDocumentSave() {
    if (this.isValidDocumentType) {
      await this.sendChangesToApi();
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
      throw error;
    }
  }

  async sendChangesToApi() {
    if (!this.isValidDocumentType) {
      await this.showErrorMessage("檔案類型無效，不進行資料上傳");
      return;
    }

    try {
      const changes = await this.checkForChanges();
      if (changes) {
        await this.apiService.sendChanges(this.workbookName, changes);
        this.changesStore.clear();
        console.log("數據已成功上傳到 API");

        // 上傳完成後捕獲新快照
        await this.excelService.captureSnapshot();
        console.log("已捕獲新快照");
      }
    } catch (error) {
      console.error("錯誤：", error);
      await this.showErrorMessage("上傳資料時發生錯誤");
    }
  }

  async syncTableWithApi() {
    if (!this.isValidDocumentType) {
      await this.showErrorMessage("檔案類型無效，不進行資料同步");
      return;
    }

    try {
      const data = await this.apiService.fetchData();
      const workbookData = data[this.workbookName];

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
          const rowIndex = colRange.values.findIndex((row) => row[0] === item.id);
          if (rowIndex === -1) continue;

          const row = rowIndex + 2;
          for (const field of item.items) {
            const colIndex = rowRange.values[0].findIndex((col) => col === field.header);
            if (colIndex === -1) continue;

            const col = String.fromCharCode(66 + colIndex);
            const cell = sheet.getRange(`${col}${row}`);
            cell.values = [[field.value]];
            cell.format.fill.color = "yellow";
          }
        }
        await context.sync();

        // 同步完成後捕獲新快照
        await this.excelService.captureSnapshot();
      });
    } catch (error) {
      console.error("同步表格資料時發生錯誤：", error);
      await this.showErrorMessage("同步資料時發生錯誤");
    }
  }

  async showErrorMessage(message) {
    // 這裡可以根據您的UI框架來實現錯誤訊息顯示
    // 例如使用自定義的彈窗組件或是Office UI Fabric的對話框
    try {
      await Excel.run(async (context) => {
        context.workbook.app.showMessage(message);
        await context.sync();
      });
    } catch (error) {
      console.error("顯示錯誤訊息時發生錯誤：", error);
    }
  }
}

// 當 Office 準備就緒時初始化
Office.onReady(() => {
  const taskPane = new TaskPane();
  taskPane.initialize();
});
