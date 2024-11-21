import { ExcelService } from "../services/excel.service";
import { ApiService } from "../services/api.service";

class TaskPane {
  constructor() {
    this.excelService = new ExcelService();
    this.apiService = new ApiService();
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
    document.getElementById("restart").onclick = () => this.restartWorkbook();
  }

  // async checkForChanges(changes) {
  //   if (!this.isValidDocumentType) {
  //     console.log("檔案類型無效，不進行變更檢查");
  //     return null;
  //   }

  //   try {
  //     const changes = await this.excelService.compareWithSnapshot();
  //     if (changes && Object.keys(changes).length > 0) {
  //       this.changesStore.setChanges(changes);
  //       console.log("變更已記錄:", changes);
  //       return changes;
  //     }
  //     return null;
  //   } catch (error) {
  //     console.error("檢查變更時發生錯誤：", error);
  //     await this.showNotification("檢查變更時發生錯誤", "error");
  //     throw error;
  //   }
  // }

  groupChangesByType(changes) {
    const groupedChanges = {
      1: {}, // 加工件
      2: {}, // 市購件
      3: {}, // 檔案類型是1或2
      4: {}, // 其餘
    };

    Object.entries(changes).forEach(([id, data]) => {
      const type = data.values.partType; // 假設 'type' 欄位是指定的分類依據

      if (type === "1" || type === "2") {
        groupedChanges[3][id] = data;
      } else if (type === "市購件") {
        groupedChanges[2][id] = data;
      } else if (type === "加工件") {
        groupedChanges[1][id] = data;
      } else {
        groupedChanges[4][id] = data;
      }
    });

    return groupedChanges;
  }

  async sendChangesToApi() {
    if (!this.isValidDocumentType) {
      await this.showNotification("檔案類型無效，不進行資料上傳", "warning");
      return;
    }

    try {
      const changes = await this.excelService.compareWithSnapshot();
      if (changes) {
        const groupChanges = this.groupChangesByType(changes);
        if (Object.keys(groupChanges[1]).length > 0) {
          await this.apiService.sendChanges(1, groupChanges[1]);
        }
        if (Object.keys(groupChanges[2]).length > 0) {
          await this.apiService.sendChanges(2, groupChanges[2]);
        }
        if (Object.keys(groupChanges[3]).length > 0) {
          await this.apiService.sendChanges(3, groupChanges[3]);
        }
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
    try {
      // 檢查文件類型是否有效
      if (!this.excelService.documentType || this.excelService.documentType > 3) {
        throw new Error("無效的文件類型");
      }

      try {
        // 使用 fetchData 獲取API數據
        const apiData = await this.apiService.fetchData(this.excelService.documentType);

        // 檢查獲取的數據是否有效
        if (!apiData || typeof apiData !== "object") {
          throw new Error("獲取的數據格式無效");
        }

        // 更新Excel表格
        await this.excelService.updateFromApiData(apiData);

        console.log(`成功同步文件類型 ${this.excelService.documentType} 的數據`);
        return {
          success: true,
          message: "數據同步成功",
        };
      } catch (apiError) {
        console.error("API請求或數據處理錯誤:", apiError);
        throw new Error(`API同步失敗: ${apiError.message}`);
      }
    } catch (error) {
      console.error("同步過程發生錯誤:", error);
      return {
        success: false,
        message: error.message,
      };
    }
  }

  async restartWorkbook() {
    try {
      const { type } = await this.excelService.determineDocumentType();

      if (type >= 1 && type <= 3) {
        this.isValidDocumentType = true;
        await this.excelService.captureSnapshot();
        await this.showNotification("檔案類型有效，已重新捕獲快照", "success");
        console.log("檔案類型有效，已捕獲新快照");

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
        await this.showNotification("檔案名稱格式不符合要求，變動將不予紀錄", "warning");
        console.log("檔案名稱類型無效");
      }
    } catch (error) {
      console.error("重啟時發生錯誤：", error);
      await this.showNotification("重啟時發生錯誤", "error");
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
