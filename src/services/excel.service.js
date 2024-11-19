// 定義表單類型的常量
const DOCUMENT_TYPES = {
  PROCESSING: 1, // 加工
  PURCHASE: 2, // 市購
  DEPARTMENT: 3, // 部門
};

// 欄位映射配置
const COLUMN_MAPPINGS = {
  [DOCUMENT_TYPES.PROCESSING]: {
    name: "C",
    type: "B",
  },
  [DOCUMENT_TYPES.PURCHASE]: {
    name: "C",
    type: "B",
    num: "D",
  },
  [DOCUMENT_TYPES.DEPARTMENT]: {
    name: "C",
    type: "B",
    brand: "D",
  },
};

export class ExcelService {
  constructor() {
    this.workbookName = "";
    this.worksheetProtected = false;
    this.currentSnapshot = null;
    this.documentType = null;
    this.departmentName = null;
    this.projectNumber = null;
    this.serialCounter = 1;
  }

  async determineDocumentType() {
    const fileName = await this.getWorkbookName();
    const wasProtected = await this.checkProtectionStatus();

    if (fileName.startsWith("加工件採購")) {
      this.documentType = 1;
    } else if (fileName.startsWith("市購件請購")) {
      this.documentType = 2;
    } else if (fileName.match(/^採購BOM表單_.*_.*/)) {
      this.documentType = 3;
      const matches = fileName.match(/^採購BOM表單_(.*)_(.*)\.[^\.]+$/);
      if (matches) {
        this.departmentName = matches[1];
        this.projectNumber = matches[2];
      }
      if (!wasProtected) {
        await this.protectWorksheet();
      }
    } else {
      this.documentType = 4;
    }

    return {
      type: this.documentType,
      departmentName: this.departmentName,
      projectNumber: this.projectNumber,
    };
  }

  async getWorkbookName() {
    try {
      let workbookName;
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name");
        await context.sync();
        workbookName = workbook.name;
      });
      this.workbookName = workbookName;
      return workbookName;
    } catch (error) {
      console.error("無法獲取檔案名：", error);
      throw error;
    }
  }

  async checkProtectionStatus() {
    try {
      let isProtected = false;
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load("protection/protected");
        await context.sync();
        isProtected = worksheet.protection.protected;
      });
      this.worksheetProtected = isProtected;
      return isProtected;
    } catch (error) {
      console.error("檢查工作表保護狀態時發生錯誤:", error);
      throw error;
    }
  }

  // 新增：保護工作表的方法
  async protectWorksheet() {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.protection.protect({
          allowInsertRows: false,
          allowDeleteRows: false,
          allowFormatCells: false,
          allowSort: false,
          allowAutoFilter: false,
        });
        await context.sync();
        this.worksheetProtected = true;
      });
    } catch (error) {
      console.error("保護工作表時發生錯誤:", error);
      throw error;
    }
  }

  // 新增：解除工作表保護的方法
  async unprotectWorksheet() {
    if (this.documentType !== 3) return;

    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.protection.unprotect();
        await context.sync();
        this.worksheetProtected = false;
      });
    } catch (error) {
      console.error("解除工作表保護時發生錯誤:", error);
      throw error;
    }
  }

  generateUniqueId() {
    const prefix = `${this.departmentName}_${this.projectNumber}_`;
    const id = `${prefix}${String(this.serialCounter).padStart(4, "0")}`;
    this.serialCounter++;
    return id;
  }

  validateId(id) {
    if (!id || typeof id !== "string") return false;

    // 如果 departmentName 和 projectNumber 都存在，檢測完整格式
    if (this.departmentName && this.projectNumber) {
      const pattern = new RegExp(`^${this.departmentName}_${this.projectNumber}_\\d{4}$`);
      return pattern.test(id);
    }

    // 如果其中之一缺失，則檢測是否符合一般的合法格式（例如 XXX_XXX_0000）
    const generalPattern = /^.+_.+_\d{4}$/;
    return generalPattern.test(id);
  }

  // 取得需要追蹤的欄位
  _getTrackingColumns() {
    const mapping = COLUMN_MAPPINGS[this.documentType];
    if (!mapping) {
      throw new Error(`未知的文件類型: ${this.documentType}`);
    }
    return Object.values(mapping);
  }

  // 取得欄位名稱對應
  _getColumnHeaders() {
    const mapping = COLUMN_MAPPINGS[this.documentType];
    if (!mapping) {
      throw new Error(`未知的文件類型: ${this.documentType}`);
    }
    return mapping;
  }

  async captureSnapshot() {
    try {
      const wasProtected = this.worksheetProtected;
      if (wasProtected) {
        await this.unprotectWorksheet();
      }

      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const columnHeaders = this._getColumnHeaders();

        // 首先獲取最後一行
        const lastRow = worksheet.getUsedRange().getLastRow();
        lastRow.load("rowIndex");
        await context.sync();

        // 構建要讀取的範圍
        const rangeAddresses = Object.values(columnHeaders).map((col) => `${col}1:${col}${lastRow.rowIndex + 1}`);
        rangeAddresses.push(`A1:A${lastRow.rowIndex + 1}`); // ID列

        // 讀取所有需要的範圍
        const ranges = rangeAddresses.map((address) => worksheet.getRange(address).load("values"));

        await context.sync();

        const snapshot = {};
        const idValues = ranges[ranges.length - 1].values; // ID列的值

        // 跳過標題行，從第二行開始
        for (let row = 1; row < idValues.length; row++) {
          let id = idValues[row][0];

          // 處理空白ID的情況
          if (!id && this.documentType === DOCUMENT_TYPES.DEPARTMENT) {
            const hasData = Object.values(columnHeaders).some((_, index) => {
              return ranges[index].values[row][0] !== "";
            });

            if (hasData) {
              id = this.generateUniqueId();
              // 更新Excel中的ID
              const cell = worksheet.getRange(`A${row + 1}`);
              cell.values = [[id]];
            }
          }

          // 驗證ID格式
          const rowRange = worksheet.getRange(`${row + 1}:${row + 1}`);

          if (id && !this.validateId(id)) {
            rowRange.format.fill.color = "red";
          } else if (id && this.validateId(id)) {
            rowRange.format.fill.clear();
          }

          if (id) {
            snapshot[id] = {
              values: {},
              timestamp: new Date().toISOString(),
              isSync: false,
            };

            // 使用欄位名稱作為key來儲存值
            Object.entries(columnHeaders).forEach(([key, col], index) => {
              const value = ranges[index].values[row][0];
              snapshot[id].values[key] = value || "";
            });
          }
        }

        await context.sync();
        this.currentSnapshot = snapshot;
      });

      // 如果之前是保護狀態，重新啟用保護
      if (wasProtected) {
        await this.protectWorksheet();
      }

      return this.currentSnapshot;
    } catch (error) {
      console.error("捕獲快照時發生錯誤:", error);
      throw error;
    }
  }

  async compareWithSnapshot() {
    try {
      const wasProtected = this.worksheetProtected;
      if (wasProtected) {
        await this.unprotectWorksheet();
      }

      const changes = {};

      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const columnHeaders = this._getColumnHeaders();

        // 首先獲取最後一行
        const lastRow = worksheet.getUsedRange().getLastRow();
        lastRow.load("rowIndex");
        await context.sync();

        // 構建要讀取的範圍
        const rangeAddresses = Object.values(columnHeaders).map((col) => `${col}1:${col}${lastRow.rowIndex + 1}`);
        rangeAddresses.push(`A1:A${lastRow.rowIndex + 1}`); // ID列

        // 讀取所有需要的範圍
        const ranges = rangeAddresses.map((address) => worksheet.getRange(address).load("values"));

        await context.sync();

        const idValues = ranges[ranges.length - 1].values; // ID列的值

        // 跳過標題行，從第二行開始
        for (let row = 1; row < idValues.length; row++) {
          let id = idValues[row][0];

          // 處理空白ID的情況
          if (!id && this.documentType === DOCUMENT_TYPES.DEPARTMENT) {
            const hasData = Object.values(columnHeaders).some((_, index) => {
              return ranges[index].values[row][0] !== "";
            });

            if (hasData) {
              id = this.generateUniqueId();
              // 更新Excel中的ID
              const cell = worksheet.getRange(`A${row + 1}`);
              cell.values = [[id]];
            }
          }

          // 驗證ID格式
          const rowRange = worksheet.getRange(`${row + 1}:${row + 1}`);

          if (id && !this.validateId(id)) {
            rowRange.format.fill.color = "red";
            continue;
          } else if (id && this.validateId(id)) {
            rowRange.format.fill.clear();
          }

          if (id) {
            const currentValues = {};

            // 使用欄位名稱作為key來獲取值
            Object.entries(columnHeaders).forEach(([key, col], index) => {
              const value = ranges[index].values[row][0];
              currentValues[key] = value || "";
            });

            // 比較與快照的差異
            if (this.currentSnapshot[id]) {
              const hasChanges = Object.keys(columnHeaders).some(
                (key) => this.currentSnapshot[id].values[key] !== currentValues[key]
              );

              if (hasChanges) {
                changes[id] = {
                  values: currentValues,
                  timestamp: new Date().toISOString(),
                };
              }
            } else {
              // 新增的記錄
              changes[id] = {
                values: currentValues,
                timestamp: new Date().toISOString(),
              };
            }
          }
        }

        await context.sync();
      });

      if (wasProtected) {
        await this.protectWorksheet();
      }

      return changes;
    } catch (error) {
      console.error("比較狀態時發生錯誤:", error);
      throw error;
    }
  }

  _columnToIndex(column) {
    return column.charCodeAt(0) - "A".charCodeAt(0);
  }

  // generateChangeReport(changes) {
  //   const report = [];
  //   report.push("=== Excel 工作表變更報告 ===");
  //   report.push(`報告時間: ${new Date().toLocaleString()}`);
  //   report.push(`文件類型: ${this.documentType}`);
  //   if (this.documentType === 3) {
  //     report.push(`部門: ${this.departmentName}`);
  //     report.push(`專案號: ${this.projectNumber}`);
  //   }
  //   report.push("");

  //   Object.entries(changes).forEach(([id, change]) => {
  //     report.push(`ID: ${id}`);
  //     report.push(`時間: ${change.timestamp}`);
  //     report.push("變更內容:");
  //     Object.entries(change.values).forEach(([column, value]) => {
  //       report.push(`  ${column}: ${value}`);
  //     });
  //     report.push("");
  //   });

  //   return report.join("\n");
  // }
  // 修改 _getTrackingColumns 方法
}
