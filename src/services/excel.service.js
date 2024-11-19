// 定義表單類型的常量
const DOCUMENT_TYPES = {
  PROCESSING: 1, // 加工
  PURCHASE: 2, // 市購
  DEPARTMENT: 3, // 部門
};

// 欄位映射配置
const COLUMN_MAPPINGS = {
  [DOCUMENT_TYPES.PROCESSING]: {
    name: "B",
    num: "C",
  },
  [DOCUMENT_TYPES.PURCHASE]: {
    name: "B",
    num: "D",
    brand: "C",
  },
  [DOCUMENT_TYPES.DEPARTMENT]: {
    name: "B",
    num: "C",
    brand: "E",
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

  _getTrackingColumns() {
    switch (this.documentType) {
      case 1:
        return ["D"];
      case 2:
        return ["E"];
      case 3:
        return ["B", "C"];
      default:
        return [];
    }
  }

  async captureSnapshot() {
    try {
      const wasProtected = this.worksheetProtected;
      if (wasProtected) {
        await this.unprotectWorksheet();
      }

      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const trackingColumns = this._getTrackingColumns();

        // 只讀取需要的欄位範圍
        const ranges = trackingColumns.map((col) => worksheet.getRange(`${col}:${col}`));
        const idRange = worksheet.getRange("A:A");

        // 一次性載入所有需要的資料
        ranges.forEach((range) => range.load("values"));
        idRange.load(["values", "rowCount"]);
        await context.sync();

        const snapshot = {};

        // 跳過標題行，從第二行開始
        for (let row = 1; row < idRange.rowCount; row++) {
          let id = idRange.values[row][0];

          // 處理空白ID的情況
          if (!id && this.documentType === 3) {
            const hasData = trackingColumns.some((col, index) => {
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
            };

            // 記錄追蹤欄位的值
            trackingColumns.forEach((col, index) => {
              const value = ranges[index].values[row][0];
              snapshot[id].values[col] = value || "";
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
        const usedRange = worksheet.getUsedRange();
        usedRange.load(["values", "rowCount"]);
        await context.sync();

        const trackingColumns = this._getTrackingColumns();

        for (let row = 1; row < usedRange.rowCount; row++) {
          let id = usedRange.values[row][0];

          if (!id && this.documentType === 3) {
            const hasData = trackingColumns.some((col) => usedRange.values[row][this._columnToIndex(col)] !== "");

            if (hasData) {
              id = this.generateUniqueId();
              const cell = worksheet.getRange(`A${row + 1}`);
              cell.values = [[id]];
            }
          }

          if (id && !this.validateId(id)) {
            const range = worksheet.getRange(`${row + 1}:${row + 1}`);
            range.format.fill.color = "red";
            continue;
          }

          if (id) {
            const currentValues = {};
            trackingColumns.forEach((col) => {
              currentValues[col] = usedRange.values[row][this._columnToIndex(col)] || "";
            });

            if (this.currentSnapshot[id]) {
              const hasChanges = trackingColumns.some(
                (col) => this.currentSnapshot[id].values[col] !== currentValues[col]
              );

              if (hasChanges) {
                changes[id] = {
                  values: currentValues,
                  timestamp: new Date().toISOString(),
                };
              }
            } else {
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
