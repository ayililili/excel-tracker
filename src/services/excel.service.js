export class ExcelService {
  constructor() {
    this.currentSnapshot = null;
    this.documentType = null;
    this.departmentName = null;
    this.projectNumber = null;
    this.serialCounter = 1;
  }

  async determineDocumentType() {
    const fileName = await this.getWorkbookName();

    if (fileName.startsWith("加工件採購")) {
      this.documentType = 1;
    } else if (fileName.startsWith("市購件請購")) {
      this.documentType = 2;
    } else if (fileName.match(/^採購BOM表單_.*_.*/)) {
      this.documentType = 3;
      const matches = fileName.match(/^採購BOM表單_(.*)_(.*)/);
      if (matches) {
        this.departmentName = matches[1];
        this.projectNumber = matches[2];
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

  generateUniqueId() {
    const prefix = `${this.departmentName}_${this.projectNumber}_`;
    const id = `${prefix}${String(this.serialCounter).padStart(4, "0")}`;
    this.serialCounter++;
    return id;
  }

  validateId(id) {
    if (!this.departmentName || !this.projectNumber) return true;
    const pattern = new RegExp(`^${this.departmentName}_${this.projectNumber}_\\d{4}$`);
    return pattern.test(id);
  }

  _getTrackingColumns() {
    switch (this.documentType) {
      case 1:
        return ["A", "B", "C"];
      case 2:
        return ["A", "D", "E"];
      case 3:
        return ["A", "F", "G"];
      default:
        return [];
    }
  }

  async captureSnapshot() {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = worksheet.getUsedRange();
        usedRange.load(["values", "rowCount"]);
        await context.sync();

        const trackingColumns = this._getTrackingColumns();
        const snapshot = {};

        // 跳過標題行，從第二行開始
        for (let row = 1; row < usedRange.rowCount; row++) {
          let id = usedRange.values[row][0]; // A欄位值

          // 處理空白ID的情況
          if (!id && this.documentType === 3) {
            const hasData = trackingColumns
              .slice(1)
              .some((col) => usedRange.values[row][this._columnToIndex(col)] !== "");

            if (hasData) {
              id = this.generateUniqueId();
              // 更新Excel中的ID
              const cell = worksheet.getRange(`A${row + 1}`);
              cell.values = [[id]];
            }
          }

          // 驗證ID格式
          if (id && this.documentType === 3 && !this.validateId(id)) {
            const range = worksheet.getRange(`${row + 1}:${row + 1}`);
            range.format.fill.color = "red";
            continue;
          }

          if (id) {
            snapshot[id] = {
              values: {},
              timestamp: new Date().toISOString(),
            };

            // 記錄追蹤欄位的值
            trackingColumns.forEach((col) => {
              const value = usedRange.values[row][this._columnToIndex(col)];
              snapshot[id].values[col] = value || "";
            });
          }
        }

        await context.sync();
        this.currentSnapshot = snapshot;
      });

      return this.currentSnapshot;
    } catch (error) {
      console.error("捕獲快照時發生錯誤:", error);
      throw error;
    }
  }

  async compareWithSnapshot() {
    try {
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
            const hasData = trackingColumns
              .slice(1)
              .some((col) => usedRange.values[row][this._columnToIndex(col)] !== "");

            if (hasData) {
              id = this.generateUniqueId();
              const cell = worksheet.getRange(`A${row + 1}`);
              cell.values = [[id]];
            }
          }

          if (id && this.documentType === 3 && !this.validateId(id)) {
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

      return changes;
    } catch (error) {
      console.error("比較狀態時發生錯誤:", error);
      throw error;
    }
  }

  _columnToIndex(column) {
    return column.charCodeAt(0) - "A".charCodeAt(0);
  }

  generateChangeReport(changes) {
    const report = [];
    report.push("=== Excel 工作表變更報告 ===");
    report.push(`報告時間: ${new Date().toLocaleString()}`);
    report.push(`文件類型: ${this.documentType}`);
    if (this.documentType === 3) {
      report.push(`部門: ${this.departmentName}`);
      report.push(`專案號: ${this.projectNumber}`);
    }
    report.push("");

    Object.entries(changes).forEach(([id, change]) => {
      report.push(`ID: ${id}`);
      report.push(`時間: ${change.timestamp}`);
      report.push("變更內容:");
      Object.entries(change.values).forEach(([column, value]) => {
        report.push(`  ${column}: ${value}`);
      });
      report.push("");
    });

    return report.join("\n");
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
      return workbookName;
    } catch (error) {
      console.error("無法獲取檔案名：", error);
      throw error;
    }
  }
}
