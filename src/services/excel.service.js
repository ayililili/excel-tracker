// 定義表單類型的常量
const DOCUMENT_TYPES = {
  PROCESSING: 1, // 加工
  PURCHASE: 2, // 市購
  DEPARTMENT: 3, // 部門
};

// 欄位映射配置
const COLUMN_MAPPINGS = {
  [DOCUMENT_TYPES.PROCESSING]: {
    modifiable: {
      brand: "D",
      num: "E",
    },
    nonModifiable: {
      id: "A",
      name: "C",
      type: "B",
      isRevoked: "F",
      createdAt: "G",
      modifyAt: "H",
    },
  },
  [DOCUMENT_TYPES.PURCHASE]: {
    modifiable: {
      num: "D",
    },
    nonModifiable: {
      id: "B",
      name: "C",
      type: "A",
      isRevoked: "E",
      createdAt: "F",
      modifyAt: "G",
    },
  },
  [DOCUMENT_TYPES.DEPARTMENT]: {
    modifiable: {
      name: "C",
      type: "B",
      isRevoked: "F",
    },
    nonModifiable: {
      id: "A",
      createdAt: "G",
      modifyAt: "H",
    },
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
    this.serialCounter = 0;
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
      await this.initializeSerialCounterFromHeader();
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
          allowInsertRows: true,
          allowDeleteRows: false,
          allowFormatCells: true,
          allowSort: true,
          allowAutoFilter: true,
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

  async initializeSerialCounterFromHeader() {
    try {
      const idColumn = this._getIdColumn();
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const headerCell = worksheet.getRange(`${idColumn}1`);
        headerCell.load("values");
        await context.sync();

        const headerValue = headerCell.values[0][0];

        // 匹配 ID(****) 格式，其中 **** 為四位數字
        const match = headerValue?.match(/ID\((\d{4})\)/);

        if (match) {
          const serialNum = parseInt(match[1], 10);
          if (!isNaN(serialNum)) {
            this.serialCounter = serialNum; // 設置為下一個序號
            console.log(`序號已初始化為: ${this.serialCounter}`);
          } else {
            console.warn("無法解析序號值");
            this.serialCounter = 0; // 預設值
          }
        } else {
          console.warn(`${idColumn}1 儲存格格式不符合預期 (應為 'ID(####)' 格式)`);
          this.serialCounter = 0; // 預設值
        }
      });
    } catch (error) {
      console.error("從標題初始化序號計數器時發生錯誤:", error);
      this.serialCounter = 0; // 發生錯誤時設置預設值
      throw error;
    }
  }

  generateUniqueId() {
    const prefix = `${this.departmentName}_${this.projectNumber}_`;
    const id = `${prefix}${String(this.serialCounter).padStart(4, "0")}`;
    this.serialCounter++;
    return id;
  }

  async updateHeaderCell() {
    try {
      const idColumn = this._getIdColumn(); // 獲取 ID 欄位
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const headerCell = worksheet.getRange(`${idColumn}1`);

        // 設置新的 ID 值
        const newHeaderValue = `ID(${String(this.serialCounter).padStart(4, "0")})`;
        headerCell.values = [[newHeaderValue]];

        await context.sync();
        console.log(`${idColumn}1 單元格更新為: ${newHeaderValue}`);
      });
    } catch (error) {
      console.error("更新 ID 表頭值時發生錯誤:", error);
      throw error;
    }
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

  // 取得欄位名稱對應
  _getColumnHeaders() {
    const mapping = COLUMN_MAPPINGS[this.documentType];
    if (!mapping) {
      throw new Error(`未知的文件類型: ${this.documentType}`);
    }
    return mapping;
  }

  // 動態獲取 ID 的欄位
  _getIdColumn() {
    const mapping = COLUMN_MAPPINGS[this.documentType];
    if (!mapping || !mapping.nonModifiable.id) {
      throw new Error(`無法獲取 ID 欄位，未知的文件類型或配置: ${this.documentType}`);
    }
    return mapping.nonModifiable.id; // 返回 ID 對應的列名稱
  }

  async captureSnapshot() {
    try {
      const wasProtected = this.worksheetProtected;
      if (wasProtected) {
        await this.unprotectWorksheet();
      }

      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const columnHeaders = this._getColumnHeaders().modifiable;
        const idColumn = this._getIdColumn(); // 動態獲取 ID 欄位

        // 首先獲取最後一行
        const lastRow = worksheet.getUsedRange().getLastRow();
        lastRow.load("rowIndex");
        await context.sync();

        // 構建要讀取的範圍
        const rangeAddresses = Object.values(columnHeaders).map((col) => `${col}1:${col}${lastRow.rowIndex + 1}`);
        rangeAddresses.push(`${idColumn}1:${idColumn}${lastRow.rowIndex + 1}`); // ID列

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
              const cell = worksheet.getRange(`${idColumn}${row + 1}`);
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

            // 使用欄位名稱作為key來儲存值
            Object.entries(columnHeaders).forEach(([key, _], index) => {
              const value = ranges[index].values[row][0];
              snapshot[id].values[key] = value || "";
            });
          }
        }

        if (this.documentType === DOCUMENT_TYPES.DEPARTMENT) {
          await this.updateHeaderCell();
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
        const columnHeaders = this._getColumnHeaders().modifiable;
        const idColumn = this._getIdColumn(); // 動態獲取 ID 欄位

        // 首先獲取最後一行
        const lastRow = worksheet.getUsedRange().getLastRow();
        lastRow.load("rowIndex");
        await context.sync();

        // 構建要讀取的範圍
        const rangeAddresses = Object.values(columnHeaders).map((col) => `${col}1:${col}${lastRow.rowIndex + 1}`);
        rangeAddresses.push(`${idColumn}1:${idColumn}${lastRow.rowIndex + 1}`); // ID列

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
              const cell = worksheet.getRange(`${idColumn}${row + 1}`);
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
                  isSync: false,
                };
              }
            } else {
              // 新增的記錄
              changes[id] = {
                values: currentValues,
                timestamp: new Date().toISOString(),
                isSync: false,
              };
            }
          }
        }

        if (this.documentType === DOCUMENT_TYPES.DEPARTMENT) {
          await this.updateHeaderCell();
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

  async updateFromApiData(apiData) {
    try {
      const wasProtected = this.worksheetProtected;
      if (wasProtected) {
        await this.unprotectWorksheet();
      }

      await Excel.run(async (context) => {
        console.log(apiData);
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const columnHeaders = this._getColumnHeaders();
        const idColumn = this._getIdColumn();

        // 獲取當前使用的範圍
        const usedRange = worksheet.getUsedRange();
        usedRange.load("rowCount");
        await context.sync();
        let lastRowIndex = usedRange.rowCount;

        // 獲取所有ID的值
        const idRange = worksheet.getRange(`${idColumn}2:${idColumn}${usedRange.rowCount}`);
        idRange.load("values");
        await context.sync();

        // 處理每個API數據項
        for (const [id, data] of Object.entries(apiData)) {
          // 在現有數據中查找ID
          const existingRowIndex = idRange.values.findIndex((row) => row[0] === id);

          if (existingRowIndex !== -1) {
            // 更新現有行
            const actualRowIndex = existingRowIndex + 2; // 加2是因為索引從0開始且有標題行

            // 更新可修改的欄位
            for (const [field, value] of Object.entries(data.values)) {
              const column = columnHeaders.modifiable[field] || columnHeaders.nonModifiable[field];
              if (column) {
                const cell = worksheet.getRange(`${column}${actualRowIndex}`);
                cell.values = "123";
              }
            }
          } else {
            // 在最後一行後面添加新行
            lastRowIndex = lastRowIndex + 1;

            // 先設置ID
            const idCell = worksheet.getRange(`${idColumn}${lastRowIndex}`);
            idCell.values = [[id]];

            // 設置其他欄位的值
            for (const [field, value] of Object.entries(data.values)) {
              const column = columnHeaders.modifiable[field] || columnHeaders.nonModifiable[field];
              if (column) {
                const cell = worksheet.getRange(`${column}${lastRowIndex}`);
                cell.values = "123";
              }
            }
          }
        }

        await context.sync();
      });

      // 如果之前是保護狀態，重新啟用保護
      if (wasProtected) {
        await this.protectWorksheet();
      }

      // 更新完成後重新捕獲快照
      await this.captureSnapshot();
    } catch (error) {
      console.error("更新Excel數據時發生錯誤:", error);
      throw error;
    }
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
}
