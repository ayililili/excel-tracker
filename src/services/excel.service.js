export class ExcelService {
  constructor() {
    this.initialSnapshot = null;
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

  // 捕獲工作表快照
  async captureSnapshot() {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = worksheet.getUsedRange();

        usedRange.load(["values", "rowCount", "columnCount", "formulas"]);

        await context.sync();

        this.initialSnapshot = {
          values: usedRange.values,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount,
          formulas: usedRange.formulas,
          timestamp: new Date(),
        };
      });

      console.log("初始狀態已記錄");
      return this.initialSnapshot;
    } catch (error) {
      console.error("捕獲快照時發生錯誤:", error);
      throw error;
    }
  }

  // 比較當前狀態與初始快照
  async compareWithSnapshot() {
    if (!this.initialSnapshot) {
      throw new Error("尚未記錄初始狀態，請先執行 captureSnapshot()");
    }

    try {
      let changes;
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = worksheet.getUsedRange();

        usedRange.load(["values", "rowCount", "columnCount", "formulas"]);

        await context.sync();

        const currentState = {
          values: usedRange.values,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount,
          formulas: usedRange.formulas,
        };

        changes = this._calculateChanges(this.initialSnapshot, currentState);
      });

      return changes;
    } catch (error) {
      console.error("比較狀態時發生錯誤:", error);
      throw error;
    }
  }

  // 私有方法：計算變更
  _calculateChanges(oldSnapshot, newSnapshot) {
    const changes = {
      structuralChanges: {
        rowCountDiff: newSnapshot.rowCount - oldSnapshot.rowCount,
        columnCountDiff: newSnapshot.columnCount - oldSnapshot.columnCount,
      },
      cellChanges: [],
      formulaChanges: [],
      summary: {
        totalChanges: 0,
        addedRows: 0,
        deletedRows: 0,
        addedColumns: 0,
        deletedColumns: 0,
        modifiedCells: 0,
        modifiedFormulas: 0,
      },
    };

    // 更新結構變化摘要
    if (changes.structuralChanges.rowCountDiff > 0) {
      changes.summary.addedRows = changes.structuralChanges.rowCountDiff;
    } else if (changes.structuralChanges.rowCountDiff < 0) {
      changes.summary.deletedRows = Math.abs(changes.structuralChanges.rowCountDiff);
    }

    if (changes.structuralChanges.columnCountDiff > 0) {
      changes.summary.addedColumns = changes.structuralChanges.columnCountDiff;
    } else if (changes.structuralChanges.columnCountDiff < 0) {
      changes.summary.deletedColumns = Math.abs(changes.structuralChanges.columnCountDiff);
    }

    // 檢查儲存格變化
    const minRows = Math.min(oldSnapshot.rowCount, newSnapshot.rowCount);
    const minCols = Math.min(oldSnapshot.columnCount, newSnapshot.columnCount);

    for (let i = 0; i < minRows; i++) {
      for (let j = 0; j < minCols; j++) {
        // 檢查值變化
        if (oldSnapshot.values[i][j] !== newSnapshot.values[i][j]) {
          changes.cellChanges.push({
            row: i + 1,
            column: j + 1,
            oldValue: oldSnapshot.values[i][j],
            newValue: newSnapshot.values[i][j],
            cellAddress: this._getCellAddress(j + 1, i + 1),
          });
          changes.summary.modifiedCells++;
        }

        // 檢查公式變化
        if (oldSnapshot.formulas[i][j] !== newSnapshot.formulas[i][j]) {
          changes.formulaChanges.push({
            row: i + 1,
            column: j + 1,
            oldFormula: oldSnapshot.formulas[i][j],
            newFormula: newSnapshot.formulas[i][j],
            cellAddress: this._getCellAddress(j + 1, i + 1),
          });
          changes.summary.modifiedFormulas++;
        }
      }
    }

    // 計算總變化數
    changes.summary.totalChanges =
      changes.summary.modifiedCells +
      changes.summary.modifiedFormulas +
      changes.summary.addedRows +
      changes.summary.deletedRows +
      changes.summary.addedColumns +
      changes.summary.deletedColumns;

    return changes;
  }

  // 私有方法：轉換儲存格地址
  _getCellAddress(column, row) {
    let columnName = "";
    while (column > 0) {
      let modulo = (column - 1) % 26;
      columnName = String.fromCharCode(65 + modulo) + columnName;
      column = Math.floor((column - modulo) / 26);
    }
    return `${columnName}${row}`;
  }

  // 產生變更報告
  generateChangeReport(changes) {
    const report = [];

    report.push("=== Excel 工作表變更報告 ===");
    report.push(`報告時間: ${new Date().toLocaleString()}\n`);

    // 結構變化摘要
    report.push("== 結構變化 ==");
    if (changes.structuralChanges.rowCountDiff !== 0) {
      report.push(
        `行數變化: ${changes.structuralChanges.rowCountDiff > 0 ? "新增" : "刪除"} ${Math.abs(changes.structuralChanges.rowCountDiff)} 行`
      );
    }
    if (changes.structuralChanges.columnCountDiff !== 0) {
      report.push(
        `列數變化: ${changes.structuralChanges.columnCountDiff > 0 ? "新增" : "刪除"} ${Math.abs(changes.structuralChanges.columnCountDiff)} 列`
      );
    }
    report.push("");

    // 儲存格值變化
    if (changes.cellChanges.length > 0) {
      report.push("== 儲存格值變化 ==");
      changes.cellChanges.forEach((change) => {
        report.push(`儲存格 ${change.cellAddress}: ${change.oldValue} → ${change.newValue}`);
      });
      report.push("");
    }

    // 公式變化
    if (changes.formulaChanges.length > 0) {
      report.push("== 公式變化 ==");
      changes.formulaChanges.forEach((change) => {
        report.push(`儲存格 ${change.cellAddress}: ${change.oldFormula} → ${change.newFormula}`);
      });
      report.push("");
    }

    // 變更摘要
    report.push("== 變更摘要 ==");
    report.push(`總變更數: ${changes.summary.totalChanges}`);
    report.push(`修改的儲存格數: ${changes.summary.modifiedCells}`);
    report.push(`修改的公式數: ${changes.summary.modifiedFormulas}`);

    return report.join("\n");
  }
}
