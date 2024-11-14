export class ChangesStore {
  constructor() {
    this.changes = {
      cellChanges: [],
      formulaChanges: [],
      structuralChanges: {},
      summary: {},
    };
  }

  // 增加儲存格變化
  addCellChange(change) {
    this.changes.cellChanges.push(change);
  }

  // 增加公式變化
  addFormulaChange(change) {
    this.changes.formulaChanges.push(change);
  }

  // 設定結構變化
  setStructuralChanges(structuralChanges) {
    this.changes.structuralChanges = structuralChanges;
  }

  // 設定摘要
  setSummary(summary) {
    this.changes.summary = summary;
  }

  // 一次性設置所有變更
  setChanges(newChanges) {
    this.changes = {
      cellChanges: newChanges.cellChanges || [],
      formulaChanges: newChanges.formulaChanges || [],
      structuralChanges: newChanges.structuralChanges || {},
      summary: newChanges.summary || {},
    };
  }

  // 返回所有變更
  getChanges() {
    return this.changes;
  }

  // 清空所有變更
  clear() {
    this.changes = {
      cellChanges: [],
      formulaChanges: [],
      structuralChanges: {},
      summary: {},
    };
  }
}
