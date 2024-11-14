export class CellChangeHandler {
  constructor(changesStore) {
    this.changesStore = changesStore;
  }

  async handleCellChange(eventArgs, sheet) {
    const changes = eventArgs.details.valueChanges;
    console.log(eventArgs);
    if (!changes || changes.length === 0) {
      return;
    }

    // 開始 Excel context 操作
    await Excel.run(async (context) => {
      for (const change of changes) {
        const changedCell = change.address;
        const newValue = change.valueAfter;

        const [column, row] = changedCell.match(/[A-Z]+|\d+/g);

        // 跳過首行或首列變動
        if (parseInt(row, 10) === 1 || parseInt(column, 10) === 1) {
          continue;
        }

        const idCell = `A${row}`;
        const headerCell = `${column}1`;

        const idRange = sheet.getRange(idCell);
        const headerRange = sheet.getRange(headerCell);

        idRange.load("values");
        headerRange.load("values");

        await context.sync();

        const idValue = idRange.values[0][0];
        const headerValue = headerRange.values[0][0];

        this.changesStore.addChange(idValue, headerValue, newValue);
        console.log(`儲存格 ${changedCell}（編號: ${idValue}, 項目: ${headerValue}）改為：${newValue}`);
      }
    });
  }
}
