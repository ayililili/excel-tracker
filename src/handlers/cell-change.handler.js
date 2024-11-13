export class CellChangeHandler {
  constructor(changesStore) {
    this.changesStore = changesStore;
  }

  async handleCellChange(eventArgs, sheet) {
    const changedCell = eventArgs.address;
    const newValue = eventArgs.details.valueAfter;
    const [column, row] = changedCell.match(/[A-Z]+|\d+/g);

    if (parseInt(row, 10) === 1 || parseInt(column, 10) === 1) {
      return 1;
    }

    const idCell = `A${row}`;
    const headerCell = `${column}1`;

    const idRange = sheet.getRange(idCell);
    const headerRange = sheet.getRange(headerCell);

    idRange.load("values");
    headerRange.load("values");

    await sheet.context.sync();

    const idValue = idRange.values[0][0];
    const headerValue = headerRange.values[0][0];

    this.changesStore.addChange(idValue, headerValue, newValue);
    console.log(`儲存格 ${changedCell}（編號: ${idValue}, 項目: ${headerValue}）改為：${newValue}`);
  }
}
