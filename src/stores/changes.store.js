export class ChangesStore {
  constructor() {
    this.changes = {};
  }

  addChange(id, header, value) {
    if (!this.changes[id]) {
      this.changes[id] = {};
    }
    this.changes[id][header] = value;
  }

  getChanges() {
    return this.changes;
  }

  clear() {
    this.changes = {};
  }
}
