class ApiService {
  constructor(baseUrl = "http://localhost:3001") {
    this.baseUrl = baseUrl;
  }

  async fetchData() {
    const response = await fetch(this.baseUrl);
    if (!response.ok) {
      throw new Error("無法從 API 取得資料");
    }
    return response.json();
  }

  async sendChanges(workbookName, changes) {
    const changeEntries = Object.entries(changes);
    if (changeEntries.length === 0) {
      return;
    }

    const requestBody = {
      id: workbookName,
      data: changeEntries.map(([id, items]) => ({
        id,
        items: Object.entries(items).map(([header, value]) => ({ header, value })),
      })),
    };

    const response = await fetch(this.baseUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      throw new Error(`上傳失敗，狀態碼：${response.status}`);
    }
  }
}
