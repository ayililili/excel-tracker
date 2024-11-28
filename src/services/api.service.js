export class ApiService {
  // constructor(baseUrl = "https://api.dt-boost.com/api") {
  constructor(baseUrl = "https://localhost:3002/api") {
    this.baseUrl = baseUrl;
  }

  async fetchData(type) {
    const response = await fetch(`${this.baseUrl}/${type}`);
    if (!response.ok) {
      throw new Error("無法從 API 取得資料");
    }
    return response.json();
  }

  async sendChanges(type, changes) {
    if (!changes || Object.keys(changes).length === 0) {
      console.warn("沒有變更需要上傳");
      return;
    }

    const requestBody = {
      data: changes,
    };

    const response = await fetch(`${this.baseUrl}/${type}`, {
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
