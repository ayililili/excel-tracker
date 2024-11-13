class ExcelService {
  async getWorkbookName() {
    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name");
        await context.sync();
        return workbook.name;
      });
    } catch (error) {
      console.error("無法獲取檔案名：", error);
      throw error;
    }
  }
}
