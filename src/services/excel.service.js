export class ExcelService {
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
