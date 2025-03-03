import pandas as pd
import xlwings as xw
import os
from config.settings import Settings

class ExcelService:
    def __init__(self):
        self.settings = Settings.EXCEL_FILES

    def get_data(self, book_key):
        """Excelデータ取得"""
        file_info = self.settings.get(book_key)
        if not file_info:
            raise ValueError(f"Unknown book key: {book_key}")
            
        return pd.read_excel(
            file_info['path'],
            sheet_name=file_info['sheet']
        )

    def update_book(self, file_path):
        """Excelブック更新"""
        try:
            app = xw.App(visible=True)
            wb = app.books.open(file_path)
            wb.api.RefreshAll()
            app.api.CalculateUntilAsyncQueriesDone()
            wb.save()
            wb.close()
            app.quit()
        except Exception as e:
            print(f"エラーが発生しました: {e}")
            os.system("taskkill /f /im EXCEL.EXE")
