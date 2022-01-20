import xlwings as xw
import project_conf as conf


class UpdateExcel:
    def __init__(self):
        self.file_path = conf.filepath
        self.sheet_name = conf.sheet_name

    def update_sheet(self, value, row_idx=2, col_idx=1):
        cell = chr(64 + col_idx) + str(row_idx)
        wbxl = xw.Book(self.file_path)
        wbxl.sheets[self.sheet_name].range(cell).value = value

    def get_result(self, row_idx=2, col_idx=3):
        cell = chr(64 + col_idx) + str(row_idx)
        wbxl = xw.Book(self.file_path)
        print(wbxl.sheets[self.sheet_name].range(cell).value)

updExl = UpdateExcel()
updExl.update_sheet(45)
updExl.get_result()