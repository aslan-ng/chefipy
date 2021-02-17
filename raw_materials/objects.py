import xlrd
import os


class ExcelReader:

    def __init__(self, BASE_DIR, file_name, row_length):
        self.database_row_length = row_length
        DATA_DIR = os.path.join(BASE_DIR, file_name)
        self.DATA_DIR = DATA_DIR

    def Read(self):
        data = list()
        loaded_excel_file = self.OpenExcel(self.DATA_DIR)
        for i in range(1,len(loaded_excel_file.col_values(0))):
            ''' includes header, so the range starts from 1 '''
            row = list()
            for j in range(self.database_row_length):
                cell = loaded_excel_file.cell(i,j).value
                row.append(cell)
            #row = self.refine_row(row)
            data.append(row)
        return data

    def OpenExcel(self, path):
        book = xlrd.open_workbook(path)
        loaded_excel_file = book.sheet_by_index(0)
        return loaded_excel_file


if __name__ == '__main__':
    BASE_DIR = os.getcwd()

    er = ExcelReader(BASE_DIR, file_name='raw_materials.xlsx', row_length=6)
    data = er.Read()