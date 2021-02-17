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


class Unit:
    def __init__(self, names, convert_to_common_VALUE, convert_to_common_UNIT, unit_type):
        self.names = names
        self.convert_to_common_VALUE = convert_to_common_VALUE
        self.convert_to_common_UNIT = convert_to_common_UNIT
        self.unit_type = unit_type


class UnitManager:
    def __init__(self, BASE_DIR):
        self.all_units = list()
        self.excel_reader = ExcelReader(BASE_DIR, file_name='units.xlsx', row_length=4)
        self.load_units()

    def load_units(self):
        data = self.excel_reader.Read()
        for row in data:
            unit = self.unit_from_row(row)
            self.all_units.append(unit)
    
    def unit_from_row(self, row):
        unit = Unit(
            names=row[0].split(', '), 
            convert_to_common_VALUE=row[1], 
            convert_to_common_UNIT=row[2], 
            unit_type=row[3]
            )
        return unit
    
    def find_unit(self, unit_name):
        result = list()
        for unit in self.all_units:
            if unit_name in unit.names:
                result.append(unit)
        if len(result) == 1 or len(result) == 0:
            return result
        else:
            print("Error: more than one unit type exists for this symbol")
            return None

    def exists_unit(self, unit_name):
        if len(self.find_unit(unit_name)) == 1:
            return True
        else:
            return False

    def convert(self, input_VALUE, input_UNIT, output_UNIT):
        input_unit = self.find_unit(input_UNIT)
        output_unit = self.find_unit(output_UNIT)
        
        if output_unit[0].unit_type == input_unit[0].unit_type:
            output_VALUE = input_VALUE * (input_unit[0].convert_to_common_VALUE / output_unit[0].convert_to_common_VALUE)
        else:
            output_VALUE = None
        return output_VALUE


if __name__ == '__main__':
    BASE_DIR = os.getcwd()

    #er = ExcelReader(BASE_DIR, file_name='units.xlsx', row_length=4)
    #data = er.Read()

    um = UnitManager(BASE_DIR)
    print(um.convert(2, 'liter', 'cup'))
