import openpyxl
import os


class Excel:
    def __init__(self, folder_path, output_filename):
        self.folder_path = folder_path
        self.output_filename = output_filename
        self.files = os.listdir(self.folder_path)
        self.workbook = openpyxl.Workbook()

    def combine_sheets(self):
        for file in self.files:
            if file.endswith('.xlsx'):
                wb = openpyxl.load_workbook(os.path.join(self.folder_path, file))
                for sheet in wb:
                    self.workbook.create_sheet(title=sheet.title, index=sheet.index)
                    for row in sheet.rows:
                        for cell in row:
                            self.workbook[sheet.title][cell.column][cell.row].value = cell.value
        self.workbook.save(self.output_filename)

    def update_values(self, key, new_value):
        sheets = self.workbook.sheetnames
        for sheet in sheets:
            current_sheet = self.workbook[sheet]
            for row in current_sheet.iter_rows():
                cell_value = row[0].value
                if cell_value == key:
                    row[1].value = new_value
        self.workbook.save(self.output_filename)
