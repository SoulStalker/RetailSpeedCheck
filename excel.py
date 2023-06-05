import openpyxl


class ExcelExporter:
    def __init__(self, filename):
        self.filename = filename
        self.workbook = self.create_workbook()

    def create_workbook(self):
        workbook = openpyxl.Workbook()
        return workbook

    def write_data(self, title, data, sheet):
        ws = self.workbook[sheet]

        ws.append(title)
        for row in sorted(data, key=lambda x: x[1]):
            ws.append(row)

    def save_workbook(self):
        self.workbook.save(self.filename)

    def get_sheet_name(self):
        wss = self.workbook.sheetnames
        return wss[0]

    def export_to_excel(self, title, data):
        sheet = self.get_sheet_name()
        self.write_data(title, data, sheet)
        self.save_workbook()
