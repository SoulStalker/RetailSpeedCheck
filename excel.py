import openpyxl


class ExcelExporter:
    def __init__(self, filename):
        self.filename = filename
        self.workbook = self.create_workbook()

    def create_workbook(self):
        workbook = openpyxl.Workbook()
        return workbook

    def write_data(self, data, sheet):
        ws = self.workbook[sheet]
        title = ['Номер чека', 'Номер магазина', 'Касса', 'Смена', 'Кассир', 'Начало', 'Конец', 'Сумма чека', 'Дата',
                 'Время позиции']
        ws.append(title)
        for row in data:
            ws.append(row)

    def save_workbook(self):
        self.workbook.save(self.filename)

    def get_sheet_name(self):
        wss = self.workbook.sheetnames
        return wss[0]

    def export_to_excel(self, data):
        sheet = self.get_sheet_name()
        self.write_data(data, sheet)
        self.save_workbook()
