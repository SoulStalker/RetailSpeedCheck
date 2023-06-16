from datetime import datetime, timedelta

from analyzer import DataAnalyzer
from excel import ExcelExporter


def put_to_excel(current_date, shop_num, exporter):
    """
    Выгрузка в Excel.
    :param current_date: дата анализа, будем еще именем вкладки.
    :param shop_num: номер магазина, если нужен фильтр по магазину.
    :param exporter: Класс экспорта из модуля excel.
    :return: None.
    """
    analyzer = DataAnalyzer(datetime.strftime(current_date, '%Y-%m-%d'))
    analyzer.get_results(shop_num)
    analyzer.calculate_cashier_data()
    analyzer.generate_summary_data()
    analyzer.export_summary_to_excel(exporter)


def analyze(operation_day, all_week=None, all_month=None, shop_num=None):
    """
    Основная функция для выборки и передачи данных для выгрузки excel.
    :param operation_day: день или начальный день для выгрузки
    :param all_week: выгружать ли неделю.
    :param all_month: выгружать ли месяц
    :param shop_num: номер магазина, если нужен фильтр по магазину.
    :return: None.
    """
    if all_week:
        day_count = 7
        exporter = ExcelExporter(f'{operation_day.month}-{operation_day.year}.xlsx')
        exporter.create_workbook()
        for single_date in (operation_day + timedelta(n) for n in range(day_count)):
            put_to_excel(single_date, shop_num, exporter)
            print(f'Analyze for day {datetime.strftime(single_date, "%Y-%m-%d")} is done')

    else:
        exporter = ExcelExporter(f'{datetime.strftime(operation_day, "%Y-%m-%d")}.xlsx')
        exporter.create_workbook()
        put_to_excel(operation_day, shop_num, exporter)
        exporter.save_workbook()


analyze_day = '2023-05-01'

if __name__ == '__main__':
    analyze(datetime.strptime(analyze_day, '%Y-%m-%d'), all_week=True, shop_num=1)
    print(f'Analyze period complete')
