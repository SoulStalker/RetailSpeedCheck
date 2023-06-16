from datetime import datetime, timedelta

from analyzer import DataAnalyzer
from excel import ExcelExporter


def analyze(operation_day, all_week=None, all_month=None, shop_num=None):
    # get_results можно запустить с фильтром по магазину например shop_index=23
    if all_week:
        day_count = 7
        exporter = ExcelExporter(f'{operation_day.month}-{operation_day.year}.xlsx')
        exporter.create_workbook()
        for single_date in (operation_day + timedelta(n) for n in range(day_count)):
            analyzer = DataAnalyzer(datetime.strftime(single_date, '%Y-%m-%d'))
            analyzer.get_results(shop_num)
            analyzer.calculate_cashier_data()
            analyzer.generate_summary_data()
            analyzer.export_summary_to_excel(exporter)
            print(f'Analyze for day {operation_day} is done')

    else:
        exporter = ExcelExporter(f'{datetime.strftime(operation_day, "%Y-%m-%d")}.xlsx')
        exporter.create_workbook()
        analyzer = DataAnalyzer(datetime.strftime(operation_day, '%Y-%m-%d'))
        analyzer.get_results(shop_num)
        analyzer.calculate_cashier_data()
        analyzer.generate_summary_data()
        analyzer.export_summary_to_excel(exporter)
        exporter.save_workbook()


analyze_day = '2023-06-01'

if __name__ == '__main__':
    analyze(datetime.strptime(analyze_day, '%Y-%m-%d'), shop_num=3)




