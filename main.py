from analyzer import DataAnalyzer


def main():
    operation_day = '2023-06-01'

    analyzer = DataAnalyzer(operation_day)
    analyzer.get_results()
    # get_results можно запустить с фильтром по магазину например shop_index=23
    analyzer.calculate_cashier_data()
    analyzer.generate_summary_data()
    analyzer.export_summary_to_excel()


if __name__ == '__main__':
    main()
