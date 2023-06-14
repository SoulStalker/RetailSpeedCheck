from datetime import datetime, timedelta
from sqlalchemy import and_
from excel import ExcelExporter

from database import db
from models import Checks, User, Session, Shift, Position
from config import se


class DataAnalyzer:
    def __init__(self, operation_day):
        self.operation_day = operation_day
        self.results = None
        self.cashier_data = None
        self.summary_data = None

    def get_results(self, shop_index=None):
        query = db.query(Checks, Position, Session, User, Shift).select_from(Checks). \
            join(Position, Checks.id == Position.id_purchase). \
            join(Shift, Checks.id_shift == Shift.id). \
            join(Session, Checks.id_session == Session.id). \
            join(User, and_(Session.id_user == User.tabnum, User.shop == Shift.shopindex)). \
            filter(Shift.operday == self.operation_day, Checks.checkstatus == 0)

        if shop_index is not None:
            query = query.filter(Shift.shopindex == shop_index)

        self.results = query.order_by(Checks.id).all()

    def calculate_cashier_data(self):
        self.cashier_data = {}

        for check, position, session, user, shift in self.results:
            full_name = user.lastname
            full_name += ' ' + user.firstname[0] + '.' if user.firstname else ''
            full_name += ' ' + user.middlename[0] + '.' if user.middlename else ''
            cashier_key = (user.tabnum, full_name)
            if check.id not in self.cashier_data.get(cashier_key, {}).get('checks', []):
                if cashier_key in self.cashier_data:
                    cashier_info = self.cashier_data[cashier_key]
                    cashier_info['total_check_sum'] += check.checksumend
                    cashier_info['total_check_count'] += 1
                    cashier_info['checks'].append(check.id)
                    cashier_info['check_speed'] += timedelta.total_seconds(check.datecommit - check.datecreate)
                    cashier_info['total_positions'] += position.numberfield
                    cashier_info['position_speed'] += timedelta.total_seconds(position.datecommit - check.datecreate)
                else:
                    cashier_info = {
                        'shop_num': shift.shopindex,
                        'total_check_sum': check.checksumend,
                        'total_check_count': 1,
                        'checks': [check.id],
                        'date': shift.operday,
                        'check_speed': timedelta.total_seconds(check.datecommit - check.datecreate),
                        'total_positions': position.numberfield,
                        'position_speed': timedelta.total_seconds(position.datecommit - check.datecreate)
                    }
                    self.cashier_data[cashier_key] = cashier_info

    def generate_summary_data(self):
        self.summary_data = []

        for cashier_key, cashier_info in self.cashier_data.items():
            shop_number = cashier_info['shop_num']
            cashier = cashier_key[1:]
            date = cashier_info['date']
            worked_hours = 12
            total_check_count = cashier_info['total_check_count']
            total_check_sum = cashier_info['total_check_sum'] / 100
            check_speed = round(cashier_info['check_speed'] / total_check_count, 0)
            average_check = round(total_check_sum / total_check_count, 0)
            positions = cashier_info['total_positions']
            position_speed = round(cashier_info['position_speed'] / positions, 2)

            row = cashier + (shop_number, se[shop_number], date, position_speed, check_speed, total_check_count,
                             worked_hours, total_check_sum, average_check)
            self.summary_data.append(row)

    def export_summary_to_excel(self):
        title = ['Кассир', 'Номер', 'Магазин', 'Дата', 'Средняя скорость позиции', 'Средняя скорость чека',
                 'Количество чеков', 'Отработано часов', 'Оборот руб.', 'Средний чек']

        filename = f'{self.operation_day}.xlsx'
        exporter = ExcelExporter(filename)
        exporter.export_to_excel(title, self.summary_data)


def main():
    operation_day = '2023-05-29'

    analyzer = DataAnalyzer(operation_day)
    analyzer.get_results(shop_index=23)
    analyzer.calculate_cashier_data()
    analyzer.generate_summary_data()
    analyzer.export_summary_to_excel()


if __name__ == '__main__':
    main()
