from datetime import datetime, timedelta
from sqlalchemy import and_
from excel import ExcelExporter

from database import db
from models import Checks, User, Session, Shift, Position

results = db.query(Checks, Position, Session, User, Shift).select_from(Checks). \
    join(Position, Checks.id == Position.id_purchase). \
    join(Shift, Checks.id_shift == Shift.id). \
    join(Session, Checks.id_session == Session.id). \
    join(User, and_(Session.id_user == User.tabnum, User.shop == Shift.shopindex)). \
    filter(Shift.operday == '2023-05-28', Shift.shopindex == 1, Checks.checkstatus == 0). \
    order_by(Checks.id).all()

cashier_data = {}

for check, position, session, user, shift in results:
    cashier_key = (user.tabnum, f'{user.lastname} {user.firstname[0]}.')
    if check.id not in cashier_data.get(cashier_key, {}).get('checks', []):
        if cashier_key in cashier_data:
            cashier_info = cashier_data[cashier_key]
            cashier_info['total_check_sum'] += check.checksumend
            cashier_info['total_check_count'] += 1
            cashier_info['checks'].append(check.id)
            cashier_info['check_speed'] += timedelta.total_seconds(check.datecommit - check.datecreate)
            cashier_info['total_positions'] += position.numberfield
            cashier_info['position_speed'] += timedelta.total_seconds(position.datecommit - check.datecreate)
        else:
            cashier_info = {
                'total_check_sum': check.checksumend,
                'total_check_count': 1,
                'checks': [check.id],
                'date': shift.operday,
                'check_speed': timedelta.total_seconds(check.datecommit - check.datecreate),
                'total_positions': position.numberfield,
                'position_speed': timedelta.total_seconds(position.datecommit - check.datecreate)
            }
            cashier_data[cashier_key] = cashier_info

summary_data = []

for cashier_key, cashier_info in cashier_data.items():
    cashier = cashier_key[1:]
    date = cashier_info['date']
    total_check_count = cashier_info['total_check_count']
    total_check_sum = cashier_info['total_check_sum'] / 100
    check_speed = round(cashier_info['check_speed'] / total_check_count, 0)
    average_check = round(total_check_sum / total_check_count, 0)
    positions = cashier_info['total_positions']
    position_speed = round(cashier_info['position_speed'] / positions, 2)

    row = cashier + (date, position_speed, check_speed, total_check_count,
                     total_check_sum, average_check)
    summary_data.append(row)

title = ['Кассир', 'Дата', 'Средняя скорость позиции', 'Средняя скорость чека',
         'Количество чеков', 'Оборот руб.', 'Средний чек']
summary_excel = ExcelExporter('summary.xlsx')
summary_excel.export_to_excel(title, summary_data)
