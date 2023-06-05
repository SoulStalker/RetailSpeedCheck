from datetime import timedelta, datetime
from sqlalchemy import and_
from excel import ExcelExporter

from database import db
from models import Checks, User, Session, Shift, Position

results = db.query(Checks, Position, Session, User, Shift).select_from(Checks). \
    join(Position, Checks.id == Position.id_purchase). \
    join(Shift, Checks.id_shift == Shift.id). \
    join(Session, Checks.id_session == Session.id). \
    join(User, and_(Session.id_user == User.tabnum, User.shop == Shift.shopindex)). \
    filter(Shift.operday == '2023-05-28', Shift.shopindex == 1).\
    order_by(Checks.id).all()

detailed = []
check_positions = {}
for check, position, session, user, shift in results:
    row = (
        check.id,
        check.numberfield,
        shift.shopindex,
        shift.cashnum,
        shift.numshift,
        f'{user.lastname} {user.firstname} {user.middlename}',
        round(timedelta.total_seconds(check.datecommit - check.datecreate), 0),
        check.checksumend / 100,
        shift.operday,
        datetime.strftime(position.datecommit, '%H:%M:%S'),
        position.qnty / 1000,
        position.priceend / 100,

    )
    detailed.append(row)

    if check.id in check_positions:
        check_positions[check.id] += 1
    else:
        check_positions[check.id] = 1


title = ['Номер чека', 'Номер чека в смене', 'Номер магазина', 'Касса', 'Смена', 'Кассир', 'Скорость чека в сек.', 'Сумма чека в руб.',
         'Дата',
         'Время позиции', 'Количество', 'Цена']
detailed_excel = ExcelExporter('detailed.xlsx')
detailed_excel.export_to_excel(title, detailed)

data = set()
for check, position, session, user, shift in results:
    positions_quantity = check_positions.get(check.id, 0)
    row = (
        check.id,
        check.numberfield,
        shift.shopindex,
        shift.cashnum,
        shift.numshift,
        f'{user.lastname} {user.firstname} {user.middlename}',
        round(timedelta.total_seconds(check.datecommit - check.datecreate), 0),
        check.checksumend / 100,
        shift.operday,
        positions_quantity,
    )
    data.add(row)

title = ['Номер чека', 'Номер чека в смене', 'Номер магазина', 'Касса', 'Смена', 'Кассир', 'Скорость чека в сек.', 'Сумма чека в руб.',
         'Дата',
         'Количество позиций']

data_excel = ExcelExporter('data.xlsx')
data_excel.export_to_excel(title, data)
