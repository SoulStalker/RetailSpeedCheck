from collections import Counter
from datetime import timedelta, datetime

from excel import ExcelExporter

from database import db
from models import Checks, User, Session, Shift, Position

results = db.query(Checks, Position, Session, User, Shift).select_from(Checks). \
    join(Position, Checks.id == Position.id_purchase).\
    join(Shift, Checks.id_shift == Shift.id).\
    join(Session, Checks.id_session == Session.id).\
    join(User, Session.id_user == User.tabnum).\
    filter(Shift.operday == '2023-05-26', Shift.shopindex == 1).all()


detailed = []
for check, position, session, user, shift in results:
    row = (
        check.id,
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

# detailed_excel = ExcelExporter('detailed.xlsx')
# detailed_excel.export_to_excel(detailed)

# print(detailed)
data = set()
for check, position, session, user, shift in results:
    positions_quantity = len(list(filter(lambda item: item[0] == check.id, detailed)))
    row = (
        check.id,
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
    # print(positions_quantity)

data_excel = ExcelExporter('data.xlsx')
data_excel.export_to_excel(data)