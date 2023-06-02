from excel import ExcelExporter

from database import db
from models import Checks, User, Session, Shift, Position

results = db.query(Checks, Position, Session, User, Shift).select_from(Checks). \
    join(Position, Checks.id == Position.id_purchase).\
    join(Shift, Checks.id_shift == Shift.id).\
    join(Session, Checks.id_session == Session.id).\
    join(User, Session.id_user == User.tabnum).\
    filter(Shift.operday == '2023-05-26', Shift.shopindex == 1).all()

data = []
for check, position, session, user, shift in results:
    row = (
        check.id,
        shift.shopindex,
        shift.cashnum,
        shift.numshift,
        f'{user.lastname} {user.firstname} {user.middlename}',
        check.datecreate,
        check.datecommit,
        check.checksumend,
        shift.operday,
        position.datecommit,

    )
    data.append(row)

ex = ExcelExporter('test.xlsx')
ex.export_to_excel(data)