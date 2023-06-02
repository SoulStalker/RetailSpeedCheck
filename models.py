from sqlalchemy import Column, Integer, String, DateTime, ForeignKey
from sqlalchemy.orm import declarative_base, relationship


Base = declarative_base()


class Shift(Base):
    __tablename__ = 'od_shift'

    id = Column(Integer, primary_key=True)
    cashnum = Column(Integer)
    numshift = Column(Integer)
    shiftcreate = Column(DateTime)
    shiftclose = Column(DateTime)
    shiftopen = Column(DateTime)
    shopindex = Column(Integer)
    # state = Column(Integer)
    sumcashbegin = Column(Integer)
    operday = Column(DateTime)


class User(Base):
    __tablename__ = 'od_user'

    tabnum = Column(String, primary_key=True)
    shop = Column(Integer, primary_key=True)
    firstname = Column(String)
    lastname = Column(String)
    middlename = Column(String)


class Session(Base):
    __tablename__ = 'od_session'

    id = Column(Integer, primary_key=True)
    cashnum = Column(Integer)
    shopnum = Column(Integer)
    id_user = Column(Integer, ForeignKey('od_user.tabnum'))


class Position(Base):
    __tablename__ = 'od_position'

    id = Column(Integer, primary_key=True)
    datecommit = Column(DateTime)
    priceend = Column(Integer)
    id_purchase = Column(Integer, ForeignKey('od_purchase.id'))

class Checks(Base):
    __tablename__ = 'od_purchase'

    id = Column(Integer, primary_key=True)
    datecreate = Column(DateTime)
    datecommit = Column(DateTime)
    checkstatus = Column(Integer)
    checksumend = Column(Integer)
    operationtype = Column(Integer)
    cash_operation = Column(Integer)
    id_shift = Column(Integer, ForeignKey('od_shift.id'))
    id_session = Column(Integer, ForeignKey('od_session.id'))

