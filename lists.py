from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.orm import sessionmaker, scoped_session


# database connection
engine = create_engine('sqlite:///addressbook.db', native_datetime=True)
session = scoped_session(sessionmaker(bind=engine))
metadata = MetaData()


def get_table(name):
    metadata.reflect(bind=engine, only=[name], views=True)
    return Table(name, metadata, autoload=True)


def fullname(record):
    return ' '.join(filter(None, (record.ZFIRSTNAME, record.ZLASTNAME)))


def rel(name):
    return '_$!<%s>!$_' % name.capitalize()


def main():
    records = get_table('ZABCDRECORD')
    related = get_table('ZABCDRELATEDNAME')
    phones = get_table('ZABCDPHONENUMBER')
    mails = get_table('ZABCDEMAILADDRESS')
    q = session.query
    people = {}
    for record in q(records):
        people[record.Z_PK] = people[fullname(record)] = record
    matches = q(related).filter_by(ZNAME='Hansa 07')
    for match in matches:
        child = people[match.ZOWNER]
        parents = q(related).filter_by(ZOWNER=child.Z_PK, ZLABEL=rel('child'))
        if parents.count():
            print(child.ZFIRSTNAME)
            for parent in parents:
                parent = people[parent.ZNAME]
                info = [fullname(parent)]
                info += (n for n, in q(phones)
                    .filter_by(ZOWNER=parent.Z_PK, ZLABEL=rel('mobile'))
                    .with_entities(phones.c.ZFULLNUMBER))
                info += (a for a, in q(mails)
                    .filter_by(ZOWNER=parent.Z_PK)
                    .with_entities(mails.c.ZADDRESS))
                print('\t' + ', '.join(info))
