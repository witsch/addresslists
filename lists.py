from argparse import ArgumentParser
from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.orm import sessionmaker, scoped_session
from xlwt import Workbook


# database connection
engine = create_engine('sqlite:///addressbook.db', native_datetime=True)
session = scoped_session(sessionmaker(bind=engine))
metadata = MetaData()


def get_table(name):
    metadata.reflect(bind=engine, only=[name], views=True)
    return Table(name, metadata, autoload=True)


def parse_arguments():
    parser = ArgumentParser(description='generate address lists')
    parser.add_argument('-x', '--excel', help='output excel spreadsheet')
    parser.add_argument('-f', '--filter', default='Hansa 07',
        help='filter on given relationship')
    return parser.parse_args()


def fullname(record):
    return ' '.join(filter(None, (record.ZFIRSTNAME, record.ZLASTNAME)))


def rel(name):
    return '_$!<%s>!$_' % name.capitalize()


def children(relation):
    records = get_table('ZABCDRECORD')
    related = get_table('ZABCDRELATEDNAME')
    phones = get_table('ZABCDPHONENUMBER')
    mails = get_table('ZABCDEMAILADDRESS')
    q = session.query
    people = {}
    for record in q(records):
        people[record.Z_PK] = people[fullname(record)] = record
    matches = q(related).filter_by(ZNAME=relation)
    for match in matches:
        child = people[match.ZOWNER]
        parent_info = []
        for parent in q(related).filter_by(ZOWNER=child.Z_PK, ZLABEL=rel('child')):
            parent = people[parent.ZNAME]
            info = [fullname(parent)]
            info += (n for n, in q(phones)
                .filter_by(ZOWNER=parent.Z_PK, ZLABEL=rel('mobile'))
                .with_entities(phones.c.ZFULLNUMBER))
            info += (a for a, in q(mails)
                .filter_by(ZOWNER=parent.Z_PK)
                .with_entities(mails.c.ZADDRESS))
            parent_info.append(info)
        if parent_info:
            yield fullname(child), parent_info


def excel(output, children):
    book = Workbook()
    sheet = book.add_sheet('Addresslist')
    row = 0
    for child, parents in children:
        sheet.write(row, 0, child)
        for i, info in enumerate(parents):
            for e, elem in enumerate(info):
                sheet.write(row + i, 1 + e, elem)
        row += len(parents)
    book.save(output)


def dump(children):
    for child, parents in children:
        print(child)
        for info in parents:
            print('\t' + ', '.join(info))


def main():
    args = parse_arguments()
    matches = children(args.filter)
    if args.excel:
        with open(args.excel, 'wb') as output:
            excel(output=output, children=matches)
    else:
        dump(children=matches)
