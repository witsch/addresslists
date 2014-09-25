from argparse import ArgumentParser
from os import environ, listdir, path, stat
from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.orm import sessionmaker, scoped_session
from xlwt import Workbook


def sources():
    home = environ['HOME']
    sources = '%s/Library/Application Support/AddressBook/Sources' % home
    filename = 'AddressBook-v22.abcddb'
    for name in listdir(sources):
        fullname = path.join(sources, name, filename)
        if path.exists(fullname):
            yield fullname


def addressbook():
    books = sorted((stat(name).st_mtime, name) for name in sources())
    mtime, name = books[-1]
    return name


# database connection
engine = create_engine('sqlite:///%s' % addressbook(), native_datetime=True)
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
        child_info = [fullname(child)]
        parent_info = []
        for parent in q(related).filter_by(ZOWNER=child.Z_PK, ZLABEL=rel('child')):
            parent = people[parent.ZNAME]
            info = [fullname(parent)]
            info += (n for n, in q(phones)
                .filter_by(ZOWNER=parent.Z_PK, ZLABEL=rel('mobile'))
                .with_entities(phones.c.ZFULLNUMBER))
            info += (a for a, in q(mails)
                .filter_by(ZOWNER=parent.Z_PK, ZLABEL=rel('home'))
                .with_entities(mails.c.ZADDRESS))
            parent_info.append(info)
        if parent_info:
            yield child_info, parent_info


def excel(output, children):
    book = Workbook()
    sheet = book.add_sheet('Addresslist')
    row = 0
    for child, parents in children:
        for i, elem in enumerate(child):
            sheet.write(row, i, elem)
        for i, info in enumerate(parents):
            for e, elem in enumerate(info):
                sheet.write(row + i, len(child) + e, elem)
        row += len(parents)
    book.save(output)


def dump(children):
    for child, parents in children:
        print(', '.join(child))
        for info in parents:
            print('\t' + ', '.join(info))


def main():
    args = parse_arguments()
    matches = children(args.filter.decode('utf8'))
    if args.excel:
        with open(args.excel, 'wb') as output:
            excel(output=output, children=matches)
    else:
        dump(children=matches)
