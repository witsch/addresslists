# -*- coding: utf-8 -*-

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


def fullname(record, specials=[]):
    name = ' '.join(filter(None, (record.ZFIRSTNAME, record.ZLASTNAME)))
    for special in specials:
        if name.startswith(special):
            name = special
    return name


def records(owner, table, label):
    q = session.query(get_table(table))
    return q.filter_by(ZOWNER=owner.Z_PK, ZLABEL=rel(label))


def addresses(owner):
    for addr in records(owner, 'ZABCDPOSTALADDRESS', 'home'):
        value = '%(ZSTREET)s\n%(ZZIPCODE)s %(ZCITY)s' % vars(addr)
        value = value.replace(u'Straße', 'Str.').replace(u'straße', 'str.')
        yield value


def phonenumbers(owner, label='home'):
    for nr in records(owner, 'ZABCDPHONENUMBER', label):
        number = nr.ZFULLNUMBER.replace('+49 ', '0').replace(u'+49 ', '0')
        before, sep, after = number.partition(' ')
        number = before.replace('030', '') + sep + after.replace(' ', '')
        yield number.strip()


def mailaddresses(owner):
    for item in records(owner, 'ZABCDEMAILADDRESS', 'home'):
        yield item.ZADDRESS


def first(items):
    return ''.join(list(items)[:1])


def rel(name):
    return '_$!<%s>!$_' % name.capitalize()


def children(relation):
    records = get_table('ZABCDRECORD')
    related = get_table('ZABCDRELATEDNAME')
    q = session.query
    people = {}
    for record in q(records):
        people[record.Z_PK] = people[fullname(record)] = record
    matches = q(related).filter_by(ZNAME=relation)
    for match in matches:
        child = people[match.ZOWNER]
        child_info = [
            fullname(child),
            first(addresses(child)),
            first(phonenumbers(child)),
        ]
        parents = q(related).filter_by(ZOWNER=child.Z_PK, ZLABEL=rel('child'))
        if parents.count():
            parent_info = []
            for parent in parents.order_by(related.c.ZNAME):
                parent = people.get(parent.ZNAME)
                if parent is not None:
                    label = q(related).filter_by(ZOWNER=parent.Z_PK, ZNAME=fullname(child)).first()
                    parent_info.append([
                        label.ZLABEL if label is not None else '',
                        fullname(parent, specials=['Mutter', 'Vater']),
                        first(phonenumbers(parent, 'mobile')),
                        first(mailaddresses(parent)),
                    ])
            parent_info = [info[1:] for info in sorted(parent_info, reverse=True)]
            yield child_info, parent_info


def excel(output, children):
    book = Workbook()
    sheet = book.add_sheet('Addresslist')
    for row, (child, parents) in enumerate(sorted(children)):
        for i, elem in enumerate(child):
            sheet.write(row, i, elem)
        for i, info in enumerate(zip(*parents)):
            sheet.write(row, len(child) + i, '\n'.join(info))
    book.save(output)


def dump(children):
    for child, parents in sorted(children):
        print(', '.join(child).encode('utf8'))
        for info in parents:
            print('\t' + ', '.join(info).encode('utf8'))


def main():
    args = parse_arguments()
    matches = children(args.filter.decode('utf8'))
    if args.excel:
        with open(args.excel, 'wb') as output:
            excel(output=output, children=matches)
    else:
        dump(children=matches)
