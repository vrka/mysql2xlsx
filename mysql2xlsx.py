#!/bin/python3

import mysql.connector
from openpyxl import Workbook
import click


@click.command()
@click.option('-u', '--user', help="username")
@click.option('-p', '--password', help="password")
@click.option('-h', '--hostname', help="MySql server hostnamme")
@click.option('-d', '--database', help="database name")
@click.option('-o', '--output', type=click.Path(writable=True), default='output.xlsx', help="xlsx filename")
@click.argument('sql')
def main(user, password, hostname, database, output, sql):
    """Saves output of SQL command as XLSX file"""

    db = mysql.connector.connect(user=user, password=password, host=hostname, database=database)
    cur = db.cursor()
    cur.execute(sql)

    wb = Workbook()
    ws = wb.worksheets[0]
    ws.append(cur.column_names)

    for row in cur.fetchall():
        ws.append(row)

    wb.save(output)


if __name__ == '__main__':
    main()
