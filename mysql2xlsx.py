#!/bin/python3

from copy import copy

import click
import mysql.connector
from openpyxl import Workbook, load_workbook, utils, workbook


@click.command()
@click.option('-u', '--user', help="username")
@click.option('-p', '--password', help="password")
@click.option('-h', '--hostname', help="MySql server hostnamme")
@click.option('-d', '--database', help="database name")
@click.option('-o', '--output', type=click.Path(writable=True), default='output.xlsx', help="xlsx filename")
@click.option('-t', '--template', type=click.Path(exists=True, readable=True), help="xlsx template filename")
@click.argument('sql')
def main(user, password, hostname, database, output, template, sql):
    """Saves output of SQL command as XLSX file, optionally formatted as template file"""

    db = mysql.connector.connect(user=user, password=password, host=hostname, database=database)
    cur = db.cursor()
    cur.execute(sql)

    if (template):
        wb = load_workbook(template)
        ws = wb.worksheets[0]
        name2index = {val: idx for idx, val in enumerate(cur.column_names)}
        name2col = {}

        row_idx = 3

        for row in cur.fetchall():
            ws.insert_rows(row_idx)
            for col_idx in range(1, ws.max_column + 1):
                # copy format from row 2
                src = ws.cell(2, col_idx)
                dst = ws.cell(row_idx, col_idx)
                dst._style = copy(src._style)
                if src.value[0] == "_":
                    dst.value = row[name2index[src.value[1:]]]
                    name2col[src.value[1:]] = col_idx
            row_idx += 1
        rules = list(ws.conditional_formatting._cf_rules)
        for cf_rule in rules:
            rng = copy(cf_rule.cells.ranges[0])
            rng.max_row = row_idx - 1
            for rule in ws.conditional_formatting[cf_rule]:
                ws.conditional_formatting.add(str(rng), rule)
        # process last row (replace range in formulas)
        for name in name2col:
            col_name = utils.cell.get_column_letter(name2col[name])
            new_range = workbook.defined_name.DefinedName('data_' + name, localSheetId=wb.sheetnames.index(ws.title),
                                                          attr_text=f"${col_name}$2:${col_name}${row_idx - 2}")
            wb.defined_names.append(new_range)
        ws.delete_rows(2)
    else:
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.append(cur.column_names)

        for row in cur.fetchall():
            ws.append(row)

    wb.save(output)


if __name__ == '__main__':
    main()
