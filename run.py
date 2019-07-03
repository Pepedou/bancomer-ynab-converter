import csv
import datetime
import os

import xlrd


def main():
    os.chdir(os.path.expanduser('~/Desktop'))

    with open('ynab.csv', 'w', encoding='utf8') as f:
        c = csv.writer(f)
        c.writerow(['Date', 'Payee', 'Memo', 'Outflow', 'Inflow'])

        with xlrd.open_workbook('descarga.xlsx') as wb:
            sheet = wb.sheet_by_index(0)

            for i in range(0, sheet.nrows):
                try:
                    my_date = datetime.datetime.strptime(sheet.cell_value(i, 0), '%d/%m/%Y')
                except ValueError:
                    print('Ignoring row {} "{}", because it does not start with a date.'.format(
                        i + 1,
                        sheet.cell_value(i, 0))
                    )
                    continue

                payee = sheet.cell_value(i, 1)
                memo = None
                outflow = str(sheet.cell_value(i, 2)).strip('-')
                inflow = str(sheet.cell_value(i, 3)).strip('-')

                row = [
                    datetime.datetime.strftime(my_date, '%m/%d/%y'),
                    payee,
                    memo,
                    outflow,
                    inflow,
                ]

                c.writerow(row)


if __name__ == '__main__':
    main()
