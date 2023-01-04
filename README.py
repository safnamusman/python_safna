# python_safna
import xlwt
from xlwt import Workbook

wb = Workbook(encoding="utf-8")


def main():
    sheet = wb.add_sheet('Sheet 1')
    rows = ['Customer Name', 'Mobile Number', 'CPR number', 'Email Address', 'Address line 1', 'Address Line 2',
            'City', 'Block No./Zip Code', 'Country']
    for row in rows:
        sheet.write(0, rows.index(row), row)
    number_of_entries = input("How many rows of data to insert: ")
    if number_of_entries.isnumeric():
        number_of_entries = int(number_of_entries)
        for i in range(1, number_of_entries + 1):
            print("ROW {}:".format(i))
            for row in rows:
                data = input("Enter {}:".format(row))
                sheet.write(i, rows.index(row), data)
    wb.save('write_to_excel.xls')


if __name__ == "__main__":
    main()

