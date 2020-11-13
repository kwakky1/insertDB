import pymysql
import openpyxl


def main():
    file = openpyxl.load_workbook('SampleData.xlsx')
    print(file.sheetnames)

    sheet = file.active

    total_list = []
    for row in sheet.iter_rows( min_row=2, max_row=sheet.max_row):
        value_list = []
        for cell in row:
            value_list.append(cell.value)

        total_list.append(value_list)

    print(total_list)

    connect = pymysql.connect(host='localhost', user='stest', password='1234', db='python',  charset='utf8')

    cursor = connect.cursor()

    # sql = "CREATE DATABASE python"
    # cursor.execute(sql)

    sql = '''CREATE TABLE purchase (
    id int(10) NOT NULL AUTO_INCREMENT PRIMARY KEY,
    order_date datetime,
    region varchar(10),
    rep varchar(40),
    item varchar(40),
    units int(10),
    unit_cost float(10),
    east_regulation varchar(10),
    total float(20)
    )'''

    cursor.execute(sql)


if __name__ == '__main__':
    main()