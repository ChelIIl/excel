from operator import index
from datetime import datetime
from openpyxl import load_workbook
from calendar import monthrange
import pandas as pnd


def main():
    cols = [1,2,3,4,5,6]
    exc = pnd.read_excel('task_support.xlsx', sheet_name='Tasks', usecols=cols)

    even_num(exc)  #задание 1
    simple_num(exc)  #задание 2
    more_num(exc)  #задание 3
    tue_num(exc)  #задание 4
    tue_num_2(exc)  #задание 5
    last_tue_num(exc)  #задание 6

def even_num(exc):
    res = 0
    series = exc['num1']

    for i in range(len(series) - 1):
        if series[i + 1] % 2 == 0:
            res += 1
    
    print('Четных числе в столбце 1: ', res)

def simple_num(exc):
    res = 0
    series = exc['num2']

    for i in range(len(series) - 1):
        deliter = 2
        count = 0

        while deliter < series[i + 1]:
            if series[i + 1] % deliter == 0:
                count += 1
            deliter += 1
        
        if count == 0:
            res += 1

    print('Простых числе в столбце 2: ', res)

def more_num(exc):
    res = 0
    series = exc['num3']

    for i in range(len(series) - 1):
        series[i + 1] = series[i + 1].replace(' ', '')
        series[i + 1] = series[i + 1].replace(',', '.')
        
        if float(series[i + 1]) < 0.5:
            res += 1

    print('Чисел меньше 0.5 в столбце 3: ', res)

def tue_num(exc):
    res = 0
    series = exc['date1']

    for i in range(len(series) - 1):
        value = series[i + 1]
        month = value[0]+value[1]+value[2]
        
        if month == 'Tue':
            res += 1

    print('Вторников в столбце 4: ', res)

def tue_num_2(exc):
    res = 0
    series = exc['date2']

    for i in range(len(series) - 1):
        dt = datetime.strptime(series[i + 1], '%Y-%m-%d %H:%M:%S.%f')
        d = dt.day
        m = dt.month
        y = dt.year

        if (m == 1 or m == 2):
            y -= 1
        
        m = m - 2
        if m <= 0:
            m += 12

        c = y // 100
        y = y - c * 100
            
        d = (d + ((13 * m - 1) // 5) + y + (y // 4 + c // 4 - 2 * c + 777)) % 7    
        
        if d == 2:
            res += 1

    print('Вторников в столбце 5: ', res)

def last_tue_num(exc):
    res = 0
    series = exc['date3']

    for i in range(len(series) - 1):
        dt = datetime.strptime(series[i + 1], '%m-%d-%Y')
        d = dt.day
        m = dt.month
        y = dt.year

        if (m == 1 or m == 2):
            y -= 1
        
        m = m - 2
        if m <= 0:
            m += 12

        c = y // 100
        y = y - c * 100
            
        dn = (d + ((13 * m - 1) // 5) + y + (y // 4 + c // 4 - 2 * c + 777)) % 7    
        
        if dn == 2:
            last_day = monthrange(y, m)

            if d + 7 > last_day[1]:
                res += 1

    print('Последних вторников в месяце в столбце 6: ', res)

if __name__ == '__main__':
    main()
    