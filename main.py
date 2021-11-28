from datetime import date
from datetime import datetime as dt
from pandas import read_excel, DataFrame
import os.path
from os import getcwd

def is_leap(year):
    if year % 400 == 0 or year % 40 == 0 or year % 4 == 0:
        return True
    else:
        return False


def minus_result(first_year, second_year):
    month_days = {
        1:31, 
        2:28, 
        3:31, 
        4:30, 
        5:31, 
        6:30, 
        7:31, 
        8:31, 
        9:30, 
        10:31, 
        11:30, 
        12:31
    }
    y = first_year.year - second_year.year
    m = first_year.month - second_year.month
    d = first_year.day - second_year.day
    if d < 0:
        if second_year.month == 2:
            if is_leap(second_year.year):
                month_days[2] = 29
        d += month_days[second_year.month]
        m -= 1
    if m < 0:
        m += 12
        y -= 1
    if y == 0:
        if m == 0:
            return round(d/365, 2)
        else:
            return round((m*30+d)/365, 2)
    else:
        return y


def get_years_old(birth: str):

    birth = birth.replace('/', '.').replace('-', '.').replace(' ', '.').replace('\\', '.')
    
    birthSpilt = birth.split('.')

    y = int(birthSpilt[0]) if int(birthSpilt[0]) > 1000 else int(birthSpilt[0])+1911
    m = int(birthSpilt[1])
    d = int(birthSpilt[2])

    birth = date(
        year=y,
        month=m,
        day=d
    )
    return f'{minus_result(dt.today(), birth)}'


def excel_workflow():
    df = read_excel(
        birthDay,
        usecols= 'A:B'
    )

    result_xlsx = DataFrame(
        {
            '姓名': [df.iat[idx, 0] for idx in range(df.shape[0])],
            '生日': [df.iat[idx, 1] for idx in range(df.shape[0])],
            '年齡': [get_years_old(df.iat[idx, 1]) for idx in range(df.shape[0])],
        }
    )

    for idx in range(len(df)):
        print(f"[>]姓名：{df.iat[idx, 0]}\t生日：{df.iat[idx, 1]}\t年齡：{get_years_old(df.iat[idx, 1])}")

    fileName = f"{birthDay.split('/')[-1].removesuffix('.xlsx') + '-年齡計算'}.xlsx"
    result_xlsx.to_excel(
        fileName,
        sheet_name='年齡計算',
        index=False
    )
    print('[*]' + ''.center(57, '='))
    filePath = f'{getcwd()}/{fileName}'.replace('\\', '/')
    print(f"[*]檔案路徑： {filePath}")
    input('[*]按下「Enter]鍵繼續...')

def clearConsole() -> None:
    command = "clear"
    if os.name in ("nt", "dos"):  # If Machine is running on Windows, use cls
        command = "cls"
    os.system(command)


if __name__ == '__main__':
    while True:
        print('[*]' + '年齡計算小工具'.center(50, '='))
        print('[*]說明：')
        print('[*]\t年份輸入民國或西元年皆可。')
        print('[*]如果要結束程式，直接按下「Enter」即可。')
        print('[*]' + ''.center(57, '='))
        birthDay = input('[?]請輸入出生年月日？ ').replace('"', '').removesuffix(' ')
        print('[*]' + '執行結果'.center(53, '='))

        if birthDay == '':
            break

        if os.path.isfile(birthDay):
            excel_workflow()
        else:
            try:
                print(f'[>]生日：{birthDay.replace(" ", ".")}\n[>]年齡：{get_years_old(birthDay)}')
                input('[*]按下「Enter]鍵繼續...')
            except IndexError:
                print('[!]生日資料輸入不完全！')
        clearConsole()