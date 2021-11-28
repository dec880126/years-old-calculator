import datetime

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

# dotdot = datetime.date(
#     year=int(input('请输入你的出生年份：')),
#     month=int(input('请输入你的出生月份：')),
#     day=int(input('请输入你的出生日期：'))
# )

# t = datetime.date(
#     year=int(input('請輸入出團年份：')),
#     month=int(input('請輸入出團月份：')),
#     day=int(input('請輸入出團日期：'))
# )

# birthList = ['64.02.28', '65.12.29', '48.11.17', '71.05.28']



def get_years_old(birth, day):
    
    if '.' in birth:
        birthSpilt = birth.split('.')
    elif '/' in birth:
        birthSpilt = birth.split('/')
    elif '-' in birth:
        birthSpilt = birth.split('-')
    else:
        birthSpilt = birth.split()

    y = int(birthSpilt[0]) if int(birthSpilt[0]) > 1000 else int(birthSpilt[0])+1911
    m = int(birthSpilt[1])
    d = int(birthSpilt[2])

    birth = datetime.date(
        year=y,
        month=m,
        day=d
    )
    return f'[>]年齡：{minus_result(day, birth)}'

# for birth in birthList:
#     print(get_years_old(birth))

if __name__ == '__main__':
    print('[*]' + '年齡計算小工具'.center(50, '='))
    print('[*]說明：')
    print('[*]\t年份輸入民國或西元年皆可。')
    print('[*]' + ''.center(57, '='))

    birthDay = input('[?]請輸入出生年月日？ ')
    
    try:
        print(
            get_years_old(
                birth=birthDay,
                day=datetime.datetime.today()
            )
        )
    except TypeError:
        print('[!]生日格式輸入錯誤！')