import xlsxwriter as xl
revs = [
    238115,
    128099,
    120521,
    180228,
    316834,
    622805,
    346336,
    379125,
    577627,
    631774,
    771438,
    840402,
    41874,
    100160,
    354105,
    139026,
    531986,
    894465,
    66915,
    754032
]
followers = []

import pandas as pd
file_path = r'F:\数模\比赛\influence net complex.xlsx'
file1 = pd.read_excel(file_path)
for i in revs:
    st = ' ' + str(i)
    wb = xl.Workbook(str(i)+'.xlsx')
    works = wb.add_worksheet()
    col = 0
    lis = 0
    works.write(col, lis, str(i))
    col = 1
    for j in range(len(file1['FOLLOWER_IDS'])):
        if st in file1['FOLLOWER_IDS'][j]:
            fl = []
            for k in file1.keys():
                works.write(col, lis, str(file1[k][j]))
                lis += 1
                fl.append(file1[k][j])
            followers.append(fl)
            col += 1
            lis = 0
    wb.close()