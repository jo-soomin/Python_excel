import os
import xlsxwriter
import openpyxl
from .db import item_query_P
# IXYS(미국) 회사 코드 (CMP_cd): 002009
# 공정 코드 ( PRA_cd ): T010
# 시작 날짜( ST_dt ) : 20200828080000
# 종료 날짜( END_dt ) : 20200904080000

# 프로젝트 경로
source_path = 'C:/Users/user/PycharmProjects/IXYS/IXYS(미국)/'


# 고객번호, 공정번호
CMP_cd = '002009'
CMP_NM = 'IXYS(미국)'
PRA_cd = 'T010'
# 시작날짜, 종료날짜
ST_dt = '20200828080000'
END_dt = '20200904080000'
WW = '36'

# 공정기준 아이템 수율현황 쿼리
item_query_P_result = item_query_P


item_query_P_xlsx = xlsxwriter.Workbook(os.path.abspath(source_path+CMP_NM+WW+END_dt+'.xlsx'))
item_query_P_xlsx_sheet = item_query_P_xlsx.add_worksheet('Summary')

num = 0

for a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u in item_query_P_result:
    num = num + 1

    item_query_P_xlsx_sheet.write(num, 0, a)
    '''
    write : num행 0열에 컬럼 a 작성, cell_format_data 서식 적용
    '''
    item_query_P_xlsx_sheet.write(num, 1, b)
    item_query_P_xlsx_sheet.write(num, 2, c)
    item_query_P_xlsx_sheet.write(num, 3, d)
    if (e == None):
        item_query_P_xlsx_sheet.write(num, 4, ' ')
    else:
        item_query_P_xlsx_sheet.write(num, 4, e)
    item_query_P_xlsx_sheet.write(num, 5, f)
    item_query_P_xlsx_sheet.write(num, 6, g)
    item_query_P_xlsx_sheet.write(num, 7, h)
    item_query_P_xlsx_sheet.write(num, 8, i)
    item_query_P_xlsx_sheet.write(num, 9, j)
    item_query_P_xlsx_sheet.write(num, 10, k)
    item_query_P_xlsx_sheet.write(num, 11, l)
    item_query_P_xlsx_sheet.write(num, 12, m)
    item_query_P_xlsx_sheet.write(num, 13, n)
    item_query_P_xlsx_sheet.write(num, 14, o)
    item_query_P_xlsx_sheet.write(num, 15, p)
    item_query_P_xlsx_sheet.write(num, 16, q)
    item_query_P_xlsx_sheet.write(num, 17, r)
    item_query_P_xlsx_sheet.write(num, 18, s)
    if (t == None):
        item_query_P_xlsx_sheet.write(num, 19, '-')
    else:
        item_query_P_xlsx_sheet.write(num, 19, t)
    item_query_P_xlsx_sheet.write(num, 20, u)
item_query_P_xlsx.close()