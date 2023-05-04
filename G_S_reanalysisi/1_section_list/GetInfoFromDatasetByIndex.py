#这个脚本用于将“母表”中的指定列信息抽取出来
#基于“索引表”中第一列中的关键词，把“母表”中第一行标题对应的列抽出来
#与“母表”第一列一起重新写入一个新表
import xlwt
import xlrd
import sys

path_otus = sys.argv[1]
path_index = sys.argv[2]
path_res = sys.argv[3]
otus = xlrd.open_workbook(path_otus)#母表
otusheet = otus.sheet_by_index(0)
ncols = otusheet.ncols
nrows = otusheet.nrows
print('outsheet_open done!')
q_key = xlrd.open_workbook(path_index)#索引表
q_sheet = q_key.sheet_by_index(0)
qrows = q_sheet.nrows
print('querysheet_open done!')
row0 = []
row0_info = q_sheet.col_values(0, 0, qrows)
for r in row0_info:
    row0.append(r)
print(row0)
f = xlwt.Workbook()
f_sheet: object = f.add_sheet('sheet1', cell_overwrite_ok=True)
for i in range(0, len(row0)):
    f_sheet.write(0, i+1, row0[i])

col0 = otusheet.col_values(0, 1, nrows)
#print(col0)
n0 = 1
for c0 in col0:
    f_sheet.write(n0, 0, c0)
    n0 = n0 + 1

s_db = otusheet.row_values(0, 0, ncols)
#print(s_db)
for r0 in row0:
    if r0 in s_db:
        n = 1
        c_n = row0.index(r0) + 1
        sc_n = s_db.index(r0)
        col_v = otusheet.col_values(sc_n, 1, nrows)
        print(col_v)
        for q in col_v:
            f_sheet.write(n, c_n, q)
            n = n + 1
        print(str(r0) + 'done!')


f.save(path_res)#结果表的绝对路径
