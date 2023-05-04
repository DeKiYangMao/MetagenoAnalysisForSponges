#基于一份总表，把被注释表中的数据tag对应的注释插入进来
import xlwt
import xlrd
import sys
path_ref = sys.argv[1]
path_in = sys.argv[2]
path_res = sys.argv[3]
s_d = xlrd.open_workbook(path_in)#需要被添加注释信息的表
s_dsheet = s_d.sheet_by_index(0)
nrows = s_dsheet.nrows

f = xlwt.Workbook()
f_sheet: object = f.add_sheet('sheet1', cell_overwrite_ok=True)

tox = xlrd.open_workbook(path_ref)#组装基因组物种注释表
tox_sheet = tox.sheet_by_index(0)
len_t = tox_sheet.nrows

for i in range(1, nrows):
    t_v = s_dsheet.cell_value(i, 0)
    for n in range(len_t):
        rown = tox_sheet.row_values(n)
        if t_v in rown:
            f_sheet.write(i, 0, rown[2])

for k in range(0, nrows):
    row_v = s_dsheet.row_values(k)
    c = 1
    for cell in row_v:
        f_sheet.write(k, c, cell)
        c += 1
f.save(path_res)
