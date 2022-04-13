#基于一份总表，把被注释表中的数据tag对应的注释插入进来
import xlwt
import xlrd

path = 'C:\\Users\\AMao\\Desktop\\'

s_d = xlrd.open_workbook(path + 'midl_nzero.xlsx')#需要被添加注释信息的表
s_dsheet = s_d.sheet_by_index(0)
nrows = s_dsheet.nrows

f = xlwt.Workbook()
f_sheet: object = f.add_sheet('sheet1', cell_overwrite_ok=True)

tox = xlrd.open_workbook('C:\\Users\\AMao\\Desktop\\otus_toxn.xlsx')#组装基因组物种注释表
tox_sheet = tox.sheet_by_index(0)
len_t = tox_sheet.nrows

for i in range(1, nrows):
    t_v = s_dsheet.cell_value(i, 0)
    for n in range(len_t):
        rown = tox_sheet.row_values(n)
        if t_v in rown:
            f_sheet.write(i, 1, rown[1])


f.save(path + 'midl_tox.xlsx')