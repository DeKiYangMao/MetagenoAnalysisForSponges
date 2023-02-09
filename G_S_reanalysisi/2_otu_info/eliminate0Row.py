import xlwt
import xlrd

path = 'C:\\Users\\AMao\\Desktop\\'

s_d = xlrd.open_workbook(path + 'sha2_otu.xlsx')
s_dsheet = s_d.sheet_by_index(0)
nrows = s_dsheet.nrows
ncols = s_dsheet.ncols
f = xlwt.Workbook()
f_sheet: object = f.add_sheet('sheet1', cell_overwrite_ok=True)

#row0 = s_dsheet.row_values(0, 0, ncols)
#for i in range(0, len(row0)):
    #f_sheet.write(0, i, row0[i])

row_n = 0

for k in range(nrows):
    count = 0
    for c in range(ncols):
        c_v = s_dsheet.cell_value(k, c)
        if c_v == 0:
            count = count + 1

    if count != ncols-1:
        rowi = s_dsheet.row_values(k, 0, ncols)
        for z in range(0, len(rowi)):
            f_sheet.write(row_n, z, rowi[z])
        row_n = row_n + 1

f.save(path + 'sha2l_nzero.xlsx')