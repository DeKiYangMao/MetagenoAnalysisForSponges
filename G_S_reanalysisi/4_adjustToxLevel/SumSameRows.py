#本脚本用于将一份表当中的相同“行标签（每一行的第一个值）”的多行数据合并
import xlwt
import xlrd

filename = 'depl_pylum.xlsx'#操作文件名
path = '' #文件所处文件夹的绝对路径
f = xlwt.Workbook()
f_sheet: object = f.add_sheet('sheet1', cell_overwrite_ok=True)

info_b = xlrd.open_workbook(path + filename)
info_sheet = info_b.sheet_by_index(0)
nrow = info_sheet.nrows
ncol = info_sheet.ncols

tox_dir = {}
n = 0

for r in range(1, nrow): #加和行的起始位置
    row_v = info_sheet.row_values(r, 0, ncol)
    #print(row_v)
    tox_n = row_v[0] #一行数据中的标记所处位置
    if tox_n in tox_dir:
        #tox_dir[tox_n][0] += 1
        for i in range(1, ncol): #加和列的起始位置
            #print(row_v)
            tox_dir[tox_n][i-1] += row_v[i]
    else:
        tox_dir[tox_n] = []
        for i in range(1, ncol): #加和列的起始位置
            tox_dir[tox_n].append(row_v[i])
        print(tox_dir[tox_n])
print(tox_dir)
for tox in tox_dir:
    f_sheet.write(n, 0, tox)
    for i in range(0, len(tox_dir[tox])):
        f_sheet.write(n, i+1, tox_dir[tox][i])

    n += 1
f.save(path + filename.replace('_pylum', '_sum'))







