#这个脚本用于统计物种在样本中的存在比率
#统计表格的数据分布：第一列为物种名，第一行为样本，数据为不同的物种在样本中的TPM丰度
import xlwt
import xlrd

path = 'C:\\Users\\AMao\\Desktop\\'#添加统计表格的绝对路径
f = xlwt.Workbook()
f_sheet: object = f.add_sheet('sheet1', cell_overwrite_ok=True)

info_b = xlrd.open_workbook(path+'')#添加统计表格的文件名
info_sheet = info_b.sheet_by_index(0)
nrow = info_sheet.nrows
ncol = info_sheet.ncols

tox_dir = {}
sample_num = ncol -2#注意默认减去一列物种注释列和一列

for r_n in range(1, nrow):
    otu_name = info_sheet.cell_value(r_n, 1)
    if otu_name in tox_dir:
        for c_n in range(2, ncol):
            c_v = info_sheet.cell_value(r_n, c_n)
            tox_dir[otu_name][c_n-2] += c_v
    else:
        tox_dir[otu_name] = []
        for c_n in range(2, ncol):
            c_v = info_sheet.cell_value(r_n, c_n)
            tox_dir[otu_name].append(c_v)

n = 0
for tox in tox_dir:
    f_sheet.write(n, 0, tox)

    tox_sum = 0
    tox_cover = 0
    for samp in tox_dir[tox]:
        if samp != 0:
            tox_cover += 1
            tox_sum += samp
    tox_dir[tox].append(tox_sum)
    tox_dir[tox].append(tox_cover/sample_num)

    f_sheet.write(n, 1, tox_dir[tox][sample_num])
    f_sheet.write(n, 2, tox_dir[tox][sample_num+1])
    n += 1

f.save(path+'')#保存文件名






