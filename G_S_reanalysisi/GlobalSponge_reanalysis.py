import xlwt
import xlrd # version == 1.2.0

path = 'C:\\Users\\AMao\\Desktop'
layer_file = 'sl.xlsx'
otu_for_spgs = 'otus_for_spgs.xlsx'
otus_toxn = 'otus_toxn.xlsx'
tempt = 'C:\\Users\\AMao\\Desktop\\tempt'
sample_list_path = path +'\\'+ layer_file
otu_sample_sheet_path = path +'\\'+ otu_for_spgs
otu_tox_sheet_path = path +'\\'+ otus_toxn

sample_list = xlrd.open_workbook(sample_list_path)
sample_list_sheet = sample_list.sheet_by_index(0)
qrows = sample_list_sheet.nrows

otu_sample = xlrd.open_workbook(otu_sample_sheet_path)
otu_sample_sheet = otu_sample.sheet_by_index(0)
otu_sample_sheet_nrows = otu_sample_sheet.nrows
otu_sample_sheet_ncols = otu_sample_sheet.ncols

otu_tox = xlrd.open_workbook(otu_tox_sheet_path)
otu_tox_sheet = otu_tox.sheet_by_index(0)
otu_tox_sheet_nrows = otu_tox_sheet.nrows
otu_tox_sheet_ncols = otu_tox_sheet.ncols

#part1
l_otu_row0 = []
l_otu_row0_info = sample_list_sheet.col_values(0,0,qrows)
for r_v_1 in l_otu_row0_info:
    l_otu_row0.append(r_v_1)

l_otu_f = xlwt.Workbook()
l_otu_f_sheet: object = l_otu_f.add_sheet('sheet1', cell_overwrite_ok=True)
for i in range(0, len(l_otu_row0)):
    l_otu_f_sheet.write(0,i+1,l_otu_row0[i])

otu_sample_col0 = otu_sample_sheet.col_values(0,1,otu_sample_sheet_nrows)
r_n_1 = 1
for c_n in otu_sample_col0:
    l_otu_f_sheet.write(r_n_1,0,c_n)
    r_n_1 += 1

sample_l = otu_sample_sheet.row_values(0,0,otu_sample_sheet_ncols)
for r0_v in l_otu_row0:
    if r0_v in sample_l:
        n = 1
        c_n = l_otu_row0.index(r0_v) + 1
        s_c_n = sample_l.index(r0_v)
        col_v = otu_sample_sheet.col_values(s_c_n,1,otu_sample_sheet_nrows)
        for q in col_v:
            l_otu_f_sheet.write(n, c_n, q)
            n += 1

l_otu_f_name = layer_file.replace('.xlsx','') 
l_otu_f_name = l_otu_f_name + '_otuinfo.xlsx'
l_otu_f.save(tempt + '\\' + l_otu_f_name)

#part2
l_otu_f_0 = xlrd.open_workbook(tempt + '\\' + l_otu_f_name)
l_otu_f_0_sheet = l_otu_f_0.sheet_by_index(0)
otuf_0_sheet_nrows = l_otu_f_0_sheet.nrows
otuf_0_sheet_ncols = l_otu_f_0_sheet.ncols

nozero_f = xlwt.Workbook()
nozero_f_sheet: object = nozero_f.add_sheet('sheet1', cell_overwrite_ok=True)

nozerof_row_n = 0
for k in range(otuf_0_sheet_nrows):
    count = 0
    for c in range(otuf_0_sheet_ncols):
        c_v = l_otu_f_0_sheet.cell_value(k, c)
        if c_v == 0:
            count = count + 1

    if count != otuf_0_sheet_ncols-1:
        rowi = l_otu_f_0_sheet.row_values(k, 0, otuf_0_sheet_ncols)
        for z in range(0, len(rowi)):
            nozero_f_sheet.write(nozerof_row_n, z, rowi[z])
        nozerof_row_n = nozerof_row_n + 1

nozero_f_path = tempt + '\\' + l_otu_f_name
nozero_f_path = nozero_f_path.replace('.xlsx','')
nozero_f.save(nozero_f_path + '_nozero.xlsx')

#part3
nzero_f = xlrd.open_workbook(nozero_f_path + '_nozero.xlsx')
nzero_f_sheet = nzero_f.sheet_by_index(0)
nzero_f_nrows = nzero_f_sheet.nrows

l_tox_f = xlwt.Workbook()
l_tox_sheet: object = l_tox_f.add_sheet('sheet1', cell_overwrite_ok=True)

for i in range(1, nzero_f_nrows):
    tox_v = nzero_f_sheet.cell_value(i,0)
    for n in range(otu_tox_sheet_nrows):
        row_n_tox = otu_tox_sheet.row_values(n)
        if tox_v in row_n_tox:
           l_tox_sheet.write(i,0,row_n_tox[2])

for k in range(0, nzero_f_nrows):
    row_v = nzero_f_sheet.row_values(k)
    c = 1
    for cell in row_v:
        l_tox_sheet.write(k, c, cell)
        c += 1

l_tox_f_path = nozero_f_path + '_nozero.xlsx'
l_tox_f_path = l_tox_f_path.replace('.xlsx','')
l_tox_f.save(l_tox_f_path + '_tox.xlsx')

