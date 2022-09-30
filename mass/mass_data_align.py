import xlwt
import xlrd

mass_data_path = input('the path of mass_data_film, which should be a .xls film: ')
res_folder = input('the path of the folder used to save results: ')

def summary_mass_list(data_list):#合并数据库中的同分异构体
    mass_dir = {}
    for mass in data_list:
        mass = format(mass, '.6f')
        if mass in mass_dir:
            mass_dir[mass] += 1
        else:
            mass_dir[mass] = 1
    return mass_dir
    #最终获得字典类型结果：{分子量：同分异构体数量，......}

def cut_RT(mass_RT_dict):#去除保留时间的冗余
    #数据结构为{质荷比1：[保留时间1，保留时间2，......],......}
    for mass in mass_RT_dict:
        index = 0
        while index < len(mass_RT_dict[mass]) - 1:
            former = mass_RT_dict[mass][index]
            later = mass_RT_dict[mass][index + 1]
            if later - former < 2:#如果当前保留时间与上一保留时间差在2以内，去除上一保留时间
                del mass_RT_dict[mass][index]
                index - 1
            else:
                index += 1
    return

def sam_mz_hit(sample_data, data_set):
    #sample_data{sample:{mass:[RT1,RT2...].....}
    #data_set{mass:num}
    sam_mz_hit = {}
    for sample in sample_data:
        sam_mz_hit[sample] = {}
        total_mol = 0
        known_mol = 0
        for mass in sample_data[sample]:
            total_mol += len(sample_data[sample][mass])
            for sdandard_mz in data_set:
                if judge_difference(mass, sdandard_mz):
                   known_mol += len(sample_data[sample][mass])
                   mol_of_mass = len(sample_data[sample][mass])
                   sam_mz_hit[sample][mass] = {}
                   sam_mz_hit[sample][mass]['mol_num'] = mol_of_mass
                   sam_mz_hit[sample][mass]['hit_mz'] = sdandard_mz
        sam_mz_hit[sample]['total_mol'] = total_mol
        sam_mz_hit[sample]['dataset_mol_hit'] = known_mol
    return sam_mz_hit

def judge_difference(data, sdandard_data):
    data = float(data)
    sdandard_data = float(sdandard_data)
    delta = abs(data-sdandard_data)/sdandard_data
    if delta < 0.000005:
        return True
    else:
        return False

def write_res(dict, file_name, file_path):
    f = xlwt.Workbook()
    for sample in dict:
        f_sheet: object = f.add_sheet(sample, cell_overwrite_ok=True)
        row = 0
        for line in dict[sample]:
            f_sheet.write(row, 0, line)
            col = 1
            try:
                for info in dict[sample][line]:
                    f_sheet.write(row, col, info)
                    f_sheet.write(row, col+1, dict[sample][line][info])
                    col += 2
            except:
                f_sheet.write(row, col, dict[sample][line])
            row += 1
    path = file_path + '\\' + file_name +'.xlsx'
    f.save(path)

mass_data = xlrd.open_workbook(mass_data_path)
data_sheets = mass_data.sheets()

control_sheet = mass_data.sheet_by_name('control')
control_mz = control_sheet.col_values(0, 1, control_sheet.nrows)
C_mz_list = []#负对照列表
for mz in control_mz:
    C_mz_list.append(format(mz, '.6f'))

dataset = mass_data.sheet_by_name('database')
Na_mass_data = dataset.col_values(3, 1, dataset.nrows)
H_mass_data = dataset.col_values(4, 1, dataset.nrows)
Na_mass = summary_mass_list(Na_mass_data)#数据库加Na峰列表
H_mass = summary_mass_list(H_mass_data)#数据库加H峰列表

sample_mass = {}#集合所有样本数据中的质荷比及其相应的保留时间
for sheet_index in range(0, len(data_sheets)-2):
    sheet = mass_data.sheet_by_index(sheet_index)
    sheet_name = str(data_sheets[sheet_index])
    sheet_name = sheet_name[10:16]
    sample_mass[sheet_name] = {}
    nrow = sheet.nrows
    for r in range(1, nrow):
        RT = sheet.cell_value(r, 0)
        mass = format(sheet.cell_value(r, 1), '.6f')
        mass_in_controlset = False
        for C_mz in C_mz_list:
           if judge_difference(mass, C_mz):
               mass_in_controlset = True
        if not mass_in_controlset:
           if mass not in sample_mass[sheet_name]:
              sample_mass[sheet_name][mass] = []
              sample_mass[sheet_name][mass].append(RT)
           else:
              sample_mass[sheet_name][mass].append(RT)
    cut_RT(sample_mass[sheet_name])


sam_mz_hit_withNa = sam_mz_hit(sample_mass, Na_mass)
sam_mz_hit_withH = sam_mz_hit(sample_mass, H_mass)

write_res(sam_mz_hit_withH, 'MZ_hit_in_Hdata', res_folder)
write_res(sam_mz_hit_withNa, 'MZ_hit_in_Nadata', res_folder)

sum_dict = {}
for mass in sam_mz_hit_withH:
    sum_dict[mass] = {}
    sum_dict[mass]['total_mol']=sam_mz_hit_withH[mass]['total_mol']
    sum_dict[mass]['hits_num']=sam_mz_hit_withH[mass]['dataset_mol_hit']+sam_mz_hit_withNa[mass]['dataset_mol_hit']

f = xlwt.Workbook()
f_sheet: object = f.add_sheet('hits_num_info', cell_overwrite_ok=True)
f_sheet.write(0, 1, 'total_mol')
f_sheet.write(0, 2, 'hits_num')
row = 1
for mass in sum_dict:
    f_sheet.write(row, 0, mass)
    f_sheet.write(row, 1, sum_dict[mass]['total_mol'])
    f_sheet.write(row, 2, sum_dict[mass]['hits_num'])
    row += 1
save_path = res_folder + '\\MZ_hits_sum.xlsx'
f.save(save_path)
