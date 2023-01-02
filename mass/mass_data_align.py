'''
This program is designed to compare the mass spectrum data from the sample and the standard library,
count the number of hits and record thier id.
Usage: python3 mass_data_align.py <input .xls file> <output_path>
Some notice about inputfile:
1.There can be many sheet within one .xls file but the last two tables in the .xls file of the mass 
spectrometry data must be named "database" and "control". The "database" table is the database data 
we want to compare with samples', and the "control" table is the blank control group data in the mass 
spectrometry detection process.
2.Each other sheet should store the mass data from one sample,and other sheets' name will be used as 
samples' id.
3.For the samples' sheets, the data of the first col must be the scan time/retention time, and the 
second col's should be m/z data.
4.For standard dataset sheet, the data of the first col should be compounds informations(like a ID), 
the second cols' data should be compounds' molecular mass(The ionized weight is calculated in the program, 
adding only the weight of one sodium ion or one hydrogen ion.)
5.For control group sheet, there only need one col, which contain the m/z ratio of the control set.
6.For all sheets, data in the first row will be discarded(I'd like to keep titles~)
Written by: Deqiang AMao :)
'''
import xlwt
import xlrd
import sys
import os

def summary_mass_list(data_list, compound_list):
    # to merge isomers in the database
    # data_list records all the m/z in the set and compound_list records the compounds id
    # the m/z and compound_id infomation in these two list should correspond to each other
    mass_dir = {}
    for mass in data_list:
        if mass in mass_dir:
            compound_id = compound_list[data_list.index(mass)+len(mass_dir[mass])]
            mass_dir[mass].append(compound_id)
        else:
            mass_dir[mass] = []
            compound_id = compound_list[data_list.index(mass)]
            mass_dir[mass].append(compound_id)
            # The prerequisite for merging isomers and effectively recording compound IDs
            # is that the data in the sheet has been sorted according to m/z
            # to ensure that compounds with the same m/z are continuously distributed.
    return mass_dir
    # Finally, a dictionary indexed by m/z will be returned, recording the compound id
    # {m/z_ratio_1:[compound_1,compound_2.....],.......}

def cut_RT(mass_RT_dict):
    # In the sample data set, there may be multiple retention times for the same m/z.
    # For this, we identified signals separated by more than two minutes as from two different molecules.
    # However, if there is a signal that occurs repeatedly, but with an interval of less than two minutes,
    # they will be considered the same molecule.
    for mass in mass_RT_dict:
        index = 0
        while index < len(mass_RT_dict[mass]) - 1:
            former = mass_RT_dict[mass][index]
            later = mass_RT_dict[mass][index + 1]
            if later - former < 2:
                del mass_RT_dict[mass][index]
                index - 1
            else:
                index += 1
    return

def sam_mz_hit(sample_data, data_set):
    # Count the number and information of m/z hits
    # sample_data{sample:{m/z_ratio_1:[RT1,RT2...].....}
    # data_set{m/z_ratio_1:[compound_1,compound_2.....],......}
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
                   compound_info = data_set[sdandard_mz]
                   sam_mz_hit[sample][mass] = {}
                   sam_mz_hit[sample][mass]['mol_num'] = mol_of_mass
                   sam_mz_hit[sample][mass]['hit_mz'] = sdandard_mz
                   sam_mz_hit[sample][mass]['compound_id'] = compound_info
        sam_mz_hit[sample]['total_mol'] = total_mol
        sam_mz_hit[sample]['dataset_mol_hit'] = known_mol
    return sam_mz_hit

def judge_difference(data, sdandard_data):
    # Calculate the difference between the m/z ratios in the database and the sample. 
    # If the difference is less than 5 parts per million of the standard value in the database, 
    # the two m/z ratios are considered to be from the same molecule (or isomer).
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
                    col += 1
                    try:
                        for compound_info in dict[sample][line][info]:
                            f_sheet.write(row, col, compound_info)
                            col += 1
                    except:
                         f_sheet.write(row, col, dict[sample][line][info])
                         col += 1
            except:
                f_sheet.write(row, col, dict[sample][line])
            row += 1
    path = file_path + '\\' + file_name +'.xlsx'
    f.save(path)

mass_data_path = sys.argv[1]
res_path = sys.argv[2]
mass_data = xlrd.open_workbook(mass_data_path)
data_sheets = mass_data.sheets()

# input the mass data of control set
control_sheet = mass_data.sheet_by_name('control')
control_mz = control_sheet.col_values(0, 1, control_sheet.nrows)
C_mz_list = []
for mz in control_mz:
    C_mz_list.append(mz)

# input the mass data of standard dataset,
# and biuld up the m/z ratio data list of sodium ion and hydrogen ion
dataset = mass_data.sheet_by_name('database')
compound_list = dataset.col_values(0, 1, dataset.nrows)
mass_data_list = dataset.col_values(1, 1, dataset.nrows)
Na_mass_data = [item + 22.989769 for item in mass_data_list]
H_mass_data = [item + 1.00784 for item in mass_data_list]
# Combine the information of the isomers in the dataset
Na_mass = summary_mass_list(Na_mass_data, compound_list)
H_mass = summary_mass_list(H_mass_data, compound_list)

# input the mass data of sample 
# Remove the m/z ratio that is consistent with the data in the control group, 
# and merge the signals with a detection time within two minutes.
sample_mass = {}
for sheet_index in range(0, len(data_sheets)-2):
    sheet = mass_data.sheet_by_index(sheet_index)
    sheet_name = str(data_sheets[sheet_index])
    sheet_name = sheet_name[10:16]
    sample_mass[sheet_name] = {}
    nrow = sheet.nrows
    for r in range(1, nrow):
        RT = sheet.cell_value(r, 0)
        mass = sheet.cell_value(r, 1)
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

write_res(sam_mz_hit_withH, 'MZ_hit_in_Hdata', res_path)
write_res(sam_mz_hit_withNa, 'MZ_hit_in_Nadata', res_path)

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
file_path = os.path.join(res_path,'MZ_hits_sum.xlsx')    
f.save(file_path)
