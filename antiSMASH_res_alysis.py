from bs4 import BeautifulSoup
import os
import xlwt
import xlrd

def avoid_same_name(n,namelist):
    if n in namelist:
       n = n + '*'
    else:   
       return n 
    return avoid_same_name(n, namelist)

def create_add_dic(dic, add_ob):
    if add_ob in dic:
        dic[add_ob] += 1
    else:
        dic[add_ob] = 1
    return

def write_list_to_row(sheet, list, row_n):
    col_n = 1
    for i in list:
        sheet.write(row_n, col_n, i)
        col_n += 1
    return

def write_sheet_withDIC(sheet_path, dic):
    # 目标sheet为已有索引行（0）和索引列（0）的表；dic为{列索引值：{行索引值：cell值...}...}
    rf = xlrd.open_workbook(sheet_path)
    rf_sheet = rf.sheet_by_index(0)
    nrows = rf_sheet.nrows
    ncols = rf_sheet.ncols
    BGC_index = rf_sheet.row_values(0, 1, ncols)
    f = xlwt.Workbook()
    f_sheet: object = f.add_sheet('sheet1', cell_overwrite_ok=True)
    col_n = 1
    for i in BGC_index:
        f_sheet.write(0, col_n, i)
        col_n += 1
    
    for row in range(1, nrows):
        row_index = rf_sheet.cell_value(row, 0)
        if '!!' in row_index:
            #f_sheet.write(row, 0, row_index)
            for col in range(2, ncols):
                col_index = rf_sheet.cell_value(0, col)
                try:
                    cell_v = dic[row_index][col_index]
                    f_sheet.write(row, col, cell_v/2)
                except:
                    f_sheet.write(row, col, 0)
            real_Bname = row_index.replace('!!', '')
            f_sheet.write(row, 0, real_Bname)
        else:    
            f_sheet.write(row, 0, row_index)
            for col in range(2, ncols):
                col_index = rf_sheet.cell_value(0, col)
                try:
                    cell_v = dic[row_index][col_index]
                    f_sheet.write(row, col, cell_v)
                except:
                    f_sheet.write(row, col, 0)
    f.save(sheet_path)
    return

o_dir =''#海绵Bin文件夹
list = os.listdir(o_dir)
bin_info = {}
for i in range(0, len(list)):#遍历各海绵样品文件夹
    spongeID = list[i]
    path = os.path.join(o_dir, list[i])
    for bin in os.listdir(path):#遍历该海绵中各组装基因组文件夹
        sub_path = os.path.join(path, bin)
        url_bin = sub_path + '/index.html'
        soup = BeautifulSoup(open(url_bin), 'html.parser')
        item = soup.find_all('tr', class_="linked-row odd")
        item_even = soup.find_all('tr', class_="linked-row even")
        item = item + item_even
        span = soup.find('div', class_="overview-switches")
        if span != None:
            #print('1')
            bin = bin + '!!'
        #else:
            #print('2')    
        bin_info[bin] = {}
        for region in item:
            #print(region)
            region_info = []
            conpunds = region.find_all('a', class_="external-link")
            similarity = region.find('td', class_="digits similarity-text")
            for a_moduel in conpunds:
                conpunds_info = a_moduel.text
                region_info.append(conpunds_info)
            #print(region_info)

            try:
                simi_v = similarity.text
                #print(simi_v)
                bgc_type = region_info[0]
                bgc_type = avoid_same_name(n=bgc_type, namelist=bin_info[bin])
                conpund = region_info[-1]
                bin_info[bin][bgc_type] = []
                bin_info[bin][bgc_type].append(conpund)
                bin_info[bin][bgc_type].append(simi_v) 
            except:
                bgc_type = region_info[0]
                bgc_type = avoid_same_name(n=bgc_type, namelist=bin_info[bin])
                bin_info[bin][bgc_type] = []
                bin_info[bin][bgc_type].append('no similar conpund')
                bin_info[bin][bgc_type].append('none')

#print(bin_info)
#统计已知化合物类似BGC
BCsum = xlwt.Workbook()
BCsum_sheet: object = BCsum.add_sheet('sheet1', cell_overwrite_ok=True)
bin_row = 0
#单独输出BIN的BGC总和
BinBGCsum = xlwt.Workbook()
BinBGCsum_sheet: object = BinBGCsum.add_sheet('sheet1', cell_overwrite_ok=True)
sum_row = 1
#建立一个新字典{bin1:{BGCtype1:bgc_num, BGCtype2:bgc_num...}, bin2:{...}...},以及总BGCtype列表
BinandBGCtype = {}
BGCrow = ['sum']

#读取bin_info中的{bin[region[bgctype...similar conpounds, similarity]]
for each_bin in bin_info:
    BinandBGCtype[each_bin] = {}#建立bin的子字典
    #两张表填入bin
    BCsum_sheet.write(bin_row, 0, each_bin)
    BinBGCsum_sheet.write(sum_row, 0, each_bin)
    #统计总bgc数量
    #BCsum_sheet.write(bin_row, 1, str(len(bin_info[each_bin])/2))
    #bin_row += 1
    BinBGCsum_sheet.write(sum_row, 1, str(len(bin_info[each_bin])/2))
    sum_row += 1
    #以下为具体BGC类型
    bgc_col = 1#已知化合物表BGCtype填写起始列
    for each_bgc in bin_info[each_bin]:
        write_bgc_type = each_bgc.replace('*', '')#去除重复BGC的‘*’tag

        if write_bgc_type not in BGCrow:#为BIN-BGCtypes丰度表构建索引行
            BGCrow.append(write_bgc_type)

        create_add_dic(dic=BinandBGCtype[each_bin], add_ob=write_bgc_type)
        BCsum_sheet.write(bin_row, bgc_col, write_bgc_type)
        cp_row = bin_row + 1#已知化合物填写对应BGCtype下一行同一列
        
        #以下为相似化合物类型
        for cp_info in bin_info[each_bin][each_bgc]:
            BCsum_sheet.write(cp_row, bgc_col, cp_info)
            cp_row += 1
        bgc_col += 1
    bin_row += 3

#print(BinandBGCtype)
#填写 BIN -- BGCtypes 丰度表索引行
write_list_to_row(sheet=BinBGCsum_sheet, list=BGCrow, row_n=0)
BCsum.save(o_dir+'/BGCandConpounds_sum.xlsx')
BinBGCsum.save(o_dir+'/BinBGCsum.xlsx')
#填写BIN -- BGCtypes丰度表的值
temp_sheet = o_dir+'/BinBGCsum.xlsx'
write_sheet_withDIC(sheet_path=temp_sheet, dic=BinandBGCtype)

