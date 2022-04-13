#这个脚本用于的文件结构：总文件夹（各个样本文件夹（各个Bin的antismash输出文件夹））
#本来是用于统计各个海绵宏基因组中高质量Bin在物种上的分布，以及SMBGC分布
#但是antisamsh爬取部分有问题（后来写了个别的），不过物种统计还是对的，懒得改了凑合用吧
import os
import xlwt
import xlrd
from bs4 import BeautifulSoup


o_dir ='F:\\sponge metagenome binning data\\18X sponoge MG data'#海绵Bin文件夹
list = os.listdir(o_dir)

for i in range(0, len(list)):#遍历各海绵样品文件夹
    spongeID = list[i]
    path = os.path.join(o_dir, list[i])


    f = xlwt.Workbook()  # 创建文件夹保存该海绵中各Bin信息
    f_sheet: object = f.add_sheet('sheet1', cell_overwrite_ok=True)
    row0 = ['BinID', 'phylum', 'ctg_length']#记录Bin门类、contig总长度、SMBGCs
    n = 0
    for i in range(0, len(row0)):
        f_sheet.write(0, i, row0[i])


    for item in os.listdir(path):#遍历该海绵中各组装基因组文件夹

        sub_path = os.path.join(path, item)
        if str(item) != 'BGCsSUM' and str(item) != 'genome':
            n = n + 1
            f_sheet.write(n, 0, str(item))

            contig = xlrd.open_workbook('C:\\Users\\AMao\\Desktop\\contig.xls')  # 抓取该bin组装序列总长度
            contig_sheet = contig.sheet_by_name("contig")
            len_c = contig_sheet.nrows
            for cn in range(len_c):
                rowc = contig_sheet.row_values(cn)
                if str(item) in rowc[0]:
                    contig_len = rowc[2]
                    f_sheet.write(n, 2, contig_len)

            phylum = xlrd.open_workbook('C:\\Users\\AMao\\Desktop\\phylum.xls')  # 抓取该Bin鉴定所属门类（这个表基于物种注释表把纲级别以下删掉了）
            phylum_sheet = phylum.sheet_by_name("phylum")
            len_p = phylum_sheet.nrows
            for pn in range(len_p):
                rowp = phylum_sheet.row_values(pn)
                if str(item) in rowp[0]:
                    phylun_type = rowp[1]
                    f_sheet.write(n, 1, phylun_type)

            url_bin = sub_path + '\\index.html'  # 爬取antismash结果文件中该Bin所含的SMBGCs
            soup = BeautifulSoup(open(url_bin), 'html.parser')
            item = soup.find_all('tr', class_="linked-row")
            num = len(item) / 2  # 由于一些不清楚的原因爬取结果会重复一遍
            bgc_sum = {}
            for b in range(int(num)):
                bgc_info = item[b]
                bgcinfo = bgc_info.find('a', class_="external-link")
                bgc_type = bgcinfo.text#单次获取该基因组内涵的一个基因簇类型计数

                # 建立该基因组基因簇类型及个数字典
                if bgc_type in bgc_sum:
                    bgc_sum[bgc_type] = bgc_sum[bgc_type] + 1
                else:
                    bgc_sum[bgc_type] = 1

                #判断是否有新的BGC出现,有则添加新的BGC列
                if bgc_type in row0:
                    existen = True
                else:
                    row0.append(bgc_type)
                    f_sheet.write(0, len(row0)-1, bgc_type)

            for k in range(3, len(row0)):
                try:
                    f_sheet.write(n, k, bgc_sum[row0[k]])
                except:
                    bgc_sum[row0[k]] = 0

    f.save('C:\\Users\\AMao\\Desktop\\' + spongeID + '\\Bins.xls')



