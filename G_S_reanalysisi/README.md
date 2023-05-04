Manually delete seawater and sediment samples from the "sponge info.xls" file, and then intercept the list according to the depth "0-30m", "30-100m", ">100m", such as "1_section_list/dl.xlsx" ("dl" is short for "deep-sea sponge list").

Step 1:
Extract the otu information of relevant samples from the OTU distribution table according to the sample list.
"python3 PATH/1_section_list/GetInfoFromDatasetByIndex.py PATH/otus_of_spgs.xlsx PATH/1_section_list/dl.xlsx PATH/2_otu_info/dep_otu.xlsx"
python3 GetInfoFromDatasetByIndex.py <OTU table> <list of samples to grab> <output path>

Step 2:
Since some OTUs are not distributed in the crawled samples, these OTUs should be removed.
"python3 PATH/2_otu_info/eliminate0Row.py PATH/2_otu_info/dep_otu.xlsx PATH/3_nozerodata/depl_nzero.xlsx"
python3 eliminate0Row.py <OTU table> <output path>

Step 3:
According to the species annotation table, add the species information to the obtained OTU table.
"python3 PATH/3_nozerodata/addTOXtoABinfo.py PATH/otus_toxn.xlsx PATH/3_nozerodata/depl_nzero.xlsx PATH/4_addToxinfo/depl_tox.xlsx"
python3 addTOXtoABinfo.py <otu annotation table> <otu table> <output path>

Step 4:
The species annotations of otu are often not at the same species level. Here, EXCEL is used to manually adjust the annotated species level (only keep phylum level information), and save it in PATH/5_adjustToxLevel/depl_pylum.

Step 5:
Add up the abundance of OTUs of the same taxonomy in each sample.
"python3 PATH/5_adjustToxLevel/SumSameRows.py PATH/5_adjustToxLevel/depl_pylum.xlsx PATH/6_sumInATox/depl_sum.xlsx"
python3 SumSameRows.py <table to be summed> <output path>

Step 6:
Considering that there is sample specificity in the abundance of microorganisms, we perform a secondary calculation on the abundance of home and post:
relative abundance = (abundance of a taxon in one sample/total microbial abundance of this sample)*1000
"python3 PATH/6_sumInATox/Calcu_R_abundence.py PATH/6_sumInATox/depl_sum.xlsx PATH/7_RAbundance/depl_ab.xlsx"

Step 7:
Find the average relative abundance and sample occupancy of each taxon sample
final abundance = sum(relative abundance of each sample)/num of samples
coveratio = num of samples harboring the taxonomy / num of samples
"python3 PATH/7_RAbundance/calcu_sum&coveratio.py PATH/7_RAbundance/depl_ab.xlsx PATH/8_coveratio/depl_coveratio.xlsx"

Step 8:
Manually adjust the taxon to be kept as a label, such as "PATH/9_plot/dep_plot.xlsx", and use ggplot2 for drawing.

中文解释：
手动地从“sponge info.xls”文件中删除海水与沉积物采样，再根据深度“0-30m”、“30-100m”、“>100m”来截取了列表，如“1_section_list/dl.xlsx”（“dl”是“deep-sea sponge list”的缩写）。
第一步：
从OTU分布表中按照样本列表把相关样本的otu信息抽取出来。
“python3 PATH/1_section_list/GetInfoFromDatasetByIndex.py PATH/otus_of_spgs.xlsx PATH/1_section_list/dl.xlsx PATH/2_otu_info/dep_otu.xlsx”
python3 GetInfoFromDatasetByIndex.py <OTU表> <要抓取的样本名单> <输出路径>
第二步：
由于有些OTU在爬取的样本的中没有分布，要去除这部分OTU。
“python3 PATH/2_otu_info/eliminate0Row.py PATH/2_otu_info/dep_otu.xlsx PATH/3_nozerodata/depl_nzero.xlsx”
python3 eliminate0Row.py <OTU表> <输出路径>
第三步：
按照物种注释表，把物种信息添加到获得的OTU表中。
“python3 PATH/3_nozerodata/addTOXtoABinfo.py PATH/otus_toxn.xlsx PATH/3_nozerodata/depl_nzero.xlsx PATH/4_addToxinfo/depl_tox.xlsx”
python3 addTOXtoABinfo.py <otu注释表> <otu表> <输出路径>
第四步：
otu的物种注释往往不在同一物种水平，这里使用EXCEL手动的调整了注释的物种水平（只保留phylum级别的信息），并保存在PATH/5_adjustToxLevel/depl_pylum.
第五步：
把同一分类的otu在各个样本中的丰度加起来。
“python3 PATH/5_adjustToxLevel/SumSameRows.py PATH/5_adjustToxLevel/depl_pylum.xlsx PATH/6_sumInATox/depl_sum.xlsx”
python3 SumSameRows.py <需要被加和的表> <输出路径>
第六步：
考虑到微生物的丰度存在样本特异性，所以我们对家和后的丰度进行的二次计算：
relative abundance = （abundance of a taxon in one sample/total microbial abundance of this sample）*1000
“python3 PATH/6_sumInATox/Calcu_R_abundence.py PATH/6_sumInATox/depl_sum.xlsx PATH/7_RAbundance/depl_ab.xlsx”
第七步：
求取每个taxon的样本平均相对丰度以及样本占有率
final abundance = sum(relative abundance of each sample)/num of samples
coveratio = num of samples harboring the taxonomy / num of samples
“python3 PATH/7_RAbundance/calcu_sum&coveratio.py PATH/7_RAbundance/depl_ab.xlsx PATH/8_coveratio/depl_coveratio.xlsx”
第八步：
手动调整要保留为标签的taxon，如“PATH/9_plot/dep_plot.xlsx”，使用ggplot2作图。
