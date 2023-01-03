This program is designed to compare the mass spectrum data from the sample and the standard library,count the number of hits and record thier id.

Usage: python3 mass_data_align.py <input .xls file> <output_path>

My inputfile is 'mass_data.xls', which is uploaded to this folder, and results is stored in 'res' folder.

Some notice about inputfile:

1.There can be many sheet within one .xls file but the last two tables in the .xls file of the mass spectrometry data must be named "database" and "control". The "database" table is the database data we want to compare with samples', and the "control" table is the blank control group data in the mass spectrometry detection process.

2.Each other sheet should store the mass data from one sample,and other sheets' name will be used as samples' id.

3.For the samples' sheets, the data of the first col must be the scan time/retention time, and the second col's should be m/z data.

4.For standard dataset sheet, the data of the first col should be compounds informations(like a ID), the second cols' data should be compounds' molecular mass(The ionized weight is calculated in the program, adding only the weight of one sodium ion or one hydrogen ion). And one more thing about dataset sheet, which is important, is that the data in the table should be arranged in ascending or descending order according to the m/z ratio.

5.For control group sheet, there only need one col, which contain the m/z ratio of the control set.

6.For all sheets, data in the first row will be discarded.

Written by: Deqiang AMao :)
