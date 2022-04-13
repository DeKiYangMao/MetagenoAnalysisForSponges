#!/bin/python3

'''
This program is designed to process CheckM output file (e.g. bin_stats_ext.tsv) to easy-to-read text format (.txt)
Usage: python3 checkm_summary.py <inputfile> <outputfile>
Written by: Emmett Peng
'''
import json
import sys

with open(sys.argv[1], 'r') as f:
	Load = {}
	for line in f:
		line = line.replace('\'','\"')
		line = line.split('\t')
		line[0] = 'bin_' + line[0].replace('.bin.', '_')
		#print(line[0])
		#line[0]:Bin Id; line[1]:Bin information
		#exec(f"line[0] = json.loads(line[1])")
		Load[line[0]] = json.loads(line[1])
		print(Load)
		#print(text)

with open(sys.argv[2], 'w+') as output:
	output.write('Bin Id\tMarker Lineage\tGenomes\tMarkers\tMarker Sets\t0\t1\t2\t3\t4\t5+\tCompleteness\tContamination\tGC\tGC std\tGenome Size\tAmbiguous Bases\tScaffolds\tContigs\tTranslation Table\tPredicted Genes\n')
	for key in Load:
		output.write(key + '\t')
		output.write(Load[key]['marker lineage'] + '\t')
		output.write(str(Load[key]['# genomes']) + '\t')
		output.write(str(Load[key]['# markers']) + '\t')
		output.write(str(Load[key]['# marker sets']) + '\t')
		output.write(str(Load[key]['0']) + '\t')
		output.write(str(Load[key]['1']) + '\t')
		output.write(str(Load[key]['2']) + '\t')
		output.write(str(Load[key]['3']) + '\t')
		output.write(str(Load[key]['4']) + '\t')
		output.write(str(Load[key]['5+']) + '\t')
		output.write(str(Load[key]['Completeness']) + '\t')
		output.write(str(Load[key]['Contamination']) + '\t')
		output.write(str(Load[key]['GC']) + '\t')
		output.write(str(Load[key]['GC std']) + '\t')
		output.write(str(Load[key]['Genome size']) + '\t')
		output.write(str(Load[key]['# ambiguous bases']) + '\t')
		output.write(str(Load[key]['# scaffolds']) + '\t')
		output.write(str(Load[key]['# contigs']) + '\t')
		output.write(str(Load[key]['Translation table']) + '\t')
		output.write(str(Load[key]['# predicted genes']) + '\t')
		output.write('\n')