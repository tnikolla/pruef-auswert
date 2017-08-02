#!/usr/bin/python
import sys
sys.path.insert(1,'/home/ertel/bin/xlutils-2.0.0')
import mod1
import numpy as np
import matplotlib.pyplot as plt
import argparse
parser = argparse.ArgumentParser()
parser.add_argument('input', nargs='+', help='name of xls input files')
parser.add_argument('max_pts', help='maximum reachable points in the exam', type=float)
parser.add_argument('pmax', help='minimum points for 1.0', type=float)
parser.add_argument('pmin', help='minimum points for 4.0', type=float)
parser.add_argument('-b', '--bins', help='number of bins in the histogram', 
                    default=30, type=int)
args = parser.parse_args()

#Arguments
max_pts = args.max_pts
pmax = args.pmax
pmin = args.pmin
bins = args.bins
diff = pmax - pmin

#Reads the pts from a xls-file and converts them to grds.
#Writes the xls-file, ready to be imported in lsf
result =  []
for i in args.input:
     result.append(mod1.ptsToGrds(i,pmax,pmin,diff))

for i in range(len(args.input)):
    if 1==len(args.input):
        punkte = result[0][0]
        noten = result[0][1]
        break
    if i==0:
        punkte = np.concatenate((result[i][0],result[i+1][0]))
        noten  = np.concatenate((result[i][1],result[i+1][1]))
    if i>0 and i+1<len(args.input):
        punkte = np.concatenate((punkte,result[i+1][0]))
        noten  = np.concatenate((noten,result[i+1][1]))
        
for i in range(len(args.input)):
    mod1.hist(args.input[i], result[i][0], result[i][1], 
              result[i][3],result[i][4], max_pts, pmax, pmin, bins)

result=np.array(result)
stud_ges=result[:,2]
mean = np.round(np.mean(noten),1)
failure = noten > 4.0
failure_nr=np.sum(failure)
failure = str(int(np.round(np.mean(failure)*100)))
string = (str(np.sum(stud_ges))+' Studenten,'
    +' 4.0 bei ' +str(int(pmin)) +' Punkte,'
    +' 1.0 bei ' +str(int(pmax)) +' Punkte,'
    +' Schnitt ' +str(mean) + ','
    +' Durchgefallen ' +str(int(failure_nr)) +' (' +failure +'%)')
print string

mod1.hist('Allgemeines', punkte, noten, mean, failure, max_pts,
          pmax, pmin, bins)

plt.show()