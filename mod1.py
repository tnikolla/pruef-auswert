import sys
sys.path.insert(1,'/home/ertel/bin/xlutils-2.0.0')
import numpy as np
from xlrd import open_workbook
from xlutils.copy import copy
import matplotlib.pyplot as plt

def ptsToGrds(input_xls,pmax,pmin,diff):
    
    #Create an object that can read the inputfile
    rb = open_workbook(input_xls)
    rs = rb.sheet_by_index(0)

    #Fill a list of the exam points
    punkte_lsf = []
    for row in range(rs.nrows-1)[2:] :
        punkte_lsf.append(rs.cell(row,11).value)
    
    punkte=[]
    for i in range(len(punkte_lsf)):
        if punkte_lsf[i] == '': punkte_lsf[i]='5U'
        else:
            punkte_lsf[i]=float(punkte_lsf[i])
            punkte.append(punkte_lsf[i])
    
    #Convert the points into floats or 5U
    #Map the points into grades
    noten_lsf=[]
    noten=[]
    stud_ges = 0
    durchgefallen = 0
    intercept = 1.0 + 3.0/diff*pmax
    for i in range(len(punkte_lsf)):
        if type(punkte_lsf[i])==float:
            temp=-3.0/diff*punkte_lsf[i] + intercept
            temp=np.floor(temp*10)/10
            if temp<1.0: temp=1.0
            if temp>5.0:
                temp=5.0
                durchgefallen=durchgefallen+1
            stud_ges = stud_ges+1
            noten.append(round(temp, 1))
            noten_lsf.append(round(temp*100,-1))
        else: noten_lsf.append(punkte_lsf[i])
    
    punkte=np.array(punkte)
    noten=np.array(noten)
    mean = str(np.round(np.mean(noten),1))
    #failure = noten > 4.0
    failure_rate = str(int(np.round(np.mean(noten>4)*100)))
    
    #Create a new xls object; a copy of the old one
    #and write the grades
    wb = copy(rb)
    ws = wb.get_sheet(0)
    for row in range(rs.nrows-1)[2:]:
        ws.write(row,11,noten_lsf[row-2])
    
    #Create the name of the output file
    output_xls = list(input_xls)
    del output_xls[-4:]
    output_xls += '_noten.xls'
    output_xls = ''.join(output_xls)
    
    #Save the xls object in an xls file(a real file)
    wb.save(output_xls)

    #Creating the new txt file
    #Create an object that can read the new xls file
    rb = open_workbook(output_xls)
    rs = rb.sheet_by_index(0)

    #Create the name of the output text file
    output_txt = list(input_xls)
    del output_txt[-4:]
    output_txt += '_noten.txt'
    output_txt = ''.join(output_txt)

    #Fill the resulting(output) xls
    f = open(output_txt, 'w')
    string = (str(stud_ges)+' Studenten (ohne 5U),'
        +' 4.0 bei ' +str(pmin) +' Punkte,'
        +' 1.0 bei ' +str(pmax) +' Punkte,'
        +' Schnitt ' +mean + ','
        +' Durchgefallen ' +str(durchgefallen) +' (' +failure_rate+'%)'
        +'\n')
    f.write(string)
    for row in range(rs.nrows-1)[2:]:
        line=[]
        line.append(rs.cell(row,3).value)#D
        line.append('{0:15}'.format(unicode(rs.cell(row,5).value).encode("utf-8")))#F
        line.append('{0:15}'.format(unicode(rs.cell(row,6).value).encode("utf-8")))#G
        line.append(rs.cell(row,7).value)#H
        line.append(rs.cell(row,8).value)#I
        line.append(rs.cell(row,15).value)#P
        line.append(rs.cell(row,13).value)#N
        line.append(rs.cell(row,11).value)#L
        line.append('\n')
        string = ', '.join([str(i) for i in line])
        f.write(string)

    return punkte, noten, stud_ges, mean, failure_rate



def hist(ifile, punkte, noten, mean, failure, max_pts, pmax, pmin, bins):
    plt.figure(ifile)
    
    # Plot the histograms
    plt.subplot(2,1,1)
    plt.hist(noten, bins, range=[1,5])
    plt.axvline(4, color='r', linestyle='--', linewidth=3)
    plt.xlabel('Noten')
    plt.ylabel('Studentenzahl')
    plt.title('Durchschnittsnote: ' + str(mean) + ',   '
        + 'Durchfallquote: ' + failure + '%')
    plt.subplot(2,1,2)
    plt.hist(punkte, bins, range=[0,max_pts])
    plt.xlabel('Punkte')
    plt.ylabel('Studentenzahl')
    plt.axvline(pmin, color='r', linestyle='--', linewidth=3)
    plt.axvline(pmax, color='r', linestyle='--', linewidth=3)