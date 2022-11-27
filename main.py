from openpyxl import Workbook, load_workbook
import os
from datetime import timedelta,datetime

origcurdir=os.curdir
timelogdir=os.path.join(os.getcwd(),'Time Tracker','timelog')
timelogfile=os.path.join(timelogdir,'TimeLog.xlsx')
print(timelogdir)
print(timelogfile)
print(os.path.normpath(timelogfile))
os.chdir(os.path.normpath(timelogdir))
wb = load_workbook(timelogfile,data_only=True)
ws = wb.worksheets[0]

for row in ws.iter_rows(min_row=1, max_row=1,values_only=True):
    rowheader=list(row)
print(rowheader)
rownum=0
allrows={}
for row in ws.iter_rows(min_row=2,values_only=True):
    rowid=str(rownum)
    rowdata=list(row)
    zippedrow={rowheader:rowdata for rowheader,rowdata in zip(rowheader,rowdata)}
    ##allrows.append(zippedrow)
    edur=zippedrow['Clock Out']-zippedrow['Clock In']
    zippedrow['Event duration']=edur.seconds/3600
    allrows[rowid]=zippedrow
    print (rowid,'=',zippedrow)
  
    ##allrows{rownum}=zippedrow
    rownum+=1
print (allrows)




