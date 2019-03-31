# -*- coding: utf-8 -*-
"""
Created on Sat Mar  9 12:41:36 2019

@author: ericliang
"""

## Set Environment and import packages ##
import re
import xlsxwriter
import os

## Set the log files ##

''' Select one log files, if needed we can use this for debug
logfile = "D:/Py_Projects/000_Source_Data/001_APT_Logs/wapprobotlog 0307/host-robot-service.log" 
'''

targetdir = "D:/Py_Projects/000_Source_Data/001_Robot_Logs/robotlog 0322"
logdate = targetdir[-4:]
targetfile = "D:/Py_Projects/000_Source_Data/001_Robot_Logs/Output/RobotLog" + " " + logdate + ".xlsx"
fulllogfile = "D:/Py_Projects/000_Source_Data/001_Robot_Logs/robotlog 0322/RobotLog" + " " + logdate + ".txt"

## Combine all logs ##
if(os.path.exists(fulllogfile)):
    os.remove(fulllogfile) ## clean up the full log 

def file_name(file_dir):
    for rootdir,subdir,dirfiles in os.walk(file_dir):
        return dirfiles

files = []
for n in file_name(targetdir):  ## create all the source file in the file list #
    fn = targetdir + "/" + n
    files.append(fn)
    
dataset = [['Ref','StartTime','EndTime','Inv_Count','PO_Count']]

flog = open(fulllogfile,'a+')
for file in files:    
    logcontent = open(file).read()
    flog.write(logcontent+"\n")
flog.close()

fulllog = open(fulllogfile).read()
    
## Generate the logcontent ##
## Use "Recevied" part for starting time and PO/INV counts ##
str1 = '2019.*?Received[\s\S]*?\[2019'
rst1 = re.findall(str1,fulllog)

## Use "Send" part for ending time 
str2 = '2019.*?Send:[\s]\[{"id":"\d\d\d\d\d\d\d\d\d\d"'
rst2 = re.findall(str2,fulllog)

# define couple functions to obtian the invoice and po number, format is fixed ##
def obtid(x):
    fid = []
    allids = re.findall('"id":"\d\d\d\d\d\d\d\d\d\d"',x)
    if len(allids) != 0:
        fid = allids[0][-11:-1]
    else:
        fid = "None"
    return fid

def obtstime(x):
    return x[0:19]

def obtedata(x):
    return x[0:19]

def obtinv(x):
    invtemp = re.findall('"inv":"\d\d\d\d\d\d\d\d"',x)
    if len(invtemp) <= 1:
        invcount = 1
    elif invtemp[0] == invtemp[1]:
        invcount = 1
    else:
        invcount = len(invtemp)
    return invcount

def obtpo(x):
    potemp = re.findall('"poNbr":"\d\d\d\d\d\d\d\d\d\d"',x)
    if len(potemp) <= 1:
        pocount = 1
    elif potemp[0] == potemp[1]:
        pocount = 1
    else:
        pocount = len(potemp)
    return pocount

def obteid(x):
    return x[-11:-1]

def obtetime(x):
    return x[0:19]

## apply map to obtain "received" part ##
iddata = list(map(obtid,rst1))
stimedata = list(map(obtstime,rst1))
invdata = list(map(obtinv,rst1))
podata = list(map(obtpo,rst1))

## apply function and create a dictionary to obtain "Send" part ##
eid = list(map(obteid,rst2))
etime = list(map(obtetime,rst2))

etimedit = {}
for n in range(0,len(eid)):
    etimedit[eid[n]] = etime[n]
  
## put all data into list ##
for n in range(0,len(iddata)):
    dataset.append([iddata[n],
                    stimedata[n],
                    etimedit.get(iddata[n],'None'),
                    invdata[n],
                    podata[n]])

''' Data Extraction '''
    
# Input data into target file ##
workbook = xlsxwriter.Workbook(targetfile)
worksheet = workbook.add_worksheet()
font = workbook.add_format({"font_size":9})
for i in range(len(dataset)):
    for j in range(len(dataset[i])):
        worksheet.write(i, j, dataset[i][j], font)

# Close workbook ##
workbook.close()
