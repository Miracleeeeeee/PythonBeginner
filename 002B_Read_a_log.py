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

targetdir = "D:/Py_Projects/000_Source_Data/001_Robot_Logs/robotlog 0318"
logdate = targetdir[-4:]
targetfile = "D:/Py_Projects/000_Source_Data/001_Robot_Logs/Output/RobotLog" + " " + logdate + ".xlsx"

def file_name(file_dir):
    for rootdir,subdir,dirfiles in os.walk(file_dir):
        return dirfiles

files = []
for n in file_name(targetdir):  ## create all the source file in the file list #
    fn = targetdir + "/" + n
    files.append(fn)
    
dataset = [['Ref','StartTime','EndTime','Inv_Count','PO_Count']]

# define two functions to obtian the invoice and po number, format is fixed ##
def obtstime(x):
    return x[0:19]

def obtetime(x):
    return x[-74:-55]
    
def obtref(x):
    return x[-4:len(x)]

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

## search the key words and get result for each file##
## focus on one invoice match with one PO ##
## Log format like inv:xxxxxx and then followed by poNbr:xxxxxx ##

for file in files:    
    logcontent = open(file).read()
    str = '2019.*?Received[\s\S]*?Send'
    rst = re.findall(str,logcontent)
    
    ## apply to map and generate a new data set ##
    stimedata = list(map(obtstime,rst))
    etimedata = list(map(obtetime,rst))
    refdata = list(map(obtref,rst))
    invdata = list(map(obtinv,rst))
    podata = list(map(obtpo,rst))
  
    ## put all data into list ##
    for n in range(0,len(refdata)):
        dataset.append([n+1,stimedata[n],etimedata[n],invdata[n],podata[n]])


# Input data into target file ##
workbook = xlsxwriter.Workbook(targetfile)
worksheet = workbook.add_worksheet()
font = workbook.add_format({"font_size":9})
for i in range(len(dataset)):
    for j in range(len(dataset[i])):
        worksheet.write(i, j, dataset[i][j], font)

# Close workbook ##
workbook.close()

