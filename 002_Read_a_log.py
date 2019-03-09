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

targetdir = "D:/Py_Projects/000_Source_Data/001_Robot_Logs/robotlog 0308"
targetfile = "D:/Py_Projects/000_Source_Data/001_Robot_Logs/Output/RobotLog.xlsx"

def file_name(file_dir):
    for rootdir,subdir,dirfiles in os.walk(file_dir):
        return dirfiles

files = []
for n in file_name(targetdir):  ## create all the source file in the file list #
    fn = targetdir + "/" + n
    files.append(fn)
    
dataset = [['Invoice_Num','PO_Num']]
invoicecount = [['File_Index','Invoice_Count']]
fc=1  ## this is used to count the invoice number ##

# define two functions to obtian the invoice and po number, format is fixed ##
def obtinv(x):
    return x[7:15]
    
def obtpo(x):
    return x[-11:len(x)-1]

## search the key words and get result for each file##
## focus on one invoice match with one PO ##
## Log format like inv:xxxxxx and then followed by poNbr:xxxxxx ##

for file in files:    
    logcontent = open(file).read()
    str = '"inv"[^\s]*"payCode"[^\s]*"poNbr":"\d\d\d\d\d\d\d\d\d\d"'
    rst = re.findall(str,logcontent)
    
    ## apply to map and generate a new data set ##
    invdata = list(map(obtinv,rst))
    podata = list(map(obtpo,rst))
    invoicecount.append([fc,len(invdata)])
    fc+=1 ## for counting invoice num to reconclie #

    for n in range(0,len(invdata)):
        dataset.append([invdata[n],podata[n]])


# Input data into target file ##
workbook = xlsxwriter.Workbook(targetfile)
worksheet = workbook.add_worksheet()
font = workbook.add_format({"font_size":9})
for i in range(len(dataset)):
    for j in range(len(dataset[i])):
        worksheet.write(i, j, dataset[i][j], font)

# Close workbook ##
workbook.close()

print(invoicecount)
