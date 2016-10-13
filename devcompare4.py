import tkinter.filedialog
import xlrd
from xlwt import Workbook
import sys
import json
import re
from time import clock
#import pdb

#读取文件名
allSheets={}
def getexcelfiledict(readFile):
    excelFileDict={}
    for iter in readFile:
        excelFileDict[iter]=''
    return excelFileDict

#读取每个文件中的表名
def getsheetsdict(excelFileDict,setting):
    allFilesDict={}
    #excelfiledict从列表改为字典
    if setting:
        for key in excelFileDict.keys():
            wb=xlrd.open_workbook(key)
            oneSheet={}
            for setting_key,setting_value in setting.items():
                if setting_value!='':
                    for sheet in wb.sheets():
                        if sheet.name==setting_key:
                            allCellsinOneSheet=readcelltodict(key,sheet,setting_value)
                            oneSheet[sheet.name]=allCellsinOneSheet
                            allCellsinOneSheet_keys=allCellsinOneSheet.keys()
                            allSheets[sheet.name]=dict.fromkeys(allCellsinOneSheet_keys,'')
                            break
                else:
                    for sheet in wb.sheets():
                        if sheet.name==setting_key:
                            allCellsinOneSheet=readcelltodict(sheet)
                            allCellsinOneSheet_keys=allCellsinOneSheet.keys()
                            try:
                                allSheets_keys=allSheets[sheet.name].keys()
                            except KeyError:
                                allSheets[sheet.name]=dict.fromkeys(allCellsinOneSheet_keys,'')
                            else:
                                allCellinSheet_keys=list(set(allCellsinOneSheet_keys).union(set(allSheets_keys)))
                                allSheets[sheet.name]=dict.fromkeys(allCellinSheet_keys)
                                
                            for allSheet_value_key in allSheets[sheet.name].keys():
                                if allSheet_value_key in allCellsinOneSheet_keys:
                                    pass
                                else:
                                    allCellsinOneSheet[allSheet_value_key]=''
                                    
                            for allFilesDict_key,allFilesDict_value in allFilesDict.items():
                                if sheet.name in allFilesDict_value.keys():
                                    for allSheet_value_key in allSheets[sheet.name].keys():
                                        if allSheet_value_key in allFilesDict_value[sheet.name].keys():
                                            pass
                                        else:
                                             allFilesDict_value[sheet.name][allSheet_value_key]=''
                                else:
                                    pass 
                            oneSheet[sheet.name]=allCellsinOneSheet   
                            break
            allFilesDict[key]=oneSheet
    else:
        for key in excelFileDict.keys():
            wb=xlrd.open_workbook(key)
            oneSheet={}
            for sheet in wb.sheets():
                allCellsinOneSheet=readcelltodict(sheet)
                allCellsinOneSheet_keys=allCellsinOneSheet.keys()
                try:
                    allSheets_keys=allSheets[sheet.name].keys()
                except KeyError:
                    allSheets[sheet.name]=dict.fromkeys(allCellsinOneSheet_keys)
                else:
                    allCellinSheet_keys=list(set(allCellsinOneSheet_keys).union(set(allSheets_keys)))
                    allSheets[sheet.name]=dict.fromkeys(allCellinSheet_keys)
                oneSheet[sheet.name]=allCellsinOneSheet
            allFilesDict[key]=oneSheet
        for allSheet_key,allSheet_value in allSheets.items():
            for allFilesDict_key,allFilesDict_value in allFilesDict.items():
                if allSheet_key in allFilesDict_value.keys():
                    for allSheet_value_key in allSheet_value.keys():
                        if allSheet_value_key in allFilesDict_value[allSheet_key].keys():
                            pass
                        else:
                            allFilesDict_value[allSheet_key][allSheet_value_key]=''
                else:
                    pass
    return allFilesDict

def readcelltodict(key,sheet,setting=None):
    allCellsinOneSheet={}
    xlsmaxrow=65536
    xlsmaxcolumn=256
    xlsxmaxrow=16384
    xlsxmaxcolumn=1048576
    if key[-1]=='s':
        maxrow=xlsmaxrow
        maxcolumn=xlsmaxcolumn
    else:
        maxrow=xlsxmaxrow
        maxcolumn=xlsxmaxcolumn
    if setting:
        zone=convertstrtonumber(setting,maxrow,maxcolumn)
        rowstart=zone['startRow']
        rowend=zone['endColumn']+1
        colstart=zone['startColumn']
        colend=zone['endRow']+1
        for row in range(rowstart,rowend):
            for col in range(colstart,colend):
                try:
                    cellvalue=sheet.cell(row,col).value
                    allCellsinOneSheet[xlrd.cellname(row,col)]=cellvalue
                except IndexError:
                    allCellsinOneSheet[xlrd.cellname(row,col)]=''
    else:
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                allCellsinOneSheet[xlrd.cellname(row,col)]=sheet.cell(row,col).value
    return allCellsinOneSheet

def convertstrtonumber(setting,maxrow,maxcolumn):
    cellRange=setting.split(':')
    zone={}
    for cellAdrress in cellRange:
        rowAlphabet=re.match("^\w[A-Z]*",cellAdrress).group()
        rowAlphabet=rowAlphabet.upper()
        rowNumber=convertalphabettonumber(rowAlphabet)-1
        colNumber=re.search("\d[0-9]*",cellAdrress).group()
        colNumber=int(colNumber)-1
        if rowNumber>maxrow:
            rowNumber=maxrow
        if colNumber>maxcolumn:
            colNumber=maxcolumn
        if zone:
            zone['endRow']=rowNumber
            zone['endColumn']=colNumber
        else:
            zone['startRow']=rowNumber
            zone['startColumn']=colNumber
    return zone
alphabet={'A':1,'B':2,'C':3,'D':4,'E':5,'F':6,'G':7,'H':8,'I':9,'J':10,'K':11,'L':12,'M':13,'N':14,'O':15,'P':16,'Q':17,'R':18,'S':19,'T':20,'U':21,'V':22,'W':23,'X':24,'Y':25,'Z':26}
def convertalphabettonumber(rowAlphabet):
    length=len(rowAlphabet)
    rowNumber=0
    for i in range(0,length):
        rowNumber=rowNumber+alphabet[rowAlphabet[i]]*(26**(length-i-1))
    return rowNumber
    
def readfilesetting(choice='n'):
    setting={}
    if choice=='y':
        try:
            settingFile=open('setting.json','r')
            #pdb.set_trace()
            setting=json.load(settingFile)
        except IOError:
            print('open file error')
        except FileNotFoundError:
            print('can not find setting.json')
    readFile=tkinter.filedialog.askopenfilenames()
    excelFileDict=getexcelfiledict(readFile)
    allFileDict=getsheetsdict(excelFileDict,setting)
    outputDict=compare(allFileDict)
    output(outputDict)
def output(outputDict):
    #pdb.set_trace()
    result=[]
    line={}
    headLineStr='FileName'
    SEPARATE=','
    for outputDict_key,outputDict_value in outputDict.items():
        headLineStr=headLineStr+ SEPARATE + outputDict_key
        if line:
            pass
        else:
            line=dict.fromkeys(outputDict_value.keys(),'')
        for line_key,line_value in line.items():
            if line_value!='':
                tempstr=line_value
                line[line_key]=tempstr+SEPARATE+outputDict_value[line_key]
            else:
                line[line_key]=outputDict_value[line_key]
    book=Workbook()
    sheet1=book.add_sheet('Sheet1')
    headLineStrlist=headLineStr.split(',')
    row=0
    col=0
    for iter in headLineStrlist:
        sheet1.write(row,col,iter)
        col=col+1
    row=row+1
    for line_key,line_value in line.items():
        col=0
        tempstrlist=line_value.split(',')
        sheet1.write(row,col,line_key)
        for tempstrlist_iter in tempstrlist:
            col=col+1
            sheet1.write(row,col,tempstrlist_iter)
        row=row+1
    book.save('result.xls')
    print('output to result.xls')
def compare(allFilesDict):
    outputDict=dict.fromkeys(allFilesDict.keys())
    for ouputDict_key,ouputDict_value in outputDict.items():
        outputDict[ouputDict_key]=dict.fromkeys(allSheets.keys(),'')
    for allSheets_key,allSheets_value in allSheets.items():          #allSheets_key表名，allSheets_value单元格字典
        temp={}
        sameCellinEachFiledDict={}
        for allFilesDict_key,allFilesDict_value in allFilesDict.items():    #allFilesDict_key文件名,allFilesDict_value表字典
            if allSheets_key in allFilesDict_value:     #如果表在此文件的表字典中
                if allSheets_value.keys():
                    for allSheets_value_key in allSheets_value.keys():
                        cell=[]
                        cell.append(allFilesDict_key)
                        cell.append(allFilesDict_value[allSheets_key][allSheets_value_key])
                        if allSheets_value_key in sameCellinEachFiledDict:
                            sameCellinEachFiledDict[allSheets_value_key].append(cell)
                        else:
                            allFileCell=[]
                            allFileCell.append(cell)
                            sameCellinEachFiledDict[allSheets_value_key]=allFileCell
                else:
                    outputDict[allFilesDict_key][allSheets_key]='Empty'
            else:
                outputDict[allFilesDict_key][allSheets_key]='None'
                
        for sameCellinEachFiledDict_key,sameCellinEachFiledDict_value in sameCellinEachFiledDict.items():
            for i in range(0,len(sameCellinEachFiledDict_value)-1):
                for j in range(i+1,len(sameCellinEachFiledDict_value)):
                    if sameCellinEachFiledDict_value[i][1]!=sameCellinEachFiledDict_value[j][1]:
                        for iter in sameCellinEachFiledDict_value:
                            value=outputDict[iter[0]][allSheets_key]
                            if value!='':
                                outputDict[iter[0]][allSheets_key]=value+';'+sameCellinEachFiledDict_key+'单元格:'+str(iter[1])
                            else:
                                outputDict[iter[0]][allSheets_key]=sameCellinEachFiledDict_key+'单元格:'+str(iter[1])
                        break
                    else:
                        if j==len(sameCellinEachFiledDict_value)-1:
                            break
                break
    return outputDict

#主函数
def main():
    while True:
        try:
            choice=input('''Compare excel files with setting.json or not(y or n):
e to exit, h to help\r\n''')
            if choice!='y' and choice!='Y' and choice!='n' and choice!='N' and choice!='e' and choice!='E' and choice!='h' and choice!='H':
                raise ValueError
            else:
                if choice=='y'or choice=='Y' or choice=='n' or choice=='N':
                    readfilesetting(choice)
                    finish=clock()
                    print(finish-start)
                elif choice=='h' or choice=='H':
                    print('''Introduction
This program used to compare excel files.Out put the difference between mutli files.
===================================================================================
Paramaters
y   you can modify setting.json file,to specify sheets and cells you want to compare.
    So Before you input y, you must modify setting.json first.
n   you just select excel files,the program will compare each cell of each sheet of each file.
e   quit program.
h   help.
''')
                elif choice=='e' or choice=='E':
                    sys.exit()
        except ValueError:
            print('Please enter y or n or e')
start=clock()
main()

