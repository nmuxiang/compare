import tkinter.filedialog
import xlrd
from xlwt import Workbook
import sys
import re
import cProfile
import time
import collections
import string
#all sheets aggregation of excel files
allSheets={}

#读取文件名
def getexcelfiledict(readFile):
    excelFileDict={}
    for iter in readFile:
        excelFileDict[iter]=''
    return excelFileDict

#读取每个文件中的表名
def getsheetsdict(excelFileDict,setting):
    allFilesDict={}
    #pdb.set_trace()
    #excelfiledict从列表改为字典
    if setting:
        for key in excelFileDict.keys():
            wb=xlrd.open_workbook(key)
            oneSheet={}
            for setting_key,setting_value in setting.items():
                if setting_value!='':
                    zone=convertstrtonumber(key,setting_value)
                    for sheet in wb.sheets():
                        if sheet.name==setting_key:
                            allCellsinOneSheet=readcelltodict(key,sheet,zone)
                            oneSheet[sheet.name]=allCellsinOneSheet
                            allCellsinOneSheet_keys=allCellsinOneSheet.keys()
                            allSheets[sheet.name]=dict.fromkeys(allCellsinOneSheet_keys,'')
                            break
                else:
                    for sheet in wb.sheets():
                        if sheet.name==setting_key:
                            allCellsinOneSheet=readcelltodict(key,sheet)
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
                allCellsinOneSheet=readcelltodict(key,sheet)
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

def readcelltodict(key,sheet,zone=None):
    allCellsinOneSheet={}
    if zone:
        rowstart=zone['startRow']
        rowend=zone['endRow']+1
        colstart=zone['startColumn']
        colend=zone['endColumn']+1
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

def convertstrtonumber(key,setting_value):
    xlsmaxrow=65536
    xlsmaxcolumn=256
    xlsxmaxrow=16384
    xlsxmaxcolumn=1048576
    if key[-1]=='s' or key[-1]=='S':
        maxrow=xlsmaxrow
        maxcolumn=xlsmaxcolumn
    else:
        maxrow=xlsxmaxrow
        maxcolumn=xlsxmaxcolumn
    cellRange=setting_value.split(':')
    zone={}
    for cellAdrress in cellRange:
        #pdb.set_trace()
        colAlphabet=re.match("^\w[a-z,A-Z]*",cellAdrress).group()
        colAlphabet=colAlphabet.upper()
        colNumber=convertalphabettonumber(colAlphabet)-1
        rowNumber=re.search("\d[0-9]*",cellAdrress).group()
        rowNumber=int(rowNumber)-1
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
def convertalphabettonumber(colAlphabet):
    length=len(colAlphabet)
    colNumber=0
    for i in range(0,length):
        colNumber=colNumber+alphabet[colAlphabet[i]]*(26**(length-i-1))
    return colNumber
    
def settingtojson(settingtext):
    sheetcellsdict={}
    #re1=''
    for setting_element in settingtext:
        if "," in setting_element and ":" in setting_element:
            sheetandcells=setting_element.split(',')
            sheet=sheetandcells[0]
            cells=sheetandcells[1]
            sheetcellsdict[sheet]=cells
        elif ":" not in setting_element:
            if setting_element[-1:]=='\n' :
                setting_element=setting_element[:-1]
            else:
                pass
            if setting_element[-1:]==',' :
                setting_element=setting_element[:-1]
                sheetcellsdict[setting_element]=""
            elif "," not in setting_element:
                sheetcellsdict[setting_element]=""
           # print("Format error. Please check setting.txt")   
    setting=sheetcellsdict
    return setting
def readfilesetting(choice='n'):
    setting={}
    if choice=='y':
        try:
            readFile=tkinter.filedialog.askopenfilename()
            settingFile=open(readFile,'r')
            settingtext=settingFile.readlines()
            setting=settingtojson(settingtext)
            #setting=json.load(setting)
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
    result=[]
    line={}
    SEPARATE=','
    headLineStr='Sheet'+SEPARATE+'Cell'
    for outputDict_key,outputDict_value in outputDict.items():
        headLineStr=headLineStr+ SEPARATE + outputDict_key
        if line:
            pass
        else:
            line=dict.fromkeys(outputDict_value.keys(),'')
        for outputDict_value_key,outputDict_value_value in outputDict_value.items():
            if line[outputDict_value_key]!='':
                for outputDict_value_value_key,outputDict_value_value_value in outputDict_value_value.items():
                    if outputDict_value_value_key in line[outputDict_value_key]:
                        temp=[]
                        temp= line[outputDict_value_key][outputDict_value_value_key]
                        temp.extend(outputDict_value_value_value)
                        line[outputDict_value_key][outputDict_value_value_key]=temp
                    else:
                        line[outputDict_value_key][outputDict_value_value_key]=outputDict_value_value_value
            else:
                line[outputDict_value_key]=outputDict_value_value
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
        sheet1.write(row,col,line_key)
        row=row+1
        colsheet=col+1
        for line_value_key,line_value_value in line_value.items():
            if line_value_key=='NoneorEmpty':
                col=colsheet+1
            else:
                col=colsheet
                sheet1.write(row,col,line_value_key)
                col=col+1
            for iter in line_value_value:
                sheet1.write(row,col,iter)
                col=col+1
            row=row+1
    book.save('result.xls')
    print("output to result.xls")


def sort(tempoutput):
    cellslist=tempoutput.keys()
    sortlist=[]
    for iter in cellslist:
        colAlphabet=re.match("^\w[a-z,A-Z]*",iter).group()
        rowNumber=int(re.search("\d[0-9]*",iter).group())
        tup=(colAlphabet,rowNumber)
        sortlist.append(tup)
    sortlist.sort(key=lambda x:(x[0],x[1]))
    sorteddict=collections.OrderedDict()
    for iter in sortlist:
        strkey=iter[0]+str(iter[1])
        sorteddict[strkey]=tempoutput[strkey]
    return sorteddict

def compare(allFilesDict):
    outputDict=dict.fromkeys(allFilesDict.keys())
    for ouputDict_key,ouputDict_value in outputDict.items():
        outputDict[ouputDict_key]=dict.fromkeys(allSheets.keys(),{})
    for allSheets_key,allSheets_value in allSheets.items():          #allSheets_key表名，allSheets_value单元格字典
        temp={}
        sameCellinEachFiledDict={}
        for allFilesDict_key,allFilesDict_value in allFilesDict.items():    #allFilesDict_key文件名,allFilesDict_value表字典
            if allSheets_key in allFilesDict_value:     #如果表在此文件的表字典中
                if allSheets_value.keys():
                    for allSheets_value_key in allSheets_value.keys():
                        cell=[]
                        #################
                        cell.append(allFilesDict_key)
                        cell.append(allFilesDict_value[allSheets_key][allSheets_value_key])
                        if allSheets_value_key in sameCellinEachFiledDict:
                            sameCellinEachFiledDict[allSheets_value_key].append(cell)
                        else:
                            allFileCell=[]
                            allFileCell.append(cell)
                            sameCellinEachFiledDict[allSheets_value_key]=allFileCell
                        #################
                else:
                    temp={}
                    temp['NoneorEmpty']=['Empty']
                    outputDict[allFilesDict_key][allSheets_key]=temp
            else:
                temp={}
                temp['NoneorEmpty']=['None']
                outputDict[allFilesDict_key][allSheets_key]=temp

        for sameCellinEachFiledDict_key,sameCellinEachFiledDict_value in sameCellinEachFiledDict.items():
            for i in range(0,len(sameCellinEachFiledDict_value)-1):
                for j in range(i+1,len(sameCellinEachFiledDict_value)):
                    if sameCellinEachFiledDict_value[i][1]!=sameCellinEachFiledDict_value[j][1]:
                        for iter in sameCellinEachFiledDict_value:
                            temp=collections.OrderedDict()
                            templist=[]
                            templist.append(iter[1])
                            temp[sameCellinEachFiledDict_key]=templist
                            tempoutput=outputDict[iter[0]][allSheets_key]
                            if tempoutput!={}:
                               templist=[]
                               templist.append(iter[1])
                               temp[sameCellinEachFiledDict_key]=templist
                               tempoutput[sameCellinEachFiledDict_key]=templist
                               ##############################################################
                               tempoutput=sort(tempoutput)
                               outputDict[iter[0]][allSheets_key]=tempoutput
                            else:
                                outputDict[iter[0]][allSheets_key]=temp
                        break
                    else:
                        if j==len(sameCellinEachFiledDict_value)-1:
                            break
                break
    #outputDict=sorted(outputDict.items(),key=lambda d:d[0],reverse=False)
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
main()

