import tkinter.filedialog
import xlrd
import sys
import json
import re
#import pdb
#读取文件名
allSheets={}
def getexcelfiledict(readFile):
    excelFileDict={}
    for iter in readFile:
        excelFileDict[item]=''
    return excelFileDict

#读取每个文件中的表名
def getsheetsdict(excelFileDict,setting):
    allFiles=[]
    #excelfiledict从列表改为字典
    if setting:
        for key in excelFileDict.keys():
            oneFile={}
            wb=xlrd.open_workbook(key)
            oneSheet={}
            for setting_key,setting_value in setting.items():
                if setting_value!='':
                    for sheet in wb.sheets():
                        if sheet.name==setting_key:
                            allCellsinOneSheet=readcelltodict(sheet,setting_value)
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
                            for oneFile_key,oneFile_value in oneFile.items():
                                if sheet.name in oneFile_value.keys():
                                    for allSheet_value_key in allSheets[sheet.name].keys():
                                        if allSheet_value_key in oneFile_value[sheet.name].keys():
                                            pass
                                        else:
                                             oneFile_value[sheet.name][allSheet_value_key]=''
                                else:
                                    pass 
                            oneSheet[sheet.name]=allCellsinOneSheet   
                            break
            oneFile[key]=oneSheet
    else:
        for key in excelFileDict.keys():
            oneFile={}
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
            oneFile[key]=oneSheet
        for allSheet_key,allSheet_value in allSheets.items():
            for oneFile_key,oneFile_value in oneFile.items():
                if allSheet_key in oneFile_value.keys():
                    for allSheet_value_key in allSheet_value.keys():
                        if allSheet_value_key in oneFile_value[allSheet_key].keys():
                            pass
                        else:
                            oneFile_value[allSheet_key][allSheet_value_key]=''
                else:
                    pass
    allFiles.append(oneFile)
    return allFiles

def readcelltodict(sheet,setting=None):
    allCellsinOneSheet={}
    if setting:
        c=convertstrtonumber(setting)
        rowstart=c[0][1]
        rowend=c[1][1]+1
        colstart=c[0][0]
        colend=c[1][0]+1
        for row in range(rowstart,rowend):
            for col in range(colstart,colend):
                try:
                    cellvalue=a.cell(row,col).value
                    b[xlrd.cellname(row,col)]=cellvalue
                except IndexError:
                    b[xlrd.cellname(row,col)]=''
    else:
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                allCellsinOneSheet[xlrd.cellname(row,col)]=sheet.cell(row,col).value
    return allCellsinOneSheet

def convertstrtonumber(setting):
    cellRange=setting.split(':')
    e=[]
    for cellAdrress in cellRange:
        g=[]
        rowAlphabet=re.match("^\w[A-Z]*",cellAdrress).group()
        f=convertalphabettonumber(rowAlphabet)
        g.append(f)
        d=re.search("\d[0-9]*",cellAdrress).group()
        d=int(d)-1
        g.append(d)
        e.append(g)
    return e

alphabet={'A':0,'B':1,'C':2,'D':3,'E':4,'F':5,'G':6,'H':7,'I':8,'J':9,'K':10,'L':11,'M':12,'N':13,'O':14,'P':15,'Q':16,'R':17,'S':18,'T':29,'U':20,'V':21,'W':22,'X':23,'Y':24,'Z':25}
def convertalphabettonumber(rowAlphabet):
    length=len(rowAlphabet)
    rowNumber=0
    for i in range(0,length):
        rowNumber=rowNumber+alphabet[rowAlphabet[i]]*(26**(length-1-i))
    return rowNumber
    
def readfilesetting(choice='n'):
    setting={}
    if choice=='y':
        try:
            settingFile=open('setting.json','r')
            #ff=f.read()
            setting=json.load(settingFile)
        except IOError:
            print('open file error')
        except FileNotFoundError:
            print('can not find setting.json')
    readFile=tkinter.filedialog.askopenfilenames()
    excelFileDict=getexcelfiledict(readFile)
    allFileDict=getsheetsdict(excelFileDict,setting)
    result=compare(allfiledict)
    output(result)
def output(result):
    #pdb.set_trace()
    for item in d: 
        str=''      
        for item2 in range(0,len(item)):
            if item2!=len(item) and item2!=0:
                str=str+';'+item[item2]
            else:
                str=str+item[item2]
        print(str)
def compare(a):
    notin=[]
    filename=[]
    filename.append('File Name')
    yn=True
    #pdb.set_trace()
    for filekey,filevalue in file.items():          #filekey表名，filevalue单元格字典
        diff=[]                             #记录表之间的不同
        diff.append(filekey)
        jj=0
        cell=[]      #记录每个表中所有单元格的不同
        if len(filevalue)!=0:
            for filevaluekey in filevalue.keys():       #filevaluekey单元格名                
                temp=[]                     #每个单元格的不同
                ii=0
                for i in a:
                    ii=ii+1
                    for key,value in i.items():       #key文件名，value表名字典
                        if yn==True:                    #产生表名行
                            filename.append(key)
                        if filekey in value:
                           temp.append(filevaluekey+'单元格'+str(value[filekey][filevaluekey]))
                           break
                        else:
                            temp.append('没有'+filekey)
                    if ii==len(a):
                        if yn==True:
                            yn=False
                            notin.append(filename)
                        else:
                            pass
                        for s in range(0,len(temp)-1):
                            for r in range(s+1,len(temp)):
                                if temp[s]!=temp[r]:
                                    for aa in range(0,len(temp)):
                                        if cell:
                                            if cell[aa][:2]!='没有':
                                                cell[aa]=cell[aa]+','+temp[aa]
                                            else:
                                                pass
                                        else:
                                            cell=temp
                                            break
                                    break
                                else:
                                    if s==len(temp)-2 and r==len(temp)-1:
                                        break
                            break
                        break
                    continue
            for ab in cell:
                diff.append(ab)
            notin.append(diff)
        else:
            temp1=[]
            for i in a:
                jj=jj+1
                for key,value in i.items():       #key文件名，value表名字典
                    if yn==True:                    #产生表名行
                        filename.append(key)
                    if filekey in value:
                        temp1.append('空表')
                        break
                    else:
                        temp1.append('没有'+filekey)
                if jj==len(a):
                    if yn==True:
                        yn=False
                        notin.append(filename)
                    else:
                        pass
                    for s in range(0,len(temp1)-1):
                        for r in range(s+1,len(temp1)):
                            if temp1[s]!=temp1[r]:
                                for aa in range(0,len(temp1)):
                                    if cell:
                                        cell[aa]=cell[aa]+','+temp1[aa]
                                    else:
                                        cell=temp1
                                        break
                                break
                            else:
                                if s==len(temp1)-2 and r==len(temp1)-1:
                                    break
                        break
                    break
                continue
            if cell:
                for ab in cell:
                    diff.append(ab)
                notin.append(diff)
    return notin


#主函数
def main():
    while True:
        try:
            choice=input('''Compare excel files with setting.json or not(y or n):
e to exit, h to help\r\n''')
            if choice!='y' and choice!='n' and choice!='e' and choice!='h':
                raise ValueError
            else:
                if choice=='y' or choice=='n':
                    readfilesetting(choice)
                elif choice=='h':
                    print('''Introduction
This program used to compare excel files.Out put the difference between mutli files.
===================================================================================
Paramaters
y   y means you can modify setting.json file,to specify sheets and cells you want to compare.
    So Before you input y, you must modify setting.json first.
n   n means you just select excel files,the program will compare each cell of each sheet of each file.
e   e means quit program.
h   h means help.
''')
                elif choice=='e':
                    sys.exit()
        except ValueError:
            print('Please enter y or n or e')
main()
