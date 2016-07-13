import tkinter.filedialog
import xlrd
import sys
import json
import re

#读取文件名
def tupletodict(a):
    b={}
    for item in a:
        b[item]=''
    return b

#读取每个文件中的表名
def sheetstodict(a):
    b={}
    for key,value in a.items():
        wb=xlrd.open_workbook(key)
        shts={}
        for s in wb.sheets():
            c=readcelltodict(s)
            shts[s.name]=c
        b[key]=shts
    return b

#不载入配置文件读取每个表中单元格的值和位置
def readcelltodict(a):
    b={}
    for row in range(a.nrows):
        for col in range(a.ncols):
            b[xlrd.cellname(row,col)]=a.cell(row,col).value
    return b

#载入配置文件读取每个表中单元格的值和位置
def readcelltodict(a,d):
    b={}
    for key,value in d.items():
        if a.name==d.key:
            for row in range(a.nrows):
                for col in range(a.ncols):
                    b[xlrd.cellname(row,col)]=a.cell(row,col).value
    return b

def convertstrtonumber(a):
    b=a.split(':')
    for i in b:
        c=re.match("^\w[A-Z]*",i)
        d=re.match("^\d[0-9]*",i)
def readfilenosetting(a='n'):
    readFile=tkinter.filedialog.askopenfilenames()
    b=tupletodict(readFile)
    c=sheetstodict(b)

    
def readfilesetting(a='y'):
    try:
        d=json.load(open('/setting.json','r'))
    except IOError:
        print('open file error')
    except FileNotFoundError:
        print('can not find setting.json')
    readFile=tkinter.filedialog.askopenfilenames()
    b=tupletodict(readFile)
    c=sheetstodict(b)

#主函数
def main():
    while True:
        try:
            a=input('''Load setting or not(y or n):
e to exit ''')
            if a!='y' and a!='n' and a!='e':
                raise ValueError
            else:
                if a=='y':
                    readfilesetting()
                elif a=='n':
                    readfilenosetting()
                elif a=='e':
                    sys.exit()
        except ValueError:
            print('Please enter y or n or e')


main()

