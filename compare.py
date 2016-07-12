import tkinter.filedialog
import xlrd
import sys
import json

def tupletodict(a):
    b={}
    for item in a:
        b[item]=''
    return b

def sheetstodict(a):
    b={}
    for key,value in a.items():
        wb=xlrd.open_workbook(key)
        shts={}
        for s in wb.sheets():
           shts[s.name]=''
        b[key]=shts
    return b

def readcelltodict(a):
    b={}
    for key,value in a.items():
        for row in range(a.nrows):
            for col in range(a.ncols):
                b[xlrd.cellname(row,col)]=a.cell(row,col).value
    return b

def readcelltodict(a,d):
    b={}
    for key,value in a.items():
        c=d[value]
        print(c)
        for row in range(a.nrows):
            for col in range(a.ncols):
                b[xlrd.cellname(row,col)]=a.cell(row,col).value
    return b
        
def readfilenosetting(a='n'):
    readFile=tkinter.filedialog.askopenfilenames()
    b=tupletodict(readFile)
    c=sheetstodict(b)
    d=readcelltodict(c)
    
def readfilesetting(a='y'):
    if a=='n':
        pass
    elif a=='y':
        try:
            d=json.load(open('/setting.json','r'))
        except IOError:
            print('open file error')
        except FileNotFoundError:
            print('can not find setting.json')

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

