import tkinter.filedialog
import xlrd
import sys
import json

def tupletodict(a):
    b={}
    i=1
    for item in a:
        b[i]=item
        i=i+1
    return b

def sheetstodict(a):
    b={}
    for key,value in a.items():
        wb=xlrd.open_workbook(value)
        shts={}
        i=0
        for s in wb.sheets():
           shts[i]=s.name
           i=i+1
           c=readcelltodict(s)
        b[key]=shts
        
    return b

##def readcelltodict(a,d):
##    b={}
##    for key,value in a.items():
##        c=d[value]
##        print(c)
##        for row in range(a.nrows):
##            for col in range(a.ncols):
##                b[xlrd.cellname(row,col)]=a.cell(row,col).value
##    return b
        

try :
    readFile=tkinter.filedialog.askopenfilenames()
    d=json.load(open('/setting.json','r'))
    print(d)
    b=tupletodict(readFile)
    c=sheetstodict(b)
    #d=readcelltodict(c)
except FileNotFoundError:
    sys.exit()

