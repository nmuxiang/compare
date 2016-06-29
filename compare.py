import tkinter.filedialog
import xlrd
import sys

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
        b[key]=shts
    return b
try:
    readFile=tkinter.filedialog.askopenfilenames()
    b=tupletodict(readFile)
    c=sheetstodict(b)
    print(c)
except FileNotFoundError:
    sys.exit()

