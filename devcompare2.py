import tkinter.filedialog
import xlrd
import sys
import json
import re

#读取文件名
def tupletodict(a):
    bb=[]
    for item in a:
        b={}
        b[item]=''
        bb.append(b)
    return bb

#读取每个文件中的表名
def sheetstodict(a,d):

    aa=[]
    for iter in a:
        for key,value in iter.items():
            b={}
            wb=xlrd.open_workbook(key)
            shts={}
            if d:
                for s in wb.sheets():
                    c=readcelltodict(s,d)
                    shts[s.name]=c
                    b[key]=shts
            for s in wb.sheets():
                c=readcelltodict(s)
                shts[s.name]=c
            b[key]=shts
            aa.append(b)
            break
    return aa

#不载入配置文件读取每个表中单元格的值和位置
#def readcelltodict(a):
#    b={}
#    for row in range(a.nrows):
#        for col in range(a.ncols):
#            b[xlrd.cellname(row,col)]=a.cell(row,col).value
#    return b

#载入配置文件读取每个表中单元格的值和位置
def readcelltodict(a,d=None):
    b={}
    if d:
        for key,value in d.items():
            if a.name==d.key:
                c=convertstrtonumber(d.value)
                for row in range(c[0][1],c[1][1]):
                    for col in range(c[0][0],c[1][0]):
                        b[xlrd.cellname(row,col)]=a.cell(row,col).value
    for row in range(a.nrows):
        for col in range(a.ncols):
            b[xlrd.cellname(row,col)]=a.cell(row,col).value
    return b

def convertstrtonumber(a):
    b=a.split(':')
    e=[]
    for i in b:
        g=[]
        c=re.match("^\w[A-Z]*",i)
        f=convertalphabettonumber(c)
        g.add(f)
        d=re.match("\d[0-9]*",i)
        g.add(d)
        e.add(g)
    return e

alphabet={'A':1,'B':2,'C':3,'D':4,'E':5,'F':6,'G':7,'H':8,'I':9,'J':10,'K':11,'L':12,'M':13,'N':14,'O':15,'P':16,'Q':17,'R':18,'S':19,'T':20,'U':21,'V':22,'W':23,'X':24,'Y':25,'Z':26}
def convertalphabettonumber(a):
    l=len(a)
    s=0
    for i in range(1,l):
        s=s+a[i]*(l-i)*26
    return s

#def readfilenosetting(a='n'):
#    readFile=tkinter.filedialog.askopenfilenames()
#    b=tupletodict(readFile)
#    c=sheetstodict(b)

    
def readfilesetting(a='n'):
    d={}
    if a=='y':
        try:
            d=json.load(open('/setting.json','r'))
        except IOError:
            print('open file error')
        except FileNotFoundError:
            print('can not find setting.json')
    readFile=tkinter.filedialog.askopenfilenames()
    b=tupletodict(readFile)
    c=sheetstodict(b,d)
    d=compare(c)
    print(d)
def compare(a):
    g={}
    notin=[]
    str=''
    strlist=[]
    for i in range(0,(len(a)-1)):
        b=a[i]
        c=a[i+1]   
        for bkey,bvalue in b.items():       #bkey文件名，bvalue表名字典
            for bvaluekey,bvaluevalue in bvalue.items():      #bvaluekey表名，bvaluevalue单元格字典
                for ckey,cvalue in c.items():       #ckey文件名，cvalue表名字典
                    if bvaluekey in cvalue:
                        if cvalue[bvaluekey]==bvalue[bvaluekey]:
                            pass
                        else:         
                            for bvaluevaluekey,bvaluevaluevalue in bvaluevalue.items():     #bvaluevaluekey单元格名，bvaluevaluevalue单元格值
                                for cvaluevaluekey,cvaluevaluevalue in cvalue[bvaluekey].items():
                                    if bvaluevaluekey in cvalue[bvaluekey]:
                                        if bvaluevaluevalue==cvalue[bvaluekey][bvaluevaluekey]:
                                            pass
                                        else:
                                            str=bkey+'文件'+bvaluekey+'表'+cvaluevaluekey+'不等于'+ckey+'文件'+bvaluekey+'表'+cvaluevaluekey
                                            notin.append(str)
                                            #nn={}
                                            #nnn={}
                                            #nn[bvaluevaluekey]=cvalue[bvaluekey][bvaluevaluekey]
                                            #nnn[bvaluekey]=nn
                                            #notin[ckey]=nn
                                    else:
                                        str=ckey+'文件'+bvaluekey+'表中没有'+bvaluevaluekey
                                        notin.append(str)
                                        #nn={}
                                        #nnn={}
                                        #nn[bvaluevaluekey]=''
                                        #nnn[bvaluekey]=nn
                                        #notin[ckey]=nnn
                                    if cvaluevaluekey not in bvaluevalue:
                                        str=bkey+'文件'+bvaluekey+'表中没有'+cvaluevaluekey
                                        notin.append(str)
                                        #nn={}
                                        #nnn={}
                                        #nn[bvaluevaluekey]=''
                                        #nnn[bvaluekey]=nn
                                        #notin[ckey]=nnn
                    else:
                        str=ckey+'文件中没有'+bvaluekey+'表'
                        notin.append(str)
                        nn={}
                        nn[bvaluekey]=''
                        notin[ckey]=nn
    return notin


#主函数
def main():
    while True:
        try:
            a=input('''Load setting or not(y or n):
e to exit ''')
            if a!='y' and a!='n' and a!='e':
                raise ValueError
            else:
                if a=='y' or a=='n':
                    readfilesetting()
                elif a=='e':
                    sys.exit()
        except ValueError:
            print('Please enter y or n or e')


main()

