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
def sheetstodict(a,d):
    b={}
    for key,value in a.items():
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
    return b

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
    c={}
    b={}
    g={}
    notin={}
    i=0
    str=''
    strlist=[]
##    for key,value in a.items():
##        f={}
##        h={}
##        if i==0:
##            for t in value:
##                h[t]=''    
##            b[key]=h
##            i=1
##        else:
##             for t in value:
##                h[t]=''    
##             c[key]=h
    for key,value in a.items():
        if i==0:
            b[key]=value
            c=a
            del c[key]
    
    
    for bkey,bvalue in b.items():
        for key in value.keys():
            for iterkey,itervalue in c.items():
                if key in itervalue.keys():
                    for key1,value1 in itervalue.items():
                        del itervalue[key]
                        break
                else:
                    no={}
                    no[key]=""
                    str=str+iterkey.split('/',)[-1]+key
                    strlist.append(str)
                    if bkey not in notin.keys():
                        notin[bkey]=no
                    else:
                        nn=notin[bkey]
                        if key not in nn.keys():
                            mm=notin[bkey].copy()
                            nn=mm.update(no)
                            notin[bkey]=nn
    d={}
    for key,value in c.items():
        if len(value)!=0:
            d[key]=value
            if key not in notin.keys():
                notin[key] = value
            else:
                for key1 in value.keys():
                    if key1 not in notin[key]:
                        nn = notin[key].copy()
                        nn = nn.update(value[key1])
                        notin[key] = nn
    print(notin)
    if len(d)!=0:
        for key,value in d.items():
            for key1,value1 in a.items():
                if key!=key1:
                    str=key1+"中没有"
                    for key2 in value.keys():
                        str+=key2+'表,'
                    if str[-1]==',':
                        str=str[:-1]
                    strlist.append(str)

    str='\n'.join(strlist)
    return(str)



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

