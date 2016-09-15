import tkinter.filedialog
import xlrd
import sys
import json
import re
import math
#import pdb
#读取文件名
file={}
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
    #pdb.set_trace()
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
                dd=c.keys()
                try:
                    cc=file[s.name].keys()
                except KeyError:
                    file[s.name]=dict.fromkeys(dd,'')
                else:
                    ll=list(set(dd).union(set(cc)))
                    file[s.name]=dict.fromkeys(ll)
            b[key]=shts
            aa.append(b)
    #pdb.set_trace()
    for iter in aa:
        for key,value in iter.items():
            for key1,value1 in value.items():
                for key2,value2 in file[key1].items():
                    if key2 in value1:
                        pass
                    else:
                        value1[key2]=''
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
    output(d)
def output(d):
    #pdb.set_trace()
    for item in d: 
        str=''      
        for item2 in item:
            str=str+item2
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
                                            cell[aa]=cell[aa]+','+temp[aa]
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
