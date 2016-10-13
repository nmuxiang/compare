import tkinter.filedialog
import xlrd
import sys
import json
import re
from time import clock 
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
                for dkey,dvalue in d.items():
                    if dvalue!='':
                        for s in wb.sheets():
                            if s.name==dkey:
                                c=readcelltodict(s,dvalue)
                                shts[s.name]=c
                                b[key]=shts
                                dd=c.keys()
                                file[dkey]=dict.fromkeys(dd,'')
                                break
                    else:
                        for s in wb.sheets():
                            if s.name==dkey:
                                c=readcelltodict(s)
                                dd=c.keys()
                                try:
                                    cc=file[dkey].keys()
                                except KeyError:
                                    file[dkey]=dict.fromkeys(dd,'')
                                else:
                                    ll=list(set(dd).union(set(cc)))
                                    file[dkey]=dict.fromkeys(ll)
                                for cellkey in file[dkey]:
                                    if cellkey in c:
                                        pass
                                    else:
                                        c[cellkey]=''
                                    for aaitem in aa:
                                        for aaitemkey,aaitemvalue in aaitem.items():
                                            if cellkey in aaitemvalue[s.name]:
                                                pass
                                            else:
                                                aaitemvalue[dkey][cellkey]=''
                                shts[s.name]=c
                                b[key]=shts
                                break
            else:
                for s in wb.sheets():
                    c=readcelltodict(s)
                    dd=c.keys()
                    try:
                        cc=file[s.name].keys()
                    except KeyError:
                        file[s.name]=dict.fromkeys(dd,'')
                    else:
                        ll=list(set(dd).union(set(cc)))
                        file[s.name]=dict.fromkeys(ll)
                    for cellkey in file[s.name]:
                        if cellkey in c:
                            pass
                        else:
                            c[cellkey]=''
                        for aaitem in aa:
                            for aaitemkey,aaitemvalue in aaitem.items():
                                if cellkey in aaitemvalue[s.name]:
                                    pass
                                else:
                                    aaitemvalue[s.name][cellkey]=''
                    shts[s.name]=c
                b[key]=shts
            aa.append(b)
            #for iter in aa:
            #    for key,value in iter.items():
            #        for key1,value1 in value.items():
            #            for key2,value2 in file[key1].items():
            #                if key2 in value1:
            #                    pass
            #                else:
            #                    value1[key2]=''
    return aa   

def readcelltodict(a,d=None):
    b={}
    if d:
        c=convertstrtonumber(d)
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
        for row in range(a.nrows):
            for col in range(a.ncols):
                b[xlrd.cellname(row,col)]=a.cell(row,col).value
    return b

def convertstrtonumber(a):
    b=a.split(':')
    e=[]
    for cellloc in b:
        g=[]
        c=re.match("^\w[A-Z]*",cellloc).group()
        f=convertalphabettonumber(c)
        g.append(f)
        d=re.search("\d[0-9]*",cellloc).group()
        d=int(d)-1
        g.append(d)
        e.append(g)
    return e

alphabet={'A':0,'B':1,'C':2,'D':3,'E':4,'F':5,'G':6,'H':7,'I':8,'J':9,'K':10,'L':11,'M':12,'N':13,'O':14,'P':15,'Q':16,'R':17,'S':18,'T':29,'U':20,'V':21,'W':22,'X':23,'Y':24,'Z':25}
def convertalphabettonumber(a):
    ll=len(a)
    s=0
    for i in range(0,ll):
        s=s+alphabet[a[i]]*(26**(ll-1-i))
    return s
    
def readfilesetting(a='n'):
    d={}
    if a=='y':
        try:
            f=open('setting.json','r')
            #ff=f.read()
            d=json.load(f)
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
            a=input('''Load setting or not(y or n):
e to exit ''')
            if a!='y' and a!='n' and a!='e':
                raise ValueError
            else:
                if a=='y' or a=='n':
                    readfilesetting(a)
                    finish=clock()
                    print(finish-start)
                elif a=='e':
                    sys.exit()
        except ValueError:
            print('Please enter y or n or e')
start=clock()
main()
