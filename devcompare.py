import tkinter.filedialog
import xlrd

readFile=tkinter.filedialog.askopenfilenames()
a=0
m=[]
n=[]
b=None
for i in readFile:
    wb=xlrd.open_workbook(i)
    if a==0:
        b=wb
        a=1
    else:
        m.append(wb)
c=0
d=0
for sht in b.sheets():
    for i in m:
        for shti in i.sheets():
            if sht.name==shti.name:
                del b.sheets()[c]
                del i.sheets()[d]
                print(b.sheets())
                print(i.sheets())
        d=d+1
    c+=1
