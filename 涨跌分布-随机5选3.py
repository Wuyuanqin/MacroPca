import numpy as np
import xlrd
import xlwt
import matplotlib.pyplot as plt
import math

def read_excel(name,a):
    table = name.sheets()[0]
    row = table.nrows#得到每个表格的行数
    col = table.ncols#得到每个表格的列数
    row=120+a#10年的数据,1年检验
    #print(row)
    data_list,t,data_list1,t1 = [],[],[],[]
    for i in range(a,row):
        t.append(table.cell(i,7).value)
        for j in range(1,col-1-1):
            data_list.append(table.cell(i,j).value)
    array = np.array(data_list).reshape(120,col-1-1-1)#将所有数据存进矩阵
    T=np.array(t).reshape(120,1)
    #检验数据，一年
    for i in range(row,row+12):
        t1.append(table.cell(i,7).value)
        for j in range(1,col-1-1):
            data_list1.append(table.cell(i,j).value)
    array1 = np.array(data_list1).reshape(12,col-1-1-1)#将所有数据存进矩阵
    T1=np.array(t1).reshape(12,1)
    return row,col,array,T,array1,T1

def zhd(t,d,w,m,n):
    if t[d]>=0:
        w.append(m)
    else:
        w.append(n)
    return w,m,n

def sum5(e,t,i,w):
    if e.index(0)==0:
        q=e[1:len(e)]
        if q.index(0)==0:
            w,m,n=zhd(t,i,w,12,13)
        elif q.index(0)==1:
            w,m,n=zhd(t,i,w,14,15)
        elif q.index(0)==2:
            w,m,n=zhd(t,i,w,16,17)
        elif q.index(0)==3:
            w,m,n=zhd(t,i,w,18,19)
    elif e.index(0)==1:
        q=e[2:len(e)]
        if q.index(0)==0:
            w,m,n=zhd(t,i,w,20,21)
        elif q.index(0)==1:
            w,m,n=zhd(t,i,w,22,23)
        elif q.index(0)==2:
            w,m,n=zhd(t,i,w,24,25)
    elif e.index(0)==2:
        q=e[3:len(e)]
        if q.index(0)==0:
            w,m,n=zhd(t,i,w,26,27)
        elif q.index(0)==1:
            w,m,n=zhd(t,i,w,28,29)
    elif e.index(0)==3:
        w,m,n=zhd(t,i,w,30,31)
    return w,m,n

def sum6(e,t,i,w):
    if e.index(1)==0:
        q=e[1:len(e)]
        if q.index(1)==0:
            w,m,n=zhd(t,i,w,32,33)
        elif q.index(1)==1:
            w,m,n=zhd(t,i,w,34,35)
        elif q.index(1)==2:
            w,m,n=zhd(t,i,w,36,37)
        elif q.index(1)==3:
            w,m,n=zhd(t,i,w,38,39)
    elif e.index(1)==1:
        q=e[2:len(e)]
        if q.index(1)==0:
            w,m,n=zhd(t,i,w,40,41)
        elif q.index(1)==1:
            w,m,n=zhd(t,i,w,42,43)
        elif q.index(1)==2:
            w,m,n=zhd(t,i,w,44,45)
    elif e.index(1)==2:
        q=e[3:len(e)]
        if q.index(1)==0:
            w,m,n=zhd(t,i,w,46,47)
        elif q.index(1)==1:
            w,m,n=zhd(t,i,w,48,49)
    elif e.index(1)==3:
        w,m,n=zhd(t,i,w,50,51)
    return w,m,n

def sum3(e,t,i,w):
    #print(e)
    if e.index(0)==0:
        w,m,n=zhd(t,i,w,2,3)
    elif e.index(0)==1:
        w,m,n=zhd(t,i,w,4,5)
    elif e.index(0)==2:
        w,m,n=zhd(t,i,w,6,7)
    elif e.index(0)==3:
        w,m,n=zhd(t,i,w,8,9)
    if e.index(0)==4:
        w,m,n=zhd(t,i,w,10,11)
    return w,m,n

def sum2(e,t,i,w):
    if e.index(0)==0:
        q=e[1:len(e)]
        if q.index(0)==0:
            w,m,n=zhd(t,i,w,10,11)
        elif q.index(0)==1:
            w,m,n=zhd(t,i,w,12,13)
        elif q.index(0)==2:
            w,m,n=zhd(t,i,w,14,15)
    elif e.index(0)==1:
        q=e[2:len(e)]
        if q.index(1)==0:
            w,m,n=zhd(t,i,w,16,17)
        elif q.index(1)==1:
            w,m,n=zhd(t,i,w,18,19)
    elif e.index(0)==2:
        w,m,n=zhd(t,i,w,20,21)
    return w,m,n

def sum1(e,t,i,w):
    if e.index(1)==0:
        #w=zhd(t,i,w,2,3)
        w,m,n=zhd(t,i,w,22,23)
    elif e.index(1)==1:
        #w=zhd(t,i,w,4,5)
        w,m,n=zhd(t,i,w,24,25)
    elif e.index(1)==2:
        w,m,n=zhd(t,i,w,26,27)
    elif e.index(1)==3:
        w,m,n=zhd(t,i,w,28,29)
    return w,m,n

def numsignal(l,w,t,i,e):
    if len(l)==1:#一信号
        w,m,n=zhd(t,i,w,0,1)
    elif len(l)==2:#两个信号
        if sum(e)==2:
            w,m,n=zhd(t,i,w,0,1)
        elif sum(e)==1:
            if e.index(0)==0:
                w,m,n=zhd(t,i,w,2,3)
            elif e.index(0)==1:
                w,m,n=zhd(t,i,w,4,5)
        elif sum(e)==0:
            w,m,n=zhd(t,i,w,6,7)
    elif len(l)==3:#三个信号
        if sum(e)==3:
            w,m,n=zhd(t,i,w,0,1)
        elif sum(e)==2:
            if e.index(0)==0:
                w,m,n=zhd(t,i,w,2,3)
            elif e.index(0)==1:
                w,m,n=zhd(t,i,w,4,5)
            elif e.index(0)==2:
                w,m,n=zhd(t,i,w,6,7)
        elif sum(e)==1:
            if e.index(1)==0:
                w,m,n=zhd(t,i,w,8,9)
            elif e.index(1)==1:
                w,m,n=zhd(t,i,w,10,11)
            elif e.index(1)==2:
                w,m,n=zhd(t,i,w,12,13)
        elif sum(e)==0:
            w,m,n=zhd(t,i,w,14,15)
    elif len(l)==4:#四个信号
        if sum(e)==4:
            w,m,n=zhd(t,i,w,0,1)
        elif sum(e)==3:
            w,m,n=sum3(e,t,i,w)
        if sum(e)==2:
            #w=zhd(t,i,w,0,1)
            w,m,n=sum2(e,t,i,w)
        elif sum(e)==1:
            w,m,n=sum1(e,t,i,w)
        elif sum(e)==0:
            #w=zhd(t,i,w,6,7)
            w,m,n=zhd(t,i,w,30,31)
    elif len(l)==5:#五个信号
        if sum(e)==5:
            w,m,n=zhd(t,i,w,0,1)
        elif sum(e)==4:
            w,m,n=sum3(e,t,i,w)
        elif sum(e)==3:
                w,m,n=sum5(e,t,i,w)
        elif sum(e)==2:
            w,m,n=sum6(e,t,i,w)
        elif sum(e)==1:
            if e.index(1)==0:
                w,m,n=zhd(t,i,w,52,53)
            elif e.index(1)==1:
                w,m,n=zhd(t,i,w,54,55)
            elif e.index(1)==2:
                w,m,n=zhd(t,i,w,56,57)
            elif e.index(1)==3:
                w,m,n=zhd(t,i,w,58,59)
            elif e.index(1)==4:
                w,m,n=zhd(t,i,w,60,61)
        elif sum(e)==0:
            w,m,n=zhd(t,i,w,62,63)
    return w,m,n        
"===============================================组合策略================================"
def zd(x,t,row,col,l):
    w=[]
    for i in range(row):
        e=[]
        for h in l:
            if x[:,h][i]>0:
                e.append(1)
            else:
                e.append(0)
        #判断几个信号
        w,m,n=numsignal(l,w,t,i,e)
    z,d=[],[]
    shang1=0
    for i in range(0,np.power(2,len(l))*2,2):
        if w.count(i)==0 and w.count(i+1)==0:
            print("组合 %s(%s,%s) 没有出现" %(int(i/2+1),i,(i+1)))
            z.append(0)
            d.append(0)
        else:
            c=w.count(i)+w.count(i+1)
            c1=w.count(i)/c
            c2=w.count(i+1)/c
            print("组合 %s(%s,%s) 的涨、跌概率：%s, %s" %(int(i/2+1),i,(i+1),'{:.2%}'.format(c1),'{:.2%}'.format(c2)))
            z.append(c1)
            d.append(c2)
            if c1>0 and c2>0 :
                shang1=(c1*(math.log(c1))*w.count(i))/120+(c2*(math.log(c2))*w.count(i+1))/120+shang1
    return w,z,d,shang1

"===============================================检验策略================================"
def test(shang,shang1,zd2,p3,o,k,p,z,d,l,x,t,w,row,h):
    p1,p2,zd1=[],[],[]
    for i in range(row):
            e=[]
            for b in l:
                if x[:,b][i]>0:
                    e.append(1)
                else:
                    e.append(0)
            w,m,n=numsignal(l,w,t,i,e)
            if z[int(m/2)]==0 and d[int(m/2)]==0:
                print("Not exist")
            elif z[int(m/2)]>=d[int(m/2)]:
                zd1.append("涨")
                if t[i]>=0:
                    p1.append(1)
            elif z[int(m/2)]<d[int(m/2)]:
                zd1.append("跌")
                if t[i]<0:
                    p2.append(-1)
    p3.append(('{:.2%}'.format((len(p1)+len(p2))/12)))
    print(("正确率：%s" %('{:.2%}'.format((len(p1)+len(p2))/12))))
    if k==1:
        shang=-shang1
    if -shang1<=shang:
        zd2=[]
        shang=-shang1
        h=h+1
        o=k
        zd2.append(zd1)
        p4.append(('{:.2%}'.format((len(p1)+len(p2))/12)))
    return p,o,p3,zd2,p4,h,shang
    
if __name__=="__main__":
    #读取数据
    a=1+12+12+12+12+12+12
    name = xlrd.open_workbook("after_data.xls")
    row, col, x,t ,x1,t1=read_excel(name,a)
    row=120
    col=col-3
    l1=[0,1,2]
    k,p,o,h,shang=0,0,0,0,0
    p3,zd2,p4=[],[],[]
    for s in l1:
        l=[]
        l.append(s)
        if s==1:
            l2=[2,3]
            l.append(0)
        elif s==2:
            l2=[3]
            l.append(0)
        else:
            l2=[1,2,3]
        for s1 in l2:
            if s1==2:
                l.pop(-1)
                l.append(s1)
                l3=[3,4]
            elif s1==3:
                l.pop(-1)
                l.append(s1)
                l3=[4]
            else:
                l.append(s1)
                l3=[2,3,4]
            for s2 in l3:
                l.append(s2)
                k+=1
                print("\n==================组合策略：5选3，第 %s 次 %s==================" %(k,l))
                w,z,d,shang1=zd(x,t,row,col,l)
                print("信息熵：%s" %-shang1)
                p,o,p3,zd2,p4,h,shang=test(shang,shang1,zd2,p3,o,k,p,z,d,l,x1,t1,w,12,h)
                l.pop(-1)
    print("\n==============================策略检验结果==============================")
    print("10个组合每次的正确率：%s" %p3)
    print("按照信息熵最小选择：第 %s 次的组合结果最好(正确率：%s)" %(o,p4[h-1]))
    print(zd2)
