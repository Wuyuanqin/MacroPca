import numpy as np
import xlrd
import xlwt
import matplotlib.pyplot as plt

def read_excel(name):
    table = name.sheets()[0]
    row = table.nrows
    col = table.ncols
    #print(row,col)
    data_list,couponsid,tradtime_lis = [],[],[]
    for i in range(1,row):
        for j in range(1,col):
            data_list.append(table.cell(i,j).value)
    array = np.array(data_list).reshape(row-1,col-1)#将所有数据存进矩阵

    #相对强度
    t=np.zeros((row-1,col-1))
    for j in range(col-1):
        a=36
        for i in range(0,row-1,a):
            if i==216:
                a=30
            for l in range(a):
                input=array[i:i+a,j]
                t[i+l,j]=getPer(input,l)
##    newdata = xlwt.Workbook()
##    newtable1 = newdata.add_sheet('t')#创建新sheet
##    for i in range(row-1):
##        for j in range(col-1):
##            newtable1.write(i,j,t[i,j])
##    newdata.save("相对强度后数据.xls")#保存文件
    return t,row

def getPer(input,index):
    if (isinstance(input,np.ndarray)):
        input = input.tolist()
    if (index==-1):
        targetVal = input[-1]
    else:
        targetVal = input[index]
    input.sort()
    newindex = input.index(targetVal)
    per = (newindex+1)/len(input)
    return per

def pca(X,k):#k is the components you want
    n_samples, n_features = X.shape#mean of each feature
    #print(n_samples, n_features)
    mean=np.array([np.mean(X[:,i]) for i in range(n_features)])  #normalization
    norm_X=X-mean
    covX = np.cov(X.T)
    eig_val, eig_vec = np.linalg.eig(covX)
    print("eigenvector:")
    print(eig_vec)
    print("eigenvalue: \n%s\n" %eig_val)
    m,n=eig_vec.shape
    eig_pairs = [(np.abs(eig_val[i]), eig_vec[:,i]) for i in range(n_features)]  # sort eig_vec based on eig_val from highest to lowest
    eig_pairs.sort(reverse=True)  # select the top k eig_vec
    feature=np.array([ele[1] for ele in eig_pairs[:k]])  #get new data
    data=np.dot(norm_X,np.transpose(feature))#feature是行为向量
    data2=np.dot(data,feature)#返回PCA

    #MSE
    sse=0
    mse_in=np.array([])
    for d in range(n_samples):
        for j in range(n_features):
            sse=np.square((norm_X[d,j]- data2[d,j]))
            mse_in=np.append(mse_in,sse)
    mse_in=mse_in.reshape(n_samples, n_features)
    meancols=mse_in.mean(axis=0)
    print("MSE of 26 Indicators : \n%s" %meancols)
    
    #画图
    t_train=norm_X
    w=data2
    fig=plt.figure(figsize=(8, 6))
    q=1#q改变指标的画图
    n=np.linspace(1,n_samples, n_samples)
    ax1 = fig.add_subplot(1, 1, 1)
    plt.title("N = %s, Indicator: %s" %(k,q))
    plt.ylim(-0.75, 0.75)
    plt.plot(n, t_train[:,q-1], c="orange", linewidth=3)
    lines2=ax1.plot(n, w[:,q-1], color='r',linewidth=2)
    plt.show()
    
    #保存data
    newdata = xlwt.Workbook()
    newtable = newdata.add_sheet('PCA')#创建新sheet
    for i in range(n_samples):
        for j in range(k):
            newtable.write(i,j,data[i,j])

    newtable1 = newdata.add_sheet('eig_vec')#创建新sheet
    for i in range(n_features):
        for j in range(n_features):
            newtable1.write(i,j,eig_vec[i,j])
    newdata.save("after_PCA.xls")#保存文件
    return data

def ma(array,col,row):#moving average
    p=np.zeros((row,col))
    for j in range(col):
        for i in range(row):
            input=array[i:i+12,j]
            p[i,j]=np.sum(input)/12
    #save
    newdata = xlwt.Workbook()#创建新excel文件，注意workbook的W为大写
    newtable = newdata.add_sheet('ma')#创建新sheet
    for i in range(row):
        for j in range(col):
            newtable.write(i,j,p[i,j])
    newdata.save("moving_average.xls")#保存文件


if __name__=="__main__":
    #读取数据
    name = xlrd.open_workbook("before_PCA.xls")
    t,row=read_excel(name)
    k=5
    data=pca(t,k)#PCA
    ma(data,k,row-1)#moving average
