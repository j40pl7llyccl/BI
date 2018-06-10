import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import Series, DataFrame
from pandas.core.frame import DataFrame
import os
import datetime
import time

start = time.clock()

Adu_path = input("Please type your Adu_path :")
Adu_path = str(Adu_path)
Re_path = input ("Please type your Re_path  :")
Re_path = str(Re_path)
name_D = input("type your file name_D:")
name_D = str(name_D)
name_L = input("type your file name_L:")
name_L = str(name_L)

#get now position
pwd = os.getcwd()
pwd = str(pwd)
# In[46]:


###2. sorte
df_Adu = pd.read_excel(Adu_path)
df_Adu_DandE = df_Adu.loc[((df_Adu["reason"] == "D") | (df_Adu["reason"] == "E") )&(df_Adu["disposition"] == "SELLABLE")&(pd.notnull(df_Adu["sku"])) ]
df_Adu_DandE.index = range(len(df_Adu_DandE))
df_Adu_DandE = df_Adu_DandE.reset_index() 
df_Adu_DandE


# In[47]:


df_Re = pd.read_excel(Re_path)
df_Re_Damaged_Warehouse = df_Re.loc[((df_Re["reason"] == "Damaged_Warehouse")&(pd.notnull(df_Re["sku"])))]
df_Re_Damaged_Warehouse.index = range(len(df_Re_Damaged_Warehouse))
df_Re_Damaged_Warehouse = df_Re_Damaged_Warehouse.reset_index()
df_Re_Damaged_Warehouse


# In[48]:


df_Adu_list_DandE = list(df_Adu_DandE["sku"])
df_Adu_list_DandE = sorted(df_Adu_list_DandE, reverse = True)

len(df_Adu_list_DandE)


dict_Adu = {}
for a in range(0, len(df_Adu_list_DandE)):
    dict_Adu[df_Adu_list_DandE[a]] = '{}'.format(a)
len(dict_Adu)


set_Adu_list_DandE = set(df_Adu_list_DandE)
len(set_Adu_list_DandE)
list(set_Adu_list_DandE)

list_s1 = []
for k in range(0, len(dict_Adu)):
    list_s1.append(list(set_Adu_list_DandE)[k])
    if k == len(dict_Adu):
        print("ok")
    else:
        continue
list_s1.sort()
list_s1
list_s1 = list_s1[::-1]
list_s1


df_Re_list_Damanged_Warehouse = list(df_Re_Damaged_Warehouse["sku"])
df_Re_list_Damanged_Warehouse = sorted(df_Re_list_Damanged_Warehouse, reverse = True)

len(df_Re_list_Damanged_Warehouse)

dict_Re = {}
for a in range(0, len(df_Re_list_Damanged_Warehouse)):
    dict_Re[df_Re_list_Damanged_Warehouse[a]] = '{}'.format(a)
len(dict_Re)

set_Re_list_Damanged_Warehouse = set(df_Re_list_Damanged_Warehouse)
len(set_Re_list_Damanged_Warehouse)
list(set_Re_list_Damanged_Warehouse)

list_s2 = []
for k in range(0, len(dict_Re)):
    list_s2.append(list(set_Re_list_Damanged_Warehouse)[k])
    if k == len(dict_Re):
        print("ok")
    else:
        continue
list_s2.sort()
list_s2
list_s2 = list_s2[::-1]
list_s2


# In[50]:


list_total = [list_s1+list_s2][0]
list_total = set(list_total)
list_total = list(list_total)
list_total.sort()
list_total = list_total[::-1]
list_total


# In[51]:


lower  = 0
higher = 0



list_t = []
for i, l in enumerate(list_total):
    #print(i,l)

    Adu_len = sum(df_Adu_DandE[df_Adu_DandE["sku"]==list_total[i]]['quantity'])
    Re_len = sum(df_Re_Damaged_Warehouse[df_Re_Damaged_Warehouse["sku"]==list_total[i]]['quantity-reimbursed-total'])

    list_ts = [list_total[i], Adu_len, Re_len,Adu_len+Re_len]
    list_t.append(list_ts)


# In[52]:


list_t


# In[53]:


list_t.sort()
list_t
data=DataFrame(list_t)
data = data.sort_values(by=3 , ascending = False)
data


# ### check total list
def mkdir(path):
    
    import os
 
    
    path=path.strip()
    
    path=path.rstrip("\\")
 

    isExists=os.path.exists(path)
 

    if not isExists:
        
        print (path+' file create ok')
        
        os.makedirs(path)
        return True
    else:
        
        print (path+' file already exist')
        return False
 

mkpath = pwd +"\\{}_List".format(name_D)
mkdir(mkpath)

os.chdir(mkpath)
os.getcwd()
# In[54]:

writer = pd.ExcelWriter('totalList_Damaged_{}.xlsx'.format(name_D), engine='xlsxwriter')

data = data.sort_values(by=3 , ascending = False)
data.to_excel(writer, sheet_name='Sheet3')

writer.save()
os.chdir(pwd)
os.getcwd()

# In[55]:


#global lower,higher
lower  = 0
higher = 0

#print(data)
data=data.reset_index(drop=True)
for k, v in enumerate(data[0]):

    mkpath = pwd + "\\{}".format(name_D)
    mkdir(mkpath)
    os.chdir(mkpath)
    os.getcwd()
    writer = pd.ExcelWriter('{}_Damaged_{}_{}.xlsx'.format(name_D,k+1,data[0][k] ), engine='xlsxwriter')
        #if total != 0:
    SKU = ["SKU","Amazon damaged inventory","Reimbursed damanged inventory","Damanged inventory not yet reimbursed"]
    vSKU = [data[0][k],data[1][k],data[2][k],data[3][k]]
    dict_test = {
        "SKU":SKU,
        "SKU_{}".format(k+1):vSKU
    }
    dict_test = pd.DataFrame(dict_test)
    dict_test.drop(dict_test.index[0],inplace=True)
    dict_test.to_excel(writer,sheet_name='Sheet3')

    adu_sku_filter = df_Adu.loc[((df_Adu["reason"] == "D") | (df_Adu["reason"] == "E") )&(df_Adu["disposition"] == "SELLABLE")&(df_Adu["sku"] == data[0][k])]
    Adu_len = len(adu_sku_filter)
    #print(adu_sku_filter)
    adu_sku_filter.to_excel(writer, sheet_name='Sheet3', startrow=6)
    
    re_sku_filter = df_Re.loc[((df_Re["reason"] == "Damaged_Warehouse")&(df_Re["sku"] == data[0][k]))]
    #print(re_sku_filter)
    re_sku_filter.to_excel(writer, sheet_name='Sheet3', startrow=9+Adu_len)
    
#os.system('ls')

writer.save()
os.chdir(pwd)
os.getcwd()
end = time.clock() 
print ("程式實際執行",end-start,"seconds ")