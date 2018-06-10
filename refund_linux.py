#headers
import pandas as pd
import numpy as np
from pandas import ExcelWriter
import time
import os
import os.path
import zipfile

#set start time 
start = time.clock()

#set path
path = "/data/fileuplode"
os.chdir(path)

#input Big Excel and read the file under this directory
reFund_path = "/data/fileuplode/filedata"
 
allFile = os.walk(reFund_path)
for file in allFile:
    print(file)

filename = file[2][0]
reFund_path_name = reFund_path+"/"+filename

#get the name of excel
name_D = filename
name_D = str(name_D)

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

def zip_dir(dirname,zipfilename):
    filelist = []
    if os.path.isfile(dirname):
        filelist.append(dirname)
    else :
        for root, dirs, files in os.walk(dirname):
            for name in files:
                filelist.append(os.path.join(root, name))
         
    zf = zipfile.ZipFile(zipfilename, "w", zipfile.zlib.DEFLATED)
    for tar in filelist:
        arcname = tar[len(dirname):]

        zf.write(tar,arcname)
    zf.close()

#make file dir
mkpath = path +"/{}".format(name_D)
mkpath = str(mkpath)
mkdir(mkpath)
#now at mkpath /data/fileuplode/filedata
mkpath = os.chdir(mkpath)
mkpath = str(mkpath)

print('now under /data/fileuplode/{}------>'.format(name_D)+os.getcwd())

df = pd.read_excel(reFund_path_name)

df = df.reset_index()

df_list = df["index"].tolist()

df_list.pop()

df_list

number_row = input("input the rows that you want cut:")

number_row = int(number_row)

df_list = df_list[::number_row]

list(enumerate(df_list, start=1))

for i, ele in enumerate(df_list):

    lower  = df_list[i]
    higher_len = (number_row-1)
    higher = df_list[i] + higher_len
    
    df = pd.read_excel(reFund_path_name)
    df = df.reset_index()
    
    df = df[(df["index"] >= lower)&(df["index"] <= higher)].head()
    
    df = df.drop(['index'], axis=1)
    
    df = df.to_excel('refund_{}-from page.{} to page.{}.xls'.format(str(i+1),lower+1,str(higher+1)),index=False,header=True) 

#set now position
nowPos = os.getcwd()
nowPos = str(nowPos)

zip_dir(nowPos,nowPos+"/{}.zip".format(name_D))

end = time.clock() 
print ("The python scripts totally have run at------>",end-start,"seconds ")
#call the API function
#then clear the toatl file
time.sleep(5) #wait 5s
if os.path.isdir(mkpath):
    startClearTime = time.clock()
    print ("--path--:"+" "+ mkpath + " " + "is exist!")
    #let all file in the list
    filelist = os.listdir(mkpath)
    print('all file in this dir is  ',filelist)
    for file in filelist:
        try:
            os.remove(mkpath+"/"file)
        except Exception as e:
            e.printstack()
    print("delete the file in this directory clearly!!")
    endClearTime = time.clock()
    print("clear all the file in this directory totally  cost ",endClearTime - startClearTime, "seconds ")
else:
    print("{}".format(mkpath)+"is not exit!!")
    