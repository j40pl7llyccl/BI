# -*- coding:utf-8 -*-
import pandas as pd
import numpy as np
from pandas import ExcelWriter
import time
import os
import os.path
import zipfile
import sys
import pycurl
import time

#set start time 
start = time.clock()

#set path
path = "/data/fileupload"
nowPos = os.chdir(path)
print("now is uder  ",nowPos) #在/data/fileupload下

#input Big Excel and read the file under this directory
reFund_path = "/data/fileupload/filedata"
reFund_path = str(reFund_path) 
allFile = os.walk(reFund_path)
for file in allFile:
    print(file)

filename = file[2][0]
reFund_path_name = reFund_path+"/"+filename

#get the name of excel
name_D = filename


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
#mkpath = path +"/{}".format(name_D)
#mkpath = path +"/filezip"

#mkpath = os.chdir(mkpath)

#print('now under /data/fileuplode/filezip------>')
#####################################################
os.chdir(reFund_path)
reFund_path = os.chdir(reFund_path)
print('now under /data/fileupload/filedata------>')
os.getcwd()

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
print("now is under /data/fileupload/filedata------>")
#zipPath = "/data/fileupload/filezip"
#os.chdir(zipPath)
#nowPos = os.getcwd()
#print("now is under /data/fileupload/filezip------>")
zip_dir(nowPos,nowPos+"/{}.zip".format(name_D))

end = time.clock() 
print ("The python scripts totally have run at------>",end-start,"seconds ")
#call the API function
'''let user download
刚才忘了把接口url 发给你了 39.108.7.40:8010/file/splitFile  
参数的话就传json格式{fileName:"",filePath:"",originalFilename:""}
fileName：拆分后的文件名
filePath：拆分后文件的绝对路径
originalFilename：拆分之前的文件名
接口的请求的Headers的content_type设置为这种的 application/x-www-form-urlencoded
'''

'''let user upload
39.108.7.40:8010/file/upload
'''
#then clear the toatl file
time.sleep(3) #wait 5s
reFund_path = "/data/fileupload/filedata"
os.chdir(reFund_path)

#調用api
start = time.time()
url = "http://39.108.7.40:8010/file/splitFile?fileName={}&filePath={}&originalFilename={}".format("{}.zip".format(name_D),"/data/fileupload/filedata"+"/{}.zip".format(name_D),filename)
c = pycurl.Curl()
c.setopt(c.URL, url)
c.perform()
end = time.time()
duration = end - start

print c.getinfo(pycurl.HTTP_CODE), c.getinfo(pycurl.EFFECTIVE_URL)
c.close()
print 'pycurl takes %s seconds to get %s ' % (duration, url)
'''
if os.path.isdir(reFund_path):
    startClearTime = time.clock()
    print ("--path--:"+" "+ reFund_path + " " + "is exist!")
    #let all file in the list
    filelist = os.listdir(reFund_path)
    print('all file in this dir is  ',filelist)
    for file in filelist:
        try:
            os.remove(reFund_path+"/"+file)
        except Exception as e:
            e.printstack()
    print("delete the file in this directory clearly!!")
    endClearTime = time.clock()
    print("clear all the file in this directory totally  cost ",endClearTime - startClearTime, "seconds ")
else:
    print("{}".format(reFund_path)+"is not exit!!")
'''
    