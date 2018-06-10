import os,time

class defFiles():
    def delFile():
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