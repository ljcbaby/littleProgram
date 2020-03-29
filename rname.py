# coding=UTF-8
#!/usr/bin/python

import os
import sys

if (sys.argv[1] != "list") and (sys.argv[1] != "do"):
    print("ERROR Option.\nOnly \"list\" and \"do\" can be accepted.")
    exit()
file_list = os.listdir(os.getcwd())
for i in file_list:
    shu = i.rfind(".EP")
    if shu != -1 :
        new_name = i[shu+1:shu+5]
        shu = i.rfind('.')
        new_name = new_name + i[shu:]
        if sys.argv[1] == "list":
            print(i," --> ",new_name,"\n")
        else:
            os.rename(i,new_name)
