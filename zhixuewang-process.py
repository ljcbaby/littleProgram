#!/usr/bin/python3

#需安装json和os两个支持库
import json
import os

#需处理的json文件路径
file_url = "./1.json"

#输出文件路径
file_output = "./res.txt"

#读入文件
def index():
    try:
        open_json = open(file_url,'rb')
        load_json = json.loads(open_json.read())
        clear(load_json)
    except:
        print("error\n")
    finally:
        open_json.close()

#处理读入的数据，提取成一个列表(list)
def clear(jsons):
    L=[]

    #jsons['result']['studentAnswerRecords'] 为一个字典(dict)
    for index in jsons['result']['studentAnswerRecords']:
        l=[]
        l.append(int(index['studentNo']))
        l.append(index['score'])
        L.append(l)

    sto(L)

#存储成文件
def sto(L):
    stdOut = open(file_output,'w',encoding = "utf-8")
    L.sort(key=takeFirst)
    cache=[]
    for i in L:
        #本段if_elif为特化处理
        if(i[0]!=20210104 and i[0]!=20210150 and i[0] != 20210120):
            #通用处理仅需保留下一行stdOut
            if(i[1]!=0):
                stdOut.write("%d\n"%(i[1]))
            else:
                stdOut.write("\n")
            if(i[0]==20210124):
                if(cache[1]!=0):
                    stdOut.write("%d\n"%(cache[1]))
                else:
                    stdOut.write("\n")
        elif(i[0] == 20210104):
            cache = i
        elif(i[0] == 20210150):
            if(i[1]!=0):
                stdOut.write("%d"%(i[1]))
            else:
                stdOut.write("")
            
    stdOut.close()

#排序用函数
def takeFirst(L):
    return L[0]

if __name__ == '__main__':
    index()
