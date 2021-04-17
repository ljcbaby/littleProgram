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
    # try:
        open_json = open(file_url,'rb')
        load_json = json.loads(open_json.read())
        clear(load_json)
    # except:
    #     print("error\n")
    # finally:
        open_json.close()

#处理读入的数据，提取成一个列表(list)
def clear(jsons):
    count = 0
    for index in jsons['result']['allSubjectList']:
        count+=1
    L=[]

    #jsons['result']['studentAnswerRecords'] 为一个字典(dict)
    for index in jsons['result']['studentScoreDetailDTO']:
        l=[]
        l.append(index['userName'])
        for subject in index['scoreInfos']:
            if (subject['score']=='未扫，不计排名'):
                l.append('-1')
            else:
                l.append(subject['score'])
        L.append(l)
    sto(L)

#存储成文件
def sto(L):
    stdOut = open(file_output,'w',encoding = "utf-8")
    for i in L:
        for o in i:
            stdOut.write("%s\t"%(o))
        stdOut.write("\n")
    stdOut.close()

if __name__ == '__main__':
    index()
