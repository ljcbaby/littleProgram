import requests
import json

file_output = "./res.txt"
stdOut = open(file_output,'w',encoding = "utf-8")

def index():
    for id in range(1,330):
        try:
            parse_json(id,json.loads(get_json("http://172.16.1.44/json/videoId.action?id=" + str(id))))
        except:
            stdOut.write(str(id) + " Parse Error\n")
    stdOut.close()

def parse_json(id,jsons):
    if(jsons["success"]):
        stdOut.write(str(id) + " " + jsons["data"]["title"] + "\n")
    else:
        stdOut.write(str(id) + " None\n")

def get_json(url):
    # 模拟浏览器发送http请求
    response = requests.get(url)
    # 编码方式
    response.encoding='utf-8'
    # 目标小说主页的网页源码
    return response.text

if __name__ == '__main__':
    index()
