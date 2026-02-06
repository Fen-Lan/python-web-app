from time import sleep

from suds.client import Client
import copy
import imghdr
import os,time
import base64
import json
import requests
import hashlib
import math
from CLPC.framework import FRAME


'''
    支持的markdown语法[强]微表情不支持，换行后不要空格
    目前应用消息中支持的markdown语法是如下的子集：

    标题 （支持1至6级标题，注意#与文字中间要有空格）
    # 标题一
    ## 标题二
    ### 标题三
    #### 标题四
    ##### 标题五
    ###### 标题六
    加粗
    **bold**
    链接
    [这是一个链接](http://work.weixin.qq.com/api/doc)
    行内代码段（暂不支持跨行）
    `code`
    引用
    > 引用文字
    字体颜色(只支持3种内置颜色)
    <font color="info">绿色</font>
    <font color="comment">灰色</font>
    <font color="warning">橙红色</font>"""
        res=wechat.sendMarkdownMsg(userStrs,msg1)
''' 
    

class MsgSendByMobile():
    """
    通过 手机号/身份证号 发送企业微信消息
    """

    def __init__(self):
        self.url = "http://9.0.17.53:8081/wechatBase-service/services/ExternalMessageService?wsdl" # 信创接口
        self.sys_id = "T0190000"
        self.sys_id_pass = "9999"
        self.sys_flag = "T0190000"
        self.baseInput = {"sys_id":self.sys_id,"sys_id_pass":self.sys_id_pass,"sys_flag":self.sys_flag}
        self.client = Client(self.url)


    def sendMarkdownMsg(self, userNum="", msg=""):
        """
        Params:\n
        userNum: 用户 手机号/身份证号,多个 手机号/身份证号 用‘ , ’分开\n
        msg: 要发送的消息
        """
        res1 = ""
        res2 = ""
        
        if msg:
            print("开始发送消息：{}".format(msg))
            InputTextMsg = copy.deepcopy(self.baseInput)
            if len(userNum.split(",")[0]) == 11:
                InputTextMsg["userType"] = "1"
            elif len(userNum.split(",")[0]) == 18:
                InputTextMsg["userType"] = "2"
            
            InputTextMsg["content"] = msg
            InputTextMsg["userNum"] = userNum
            res1 = self.client.service.sendMarkdownMsg(InputTextMsg)
            if res1.errCode == "000000" and res1.errMsg is None:
                print("消息发送成功！！！")
            else:
                print("消息发送失败，失败原因：{}".format(res1.errMsg))

    def sendMsg(self, userNum="", msg="", file_path=[]):
        """
        Params:\n
        userNum: 用户 手机号/身份证号,多个 手机号/身份证号 用‘ , ’分开\n
        msg: 要发送的消息\n
        file_path: 发送的文件/图片路径 列表\n\n
        """
        res1 = ""
        res2 = ""
        
        if msg:
            print("开始发送消息：{}".format(msg))
            InputTextMsg = copy.deepcopy(self.baseInput)
            if len(userNum.split(",")[0]) == 11:
                InputTextMsg["userType"] = "1"
            elif len(userNum.split(",")[0]) == 18:
                InputTextMsg["userType"] = "2"
            
            InputTextMsg["content"] = msg
            InputTextMsg["userNum"] = userNum
            res1 = self.client.service.sendTextMsg(InputTextMsg)
            if res1.errCode == "000000" and res1.errMsg is None:
                print("消息发送成功！！！")
            else:
                print("消息发送失败，失败原因：{}".format(res1.errMsg))

        if file_path:
            for f in file_path:
                size = os.path.getsize(f)
                print("文件大小：{:.2f}M".format(size/1024/1024))
                if size>6000*999*1024:
                    raise Exception('文件太大，不支持发送')
                # if size > 8000000:        # 企业微信那边说支持8M的，但测试下来6M就发不了了
                if size > 6044000:      # 5.859兆=6000*1024
                    file_list = split_file_with_7z(f)
                    if []==file_list:
                        raise Exception('文件大于6000KB,在拆分文件时失败，请检查文件：例如不要把文件放置在磁盘根目录')
                    self.send(userNum, file_list)
                else:
                    self.send(userNum, [f])

    # 发送图片或文件
    def send(self, userNum, file_path):
        for f in file_path:
            imgType_list = {'jpg','bmp','png','jpeg','rgb','tif'}
            InputFileMsg = copy.deepcopy(self.baseInput)
            if len(userNum.split(",")[0]) == 11:
                InputFileMsg["userType"] = "1"
            elif len(userNum.split(",")[0]) == 18:
                InputFileMsg["userType"] = "2"
            InputFileMsg["baseStr"] = self.file_to_base64(f)
            InputFileMsg["fileName"] = os.path.basename(f)
            InputFileMsg["userNum"] = userNum
            size = os.path.getsize(f)
            InputFileMsg["fileLength"] = size

            if imghdr.what(f) in imgType_list:
                print("开始发送图片：{}".format(f))
                InputFileMsg["mediaType"] = "image"
                res2 = self.client.service.sendImgMsg(InputFileMsg)
                if res2.errCode == "000000" and res2.errMsg is None:
                    print("图片发送成功！！！")
                else:
                    print("图片发送失败，失败原因：{}".format(res2.errMsg))
            else:
                print("开始发送文件：{}".format(f))
                InputFileMsg["mediaType"] = "file"
                res2 = self.client.service.sendFileMsg(InputFileMsg)
                if res2.errCode == "000000" and res2.errMsg is None:
                    print("文件发送成功！！！")
                else:
                    print("文件发送失败，失败原因：{}".format(res2.errMsg))
            time.sleep(1)

    def file_to_base64(self,file_path):
        if not file_path:
            return "请传入图片地址"
        if not os.path.exists(file_path):
            return "文件不存在，请核实文件地址"           
        
        with open(file_path,'rb') as f:
            base64_data = base64.b64encode(f.read())
        # print(base64_data)
        return str(base64_data,"utf8")


class MsgSendByRobot():
    """
    通过群机器人发送企业微信消息\n
    初始化时需提供机器人webhook_id
    """

    def __init__(self, webhook_id):
        self.url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key={}".format(webhook_id)
        self.id_url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={}&type=file".format(webhook_id)

    # 调用企业微信群机器人向群内发送消息
    def sendMarkdownMsg(self, msg=""):
        """
        Params:
        msg: 要发送的markdown消息
        """
        res1 = ""
        if msg:
            print("开始发送消息！！！")
            data = json.dumps({"msgtype": "markdown","markdown":{"content": msg}})

            exec_env = FRAME.get_env_arg("RPA_ENV", "")
            if exec_env == "prod" or exec_env == 'test':  # 服务器环境，使用代理发送群信息
                proxies = {'http': 'http://10.20.102.61:8080', 'https': 'http://10.20.102.61:8080'}
                result = requests.post(self.url, data, auth=('Content-Type', 'application/json'), proxies=proxies)
            else:  # 本地环境，直接发送群信息
                result = requests.post(self.url, data, auth=('Content-Type', 'application/json'))

            res1 = result.json()
            if res1.get("errcode") == 0 and res1.get("errmsg") == "ok":
                print("消息发送成功！！！")
            else:
                print("消息发送失败！失败原因：{}".format(res1.get("errmsg")))

    # 调用企业微信群机器人向群内发送消息
    def sendMsg(self, msg="", file_path=[], mentioned_mobile_list=[]):
        """
        Params:
        msg: 要发送的文本消息
        file_path: 要发送的文件/图片 列表
        """
        res1 = ""
        res2 = ""
        if isinstance(mentioned_mobile_list,list) is not True:
            print("@信息格式不正确")
            mentioned_mobile_list = []
        if msg:
            print("开始发送消息！！！")
            data = json.dumps(
                {   "msgtype": "text",
                    "text":
                     {  "content": msg,
                        "mentioned_mobile_list": mentioned_mobile_list
                      }
                 })

            exec_env = FRAME.get_env_arg("RPA_ENV", "")
            if exec_env == "prod" or exec_env == 'test':  # 服务器环境，使用代理发送群信息
                proxies = {'http': 'http://10.20.102.61:8080', 'https': 'http://10.20.102.61:8080'}
                result = requests.post(self.url, data, auth=('Content-Type', 'application/json'), proxies=proxies)
            else:  # 本地环境，直接发送群信息
                result = requests.post(self.url, data, auth=('Content-Type', 'application/json'))

            res1 = result.json()
            if res1.get("errcode") == 0 and res1.get("errmsg") == "ok":
                print("消息发送成功！！！")
            else:
                print("消息发送失败！失败原因：{}".format(res1.get("errmsg")))

        if file_path:

            imgType_list = {'jpg','bmp','png','jpeg','rgb','tif'}

            for f_path in file_path:

                if imghdr.what(f_path) in imgType_list and os.path.getsize(f_path) < 2097152:       # 2兆以内
                    print("开始发送图片：{}".format(f_path))

                    with open(f_path, 'rb') as f:  # 转换图片成base64格式
                        data = f.read()
                        encodestr = base64.b64encode(data)
                        image_data = str(encodestr, 'utf-8')

                    with open(f_path, 'rb') as f:  # 图片的MD5值
                        md = hashlib.md5()
                        md.update(f.read())
                        image_md5 = md.hexdigest()

                    headers = {"Content-Type": "application/json"}
                    data = {
                        "msgtype": "image",
                        "image": {
                            "base64": image_data,
                            "md5": image_md5
                        }
                    }
                    # result = requests.post(self.url, headers=headers, json=data)

                    exec_env = FRAME.get_env_arg("RPA_ENV", "")
                    if exec_env == "prod" or exec_env == 'test':  # 服务器环境，使用代理发送群信息
                        proxies = {'http': 'http://10.20.102.61:8080', 'https': 'http://10.20.102.61:8080'}
                        result = requests.post(self.url, headers=headers, json=data, auth=('Content-Type', 'application/json'),
                                               proxies=proxies)
                    else:  # 本地环境，直接发送群信息
                        result = requests.post(self.url, headers=headers, json=data, auth=('Content-Type', 'application/json'))

                    res2 = result.json()
                    print(res2)
                    if res2.get("errcode") == 0 and res2.get("errmsg") == "ok":
                        print("==========图片发送成功！！！==========")
                    else:
                        print("图片发送失败！失败原因：{}".format(res2.get("errmsg")))
                    
                else:
                    print("开始发送文件：{}".format(f_path))
                    
                    size = os.path.getsize(f_path)
                    print("文件大小：{:.2f}M".format(size/1024/1024))
                    if size>6000*999*1024:
                        raise Exception('文件太大，不支持发送')
                    if size > 6044000:      # 5.859兆=6000*1024
                        file_list = split_file_with_7z(f_path)
                        if []==file_list:
                            raise Exception('文件大于6000KB,在拆分文件时失败，请检查文件：例如不要把文件放置在磁盘根目录')
                        for f in file_list:
                            self.send(f)
                    else:
                        self.send(f_path)

    def send(self,f_path):
        fname = f_path.split('\\')[-1]
        headers = {
            'Content-Type': 'multipart/form-data',
            }
        
        data = {
            "filename": (open(f_path, 'rb'))
            }
        print(data)
        # 先获取上传后文件的media_id，再发送消息

        exec_env = FRAME.get_env_arg("RPA_ENV", "")
        if exec_env == "prod" or exec_env == 'test':  # 服务器环境，使用代理发送群信息
            proxies = {'http': 'http://10.20.102.61:8080', 'https': 'http://10.20.102.61:8080'}
            response = requests.post(url=self.id_url, files=data, headers = headers, auth=('Content-Type', 'application/json'), proxies=proxies)
        else:
            response = requests.post(url=self.id_url, files=data, headers = headers)

        json_res = response.json()
        # print(json_res)
        media_id = json_res.get('media_id')
        data = {
            "msgtype": "file",
            "file": {
                "media_id": media_id
            }
        }
        if exec_env == "prod" or exec_env == 'test':  # 服务器环境，使用代理发送群信息
            proxies = {'http': 'http://10.20.102.61:8080', 'https': 'http://10.20.102.61:8080'}
            result = requests.post(self.url, json=data, auth=('Content-Type', 'application/json'), proxies=proxies)
        else:
            result = requests.post(self.url, json=data)
        res2 = result.json()
        if res2.get("errcode") == 0 and res2.get("errmsg") == "ok":
            print("==========文件发送成功！！！==========")
        else:
            print("文件发送失败！失败原因：{}".format(res2.get("errmsg")))


class MsgSendByRobotIntranet():
    """
    内网，通过群机器人发送企业微信消息\n
    初始化时需提供机器人id和password，事先通过OA运维单向企业微信团队申请
    """

    def __init__(self, robot_id,robot_pwd):
        self.url = "http://9.0.17.52:8081/wechatBase-web/rebot/sendMessage?id={}&password={}".format(robot_id,robot_pwd) # 信创接口



    def sendMarkdownMsg(self, msg=""):
        """
        Params:\n
        msg: 要发送的消息
        """
        res1 = ""
        
        if msg:
            print("开始发送消息！！！")
            data = json.dumps({"msgtype": "markdown", "markdown": {"content": msg}})
            result = requests.post(self.url, data, auth=('Content-Type', 'application/json'))
            res1 = result.json()
            if res1.get("errcode") == '0' and res1.get("errmsg") == "ok":
                print("消息发送成功！！！")
            else:
                print("消息发送失败！失败原因：{}".format(res1.get("errmsg")))
                raise Exception("消息发送失败！失败原因：{}".format(res1.get("errmsg")))

    # 调用企业微信群机器人向群内发送消息
    def sendMsg(self, msg="", file_path=[]):
        """
        Params:
        msg: 要发送的文本消息
        file_path: 要发送的图片(仅'jpg','jpeg','png') 列表，单个文件小于2M
        """
        res1 = ""
        res2 = ""

        if msg:
            print("开始发送消息！！！")
            data = json.dumps({"msgtype": "text", "text": {"content": msg}})
            result = requests.post(self.url, data, auth=('Content-Type', 'application/json'))
            res1 = result.json()
            if res1.get("errcode") == '0' and res1.get("errmsg") == "ok":
                print("消息发送成功！！！")
            else:
                print("消息发送失败！失败原因：{}".format(res1.get("errmsg")))
                raise Exception("消息发送失败！失败原因：{}".format(res1.get("errmsg")))

        if file_path:

            imgType_list = {'jpg','jpeg','png'}

            for f_path in file_path:
                print(imghdr.what(f_path),os.path.getsize(f_path))
                if imghdr.what(f_path) in imgType_list and os.path.getsize(f_path) < 2097152:       # 2兆以内
                    print("开始发送图片：{}".format(f_path))

                    with open(f_path, 'rb') as f:  # 转换图片成base64格式
                        data = f.read()
                        encodestr = base64.b64encode(data)
                        image_data = str(encodestr, 'utf-8')

                    with open(f_path, 'rb') as f:  # 图片的MD5值
                        md = hashlib.md5()
                        md.update(f.read())
                        image_md5 = md.hexdigest()

                    headers = {"Content-Type": "application/json"}
                    data = {
                        "msgtype": "image",
                        "image": {
                            "base64": image_data,
                            "md5": image_md5
                        }
                    }
                    result = requests.post(self.url, headers=headers, json=data)
                    res2 = result.json()
                    print(res2)
                    if res2.get("errcode") == '0' and res2.get("errmsg") == "ok":
                        print("==========图片发送成功！！！==========")
                    else:
                        print("图片{}发送失败！失败原因：{}".format(f_path,res2.get("errmsg")))
                        raise Exception("图片发送失败！失败原因：{}".format(res2.get("errmsg")))
                    
                else:
                    print("发送失败。不是图片，或者不是jpg、jpeg、png的图片类型，或者图片超过2M：{}".format(f_path))
                    raise Exception("发送失败。不是图片，或者不是jpg、jpeg、png的图片类型，或者图片超过2M")
                    
    
# 将大文件拆分为多个多个文件，并返回文件路径列表
def create_sfx(file_path):
    size = os.path.getsize(file_path)/1024
    n = math.ceil(size/6000)
    for i in range(n):
        f = file_path + ".v." + "{0:0>3}".format(1+i)
        if os.path.exists(f):
            os.remove(f)
    if "7-Zip" in os.listdir("C:\\Progra~1"):
        p = "C:\\Progra~1\\7-Zip\\7z"
    elif "7-Zip" in os.listdir("C:\\Progra~2"):
        p = "C:\\Progra~2\\7-Zip\\7z"
    else:
        raise Exception("请安装7-zip后尝试...")
    cmd = p + " a -tzip {} {} -v6000k".format(file_path+".v", file_path)
    # print(cmd)
    os.system(cmd)
    file_list = []
    for i in range(n):
        f = file_path + ".v." + "{0:0>3}".format(1+i)
        if os.path.exists(f):
            file_list.append(f)
    return file_list


def split_file_with_7z(file_name, part_size_mb=5):
    name, type = os.path.splitext(os.path.basename(file_name))

    sz_file = name + '.7z'
    vol_size = 1024 * 1024 * part_size_mb

    import multivolumefile
    import py7zr

    with multivolumefile.open(sz_file, mode='wb', volume=vol_size) as target_archive:
        with py7zr.SevenZipFile(target_archive, 'w') as archive:
            archive.writeall(file_name)
    print("压缩完成")
    import glob

    # 定义分卷文件的模式
    pattern = sz_file + '.*'  # 根据你的实际文件名模式进行调整
    files = glob.glob(pattern)
    files.sort()  # 确保按正确的顺序处理分卷文件
    return list(files)

