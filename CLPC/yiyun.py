from CLPC.browser_visual import Browser
from time import sleep, time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from CLPC.element import Element
from selenium.webdriver.chrome.webdriver import WebDriver
import os,traceback

from CLPC.framework import FUNC_USAGE_TRACKER
from CLPC.union_login import UNILOGIN
from CLPC.tool import TOOL


@FUNC_USAGE_TRACKER
class YIYUN(Browser):
    def __init__(self, driver: WebDriver):
        """
        登录统一工作台，统一工作台操作组件\n
        初始化后需要传入用户信息
        @driver: 已初始化的webdriver
        """
        # super().__init__(driver)
        self.driver = driver
        current_dir = os.path.dirname(os.path.abspath(__file__))
        ele_json_path = os.path.join(current_dir, 'yiyun.json')
        self.ele_map = Element(ele_json_path)  # 加载组件自己的元素文件


        self.userName = ""
        self.passWord = ""
        self.verfiCode = ""
        self.BASEURL = "https://gsyy.chinalife-p.com.cn/api/v5/app/entry/"

    def login_yiyun(self, facility:str):
        userName = self.userName
        passWord = self.passWord
        tool_get_verify_code = TOOL.get_verify_code(userName)

        if tool_get_verify_code:
            verfiCode = tool_get_verify_code
        else:
            verfiCode = self.verfiCode

        # uni.login_gateway()
        # uni.enter_subsys_from_login_center("国寿易研(低代码门户)")



        self.create("http://9.0.16.161:8080")

        self.click("内网门户-账号登录切换")
        self.input_text("用户名", userName)
        self.input_text("密码", passWord)
        self.input_text("验证码", verfiCode)
        self.click("登录")

        sleep(3)
        self.catch("国寿易研", mode='equal')
        self.click_by_js("易云入口")
        self.click_by_js("机构选择")
        self.click("机构选项", arg=facility)
        self.click("机构确定")
        sleep(3)
        self.catch("gsyy.chinalife-p.com.cn", mode="contain", type='url', timeout=60)
        self.wait_loaded("工作台")

    def enter_app(self, app_name:str):
        """
        根据应用名称进入应用，需要确保应用名称的唯一性
        :param app_name: 应用名称
        """
        self.click("应用名称", arg=app_name)

    def enter_app_use_search(self, app_name:str):
        """
        通过应用搜索的方式进入应用，固定选择第一个搜索结果
        :param app_name: 应用名称
        """
        self.input_text("应用搜索输入", app_name)
        sleep(1)
        self.click("应用搜索结果")

    def search_data_list_use_api(self, api_key:str, app_id:str, entry_id:str, limit:int=10, filter:dict=None, fields:list=None):
        """
        通过API查询多条数据，返回查询结果
        """
        result = ""
        import requests, json
        Authorization = "Bearer " + api_key
        url = self.BASEURL + 'data/list'
        body = {'app_id': app_id, 'entry_id': entry_id, 'limit': limit}
        if filter != None:
            # 如果filter是dict类型，则直接使用，否则将其转换为dict
            if isinstance(filter, dict):
                body['filter'] = filter
            else:
                body['filter'] = json.loads(filter)
        if fields and len(fields) > 0:
            body['fields'] = fields

        try:
            headers = {
                'Authorization': Authorization,
                'Content-Type': 'application/json'
            }
            print(f"===> http post  data:{body} ")
            json_data=body
            if type(body)==dict:
                json_data=json.dumps(json_data).encode("utf-8") #dict数据转成str数据
            r = requests.post(url, headers=headers, data=json_data)
            print(f"response status:{r.status_code}, reason:{r.reason}")
            result = r.json()
            return result['data']
        except Exception as e:
            print(f'===> http post err {e}')
            return None

    def search_data_use_api(self,api_key:str, app_id:str, entry_id:str, data_id:str):
        """
        通过API查询单条数据，返回查询结果
        """
        result = ""
        import requests, json
        authorization = "Bearer " + api_key
        url = self.BASEURL + 'data/get'
        data = {'app_id': app_id, 'entry_id': entry_id, 'data_id': data_id}
        try:
            headers = {
                'Authorization': authorization,
                'Content-Type': 'application/json'
            }
            print(f"===> http post  data:{data} type:{type(data)}")
            json_data=data
            if type(data)==dict:
                json_data=json.dumps(json_data) #dict数据转成str数据
            r = requests.post(url, headers=headers, data=json_data)
            print(f"response status:{r.status_code}, reason:{r.reason}")
            result = r.json()
            print(f"<=== response json:\n{result}")
            return result['data']
        except Exception as e:
            print(f'===> http post err {e}')
            return None

    def get_file_use_url(self, url, file_name, download_timeout=60):
        """
        通过下载链接下载文件，保存到当前工作目录下
        :param url: 问价下载地址，易云返回
        :param file_name: 保存的文件名
        :param download_timeout: 最大下载等待时间
        """
        before = set(os.listdir(self.download_dir))
        sleep(2)
        self.driver.execute_script("window.open('');")
        windows = self.driver.window_handles
        self.driver.switch_to.window(windows[-1])
        self.create(url)
        sleep(2)
        for ti in range(download_timeout):
            after = set(os.listdir(self.download_dir))
            if after:
                for tf in after:
                    if '.crdownload' in tf:
                        sleep(1)
                        break
                else:
                    down_files = after - before
                    break

            sleep(1)

        try:
            down_name = down_files.pop()
            print('filename:', down_name)
        except:
            print("文件下载失败")
            raise Exception('文件下载失败')

        source_path = os.path.join(self.download_dir, down_name)
        for i in range(download_timeout):
            if TOOL.file_exist(source_path) and TOOL.get_FileSize_kb(source_path) > 3:
                break
            else:
                sleep(1)
        else:
            print("文件下载失败")
            raise Exception('文件下载失败')

        current_dir = os.getcwd()  ## todo 增加支持指定目录
        target_path = os.path.join(current_dir, file_name)
        TOOL.mov_file(source_path, target_path)

        self.driver.close()
        windows = self.driver.window_handles
        self.driver.switch_to.window(windows[-1])

    def update_data_use_api(self, data, api_key:str, app_id:str, entry_id:str, data_id:str, updator:str):
        """
        通过API更新单条数据
        注意data 格式， data_creator 必填
        代码开发组件
        :param data: 更新数据 json，需要按照易云api的格式
        :param api_key: 易云的api_key
        :param app_id: 易云的appid
        :param entry_id: 易云的entryid
        :param data_id: 易云的dataid
        :param updator: 更新人身份证号
        """
        # todo 更新只更新对应的字段，data要改

        result = ""
        import requests, json
        authorization = "Bearer " + api_key
        url = self.BASEURL + 'data/update'
        dict_data = {'app_id': app_id, 'entry_id': entry_id, 'data_id': data_id, 'data': data, "data_creator": updator}
        try:
            headers = {
                'Authorization': authorization,
                'Content-Type': 'application/json'
            }
            json_data=dict_data
            if type(dict_data)==dict:
                json_data=json.dumps(json_data) #dict数据转成str数据
            print(f"headers:,{headers}, body:{json_data}")
            r = requests.post(url, headers=headers, data=json_data)
            print(f"response status:{r.status_code}, reason:{r.reason}")
            result = r.json()
            print(f"<=== response json:\n{result}")
            return result
        except Exception as e:
            print(f'===> http post err {e}')
            return None


    # def get_data_use_task_yiyun_ids(self, api_key, app_id, entry_id):
    #     """
    #     根据任务带出的易云dataid列表获取易云的表单数据，返回表单数据的列表，列表元素为单条表单的数据json
    #     """
    #     from CLPC.framework import FRAME
    #     yy_id_list = FRAME.get_yiyun_ids()
    #     result = []
    #     for data_id in yy_id_list:
    #         data = self.search_data_use_api(api_key, app_id, entry_id, data_id)
    #         result.append(data)
    #     return result