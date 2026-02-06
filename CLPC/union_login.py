import platform
from CLPC.browser_visual import Browser
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from CLPC.element import Element
from selenium.webdriver.chrome.webdriver import WebDriver

from CLPC.framework import FUNC_USAGE_TRACKER
from CLPC.tool import TOOL
import os

@FUNC_USAGE_TRACKER
class UNILOGIN(Browser):
    def __init__(self, driver: WebDriver):
        """
        内网统一登录组件\n
        初始化后需要传入用户信息
        @driver: 已初始化的webdriver
        """
        # super().__init__(driver)
        self.driver = driver

        current_dir = os.path.dirname(os.path.abspath(__file__))
        ele_json_path = os.path.join(current_dir, 'union_login.json')
        self.ele_map = Element(ele_json_path)

        self.userName = ""
        self.passWord = ""
        self.verfiCode = ""

    def __del__(self):
        pass

    def login_gateway(self):
        """
        登录内网门户首页
        """
        userName = self.userName
        password = self.passWord

        tool_get_verfy_code = TOOL.get_verify_code(userName)
        if tool_get_verfy_code:
            verfiCode = tool_get_verfy_code
        else:
            verfiCode = self.verfiCode

        menhu_url = "http://9.0.8.101/UAM/loginewm.jsp"

        for i in range(3):
            self.create(menhu_url)
            self.click("内网门户-账号登录切换")
            self.input_text("用户名", userName)
            self.input_text("密码", password)
            self.input_text("验证码", verfiCode)
            self.click("登录")
            for j in range(10):
                if self.count("登录中心") > 0:
                    break
            else:
                print(f"第{i+1}次登录失败")
                continue
            print("登录成功")
            break
        else:
            raise Exception("3次登录内网门户失败，请检查网络情况与用户信息填写是否准确")

    def enter_subsys_from_login_center(self, subsys_name):
        """
        通过登录中心，登录门户有的子系统
        """
        self.catch("中国人寿财险门户系统")
        self.click("登录中心")
        sleep(3)
        self.click("登录中心业务系统", arg=subsys_name)
        sleep(3)
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[-1])

    def enter_subsys_from_business_center(self, subsys_name):
        """
        通过业务中心，登录门户有的业务系统
        """
        self.catch("中国人寿财险门户系统")
        self.click("业务门户")
        sleep(3)
        if subsys_name in ['系统导航', '统一工作台', '运营平台']:
            self.click("业务门户-快速导航", arg=subsys_name)
        else:
            self.click("业务门户子系统", arg=subsys_name)
        sleep(3)
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[-1])

