from CLPC.browser_visual import Browser
from time import sleep, time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from CLPC.element import Element
from selenium.webdriver.chrome.webdriver import WebDriver
import os,traceback

from CLPC.framework import FUNC_USAGE_TRACKER
from CLPC.tool import TOOL

@FUNC_USAGE_TRACKER
class BDP(Browser):
    def __init__(self, driver: WebDriver):
        """
        登录统一工作台，统一工作台操作组件\n
        初始化后需要传入用户信息
        @driver: 已初始化的webdriver
        """
        # super().__init__(driver)
        self.driver = driver
        current_dir = os.path.dirname(os.path.abspath(__file__))
        ele_json_path = os.path.join(current_dir, 'bdp.json')
        self.ele_map = Element(ele_json_path)  # 加载组件自己的元素文件


        self.userName = ""
        self.passWord = ""
        self.verfiCode = ""

    def loginMainSystem(self):
        """ 进入数据分析平台\n
        用户名 密码 验证码 3个都在__init__ 初始化时可做替换
        :param userName: 用户名 :<str>\n
        :param password: 密码 :<str>\n
        :param verfiCode: 验证码 :<str>\n
        :param chrome_path : chrome.exe路径,\n
        :return: 无\n
        """
        userName = self.userName
        password = self.passWord
        tool_get_verfy_code = TOOL.get_verify_code(userName)
        if tool_get_verfy_code:
            verfiCode = tool_get_verfy_code
        else:
            verfiCode = self.verfiCode

        self.create(url="http://9.0.8.6:7021/portal/UAMLoginByEntWechat") # 创建网页对象
        self.maximize()
        sleep(1)
        # self.doubleclick("返回账号登录")
        self.click("返回账号登录", simulate=True)
        self.input_text("账号", userName)
        self.input_text("密码", password)
        self.input_text("验证码", verfiCode)
        self.click("登录")
        sleep(1)
        self.wait_loaded("BAS", timeout=30)
        if self.count("BAS") > 0:
            pass
        else:
            raise Exception("登录失败，请检查登录信息是否正确")

    def enterSubSystem(self, subSystem):
        """ 进入子系统\n

        :param subSystem: 子系统, "闪查plus" :<str>\n
        :return: 无\n
        """
        print('当前tab title:', self.get_title())
        self.wait_loaded(subSystem)
        if self.count("今天隐藏") != 0:
            self.click("今天隐藏")
        self.click(subSystem)
        sleep(1)

        if subSystem == "BAS":
            self.catch('BAS', mode='contain', timeout=30)
            sleep(1)
        elif subSystem == '车险超多维':
            self.catch('海致BDP', mode='contain', timeout=60)
            self.wait_loaded("仪表盘")
            sleep(1)
        else:
            tabs = self.driver.window_handles
            self.driver.switch_to.window(tabs[-1])

    def enterBDP(self):
        """
        进入车险超多维主页
        """
        self.loginMainSystem()
        self.enterSubSystem('车险超多维')
        self.wait_loaded("仪表盘", timeout=60)
        sleep(5)

    def enter_dash(self, route):
        """
        :param route: 文件夹>文件夹>报表名（不能重复，否则用第一个）使用 > 分隔
        """
        folder_list = route.split('>')[:-1]
        dash_name = route.split('>')[-1]
        for f in folder_list:
            isopen = self.attr("仪表盘文件目录-折叠状态","class", arg=f)
            if "open" in isopen:
                continue
            else:
                self.click('仪表盘文件目录', arg=f)
                sleep(1)
        self.click('仪表板页面', arg=dash_name)

    def refresh_dash(self, dash_name):
        """
        刷新仪表盘
        """
        self.mouse_move("标题", arg=dash_name)
        self.click('更新', arg=dash_name)

    def export_dash(self, dash_name, file_name, timeout=60):
        """
        导出仪表盘
        """
        self.scroll_into_view("标题", arg=dash_name)
        self.mouse_move("标题", arg=dash_name)
        self.scroll_into_view('更多', arg=dash_name)
        self.click('更多', arg=dash_name)
        sleep(1)
        self.scroll_into_view('导出Excel', arg=dash_name)
        self.down_by_element("导出Excel", file_name, arg=dash_name, download_timeout=timeout)

    def global_filter(self, filter_name, filter_value):
        """
        全局筛选
        :param filter_name: 筛选器
        :param filter_value: 筛选值
        """
        self.click("全局筛选下拉", arg=filter_name)
        sleep(1)
        self.click("全局下拉选项", arg=filter_value)

    def dash_filter(self, dash_name, filter_name, filter_value):
        """
        图内筛选
        :param dash_name: 仪表盘名称
        :param filter_name: 筛选器
        :param filter_value: 筛选值
        """
        self.mouse_move("标题", arg=dash_name)
        self.click("筛选", arg=dash_name)
        self.click("表内筛选下拉", arg=[dash_name, filter_name])
        sleep(1)
        self.wait_loaded("表内下拉选项", arg=[dash_name, filter_value], timeout=30)
        self.click("表内下拉选项", arg=[dash_name, filter_value])

    def wait_dash_load(self, dash_name, timeout=60):
        """
        等待仪表盘加载完成
        :param dash_name: 仪表盘名称
        :param timeout: 超时时间
        """
        for i in range(timeout):
            if self.isvisible("标题", arg=dash_name):
                break
            else:
                sleep(1)
        self.wait_invisible("报表加载", arg=dash_name, timeout=timeout)

    def enter_dash_edit(self, dash_name):
        """
        进入仪表盘详情编辑模式
        """
        self.mouse_move("标题", arg=dash_name)
        self.click("编辑", arg=dash_name)
        self.wait_loaded("返回", timeout=30)

    def edit_dash_filter_expression(self, filter_name, expression):
        """
        编辑仪表盘详情筛选
        :param filter_name: 筛选器
        :param expression: 完整表达式
        """
        self.click("编辑筛选器", arg=filter_name)
        self.click("勾选表达式")
        self.clean_expression()
        sleep(1)
        self.click("筛选器筛选输入")
        sleep(1)
        self.input_text_use_action(expression)
        self.click("弹窗确定")
        self.wait_invisible("筛选验证通过", timeout=30)

    def clean_expression(self):
        """
        清除筛选器的表达式
        """
        self.click("筛选器筛选输入")
        self.clear_text_use_action()

    def return_to_dash_table(self, is_save=True):
        """
        返回仪表盘主界面
        """
        self.click("返回")
        if is_save:
            self.click("弹窗确定")
        else:
            self.click("弹窗取消")
        sleep(3)

    def export_whole_dash_2_img(self, file_name, scale=1, timeout=60):
        """
        将整个仪表盘导出为图片
        :param file_name: 图片文件名
        :param scale: 图片倍数
        :param timeout: 下载超时时间
        """
        self.click("整体仪表盘-更多")
        self.click("整体仪表盘-导出图片")
        if scale == 1:
            self.click("整体仪表盘-导出1倍")
        elif scale == 2:
            self.click("整体仪表盘-导出2倍")
        else:
            raise Exception("图片倍数填写有误，只能1或2")
        self.down_by_element("弹窗确定", file_name, download_timeout=timeout)

    def export_whole_dash_2_PDF(self, file_name, timeout=60):
        """
        将整个仪表盘导出为图片
        :param file_name: 图片文件名
        :param scale: 图片倍数
        :param timeout: 下载超时时间
        """
        self.click("整体仪表盘-更多")
        self.click("整体仪表盘-导出PDF")
        self.down_by_element("弹窗确定", file_name, download_timeout=timeout)
