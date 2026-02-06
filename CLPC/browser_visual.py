import time
from time import sleep

import selenium
from selenium import webdriver
from selenium.common.exceptions import *
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
import platform
from CLPC.element import Element
from selenium.webdriver.common.action_chains import ActionChains
import traceback, os
from CLPC.tool import *
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.remote.command import Command
from CLPC.framework import *

def get_operating_system():
    system = platform.system()

    if system == 'Windows':
        return 'Windows'
    elif system == 'Darwin':
        return 'MacOS'
    elif system == 'Linux':
        return 'Linux'
    else:
        return 'Unknown'

def generate_screenshot_driver(width = 1920)-> WebDriver:
    options = Options()

    options.add_argument('--no-sandbox')  # 解决DevToolsActivePort文件不存在的报错
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_argument('--disable-gpu')  # 谷歌文档提到需要加上这个属性来规避bug
    sys_env = get_operating_system()
    options.add_argument('--force-device-scale-factor=3')  # 设置DPI
    options.add_argument(f"--window-size={width},1080")  # 设置浏览器窗口大小
    options.add_argument('--hide-scrollbars')  # 隐藏滚动条, 应对一些特殊页面
    options.add_argument("--disable-software-rasterizer")

    exec_env = FRAME.get_env_arg("RPA_ENV", "")
    if exec_env == "prod" or exec_env == 'test':
        prefs = {
            'profile.default_content_settings.popups':  0,  # 禁止弹窗选择下载路径的弹窗
            "safebrowsing.enabled":                     False,  # 禁用安全浏览检查
            "safebrowsing.disable_download_protection": True
        }  # 设置下载路径
        options.add_experimental_option('prefs', prefs)
        options.add_argument('--headless=new')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-popup-blocking')

        b_path = FRAME.get_env_arg("RPA_CHROME_HEADLESS_SHELL_PATH")
        driver_path = FRAME.get_env_arg("RPA_CHROME_DRIVER_PATH")
        if b_path:
            options.binary_location = b_path
        else:
            options.binary_location = '/opt/rpa/browser/chrome-headless-shell-linux64/chrome-headless-shell'
        if driver_path:
            driver = webdriver.Chrome(executable_path=driver_path, chrome_options=options)
        else:
            driver = webdriver.Chrome(executable_path='/opt/rpa/driver/chrome/chromedriver', chrome_options=options)


    else:
        # 开发机配置
        options.add_argument('--headless=new')  # 提供后台运行的选项

        if sys_env == 'MacOS' or sys_env == 'Linux':


            if sys_env == 'Linux':
                options.binary_location = '/usr/bin/chrome-rpa'
                driver_path = '/usr/bin/chromedriver-rpa'
                try:
                    driver = webdriver.Chrome(executable_path=driver_path, options=options)
                except SessionNotCreatedException:
                    driver = None
                    raise Exception("chromedriver版本不匹配，请连接内网完成自动更新")
                except WebDriverException:
                    raise Exception("未检测到 chromedriver，请确认配置状态")
            else:
                try:
                    driver = webdriver.Chrome(options=options)
                except SessionNotCreatedException:
                    driver = None
                    raise Exception("chromedriver版本不匹配，请连接内网完成自动更新")
                except WebDriverException:
                    raise Exception("未检测到 chromedriver，请确认配置状态")
        else:
            try:
                driver = webdriver.Chrome(options=options)
            except SessionNotCreatedException:
                driver = None
                raise Exception("chromedriver版本不匹配，请连接内网完成自动更新")
            except WebDriverException:
                raise Exception("未检测到 chromedriver，请确认配置状态")

        # 设置允许不安全网站直接下载
    if driver:
        driver.implicitly_wait(20)  # 全局动态等待延时

    ACTIVE_CHROME_DRIVERS.append(driver) # 生产的driver 加入全局变量，用注解进行统一回收
    return driver


def generate_new_driver(driver_path=None, options=None, is_visible=True) -> WebDriver:
    """
        :param driver_path: 默认为空，读取环境变量中的chromedriver，若传值，则已传值为准，仅针对开发环境生效
        :param options: 浏览器设置,默认为空，读取默认配置。
        :param is_visible: 是否前台运行，默认为前台运行
    """

    if not options:
        options = Options()

    options.add_argument('--no-sandbox')  # 解决DevToolsActivePort文件不存在的报错
    options.add_argument('--disable-gpu')  # 谷歌文档提到需要加上这个属性来规避bug
    # options.add_argument('--hide-scrollbars')  # 隐藏滚动条, 应对一些特殊页面
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    # options.add_argument('--incognito')  # 无痕模式
    # options.add_argument('--force-first-run')
    options.add_argument('--no-first-run')  # 避免弹出首次使用提示
    options.add_argument('--ignore-certificate-errors')
    options.add_experimental_option('detach', True)
    options.add_argument('--allow-running-insecure-content')  # 允许不安全链接
    options.add_argument('--accept-insecure-certs')  # 允许不安全链接
    options.set_capability('acceptInsecureCerts', True)

    sys_env = get_operating_system()

    exec_env = FRAME.get_env_arg("RPA_ENV", "")
    if exec_env == "prod" or exec_env == 'test':
        # 服务器配置
        download_dir = './download'
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
        download_dir = download_dir
        prefs = {
            'profile.default_content_settings.popups':  0,  # 禁止弹窗选择下载路径的弹窗
            'download.default_directory':               download_dir,
            "download.directory_upgrade":               True,
            "safebrowsing.enabled":                     False,  # 禁用安全浏览检查
            "safebrowsing.disable_download_protection": True
        }  # 设置下载路径
        options.add_experimental_option('prefs', prefs)
        options.add_argument("window-size=1024x768")  # 设置浏览器窗口大小 低分辨率保障直播顺利
        options.add_argument('--headless=new')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-popup-blocking')
        options.add_argument("--start-maximized")  # 启动时最大化窗口

        # 设置录频
        debugging_port = FRAME.get_available_port()
        options.add_argument('--remote-debugging-address=0.0.0.0')
        options.add_argument('--remote-debugging-port=' + str(debugging_port))
        options.add_argument('--remote-allow-origins=*')

        b_path = FRAME.get_env_arg("RPA_CHROME_HEADLESS_SHELL_PATH")
        driver_path = FRAME.get_env_arg("RPA_CHROME_DRIVER_PATH")
        if b_path:
            options.binary_location = b_path
        else:
            options.binary_location = '/opt/rpa/browser/chrome-headless-shell-linux64/chrome-headless-shell'

        try:
            if driver_path:
                driver = webdriver.Chrome(executable_path=driver_path, chrome_options=options)
            else:
                driver = webdriver.Chrome(executable_path='/opt/rpa/driver/chrome/chromedriver', chrome_options=options)
        except SessionNotCreatedException:
            print(traceback.format_exc())
            driver = None
        is_trigger = FRAME.get_env_arg("RPA_IS_SUPPORT_SCREEN", "True")
        is_pre_check = FRAME.get_env_arg('RPA_IS_PRE_CHECK', 'False')  # 获取是否为预检任务
        is_24_hour = FRAME.get_env_arg("RPA_IF_PERSISTENT", "False")

        if is_trigger == "True" and is_pre_check == 'False' and is_24_hour == "False" : # 只有再非预检任务（即普通的业务任务）时,非24小时任务，且设置为打开录频的状态，才触发录屏
            FRAME.trigger_rim(debugging_port)  # 触发录频


    else:
        # 开发机配置
        if sys_env == 'MacOS' or sys_env == 'Linux':
            download_dir = './download'
            if not os.path.exists(download_dir):
                os.makedirs(download_dir)
            download_dir = download_dir
        else:
            download_dir = 'c:\\RPAdownloads\\'
            if not os.path.exists(download_dir):
                os.makedirs(download_dir)

        prefs = {'profile.default_content_settings.popups': 0,  # 禁止弹窗选择下载路径的弹窗
                 'download.default_directory':              download_dir,  # 设置下载路径
                 "credentials_enable_service":              False,  # 不保存密码
                 "profile.password_manager_enabled":        False,
                 'safebrowsing.enabled':                    False,
                 "safebrowsing.disable_download_protection": True
                 }
        options.add_experimental_option('prefs', prefs)
        options.add_argument("--disable-software-rasterizer")  # 禁用软件光栅化
        # options.page_load_strategy = 'eager'  # 加载策略
        '''
        normal: 页面完全加载后才会返回
        eager: 页面内容加载完成后就返回，可能会导致元素未加载完成
        none: 不等待页面加载，直接返回
        '''
        # options.add_argument('--page-load-strategy=eager')

        if not is_visible:
            options.add_argument('--headless')  # 提供后台运行的选项

        options.add_argument('--remote-debugging-address=0.0.0.0')
        options.add_argument('--remote-debugging-port=9222')
        options.add_argument('--remote-allow-origins=*')

        if sys_env == 'Linux':
            options.binary_location = '/usr/bin/chrome-rpa'
            driver_path = '/usr/bin/chromedriver-rpa'
            try:
                driver = webdriver.Chrome(executable_path=driver_path, options=options)
            except SessionNotCreatedException:
                driver = None
                raise Exception("chromedriver版本不匹配，请连接内网完成自动更新")
            except WebDriverException:
                raise Exception("未检测到 chromedriver，请确认配置状态")
        else:
            try:
                if driver_path:
                    driver = webdriver.Chrome(executable_path=driver_path, options=options)
                else:
                    driver = webdriver.Chrome(options=options)
            except SessionNotCreatedException:
                driver = None
                raise Exception("chromedriver版本不匹配，请连接内网完成自动更新")
            except WebDriverException:
                raise Exception("未检测到 chromedriver，请确认配置状态")

    # 设置允许不安全网站直接下载
    if driver:
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_dir}}
        driver.execute("send_command", params)

        driver.implicitly_wait(10)  # 全局动态等待延时
        driver.maximize_window()

    # with open('stealth.min.js') as f:
    #     js = f.read()
    #
    # driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    #     "source": js
    # })
    # sleep(2)
    ACTIVE_CHROME_DRIVERS.append(driver)
    return driver


def connect_exist_driver(driver_path=None) -> WebDriver:
    """
    连接已经打开的浏览器
    """
    FRAME.clear_log()
    sys_env = get_operating_system()

    if sys_env == 'MacOS' or sys_env == 'Linux':
        download_dir = './download'
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
        download_dir = download_dir
    else:
        download_dir = 'c:\\RPAdownloads\\'
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
    try:
        options = Options()
        options.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
        if driver_path:
            driver = webdriver.Chrome(executable_path=driver_path, options=options)
        else:
            driver = webdriver.Chrome(options=options)
    except SessionNotCreatedException:
        raise Exception("未检测到已打开的浏览器，请从第一步开始重新运行")
    if driver:
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_dir}}
        driver.execute("send_command", params)

        driver.implicitly_wait(10)  # 全局动态等待延时
    return driver


class Browser:
    sys_env = get_operating_system()

    if sys_env == 'MacOS' or sys_env == 'Linux':
        download_dir = './download'
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
        else:
            TOOL.del_file(download_dir)
    else:
        download_dir = 'c:\\RPAdownloads\\'
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
        else:
            TOOL.del_file(download_dir)

    def __init__(self, driver: WebDriver, ele_file_name='xpath.json'):
        """
        :param driver: 已经初始化的webdriver
        :param ele_file_name: 元素文件,默认为当前目录的ele.json文件
        """
        if not os.path.exists(ele_file_name):
            current_dir = os.path.dirname(os.path.abspath(__file__))
            ele_json_path = os.path.join(current_dir, 'xpath.json')
            ele_file_name = ele_json_path
        self.ele_map = Element(ele_file_name)
        self.driver = driver
        self.driver.implicitly_wait(10)  # 全局动态等待延时

    def __del__(self):
        pass
        # if self.driver:
        #     self.driver.quit()

    def get_driver(self):
        """
        获取driver对象
        """
        if self.driver:
            return self.driver

    def close_all(self):
        """
        关闭浏览器
        """
        self.driver.quit()

    def close(self):
        """
        关闭当前page；
        关闭后需要使用catch()方法手动获取新的页面
        """
        self.driver.close()

    def maximize(self):
        """
        最大化窗口
        """
        self.driver.maximize_window()

    def minimize(self):
        """
        最小化窗口
        """
        self.driver.minimize_window()

    def restore(self):
        """
        刷新页面
        """
        self.driver.refresh()

    def forward(self):
        """
        前进一页
        """
        self.driver.forward()

    def backward(self):
        """
        后退一页
        """
        self.driver.back()

    def click(self, ele_name, simulate=False, type='left', arg=None, index=0):
        """
        点击指定控件
        :param ele_name: 控件名称
        :param simulate: 是否模拟点击
        :param type: left单击 right右击
        :param arg: 元素变量，默认为空
        :param index: 元素序号，默认第一个
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)

        # self.driver.execute_script("arguments[0].click()", ele)  # 使用js方式点击

        if type == 'left':
            try:
                if simulate:
                    ActionChains(self.driver).click(ele).perform()
                else:
                    ele.click()
            except (ElementNotInteractableException, ElementClickInterceptedException):
                print(f"页面元素{ele_name}不可点击，尝试使用js点击，建议检查并优化元素")
                self.click_by_js(ele_name, arg=arg, index=index)  # 使用js方式点击
        elif type == 'right':
            ActionChains(self.driver).context_click(ele).perform()
        else:
            raise Exception('type 参数值错误')
        time.sleep(0.5)

    def click_by_js(self, ele_name, arg=None, index=0):
        """
        使用 js 方法点击元素，可以点击被遮挡的元素
        :param ele_name:
        :param
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        self.driver.execute_script("arguments[0].click()", ele)  # 使用js方式点击

    def doubleclick(self, ele_name, arg=None, index=0):
        """
        双击指定控件
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        ActionChains(self.driver).double_click(ele).perform()
        time.sleep(1)

    def create(self, url):
        """
        在当前tab打开指定链接的页面
        """
        # todo 加入埋点动作
        try:
            self.driver.get(url)
        except:
            raise Exception("服务器无法访问当前网址：{}，请确认网址输入是否正确，并确认是否为内网网址。".format(url))
        # self.send_keys('thisisunsafe')  # 用于访问 http 网址时使用
        # self.maximize()

    # def zoom_page(self, percent:float8):
    #     """
    #     缩放页面
    #     @percent: 缩放的比例,0.1-10.0
    #     """
    #     if 0.1<=percent<=10.0:
    #         pass
    #     else:
    #         print('缩放比例不支持')
    #         raise(Exception('缩放比例不支持'))
    #     js_str = "document.body.style.webkitTransform=\'scale({})\'".format(percent)

    #     self.driver.execute_script(js_str)
    #     time.sleep(0.5)

    # def zoom_reset(self):
    #     self.zoom_page(1.0)
    #     time.sleep(0.5)

    def input_text(self, ele_name: str, text: str, arg=None, index=0):
        """
        输入文本
        :param ele_name: 控件名称
        :param text: 待输入的文本
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        ele.send_keys(text)
        time.sleep(0.5)

    def input_text_use_action(self, text):
        """
        使用ActionChains输入文本
        :param text: 待输入的文本
        """
        ActionChains(self.driver).send_keys(text).perform()

        time.sleep(0.5)

    def clear_text_use_action(self):
        """
        使用ActionChains清空文本
        """
        sys_env = get_operating_system()
        action = ActionChains(self.driver)

        if sys_env == 'MacOS':
            action.key_down(Keys.COMMAND).perform()
            action.send_keys('a').key_up(Keys.COMMAND).perform()
            sleep(1)
            action.send_keys(Keys.DELETE).perform()
        else:
            action.key_down(Keys.CONTROL).perform()
            action.send_keys('a').key_up(Keys.CONTROL).perform()
            sleep(1)
            action.send_keys(Keys.DELETE).perform()
        time.sleep(1)

    def input_verify_code(self, ele_name: str, code: str, account: str, arg=None, index=0):
        """
        输入验证码
        param: ele_name: 控件名称
        param: code: 待输入的文本
        param: account: 账号
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        tool_code = TOOL.get_verify_code(account)
        if tool_code:
            ele.send_keys(tool_code)
        else:
            ele.send_keys(code)
        time.sleep(0.5)

    def clear_input(self, ele_name, arg=None, index=0):
        """
        清楚输入框
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        ele.clear()
        time.sleep(0.5)

    def send_keys(self, key, ele_name=None, arg=None, index=0):  # todo # 键位输入
        if type(key) == Keys:
            if ele_name:
                ele = self.__find_element(ele_name, arg=arg, index=index)
                ActionChains(self.driver).send_keys_to_element(ele, key)
            else:
                ActionChains(self.driver).send_keys(key)
        else:
            pass
            # 需要对输入的信息进行判断
        time.sleep(0.5)

    def isvisible(self, ele_name, arg=None, index=0):
        """
        判断元素是否可见,元素需存在才可判断,否则报错
        """
        self.driver.implicitly_wait(1)  # 全局动态等待延时
        try:
            ele = self.__find_element(ele_name, arg=arg, index=index)
        except:
            self.driver.implicitly_wait(10)  # 全局动态等待延时

            return False
        try:
            res = ele.is_displayed()
        except selenium.common.exceptions.StaleElementReferenceException:
            res = False

        self.driver.implicitly_wait(10)  # 全局动态等待延时

        return res

    def wait_loaded(self, ele_name, arg=None, timeout=10):
        """
        等待元素加载
        """
        for i in range(timeout):
            if self.count(ele_name, arg=arg) > 0:
                break
            time.sleep(1)
        else:
            print('元素：{}未加载成功'.format(ele_name))
            raise Exception("元素未加载成功，请确认流程是否正确".format(ele_name))

    def wait_disappear(self, ele_name, arg=None, timeout=10):
        """
        等待元素消失（不存在）
        """
        for i in range(timeout):
            if self.count(ele_name, arg=arg) == 0:
                break
            time.sleep(1)
        else:
            print('元素：{}未消失'.format(ele_name))
            raise Exception("元素未消失，请确认流程是否正确".format(ele_name))

    def wait_invisible(self, ele_name, arg=None, timeout=10):
        """
        等待元素不可见（不存在，或者存在但不可见）

        """
        for i in range(timeout):
            if self.count(ele_name, arg) == 0:
                break
            elif not self.isvisible(ele_name, arg=arg):
                break
            time.sleep(1)
        else:
            print('元素：{}依然可见，请确认流程是否正确'.format(ele_name))
            raise Exception("元素依然可见，请确认流程是否正确".format(ele_name))

    def count(self, ele_name, arg=None) -> int:
        """
        获取元素个数
        """
        self.driver.implicitly_wait(1)  # 全局动态等待延时

        try:
            self.driver.switch_to.default_content()
        except:
            pass
        xpath = self.ele_map.get_xpath(ele_name)

        if arg and arg != '':
            if isinstance(arg, list):  # 处理多参数
                for a in arg:
                    xpath = xpath.replace('$', a, 1)
            elif isinstance(arg, str):
                xpath = xpath.replace('$', arg)
            else:
                print("不支持的参数类型")
                raise Exception("不支持的参数类型")

        iframes = self.ele_map.get_iframes(ele_name)

        if len(iframes) == 0:
            try:
                eles = self.driver.find_elements(By.XPATH,xpath)
            except:
                eles = []
        else:
            try:
                for f in iframes:
                    tmp_frame = self.driver.find_element_by_xpath(f)
                    self.driver.switch_to.frame(tmp_frame)

                eles = self.driver.find_elements(By.XPATH,xpath)
            except:
                eles = []
        self.driver.implicitly_wait(10)  # 全局动态等待延时
        return len(eles)

    def scroll_into_view(self, ele_name: str, arg=None, index=0):
        """
        拖动页面使元素可见
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        ele.location_once_scrolled_into_view

    def text(self, ele_name: str, arg=None, index=0) -> str:
        """
        获取指定元素的文本
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        result = ele.text
        return result

    def get_attribute(self, ele_name: str, arg=None, index=0) -> str:
        ele = self.__find_element(ele_name, arg=arg, index=index)
        attribute = ele.get_attribute("value")

        return attribute

    def upload(self, ele_name: str, file_name: str, arg=None, index=0):
        """
        上传文件,需确认元素为input,且类型为 file
        文件是否上传完成,需要事后确认
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        abs_path = os.path.abspath(file_name)
        ele.send_keys(abs_path)
        time.sleep(3)

    def upload_files(self, ele_name:str, file_name_list:list, arg=None, index=0):
        """
        上传多个文件,需确认元素为input,且类型为 file
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        abs_path_list = [os.path.abspath(file_name) for file_name in file_name_list]
        spliter = str(os.linesep)
        ele.send_keys(spliter.join(abs_path_list))
        time.sleep(3)

    def attr(self, ele_name: str, type_name: str, arg=None, index=0) -> str:
        """
        获取元素的的指定属性值
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        attr = ele.get_attribute(type_name)

        return attr

    def option_by_index(self, ele_name, select_index: int, arg=None, index=0):
        """
        根据index选择对应下拉框
        元素需要为select类型
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        Select(ele).select_by_index(select_index)

    def option_by_text(self, ele_name, text: str, arg=None, index=0):
        """
        根据文本值选择对应下拉框
        元素需要为select类型
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        Select(ele).select_by_visible_text(text)

    def get_options(self, ele_name, mode='all', arg=None, index=0):
        """
        获取选项值,
        mode=all,获取全部选项；mode=selected,获取已选择选项
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        result = []
        if mode == 'all':
            result = Select(ele).options

        if mode == 'selected':
            result = Select(ele).all_selected_options
        return result

    def get_checked_state(self, ele_name, arg=None, index=0):
        """
        获取当前勾选状态
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        result = ele.is_selected()
        return result

    def set_checked_state(self, ele_name, value=True, arg=None, index=0):
        """
        设置勾选状态
        :param value  True 勾选,False 取消勾选
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        state = ele.is_selected()

        if value:
            if not state:
                ele.click()
        else:
            if state:
                ele.click()
        time.sleep(0.5)
        return

    def pos(self, ele_name, arg=None, index=0):
        """
        获取元素的坐标
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        pos = ele.location
        size = ele.size
        return pos,size
    
    def pos_ele_center(self, ele_name, arg=None, index=0):
        """
        获取元素中点的坐标
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        location = ele.location
        size = ele.size
        x = int(location['x'] + size['width'] / 2)
        y = int(location['y'] + size['height'] / 2)
        return {'x':x, 'y':y}
    
    
    def mouse_move_pos_ele_center_click(self, ele_name, arg=None, index=0):
        """
        鼠标移动到元素中点坐标点击
        """
        location = self.pos_ele_center(ele_name, arg=arg, index=index)
        x = location['x']
        y = location['y']
        ActionChains(self.driver).move_by_offset(x,y).click().perform()
        # print(x,y)
        

    def mouse_move(self, ele_name, arg=None, index=0):
        """
        鼠标移入元素中间
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        ActionChains(self.driver).move_to_element(ele).perform()
        time.sleep(0.5)

    def drag_then_drop(self, source_ele, target_ele, source_arg=None, target_arg=None, source_index=0, target_index=0):
        """
        将初始元素拖动至目标元素位置
        """
        ele_sor = self.__find_element(source_ele, arg=source_arg, index=source_index)
        ele_tar = self.__find_element(target_ele, arg=target_arg, index=target_index)
        ActionChains(self.driver).drag_and_drop(ele_sor, ele_tar).perform()
        time.sleep(0.5)
        return

    def execute_js(self, js_str, ele_name=None, arg=None):
        try:
            self.driver.switch_to.default_content()
        except:
            pass
        if ele_name:
            iframes = self.ele_map.get_iframes(ele_name)
            if len(iframes) > 0:
                for i in iframes:
                    tmp = self.driver.find_element_by_xpath(i)
                    self.driver.switch_to.frame(tmp)
                self.driver.execute_script(js_str)
                self.driver.switch_to.default_content()
        else:
            self.driver.execute_script(js_str)

    def get_table(self, ele_name, arg=None):
        """
        获取指定表格的值,返回二维数组
        使用该方法时，需要保证表格元素为 tr， td 构成
        """
        table = self.__find_element(ele_name, arg=arg)
        rows = table.find_elements_by_tag_name('tr')
        result = []
        for row in rows:
            # 在每行中获取所有单元格
            cells = row.find_elements_by_tag_name('td')
            cell_data = [cell.text for cell in cells]
            result.append(cell_data)
        return result

    def get_url(self):
        """
        获取当前页面url
        """
        return self.driver.current_url

    def get_title(self):
        """
        获取当前页面标题
        """
        return self.driver.title

    def screenshot(self, ele_name: str, png_name: str, arg=None, index=0):
        """
        对指定元素进行截图
        """
        ele = self.__find_element(ele_name, arg=arg, index=index)
        ele.screenshot(png_name)
        time.sleep(2)
        if os.path.exists(png_name):
            return
        else:
            print('截图失败')

    def __find_element(self, ele_name: str, arg=None, index=0) -> WebElement:
        """
        更具捕捉好的元素名称捕获元素
        :param ele_name: 元素名称
        :param arg: 元素参数，支持单个参数或多参数，多参数以列表形式传入
        """
        xpath = self.ele_map.get_xpath(ele_name)
        if xpath == 'None':
            raise Exception("元素：[{}] 未捕捉".format(ele_name))
        try:
            self.driver.switch_to.default_content()
        except:
            pass
        if arg and arg != '':
            if isinstance(arg, list):  # 处理多参数
                for a in arg:
                    xpath = xpath.replace('$', a, 1)
            elif isinstance(arg, str):
                xpath = xpath.replace('$', arg)
            else:
                print("不支持的参数类型")
                raise Exception("不支持的参数类型")
        iframes = self.ele_map.get_iframes(ele_name)
        if len(iframes) == 0:
            eles = self.driver.find_elements(By.XPATH,xpath)
            if len(eles) > 0:
                ele = eles[index]
            else:
                if arg:
                    info = "元素: [{}-{}] 未找到".format(ele_name, arg)
                else:
                    info = "元素: [{}] 未找到".format(ele_name)
                raise Exception(info)
        else:
            try:
                for f in iframes:
                    tmp_frame = self.driver.find_element_by_xpath(f)
                    self.driver.switch_to.frame(tmp_frame)
                eles = self.driver.find_elements(By.XPATH,xpath)
            except:
                raise Exception("元素: [{}] 未找到".format(ele_name))
            if len(eles) > 0:
                ele = eles[index]
            else:
                if arg:
                    info = "元素: [{}-{}] 未找到".format(ele_name, arg)
                else:
                    info = "元素: [{}] 未找到".format(ele_name)
                raise Exception(info)
        try:
            if not ele.is_displayed():
                ele.location_once_scrolled_into_view  # 滚动页面使元素可见
            sleep(1)
        except:
            pass
        return ele

    def down_by_element(self, ele_name, file_name, download_timeout=60, simulate=False, arg: str = None, index=0):
        """
        将文件下载至当前工作目录
        """

        before = set(os.listdir(self.download_dir))
        down_files = set()
        time.sleep(2)
        self.click(ele_name, simulate, arg=arg, index=index)

        time.sleep(2)
        for ti in range(download_timeout):
            after = set(os.listdir(self.download_dir))
            if after:
                for tf in after:
                    if '.crdownload' in tf:
                        time.sleep(1)
                        break
                else:
                    down_files = after - before
                    break

            time.sleep(1)
        else:
            print("文件下载失败")
            raise Exception('文件下载失败')
        # down_name = self.getDownLoadedFileName(1)
        down_name = down_files.pop()
        print('filename:', down_name)

        source_path = os.path.join(self.download_dir, down_name)
        for i in range(download_timeout):
            if TOOL.file_exist(source_path) and TOOL.get_FileSize_kb(source_path) > 2:
                break
            else:
                time.sleep(1)
        else:
            print("文件下载失败")
            raise Exception('文件下载失败')

        current_dir = os.getcwd()  ## todo 增加支持指定目录
        target_path = os.path.join(current_dir, file_name)
        TOOL.mov_file(source_path, target_path)

    def catch(self, value, type='title', mode='equal', timeout=10):
        """
        根据title 或者 url 切换对应tab
        type: title, url
        mode: equal, contain
        """
        # todo 需要增加埋点
        flag = False
        for i in range(timeout):
            tabs = self.driver.window_handles
            for t in tabs:
                self.driver.switch_to.window(t)
                title = self.driver.title
                url = self.driver.current_url
                if type == 'title':
                    if mode == 'equal':
                        if title == value:
                            flag = True
                            break
                    if mode == 'contain':
                        if value in title:
                            flag = True
                            break
                if type == 'url':
                    if mode == 'equal':
                        if url == value:
                            flag = True
                            break
                    if mode == 'contain':
                        if value in url:
                            flag = True
                            break
            if flag:
                break
            else:
                time.sleep(1)
        else:
            raise Exception('未找到指定页面')

    def set_date(self, ele, date_str, arg=None, index=0):
        """
        通过value设置日期
        页面元素需要有value属性
        :param ele: 页面元素
        :param date_str: 日期
        :param arg: 元素参数
        :param index: 元素序号
        """
        ele = self.__find_element(ele, arg=arg, index=index)
        self.driver.execute_script("arguments[0].value='{}'".format(date_str), ele)

    def full_screen_shot(self, png_name):
        """
        全屏截图
        :param png_name: 截图文件路径
        """
        try:
            self.driver.save_screenshot(png_name)
        except:
            print('截图失败')

    def deal_alert(self, is_accept=True):
        """
        处理浏览器上部弹窗
        :param is_accept: 是否接受弹窗 True 确认 False 取消
        """
        try:
            alert = self.driver.switch_to.alert
            if is_accept:
                alert.accept()
            else:
                alert.dismiss()
        except:
            pass

    # method to get the downloaded file name
    def getDownLoadedFileName(self, waitTime):
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[0])
        self.driver.execute_script("window.open('', '_blank');")
        self.catch('blank', type='url', mode='contain')
        self.create("chrome://downloads")

        self.catch('downloads', type='url', mode='contain')
        endTime = time.time() + waitTime
        while True:
            try:
                # get downloaded percentage
                downloadPercentage = self.driver.execute_script(
                    # "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress').value")
                    "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress')")
                print('percent', downloadPercentage)

                # check if downloadPercentage is 100 (otherwise the script will keep waiting)
                if downloadPercentage == 100:
                    # return the file name once the download is completed
                    return self.driver.execute_script(
                        "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
            except:
                print(traceback.format_exc())
                pass
            time.sleep(1)
            if time.time() > endTime:
                return self.driver.execute_script(
                    "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
    
    
    def get_target_response(self,inf_url):
        target_response = None
        target_url = inf_url
        request_ids = []  # 存储所有请求ID

        # 1. 启用网络监控
        self.driver.execute_cdp_cmd("Network.enable", {})

        # 2. 拦截请求，记录所有requestId（通过设置空header触发）
        def collect_request_ids(params):
            request_ids.append(params["requestId"])
        # 注意：Selenium 3.x 没有add_cdp_listener，需用 execute_cdp_cmd 注册事件
        self.driver.execute_cdp_cmd("Network.setRequestInterception", {
            "patterns": [{"urlPattern": "*", "resourceType": "Document", "interceptionStage": "HeadersReceived"}]
        })

        # 3. 触发页面刷新，发起请求
        self.driver.get(self.driver.current_url)

        # 4. 轮询检查目标响应（设置超时时间，避免阻塞）
        timeout = 10
        start_time = time.time()
        while time.time() - start_time < timeout:
            if not request_ids:
                time.sleep(0.5)
                continue
            # 遍历所有请求ID，查询响应信息
            for req_id in request_ids[:]:  # 浅拷贝，避免遍历时列表变化
                try:
                    # 获取响应详情
                    response = self.driver.execute_cdp_cmd("Network.getResponseDetails", {"requestId": req_id})
                    if response["response"]["url"] == target_url:
                        # 获取响应体
                        response_body = self.driver.execute_cdp_cmd("Network.getResponseBody", {"requestId": req_id})
                        target_response = {
                            "url": response["response"]["url"],
                            "status_code": response["response"]["status"],
                            "headers": response["response"]["headers"],
                            "body": response_body.get("body", ""),
                            "base64_encoded": response_body.get("base64Encoded", False)
                        }
                        request_ids.remove(req_id)
                        break  # 找到目标，退出循环
                except Exception as e:
                    # 忽略未完成的请求
                    request_ids.remove(req_id)
                    continue
            if target_response:
                break
            time.sleep(0.5)

        # 5. 关闭网络监控和拦截
        self.driver.execute_cdp_cmd("Network.disable", {})
        return target_response
    
    def simulate_area_select_and_switch(self, start_x, start_y, end_x, end_y,base_url):
        """
        模拟鼠标框选区域并切换到新打开的标签页
        :param start_x: 框选起始点X坐标
        :param start_y: 框选起始点Y坐标
        :param end_x: 框选结束点X坐标
        :param end_y: 框选结束点Y坐标
        :return: 新标签页的driver句柄，失败返回None
        """
        try:
            # 1. 模拟鼠标拖拽框选：按住左键 → 拖动到结束点 → 松开
            time.sleep(5) # 等待页面加载完成
            action = ActionChains(self.driver)
            action.move_by_offset(start_x, start_y)  # 移动到起始坐标
            action.click_and_hold()  # 按住左键
            action.move_by_offset(end_x - start_x, end_y - start_y)  # 拖动到结束坐标
            action.release()  # 松开左键
            action.perform()  # 执行动作链
            time.sleep(2)  # 等待跳转页面加载（根据实际情况调整）

            # # 2. 获取所有标签页句柄，切换到最新打开的标签页
            # all_handles = self.driver.window_handles
            # # 切换到最后一个句柄（新打开的页面）
            # self.driver.switch_to.window(all_handles[-1])
            self.catch(base_url+"/transactionList", mode="contain", type='url', timeout=60)
            
            print(f"成功框选并切换到新标签页，当前页面标题: {self.driver.title}")
            return self.driver.current_window_handle
        
        except Exception as e:
            print(f"模拟框选失败: {e}")
            return None
