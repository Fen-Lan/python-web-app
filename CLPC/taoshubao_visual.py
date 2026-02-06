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


@FUNC_USAGE_TRACKER
class Tao(Browser):
    def __init__(self, driver: WebDriver):
        """
        登录数据分析平台，多维分析操作组件\n
        初始化后需要传入用户信息
        @driver: 已初始化的webdriver
        """
        self.driver = driver
        sys = platform.system()
        if sys == 'Windows':
            self.ele_map = Element('CLPC\\taoshubao.json') # 加载组件自己的元素文件
        if sys == 'Darwin' or sys == 'Linux':
            self.ele_map = Element('CLPC/taoshubao.json') # 加载组件自己的元素文件

        self.userName = "", 
        self.password = "", 
        self.verfiCode = "",    

    def __del__(self):
        pass

    def get_taoshubao_browser(self):
        return self.browser

    def enterPage(self, name):
        """ 跳转进入指定页面\n
        @ name: '淘数宝' :<str>\n
        """

        self.loginMainSystem()
        self.enterSubSystem(name)
        sleep(3)
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[-1])
        sleep(5)



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
        password = self.password
        tool_get_verfy_code = TOOL.get_verify_code(userName)
        if tool_get_verfy_code:
            verfiCode = tool_get_verfy_code
        else:
            verfiCode = self.verfiCode

        for i in range(3):
            self.create(url="http://9.0.8.6:7021/portal/UAMLoginByEntWechat") # 创建网页对象
            #self.maximize()
            sleep(1)
            # self.doubleclick("返回账号登录")
            self.click("返回账号登录", simulate=True)
            self.input_text("账号", userName)
            self.input_text("密码", password)
            self.input_text("验证码", verfiCode)
            self.click("登录")
            sleep(2)
            for j in range(10):
                if self.count("BAS") > 0 or self.count("今天隐藏") > 0 or self.count("我已知晓") > 0:
                    break
                else:
                    sleep(1)
            else:
                print(f"第{i + 1}次登录失败")
                continue
            print("登录成功")
            break
        else:
            raise Exception("3次登录均失败, 请检查账号密码验证码是否正确")

    def enterSubSystem(self, subSystem):
        """ 进入子系统\n
        
        :param subSystem: 子系统, "淘数宝" :<str>\n
        :return: 无\n
        """        
        print('当前tab title:', self.get_title())     
        self.wait_loaded(subSystem)
        if self.count("今天隐藏") != 0:
            self.click("今天隐藏")
        self.click(subSystem)
        sleep(3)
        
        if subSystem == "BAS":
            self.catch('BAS',mode='contain', timeout=30)
            sleep(2)
        elif subSystem == '淘数宝':
            self.catch('淘数宝', timeout=30)
            self.wait_loaded("我的订单")
            sleep(3)
        else:
            tabs = self.driver.window_handles
            self.driver.switch_to.window(tabs[-1]) # 切换至最后一个tab

  
    def search(self, billNo):
        """ 搜索订单\n
        :billNo: 订单号,例如: 'd560d320' :<str>\n  
        :return: 订单已失效/订单未失效\n
        """
        self.restore()
        sleep(2)
        self.click('我的订单')
        sleep(2)
        self.wait_loaded("订单编号")
        self.input_text('订单编号',billNo)
        self.click('查询订单')
        sleep(5)
        if self.count('再来一单'):
            self.click('再来一单')
            sleep(2)
            return '订单未失效'
        else:
            return '订单已失效'

    def modifyDate(self,dateStart,dateEnd):#
        """ 修改提数日期\n
        :dateStart: 开始日期,例如: '2024-01-01' :<str>\n  
        :dateEnd: 截止日期,例如: '2024-02-01' :<str>\n        
        :return: 无\n
        """
        self.click('日期')
        sleep(1)
        self.click('日期-编辑')
        sleep(1)
        print(dateStart)
        print(dateEnd)
        self.clear_input('开始日期')
        self.input_text('开始日期',dateStart)
        sleep(1)
        try:
            self.click('开始日期-关闭')
        except:
            print('开始日期无关闭按钮')
        self.clear_input('结束日期')
        self.input_text('结束日期',dateEnd)
        sleep(1)
        try:
            self.click('结束日期-关闭')
        except:
            print('结束日期无关闭按钮')
        self.click('日期-确定')


    def order(self,taskName):
        """ 创建报表任务\n
        
        :taskName: 任务名称,例如: '车险月度数据报表' :<str>\n      
        :return: newBillNo\n
        """
        self.click('加入购物车')
        self.click('加入购物车-确定')
        self.input_text('任务名称',taskName)
        self.option_by_text('任务类型','一次性临时查询')
        self.click('提交订单')
        sleep(2)
        newBillNo=self.text('下单订单号')
        print(newBillNo)
        self.click('提交订单-确定')
        return newBillNo

    def downloadExcel(self, billNo,downPath):
        """ 下载报表\n
        
        :param downPath: 下载路径(文件路径+文件名),例如: r"d:\A\101.xlsx" :<str>\n        
        :return: 无\n
        """
        self.click('我的订单')
        sleep(2)
        self.wait_loaded("订单编号")
        self.input_text('订单编号',billNo)
        sleep(2)
        for i in range(100):
            self.click('查询订单')
            print('执行状态完成控件的数量：',self.count('执行状态-完成'))
            if self.count('执行状态-完成'):
                self.click('下载')
                sleep(2)
                self.down_by_element('我知道了',downPath)
                sleep(5)
                break
            else:
                sleep(6)
        pass

    def waitElementDisappear(self, elementName, arg:str=None, timeout=60):
        """ 等待控件消失\n
        
        :elementName: 控件名: <str>\n        
        :return: 无\n
        """     
        # rpa.console.logger.info("等待" + "\"" + elementName + "\"" + "消失")
        for i in range(timeout):
            if (self.count(elementName, arg=arg) != 0):
                sleep(1)
                continue
            else:
                return
  