import os
import platform
from CLPC.browser_visual import Browser
from time import sleep, time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from CLPC.element import Element
from selenium.webdriver.chrome.webdriver import WebDriver
import datetime

from CLPC.framework import FUNC_USAGE_TRACKER
from CLPC.tool import TOOL


@FUNC_USAGE_TRACKER
class SCPLUS(Browser):
    def __init__(self, driver: WebDriver):
        """
        登录数据分析平台，多维分析操作组件\n
        初始化后需要传入用户信息
        @driver: 已初始化的webdriver
        """
        # super().__init__(driver)
        self.driver = driver

        current_dir = os.path.dirname(os.path.abspath(__file__))
        ele_json_path = os.path.join(current_dir, 'scplus.json')  # 加载组件自己的元素文件
        self.ele_map = Element(ele_json_path)  # 加载组件自己的元素文件

        self.userName = "",
        self.passWord = "",
        self.verfiCode = "",

    def get_scplus_browser(self):
        return self.browser

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
        elif subSystem == '闪查Plus':
            self.catch('首页 - 数据自服务平台', timeout=30)
            self.wait_loaded("清单名称", arg="SF01-非车险理赔清单")
            sleep(1)
        else:
            tabs = self.driver.window_handles
            self.driver.switch_to.window(tabs[-1])  # 切换至最后一个tab

    def enterscPage(self):
        """
        进入闪查plus\n
        :return: 无\n
        """
        self.loginMainSystem()
        self.enterSubSystem('闪查Plus')
        sleep(1)
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[-1])  # 切换至最后一个tab
        for idx in range(10):
            sleep(3)
            try:
                self.wait_loaded("清单名称", arg="SF01-非车险理赔清单",timeout=60)
            except:
                pass
            if self.count("清单名称", arg="SF01-非车险理赔清单") > 0:
                break
            else:
                self.restore()

    def enterTargetInventory(self, targetInventory, anaMode):
        """ 进入目标清单\n

        :param targetInventory: 目标清单,  :<str>\n
        :param anaMode: 分析模式：清单分析模式或多维分析模式,  :<str>\n
        :return: 无\n
        """
        if isinstance(targetInventory, str) is not True:
            raise Exception(str(targetInventory) + "非字符串, 格式错误")
        self.wait_loaded('清单名称', arg=targetInventory, timeout=120)
        self.click('清单名称', arg=targetInventory)
        sleep(1)
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[-1])
        sleep(1)
        self.__wait_mask_disapper()
        for idx in range(10):
            sleep(2)
            try:
                self.wait_loaded('清单分析模式',timeout=120)
            except:
                pass
            if self.count('清单分析模式') > 0:
                break
            self.restore()
        self.__wait_mask_disapper()
        if anaMode == '清单分析模式':
            self.click('清单分析模式')
        elif anaMode == '多维分析模式':
            self.click('多维分析模式')
        else:
            pass
        self.__wait_mask_disapper()

    def inventoryModeSelectAll(self):
        """ 清单分析模式勾选全选\n

        :return: 无\n
        """
        self.click('清单分析模式-全选')

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

        for i in range(3):
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

    def define(self, filter, targetPoistion):
        """ 拖拽筛选维度到指定位置\n
        :param filter: 筛选维度,  :<str>\n
        :param targetPoistion: 目标位置：行、列、数据项,  :<str>\n
        :return: 无\n
        """
        if targetPoistion == '筛选器':
            self.drag_then_drop(source_ele='筛选字段', target_ele='筛选器', source_arg=filter)
        elif targetPoistion == '行':
            sleep(1)
            if self.count('清单分析-数据项') > 0:
                raise Exception('参数定义错误,清单分析模式请选择：数据项')
            self.drag_then_drop(source_ele='筛选字段', source_arg=filter, target_ele='多维分析模式-行')
        elif targetPoistion == '列':
            sleep(1)
            if self.count('清单分析-数据项') > 0:
                raise Exception('参数定义错误,清单分析模式请选择：数据项')
            self.drag_then_drop(source_ele='筛选字段', source_arg=filter, target_ele='多维分析模式-列')
        elif targetPoistion == '数据项':
            sleep(1)
            if self.count('多维分析-切换维度字段') > 0:
                raise Exception('参数定义错误,多维分析模式请选择：行或列')
            self.drag_then_drop(source_ele='筛选字段', source_arg=filter, target_ele='清单分析模式-数据项')
        else:
            raise Exception('参数定义错误')

        pass

    def selectByCode(self, selectTarget):
        """ 根据码值进行筛选\n
        :param selectTarget: 需要筛选的目标值,  :<list>\n
        :return: 无\n
        """
        if isinstance(selectTarget, list) is not True:
            raise Exception(str(selectTarget) + "非数组, 格式错误")
        self.__wait_mask_disapper()
        self.wait_loaded('已选匹配')
        text = self.text('已选匹配')
        if text == '码值筛选':
            pass
        else:
            self.click('未选匹配')
        self.__wait_mask_disapper()

        self.wait_loaded('码值筛选-检索输入')
        for select_target in selectTarget:
            self.input_text('码值筛选-检索输入', select_target)
            sleep(1)
            self.click('码值筛选-输入搜索')
            self.__wait_mask_disapper()
            self.wait_loaded('码值筛选-搜索结果', arg=select_target)

            sleep(1)
            self.click('码值筛选-搜索结果', arg=select_target)
            self.__wait_mask_disapper()
            self.clear_input('码值筛选-检索输入')
            self.__wait_mask_disapper()

        sleep(1)
        self.click('码值筛选-确定')
        self.__wait_mask_disapper()

        pass

    def selectByExistTime(self, existTime):
        """ 日期筛选选择近七天、本月、近一月、本年、近一年\n
        :param existTime: 选择勾选的日期,  :<str>\n
        :return: 无\n
        """
        self.__wait_mask_disapper()
        self.wait_loaded('日期-时间粒度1')
        self.wait_loaded('日期-近七天')
        if existTime == '近七天':
            self.click('日期-近七天')
        elif existTime == '本月':
            self.click('日期-本月')
        elif existTime == '近一月':
            self.click('日期-近一月')
        elif existTime == '本年':
            self.click('日期-本年')
        elif existTime == '近一年':
            self.click('日期-近一年')
        else:
            raise Exception(str(existTime) + "入参错误")
        self.__wait_mask_disapper()
        self.click('日期-确定')
        self.__wait_mask_disapper()

    def selectByTime(self, fileterRange: str, startTime: str, endTime: str):
        """ 根据开始时间 结束时间进行筛选 时间粒度\n
        :param fileterRange: 过滤维度 ：日、月、年,  :<str>\n
        :param startTime: 开始时间,  :<str>\n
        :param endTime: 结束时间,  :<str>\n
        :return: 无\n
        """
        self.__wait_mask_disapper()
        self.wait_loaded('日期-时间粒度1')
        sleep(1)
        self.click('日期-时间粒度1')
        # if fileterRange == '日':
        #     self.wait_loaded('日期-时间粒度2-日')
        #     sleep(1)
        #     self.click('日期-时间粒度2-日')
        # else:
        self.__wait_mask_disapper()
        self.wait_loaded('日期-时间粒度2', arg=fileterRange)
        sleep(1)
        self.__wait_mask_disapper()
        self.click('日期-时间粒度2', arg=fileterRange)
        sleep(1)
        end_time_ele = "日期-结束日期-{}".format(fileterRange)
        # self.clear_input(end_time_ele)
        date_page_end = self.get_attribute(end_time_ele)
        sleep(1)
        start_time_ele = "日期-开始日期-{}".format(fileterRange)
        # self.clear_input(start_time_ele)
        date_page_start = self.get_attribute(start_time_ele)
        sleep(1)
        if date_page_start is None or not date_page_start:
            self.input_text(start_time_ele, startTime)
            ActionChains(self.driver).send_keys(Keys.RETURN)
            sleep(1)
            self.input_text(end_time_ele, endTime)
            ActionChains(self.driver).send_keys(Keys.RETURN)
        else:
            if fileterRange == '日':
                date_page_start = datetime.datetime.strptime(date_page_start, "%Y-%m-%d")
                date_write_start = datetime.datetime.strptime(startTime, "%Y-%m-%d")
                date_page_end = datetime.datetime.strptime(date_page_end, "%Y-%m-%d")
                date_write_end = datetime.datetime.strptime(endTime, "%Y-%m-%d")
                if date_write_start > date_page_end:
                    self.clear_input(end_time_ele)
                    sleep(1)
                    self.input_text(end_time_ele, endTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                    sleep(1)
                    self.clear_input(start_time_ele)
                    sleep(1)
                    self.input_text(start_time_ele, startTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                elif date_write_end < date_page_start:
                    self.clear_input(start_time_ele)
                    sleep(1)
                    self.input_text(start_time_ele, startTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                    sleep(2)
                    self.clear_input(end_time_ele)
                    sleep(1)
                    self.input_text(end_time_ele, endTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                else:
                    self.clear_input(end_time_ele)
                    sleep(1)
                    self.input_text(end_time_ele, endTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                    sleep(1)
                    self.clear_input(start_time_ele)
                    sleep(1)
                    self.input_text(start_time_ele, startTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
            elif fileterRange == '月':
                date_page_start_list = date_page_start.split(' ')
                date_page_end_list = date_page_end.split(' ')
                date_write_start_list = startTime.split(' ')
                date_write_end_list = endTime.split(' ')
                if date_write_start_list[0] > date_page_end_list[0] or (
                        date_write_start_list[0] == date_page_end_list[0] and date_write_start_list[2] >
                        date_page_end_list[2]):
                    self.clear_input(end_time_ele)
                    sleep(1)
                    self.input_text(end_time_ele, endTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                    sleep(1)
                    self.clear_input(start_time_ele)
                    sleep(1)
                    self.input_text(start_time_ele, startTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                elif date_write_end_list[0] < date_page_start_list[0] or (
                        date_write_end_list[0] == date_page_start_list[0] and date_write_end_list[2] >
                        date_page_start_list[2]):
                    self.clear_input(start_time_ele)
                    sleep(1)
                    self.input_text(start_time_ele, startTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                    sleep(1)
                    self.clear_input(end_time_ele)
                    sleep(1)
                    self.input_text(end_time_ele, endTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                else:
                    self.clear_input(end_time_ele)
                    sleep(1)
                    self.input_text(end_time_ele, endTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                    sleep(1)
                    self.clear_input(start_time_ele)
                    sleep(1)
                    self.input_text(start_time_ele, startTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
            else:
                date_page_start_list = date_page_start.split(' ')
                date_page_end_list = date_page_end.split(' ')
                date_write_start_list = startTime.split(' ')
                date_write_end_list = endTime.split(' ')
                if date_write_start_list[0] > date_page_end_list[0]:
                    self.clear_input(end_time_ele)
                    sleep(1)
                    self.input_text(end_time_ele, endTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                    sleep(1)
                    self.clear_input(start_time_ele)
                    sleep(1)
                    self.input_text(start_time_ele, startTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                elif date_write_end_list[0] < date_page_start_list[0]:
                    self.clear_input(start_time_ele)
                    sleep(1)
                    self.input_text(start_time_ele, startTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                    sleep(1)
                    self.clear_input(end_time_ele)
                    sleep(1)
                    self.input_text(end_time_ele, endTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                else:
                    self.clear_input(end_time_ele)
                    sleep(1)
                    self.input_text(end_time_ele, endTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)
                    sleep(1)
                    self.clear_input(start_time_ele)
                    sleep(1)
                    self.input_text(start_time_ele, startTime)
                    ActionChains(self.driver).send_keys(Keys.RETURN)

        self.click('日期-确定')
        self.__wait_mask_disapper()

    def selectByFormula(self, matchWay, matchName, satisfyCondition):
        """ 根据条件公式进行筛选  \n
        :param matchWay: 匹配方式  :<list>\n
        :param matchName: 匹配值,  :<list>\n
        :param satisfyCondition: 满足任意条件或者全部条件,  :<str>\n
        :return: 无\n
        """
        if isinstance(matchWay, list) is not True:
            raise Exception(str(matchWay) + "非列表, 格式错误")
        if isinstance(matchName, list) is not True:
            raise Exception(str(matchName) + "非列表, 格式错误")

        self.__wait_mask_disapper()
        self.click('条件公式')
        self.__wait_mask_disapper()
        for idx in range(len(matchName)):
            match_name = matchName[idx]
            match_way = matchWay[idx]
            self.click('筛选条件-匹配方式1')
            self.__wait_mask_disapper()
            self.wait_loaded('筛选条件-匹配方式2', match_way)
            self.scroll_into_view('筛选条件-匹配方式2', arg=match_way)
            self.click('筛选条件-匹配方式2', arg=match_way)
            self.__wait_mask_disapper()
            if match_way == '空值' or match_way == '不为空':
                pass
            else:
                self.input_text('筛选条件-输入条件', match_name)
            self.click('筛选条件-新增')
            self.__wait_mask_disapper()
        self.click('筛选条件-满足条件1')
        self.__wait_mask_disapper()
        self.wait_loaded('筛选条件-满足条件2', arg=satisfyCondition)
        self.click('筛选条件-满足条件2', arg=satisfyCondition)
        self.__wait_mask_disapper()
        self.click('筛选条件-确定')
        self.__wait_mask_disapper()

        pass

    def selectByMeasure(self, measureWay, measureName):
        """ 根据度量进行筛选  \n
        :param measureWay: 度量名称  :<str>\n
        :param measureName: 度量值,  :<str>\n 若度量名称为包含，则需要输入 最小值 最小值包含 最大值 最大值包含  例如：'1,是,2,否'
        :return: 无\n
        """
        self.__wait_mask_disapper()
        self.wait_loaded('度量筛选方式1')
        self.__wait_mask_disapper()
        self.click('度量筛选方式1')
        self.wait_loaded('度量筛选方式2', arg=measureWay)
        self.__wait_mask_disapper()
        self.click('度量筛选方式2', arg=measureWay)
        if measureWay == '包含':
            measureNameList = measureName.split(',')
            self.input_text('度量输入最小值', measureNameList[0])
            sleep(1)
            if measureNameList[1] == '是':
                self.click('度量最小值包含')
            sleep(1)
            self.input_text('度量输入最小值', measureNameList[2])
            if measureNameList[3] == '是':
                self.click('度量最大值包含')
        else:
            self.input_text('度量输入数值', measureName)
            self.__wait_mask_disapper()
        self.click('度量-确定')
        self.__wait_mask_disapper()
        pass

    def executeDownFile(self,file_path):
        """ 点击执行以后直接下载文件，不管是否大于1000  \n
        :param filePath: 下载路径  :<str>\n
        :return: downOrderNo 订单编号  :<str>\n
        """
        self.click('执行')
        self.__wait_mask_disapper()
        for idx in range(60):
            if self.count('执行完成提示判断') > 0:
                if self.isvisible('执行完成提示判断'):
                    break
            if self.count('执行完成保单判断') > 0:
                if self.isvisible('执行完成保单判断'):
                    break
            sleep(3)
        else:
            print("可能执行未结束，请检查页面")
        sleep(1)
        execute_flag = False
        if self.count('提示继续')>0 and self.count('执行完成提示判断') > 0:

            self.click('提示继续')
            sleep(1)
            for idx in range(60):
                sleep(1)
                if self.count('下载按钮') > 0:
                    break
            execute_flag = True
        else:
            execute_flag = False
        if execute_flag:
            orderNo = self.getDownOrder('getOrderNum')
            self.downFileByOrder(orderNo,file_path)
            return orderNo
        else:
            self.downLT1000File(file_path)
            return ''

    def executeFile(self, executeFlag):
        """ 点击执行  返回执行成功或失败，分为\n
        :param executeFlag: 执行类型，大于1000或者小于1000,大于1000为True，小于1000为False  :<boolean>\n
        :return: 无\n
        """

        if not executeFlag:
            self.click('执行')
            for idx in range(60):
                if self.count('执行完成保单判断') > 0:
                    if self.isvisible('执行完成保单判断'):
                        break
                sleep(3)
            else:
                print("可能执行未结束，请检查页面")
        else:
            self.click('执行')
            for idx in range(60):
                if self.count('执行完成提示判断') > 0:
                    if self.isvisible('执行完成提示判断'):
                        break
                sleep(3)
            else:
                print("可能执行未结束，请检查页面")

            self.click('提示继续')
            sleep(1)
            for idx in range(60):
                sleep(1)
                if self.count('下载按钮') > 0:
                    break

    def executeFileNew(self):
        """ 点击执行  若数据大于1000条，则返回True，小于1000条则返回False
        :param
        :return: True or False\n
        """
        self.click('执行')
        self.__wait_mask_disapper()
        for idx in range(60):
            if self.count('执行完成提示判断') > 0:
                if self.isvisible('执行完成提示判断'):
                    break
            if self.count('执行完成保单判断') > 0:
                if self.isvisible('执行完成保单判断'):
                    break
            sleep(3)
        else:
            print("可能执行未结束，请检查页面")
        sleep(1)
        if self.count('提示继续')>0 and self.count('执行完成提示判断') > 0:

            self.click('提示继续')
            sleep(1)
            for idx in range(60):
                sleep(1)
                if self.count('下载按钮') > 0:
                    break
            return '是'
        else:
            return '否'



    def downLT1000File(self, filePath):
        """ 直接下载数据量为1000以下的数据  \n
        :param filePath: 下载路径  :<str>\n
        :return: 无\n
        """
        self.down_by_element(ele_name='下载按钮', file_name=filePath)
        pass

    def getDownOrder(self, orderName):
        """下载1000以上的数据 返回订单编号  \n
        :param orderName: 下载订单名称  :<str>\n
        :return: downOrder：订单编号\n
        """
        orderNo = ''
        self.click('下载按钮')
        self.wait_loaded('下载提示', timeout=300)
        self.click('保存订单离线下载')
        self.wait_loaded('订单编号')
        orderNo = self.text('订单编号')
        print('订单编号：' + str(orderNo))
        sleep(1)
        self.click('自动执行', simulate=True)
        self.input_text('订单名称输入', orderName)
        sleep(2)
        self.click('制作订单确定')
        self.wait_loaded('提交订单关闭')
        self.__wait_mask_disapper()
        self.click('提交订单关闭')
        self.wait_loaded('下载按钮', timeout=300)

        return orderNo

    def downFileByOrder(self, downOrderNo, filePath):
        """ 根据订单号下载数据  \n
        :param downOrderNo: 订单编号  :<str>\n
        :param filePath: 下载路径  :<str>\n
        :return: 无\n
        """
        down_flag = False
        self.click('我的订单')
        sleep(3)
        for idx in range(5):
            self.restore()
            self.__wait_mask_disapper()
            self.wait_loaded('订单号码输入', timeout=60)
            self.input_text('订单号码输入', downOrderNo)
            self.click('订单查询')
            sleep(5)
            self.__wait_mask_disapper()
            self.wait_loaded('订单编号出现标志', downOrderNo)
            if self.count('已完成标志') > 0:
                self.scroll_into_view('订单下载')
                sleep(1)
                self.down_by_element('订单下载', filePath)
                down_flag = True
                break
            else:
                sleep(10)
                print('订单未执行完成')
        if down_flag:
            print('订单下载完成')
        else:
            return '订单未完成'

    def getMoreOrder(self, downOrderNo):
        """ 再来一单  \n
        :param downOrder: 订单编号  :<str>\n
        :return: 无\n
        """
        self.click('我的订单')
        self.__wait_mask_disapper()
        self.wait_loaded('订单号码输入', timeout=60)
        sleep(1)
        self.input_text('订单号码输入', downOrderNo)
        self.click('订单查询')
        sleep(5)
        self.__wait_mask_disapper()
        self.wait_loaded('订单编号出现标志', downOrderNo)
        self.click('再来一单')
        sleep(1)
        self.__wait_mask_disapper()
        pass

    def __wait_mask_disapper(self):
        self.wait_disappear("加载中", timeout=60)
        self.wait_disappear("加载中2", timeout=60)

    def edit_selector(self, selector_name: str):
        """编辑筛选器
        :param selector_name: 筛选器名称
        """
        self.click("筛选字段设置", arg=selector_name)
        self.__wait_mask_disapper()

    def delete_selector(self, selector_name: str):
        """删除筛选器
        :param selector_name: 筛选器名称
        """
        self.mouse_move("筛选字段标题位置", arg=selector_name)
        sleep(1)
        self.click("筛选字段删除", arg=selector_name)
        sleep(1)

    def calculate_row_col(self,row_or_col,cal_name,cal_way,sort=''):
        """多维分析模式下  对字段进行 计数 去重等操作 以及排序等
        :param row_or_col: 字段所在位置，行或列
        :param cal_name: 字段名称
        :param cal_way: 计算操作
        :param sort: 正序或者倒序
        """
        self.__wait_mask_disapper()
        if row_or_col == '行':
            sleep(1)
            # self.mouse_move('多维行列下拉-行',cal_name)
            self.click('多维行列下拉-行',simulate=True,arg=cal_name)
        else:
            sleep(1)
            # self.mouse_move('多维行列下拉-列', cal_name)
            self.click('多维行列下拉-列',simulate=True, arg=cal_name)
        self.wait_loaded('多维行列下拉-筛选条件',arg=cal_way)
        self.click('多维行列下拉-筛选条件',arg=cal_way)
        if sort:
            if sort == '正序':
                self.click('多维行列下拉-筛选条件-排序','正序')
            else:
                self.click('多维行列下拉-筛选条件-排序', '倒叙')
