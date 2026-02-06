import os
import platform
from CLPC.browser_visual import Browser
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from CLPC.element import Element
from selenium.webdriver.chrome.webdriver import WebDriver

from CLPC.framework import FUNC_USAGE_TRACKER
from CLPC.tool import TOOL

@FUNC_USAGE_TRACKER
class Portal(Browser):
    def __init__(self, driver: WebDriver):
        """
        登录数据分析平台，多维分析操作组件\n
        初始化后需要传入用户信息
        @driver: 已初始化的webdriver
        """
        # super().__init__(driver)
        current_dir = os.path.dirname(os.path.abspath(__file__))
        ele_json_path = os.path.join(current_dir, 'portal.json')

        self.driver = driver
        self.ele_map = Element(ele_json_path) # 加载组件自己的元素文件
        # sys = platform.system()
        # if sys == 'Windows':
        #     self.ele_map = Element('CLPC\\portal.json') # 加载组件自己的元素文件
        # if sys == 'Darwin' or sys == 'Linux':
        #     self.ele_map = Element('CLPC/portal.json') # 加载组件自己的元素文件

        self.__cognosMenuList = ["一级控件", "二级控件", "三级控件", "四级控件", "五级控件","六级控件"]
        self.__cognosMenuList_new_tab = ["新页面-一级控件", "新页面-二级控件", "新页面-三级控件", "新页面-四级控件", "新页面-五级控件","新页面-六级控件"]
        self.__cognosMenuList_2 = ["一级控件-重名", "二级控件-重名", "三级控件-重名", "四级控件-重名", "五级控件-重名","六级控件-重名"]
        self.__defineZoneDic = {"row": "行", "mulRow": "嵌套的行", "col": "列", "mulCol": "嵌套的列", "context": "在上下文中过滤", "repRow": "替换_行", "repCol": "替换_列", "measure": "使用默认度量"}
        self.__defineZoneDic_new_tab = {"row": "新页面-行", "mulRow": "新页面-嵌套的行", "col": "新页面-列", "mulCol": "新页面-嵌套的列", "context": "新页面-在上下文中过滤", "repRow": "新页面-替换_行", "repCol": "新页面-替换_列", "measure": "新页面-使用默认度量"}
        self.__deleteZoneDic = {"row": "行下三角", "mulRow": "行下三角-嵌套", "col": "列下三角", "mulCol": "列下三角-嵌套", "context": "上下文过滤器下三角", "mulContext": "上下文过滤器下三角-嵌套"}
        self.__deleteZoneDic_new_tab = {"row": "新页面-行下三角", "mulRow": "新页面-行下三角-嵌套", "col": "新页面-列下三角", "mulCol": "新页面-列下三角-嵌套", "context": "新页面-上下文过滤器下三角", "mulContext": "新页面-上下文过滤器下三角-嵌套"}
        self.userName = "",
        self.passWord = "",
        self.verfiCode = "",

        BASMenuListDict = {3: ["BAS", "分类", "一级页面"], 4: ["BAS", "分类", "一级菜单","二级页面"], 5: ["BAS", "分类", "一级菜单", "二级菜单", "三级页面"], 6: ["BAS", "分类", "一级菜单", "二级菜单", "三级菜单", "四级页面"]}
        FASMenuListDict = {3: ["FAS", "FAS_分类", "FAS_页面"], 4: ["FAS", "FAS_分类", "FAS_菜单展开","FAS_页面"], 5: ["FAS", "FAS_分类", "FAS_菜单展开", "FAS_菜单展开", "FAS_页面"], 6: ["FAS", "FAS_分类", "FAS_菜单展开", "FAS_菜单展开", "FAS_菜单展开", "FAS_页面"]}
        reserveListDict = {3: ["reserve", "精算_一级菜单", "精算_二级页面"]}
        self.__menuDict = {"BAS": BASMenuListDict, "FAS": FASMenuListDict, "精算": reserveListDict}

    def __del__(self):
        pass

    def enterPage(self, name):
        """ 跳转进入指定页面\n
        
        @ name: 标题(例如："FAS>经营分析>总体经营情况>收入总体分析", "BAS>业管>车险>车险承保>多维分析>业务结构多维分析")或者"url" :<str>\n
        @ mode: "title"|"url" :<str>\n
        @ start: 从第几个位置开始点击, 例如: name = "FAS-经营分析-总体经营情况-收入总体分析", start = 2,表示fas页面已经打开,从当前页面的经营分析开始点击, 
            如果start = 3则表示经营分析已经点开, 从总体经营情况开始点击:<str>\n
        """
        def clickMethod(i):
            if subSys == "FAS":
                if (i != length - 1) & (i != 1): 
                    self.click(menuList[i], arg=subTitleList[i], index = 0)
                elif i == 1:
                    for times in range(6):
                        self.click(menuList[i], arg=subTitleList[i])
                        if self.wait_loaded('fas_flag'):
                            break
                else:
                    self.click(menuList[i], arg=subTitleList[i])

            elif (subSys == "精算") & (i == length - 1):
                    page_href = self.attr(menuList[i], arg=subTitleList[i], attrname = "href")
                    if page_href.startswitch("/reserve"):
                        page_href = "http://9.1.64.187:7001" + page_href
                    self.create(page_href)

            else: # bas 或其他系统
                self.click(menuList[i], arg=subTitleList[i])
            sleep(2)
            print(subTitleList[i] + " 已点击")
            tabs = self.driver.window_handles
            self.driver.switch_to.window(tabs[-1])

        subTitleList = name.split(">")
        subSys = subTitleList[0]
        menuListDict = self.__menuDict[subSys]
        length = len(subTitleList)
        try:
            menuList = menuListDict[length]
        except Exception as e:
            raise Exception("title 格式不对") 

        for i in range(0, length):
            if i == 0:
                if subSys == "精算":
                    self.loginReserveSystem()
                    continue
                self.loginMainSystem()
                self.enterSubSystem(subSys)

            else:
                try:
                    clickMethod(i)
                    sleep(2)
                except Exception as e:
                    clickMethod(i)

        sleep(3)
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[-1])
        sleep(5)
        if 'IBM Cognos Analysis Studio' in self.get_title():
            self.clear()
            self.wait_loaded("列表格", timeout=20)



    def loginMainSystem(self):
        """ 进入数据分析平台\n
        用户名 密码 验证码 3个都在__init__ 初始化时可做替换
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
            # self.maximize()
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
                print(f"第{i+1}次登录失败")
                continue
            print("登录成功")
            break
        else:
            raise Exception("3次登录均失败, 请检查账号密码验证码是否正确")


    def enterSubSystem(self, subSystem):
        """ 进入子系统\n
        
        :param subSystem: 子系统, "BAS"|"FAS" :<str>\n
        :return: 无\n
        """        
        print('当前tab title:', self.get_title())     
        self.wait_loaded(subSystem)
        if self.count("今天隐藏") != 0:
            self.click("今天隐藏")
        if self.count("我已知晓") != 0:
            self.click("我已知晓")
        self.click(subSystem)
        sleep(3)
        
        if subSystem == "BAS":
            self.catch('BAS', mode='contain', timeout=30)
            sleep(2)
        elif subSystem == 'FAS':
            self.catch('财务分析系统', mode='contain', timeout=30)
            # self.wait_loaded("净利润图片")
            self.wait_loaded("FAS_分类", arg="多维分析", timeout=30)
            sleep(10)
        else:
            tabs = self.driver.window_handles
            self.driver.switch_to.window(tabs[-1]) # 切换至最后一个tab

    def unpackEle(self, operationObjects, level):
        """ 展开插入对象列表里的控件\n
        
        :param operationObjects: 要展开的控件名列表(例如: ["机构","日期"],["2020"]) :<list>\n
        :param level: 控件层次,1|2|3|4,机构为1级控件,机构下方的北京分公司为2级控件,以此类推 :<int>\n
        :return: 无\n
        """ 
        if isinstance(operationObjects,list) is not True:
            raise Exception(str(operationObjects) + "非列表, 格式错误")

        self.wait_loaded("可插入对象")

        menu = self.__cognosMenuList[level-1]
        for operationObject in operationObjects:
            for i in range(10):
                if self.count(menu, arg=operationObject) == 0:
                    sleep(1)
                else:
                    break 
            else:
                raise Exception(str(operationObject) + "不存在或层次定义错误")        
            '''有时候会点击失效，一次不行，再来一次'''
            for j in range(10):
                try:
                    self.doubleclick(menu,arg=operationObject)
                    sleep(2)
                    self.wait_disappear("正在载入_1级")
                    break
                except Exception as e:
                    print(e)
        
        # 展开更多
        while True:
            nextLevelMenu = self.__cognosMenuList[level]
            if self.count(nextLevelMenu, arg='更多') == 0:
                sleep(1)
                break
            self.click(nextLevelMenu, arg='更多')
            sleep(2)
            if self.count(nextLevelMenu, arg='更多') == 0:
                sleep(1)
                break

    def unpackEle_with_duplicate_name(self, operationObjects):
        """
        展开插入对象列表里的控件,当不同的菜单中有重名的控件时使用，空间名需要附带上级菜单名\n
        :param operationObjects: 要展开的控件名列表(例如: ["机构","日期>2024"],["承保日期>2020"]) :<list>\n
        :return: 无\n
        """
        if isinstance(operationObjects, list) is not True:
            raise Exception(str(operationObjects) + "非列表, 格式错误")

        self.wait_loaded("可插入对象")

        menu = "菜单控件-重名"
        for operationObject in operationObjects:
            args = operationObject.split(">")
            for i in range(10):
                if self.count(menu, arg=args) == 0:
                    sleep(1)
                else:
                    break
            else:
                raise Exception(str(operationObject) + "不存在或层次定义错误")
            for j in range(10):
                try:
                    self.doubleclick(menu, arg=args)
                    sleep(1)
                    self.wait_disappear("正在载入_1级")
                    break
                except Exception as e:
                    print(e)

        # 展开更多
        while True:
            nextLevelMenu = "一级控件"
            if self.count(nextLevelMenu, arg='更多') == 0:
                sleep(1)
                break
            self.click(nextLevelMenu, arg='更多')
            sleep(2)
            if self.count(nextLevelMenu, arg='更多') == 0:
                sleep(1)
                break

    def define(self, operationObjects, level, clickStyle:str, definePart):
        """ 插入控件到行|列|上下文|度量\n
        
        :param operationObjects: 要插入的控件名列表(例如: ["机构","日期"],["2020"]) :<list>\n
        :param level: 控件层次,1|2|3|4,机构为1级控件,机构下方的北京分公司为2级控件,以此类推 :<int>\n
        :param clickStyle: 点击方式,"ctrl"|"shift" :<str>\n
        :param definePart: 要插入的区域,"row": "行", "mulRow": "嵌套的行", "col": "列", "mulCol": "嵌套的列", "context": "在上下文中过滤", "repRow": "替换_行", "repCol": "替换_列", "measure": "使用默认度量" :<str>\n        
        :return: 无\n
        """
        clickStyle = clickStyle.lower()
        if isinstance(operationObjects,list) is not True:
            raise Exception(str(operationObjects) + "非列表, 格式错误")         
        menu = self.__cognosMenuList[level-1]
        
        # 控制变量num,用于点击第一次CTRL,
        num = 0
        for operationObject in operationObjects:
            for i in range(10):
                if self.count(menu, arg=operationObject) == 0:
                    sleep(0.3)
                else:
                    break
            if self.count(menu, arg=operationObject) == 0:
                raise Exception(str(operationObject) + "不存在或层次定义错误")                          
            self.click(menu, arg=operationObject)  ##
            num += 1
            if (num == 1):
                if clickStyle == 'ctrl':
                    ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                elif clickStyle == 'shift':
                    ActionChains(self.driver).key_down(Keys.SHIFT).perform()

            # 当操作对象是这种[2018-01, 2018-01]时，点击之后就结束循环
            if (num == len(operationObjects)) | ((len(operationObjects) > 1) & (operationObjects[0] == operationObjects[len(operationObjects)-1])):
                if clickStyle == 'ctrl': 
                    ActionChains(self.driver).key_up(Keys.CONTROL).perform()
                elif clickStyle == 'shift':
                    ActionChains(self.driver).key_up(Keys.SHIFT).perform()
                self.click(menu, arg=operationObject, type='right')
                
        if definePart == "context":
            self.wait_loaded("在上下文中过滤")             
            self.click('在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("使用默认度量")             
            self.click('使用默认度量')
        elif definePart.startswith("rep"):
            self.click('替换')
            self.click(self.__defineZoneDic[definePart])                        
        else:
            self.click('插入')
            self.click(self.__defineZoneDic[definePart])
        self.waitStatusLoaded()

    def define_with_duplicate_name(self, operationObjects, clickStyle: str, definePart):
        """ 插入控件到行|列|上下文|度量\n

        :param operationObjects: 要插入的控件名列表(例如: ["机构","日期>2024"],["承保日期>2020"]) :<list>\n
        :param clickStyle: 点击方式,"ctrl"|"shift" :<str>\n
        :param definePart: 要插入的区域,"row": "行", "mulRow": "嵌套的行", "col": "列", "mulCol": "嵌套的列", "context": "在上下文中过滤", "repRow": "替换_行", "repCol": "替换_列", "measure": "使用默认度量" :<str>\n
        :return: 无\n
        """
        clickStyle = clickStyle.lower()
        if isinstance(operationObjects, list) is not True:
            raise Exception(str(operationObjects) + "非列表, 格式错误")
        # menu = self.__cognosMenuList_2[level - 1]
        menu = "菜单控件-重名"

        # 控制变量num,用于点击第一次CTRL,
        num = 0
        for operationObject in operationObjects:
            args = operationObject.split(">")
            for i in range(10):
                if self.count(menu, arg=args) == 0:
                    sleep(0.3)
                else:
                    break
            if self.count(menu, arg=args) == 0:
                raise Exception(str(operationObject) + "不存在或层次定义错误")
            self.click(menu, arg=args)  ##
            num += 1
            if (num == 1):
                if clickStyle == 'ctrl':
                    ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                elif clickStyle == 'shift':
                    ActionChains(self.driver).key_down(Keys.SHIFT).perform()

            # 当操作对象是这种[2018-01, 2018-01]时，点击之后就结束循环
            if (num == len(operationObjects)) | (
                    (len(operationObjects) > 1) & (operationObjects[0] == operationObjects[len(operationObjects) - 1])):
                if clickStyle == 'ctrl':
                    ActionChains(self.driver).key_up(Keys.CONTROL).perform()
                elif clickStyle == 'shift':
                    ActionChains(self.driver).key_up(Keys.SHIFT).perform()
                self.click(menu, arg=args, type='right')

        if definePart == "context":
            self.wait_loaded("在上下文中过滤")
            self.click('在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("使用默认度量")
            self.click('使用默认度量')
        elif definePart.startswith("rep"):
            self.click('替换')
            self.click(self.__defineZoneDic[definePart])
        else:
            self.click('插入')
            self.click(self.__defineZoneDic[definePart])
        self.waitStatusLoaded()

    def define_multi(self, operationObjectsList, levelList, definePart):
        """ 插入控件到行|列|上下文|度量, 可以选取不同层次的控件, 点击方式固定为CTRL\n

        :param operationObjectsList: 要插入的多个控件名列表(例如: [["机构","日期"],["2020"]]) :<list>双重列表\n
        :param levelList: 控件层次列表(例如: [1, 2]) :<list>\n
        :param definePart: 要插入的区域,"row": "行", "mulRow": "嵌套的行", "col": "列", "mulCol": "嵌套的列", "context": "在上下文中过滤", "repRow": "替换_行", "repCol": "替换_列", "measure": "使用默认度量" :<str>\n
        :return: 无\n
        """
        if isinstance(operationObjectsList, list) is not True:
            raise Exception(str(operationObjectsList) + "非列表, 格式错误")
        for i in range(len(operationObjectsList)):
            menu = self.__cognosMenuList[levelList[i] - 1]

            # 控制变量num,用于点击第一次CTRL,
            num = 0
            for operationObject in operationObjectsList[i]:
                for j in range(10):
                    if self.count(menu, arg=operationObject) == 0:
                        sleep(0.3)
                    else:
                        break
                if self.count(menu, arg=operationObject) == 0:
                    raise Exception(str(operationObject) + "不存在或层次定义错误")
                    # 有时候会点击失效，一次不行，再来一次
                for j in range(10):
                    try:
                        self.click(menu, arg=operationObject)
                        break
                    except Exception as e:
                        sleep(1)
                num += 1
                if num == 1:
                    ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                    # 当操作对象是这种[2018-01, 2018-01]时，点击之后就结束循环
                if (i == len(operationObjectsList) - 1) & (num == len(operationObjectsList[i])):
                    ActionChains(self.driver).key_up(Keys.CONTROL).perform()

                    # 有时候会点击失效，一次不行，再来一次
                    for j in range(10):
                        try:
                            self.click(menu, arg=operationObject, type="right")
                            break
                        except Exception as e:
                            sleep(1)

        if definePart == "context":
            self.wait_loaded("在上下文中过滤")
            self.click('在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("使用默认度量")
            self.click('使用默认度量')
        elif definePart.startswith("rep"):
            self.click('替换')
            self.click(self.__defineZoneDic[definePart])
        else:
            self.click('插入')
            self.click(self.__defineZoneDic[definePart])
        self.waitStatusLoaded()

    def define_cross(self, operationObjectsList, levelList, clickStyle:str, definePart):
        """跨级插入菜单，支持相同的父菜单下不同级子菜单的插入
        operationObjectsList: 要插入的多个控件名列表(例如: [["机构","日期"],["2020"]]) :<list>双重列表
        """
        if isinstance(operationObjectsList, list) is not True:
            raise Exception(str(operationObjectsList) + "非列表, 格式错误")

        for i in range(len(operationObjectsList)):
            menu = self.__cognosMenuList[levelList[i] - 1]

            # 控制变量num,用于点击第一次CTRL,

            for  index,operationObject in enumerate(operationObjectsList[i]):
                for j in range(10):
                    if self.count(menu, arg=operationObject) == 0:
                        sleep(0.3)
                    else:
                        break
                if self.count(menu, arg=operationObject) == 0:
                    raise Exception(str(operationObject) + "不存在或层次定义错误")
                    # 有时候会点击失效，一次不行，再来一次
                for j in range(10):
                    try:
                        self.click(menu, arg=operationObject)
                        if index ==0:
                            ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                            sleep(1)
                            if clickStyle == "shift":
                                ActionChains(self.driver).key_down(Keys.SHIFT).perform()
                                sleep(1)
                        break
                    except Exception as e:
                        sleep(1)
                    # 当操作对象是这种[2018-01, 2018-01]时，点击之后就结束循环
            if clickStyle == "shift":
                ActionChains(self.driver).key_up(Keys.SHIFT).perform()

        ActionChains(self.driver).key_up(Keys.CONTROL).perform()

        menu_last = self.__cognosMenuList[levelList[-1] - 1]
        object_last = operationObjectsList[-1][-1]
        self.click(menu_last, arg=object_last, type="right")



        if definePart == "context":
            self.wait_loaded("在上下文中过滤")
            self.click('在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("使用默认度量")
            self.click('使用默认度量')
        elif definePart.startswith("rep"):
            self.click('替换')
            self.click(self.__defineZoneDic[definePart])
        else:
            self.click('插入')
            self.click(self.__defineZoneDic[definePart])
        self.waitStatusLoaded()

    def define_cross_with_duplicate_name(self, operationObjectsList, clickStyle:str, definePart):
        """跨级插入菜单，支持相同的父菜单下不同级子菜单的插入,支持重名菜单
        operationObjectsList: 要插入的多个控件名列表(例如: [["机构","日期"],["2020"]]) :<list>双重列表
        """
        if isinstance(operationObjectsList, list) is not True:
            raise Exception(str(operationObjectsList) + "非列表, 格式错误")

        menu = "菜单控件-重名"

        for operationObjects in (operationObjectsList):


            # 控制变量num,用于点击第一次CTRL,

            for index, operationObject in enumerate(operationObjects):
                args = operationObject.split(">")

                for j in range(10):
                    if self.count(menu, arg=args) == 0:
                        sleep(0.3)
                    else:
                        break
                if self.count(menu, arg=args) == 0:
                    raise Exception(str(operationObject) + "不存在或层次定义错误")
                    # 有时候会点击失效，一次不行，再来一次
                for j in range(10):
                    try:
                        self.click(menu, arg=args)
                        if index ==0:
                            ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                            sleep(1)
                            if clickStyle == "shift":
                                ActionChains(self.driver).key_down(Keys.SHIFT).perform()
                                sleep(1)
                        break
                    except Exception as e:
                        sleep(1)
            if clickStyle == "shift":
                ActionChains(self.driver).key_up(Keys.SHIFT).perform()

        ActionChains(self.driver).key_up(Keys.CONTROL).perform()

        args_last = operationObjectsList[-1][-1].split(">")
        self.click(menu, arg=args_last, type="right")

        if definePart == "context":
            self.wait_loaded("在上下文中过滤")
            self.click('在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("使用默认度量")
            self.click('使用默认度量')
        elif definePart.startswith("rep"):
            self.click('替换')
            self.click(self.__defineZoneDic[definePart])
        else:
            self.click('插入')
            self.click(self.__defineZoneDic[definePart])
        self.waitStatusLoaded()


    def delete(self, deletePart):
        """ 删除行|列|上下文|度量的控件\n
        
        :param deletePart: 要删除的区域,"row": "行", "mulRow": "嵌套的行", "col": "列", "mulCol": "嵌套的列", "context": "上下文过滤器", "mulContext": "嵌套的上下文过滤器" :<str>\n        
        :return: 无\n
        """        
        delObj = self.__deleteZoneDic[deletePart]
        # self.page.wait_loaded(delObj)
        self.click(delObj)
        # self.page.wait_loaded("删除")
        self.click("删除")
        self.waitStatusLoaded()

    def clear(self):
        """
        清空页面
        """
        for i in range(3):
            try:
                sleep(2)
                # self.page.wait_loaded('文件')
                sleep(1)
                self.click("文件")
                # self.page.wait_loaded("新建" ,timeout=5)
                self.click("新建" )
                # self.page.wait_loaded("交叉空白表")
                self.click("交叉空白表")
                if self.count('空白交叉表-保存-否')>0:
                    self.click("空白交叉表-保存-否")
                break
            except:
                pass
        self.waitStatusLoaded()
    
    def showAllRows(self):
        """ 
        显示所有行\n
        """            
        # 点击更多
        sleep(3)
        self.click('更多-行', simulate=True)
        # 点击自定义
        self.click('自定义')
        # 点击可视项目下三角
        self.click('可视项目下三角')
        # 点击9999
        self.click('9999')
        # 点击确定
        self.click("显示所有行-确定")
        self.waitStatusLoaded()
        sleep(2)

    def show_all_nest_rows(self):
        """
        展开嵌套行，一次只展开一个嵌套的行，多行有嵌套，分多次展开

        """
        sleep(3)
        self.click('更多-嵌套行', simulate=True)
        self.click('自定义')
        self.click('可视项目下三角')
        self.click('9999')
        self.click("显示所有行-确定")
        self.waitStatusLoaded()
        sleep(2)

    def showAllCols(self):
        """
        显示所有列\n
        """
        # 点击更多
        sleep(3)
        if self.count("横向滚动条") > 0:
            self.drag_then_drop("横向滚动条", "横向滚动条-最右")
        sleep(1)
        self.click('更多-列', simulate=True)
        # 点击自定义
        self.click('自定义')
        # 点击可视项目下三角
        self.click('可视项目下三角')
        # 点击9999
        self.click('9999')
        # 点击确定
        self.click("显示所有行-确定")
        self.waitStatusLoaded()
        sleep(2)
    
    def initDataPresentationNums(self):
        """ 设定可视项目数为9999\n
        
        :param: 无\n        
        :return: 无\n
        """ 
        # 等待下方页面元素加载完成，否则会影响点击 "设置"

        self.wait_loaded("可插入对象")                  
        # 点击设置
        self.wait_loaded("设置")
        # 有时候会点击失效，一次不行，再来一次
        for j in range(100):
            try:
                self.click('设置')
                break
            except Exception as e:
                sleep(1)           
        # 点击设置显示数量菜单
        self.click('设定可视项目数')
        # 输入数量
        self.input_text("可视项目数输入框","9999")
        # 点击确定
        self.click("确定")
        self.waitStatusLoaded()

    def reduceWindow(self, num):
        for i in range(num):
            ActionChains(self.driver).key_down(Keys.COMMAND).send_keys('-').key_up(Keys.COMMAND).perform()
            sleep(0.5)


    def resetWindow(self):
        ActionChains(self.driver).key_down(Keys.COMMAND).send_keys('+').send_keys('+').send_keys('+').key_up(Keys.COMMAND).perform()
        # try:
        #     self.driver.switch_to.default_content()
        # except:
        #     pass
        # tmp = self.driver.find_element_by_xpath('//frame')
        # # 切换到iframe
        # self.driver.switch_to.frame(tmp)
        # self.zoom_reset()
        # self.driver.switch_to.default_content()

    def downloadExcel(self, downPath):
        """ 下载报表\n
        
        :param downPath: 下载路径(文件路径+文件名),例如: r"d:\A\101.xlsx" :<str>\n        
        :return: 无\n
        """
        currnet_url = self.get_url()
        self.click('运行')
        self.down_by_element('以 Excel 2007格式', file_name=downPath)
        self.catch('IBM Cognos Viewer', mode='contain')
        print(self.get_title())
        print(self.get_url())
        self.close()
        self.catch(currnet_url, type='url')
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
        
    def waitStatusLoaded(self):
        """ 等待页面状态加载\n
        
        :param: 无\n        
        :return: 无\n
        """             
        while(True):
            try:
                statusStyle1 = self.attr("沙漏1", "style")
                display1 = statusStyle1.split(";")[0].split(":")[1]
                statusStyle2 = self.attr("沙漏2", "style")
                display2 = statusStyle2.split(";")[0].split(":")[1]
                if (display1 != " block") and (display2 != " block"):
                    return
            except:
                return

    def waitStatusLoaded_in_new_tab(self):
        """ 等待页面状态加载\n

        :param: 无\n
        :return: 无\n
        """
        while (True):
            try:
                statusStyle1 = self.attr("新页面-沙漏1", "style")
                display1 = statusStyle1.split(";")[0].split(":")[1]
                statusStyle2 = self.attr("新页面-沙漏2", "style")
                display2 = statusStyle2.split(";")[0].split(":")[1]
                if (display1 != " block") and (display2 != " block"):
                    return
            except:
                return

    def open_folder(self, patn_route):
        """
        多维分析中，打开文件夹
        :param patn_name: 我的文件夹中的模板名称
        :return:
        """
        self.waitStatusLoaded()

        # self.wait_loaded("打开")
        # self.click("打开")
        self.click("文件")
        self.click("文件-打开")
        self.click("我的文件夹")
        patn_list = patn_route.split(">")
        patn_name = patn_list[-1]
        self.waitStatusLoaded()
        # sleep(2)
        for p in patn_list:
            # self.wait_loaded("文件夹列表", arg=p)
            self.doubleclick("文件夹列表", arg=p)
            sleep(3)

        self.wait_loaded("打开文件夹后的弹窗确定")
        self.click("打开文件夹后的弹窗确定")
        sleep(1)
        self.catch(patn_name, mode='contain', timeout=30)

    def use_search(self, search_name, search_data):
        """
        多维分析中，使用搜索功能
        :param search_name: 搜索的名称
        :param search_data: 搜索的关键字
        """
        self.click("搜索数据", arg=search_name)
        self.waitStatusLoaded()
        self.wait_loaded("搜索输入框")
        self.input_text("搜索输入框", search_data)
        self.click("搜索按钮")
        self.waitStatusLoaded()
        sleep(3)

    def define_search_data(self, operationObjects, clickStyle:str, definePart):
        """
        多维分析中，使用搜索功能，插入控件到行|列|上下文|度量
        :param operationObjects: 搜索的关键字
        :param clickStyle: 点击方式,"ctrl"|"shift" :<str>\n
        :param definePart: 要插入的区域,"row": "行", "mulRow": "嵌套的行", "col": "列", "mulCol": "嵌套的列", "context": "在上下文中过滤", "repRow": "替换_行", "repCol": "替换_列", "measure": "使用默认度量" :<str>\n
        """
        clickStyle = clickStyle.lower()
        if isinstance(operationObjects, list) is not True:
            raise Exception(str(operationObjects) + "非列表, 格式错误")

        # 控制变量num,用于点击第一次CTRL,
        num = 0
        for operationObject in operationObjects:
            for i in range(10):
                if self.count('搜索结果', arg=operationObject) == 0:
                    sleep(0.3)
                else:
                    break
            if self.count('搜索结果', arg=operationObject) == 0:
                raise Exception(str(operationObject) + "不存在或层次定义错误")
            self.click('搜索结果', arg=operationObject)  ##
            num += 1
            if (num == 1):
                if clickStyle == 'ctrl':
                    ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                elif clickStyle == 'shift':
                    ActionChains(self.driver).key_down(Keys.SHIFT).perform()

            # 当操作对象是这种[2018-01, 2018-01]时，点击之后就结束循环
            if (num == len(operationObjects)) | (
                    (len(operationObjects) > 1) & (operationObjects[0] == operationObjects[len(operationObjects) - 1])):
                if clickStyle == 'ctrl':
                    ActionChains(self.driver).key_up(Keys.CONTROL).perform()
                elif clickStyle == 'shift':
                    ActionChains(self.driver).key_up(Keys.SHIFT).perform()
                self.click('搜索结果', arg=operationObject, type='right')

        if definePart == "context":
            self.wait_loaded("在上下文中过滤")
            self.click('在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("使用默认度量")
            self.click('使用默认度量')
        elif definePart.startswith("rep"):
            self.click('替换')
            self.click(self.__defineZoneDic[definePart])
        else:
            self.click('插入')
            self.click(self.__defineZoneDic[definePart])
        self.waitStatusLoaded()

    def unpackEle_in_new_tab(self, operationObjects, level):
        """
        在新tab页中展开插入对象列表里的控件
        :param operationObjects: 要展开的控件名列表(例如: ["机构","日期"],["2020"]) :<list>\n
        :param level: 控件层次,1|2|3|4,机构为1级控件,机构下方的北京分公司为2级控件,以此类推 :<int>\n
        :return: 无\n
        """
        if isinstance(operationObjects, list) is not True:
            raise Exception(str(operationObjects) + "非列表, 格式错误")

        self.wait_loaded("新页面-可插入对象")

        menu = self.__cognosMenuList_new_tab[level - 1]
        for operationObject in operationObjects:
            for i in range(10):
                if self.count(menu, arg=operationObject) == 0:
                    sleep(1)
                else:
                    break
            else:
                raise Exception(str(operationObject) + "不存在或层次定义错误")
            '''有时候会点击失效，一次不行，再来一次'''
            for j in range(10):
                try:
                    self.doubleclick(menu, arg=operationObject)
                    sleep(2)
                    self.wait_disappear("新页面-正在载入_1级")
                    break
                except Exception as e:
                    print(e)

        # 展开更多
        while True:
            nextLevelMenu = self.__cognosMenuList_new_tab[level]
            if self.count(nextLevelMenu, arg='更多') == 0:
                sleep(1)
                break
            self.click(nextLevelMenu, arg='更多')
            sleep(2)
            if self.count(nextLevelMenu, arg='更多') == 0:
                sleep(1)
                break

    def unpackEle_with_duplicate_name_in_new_tab(self, operationObjects, level):
        """
        展开插入对象列表里的控件,当不同的菜单中有重名的控件时使用，空间名需要附带上级菜单名\n
        :param operationObjects: 要展开的控件名列表(例如: ["机构","日期>2024"],["承保日期>2020"]) :<list>\n
        :return: 无\n
        """
        if isinstance(operationObjects, list) is not True:
            raise Exception(str(operationObjects) + "非列表, 格式错误")

        self.wait_loaded("新页面-可插入对象")

        menu = "新页面-菜单控件-重名"
        for operationObject in operationObjects:
            args = operationObject.split(">")
            for i in range(10):
                if self.count(menu, arg=args) == 0:
                    sleep(1)
                else:
                    break
            else:
                raise Exception(str(operationObject) + "不存在或层次定义错误")
            for j in range(10):
                try:
                    self.doubleclick(menu, arg=args)
                    sleep(1)
                    self.wait_disappear("新页面-正在载入_1级")
                    break
                except Exception as e:
                    print(e)

        # 展开更多
        while True:
            nextLevelMenu = self.__cognosMenuList_new_tab[level]
            if self.count(nextLevelMenu, arg='更多') == 0:
                sleep(1)
                break
            self.click(nextLevelMenu, arg='更多')
            sleep(2)
            if self.count(nextLevelMenu, arg='更多') == 0:
                sleep(1)
                break

    def define_in_new_tab(self, operationObjects, level, clickStyle: str, definePart):
        """ 插入控件到行|列|上下文|度量\n

        :param operationObjects: 要插入的控件名列表(例如: ["机构","日期"],["2020"]) :<list>\n
        :param level: 控件层次,1|2|3|4,机构为1级控件,机构下方的北京分公司为2级控件,以此类推 :<int>\n
        :param clickStyle: 点击方式,"ctrl"|"shift" :<str>\n
        :param definePart: 要插入的区域,"row": "行", "mulRow": "嵌套的行", "col": "列", "mulCol": "嵌套的列", "context": "在上下文中过滤", "repRow": "替换_行", "repCol": "替换_列", "measure": "使用默认度量" :<str>\n
        :return: 无\n
        """
        clickStyle = clickStyle.lower()
        if isinstance(operationObjects, list) is not True:
            raise Exception(str(operationObjects) + "非列表, 格式错误")
        menu = self.__cognosMenuList_new_tab[level - 1]

        # 控制变量num,用于点击第一次CTRL,
        num = 0
        for operationObject in operationObjects:
            for i in range(10):
                if self.count(menu, arg=operationObject) == 0:
                    sleep(0.3)
                else:
                    break
            if self.count(menu, arg=operationObject) == 0:
                raise Exception(str(operationObject) + "不存在或层次定义错误")
            self.click(menu, arg=operationObject)  ##
            num += 1
            if (num == 1):
                if clickStyle == 'ctrl':
                    ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                elif clickStyle == 'shift':
                    ActionChains(self.driver).key_down(Keys.SHIFT).perform()

            # 当操作对象是这种[2018-01, 2018-01]时，点击之后就结束循环
            if (num == len(operationObjects)) | (
                    (len(operationObjects) > 1) & (operationObjects[0] == operationObjects[len(operationObjects) - 1])):
                if clickStyle == 'ctrl':
                    ActionChains(self.driver).key_up(Keys.CONTROL).perform()
                elif clickStyle == 'shift':
                    ActionChains(self.driver).key_up(Keys.SHIFT).perform()
                self.click(menu, arg=operationObject, type='right')

        if definePart == "context":
            self.wait_loaded("新页面-在上下文中过滤")
            self.click('新页面-在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("新页面-使用默认度量")
            self.click('新页面-使用默认度量')
        elif definePart.startswith("rep"):
            self.click('新页面-替换')
            self.click(self.__defineZoneDic_new_tab[definePart])
        else:
            self.click('新页面-插入')
            self.click(self.__defineZoneDic_new_tab[definePart])
        self.waitStatusLoaded_in_new_tab()

    def define_with_duplicate_name_in_new_tab(self, operationObjects, clickStyle: str, definePart):
        """ 插入控件到行|列|上下文|度量\n

        :param operationObjects: 要插入的控件名列表(例如: ["机构","日期>2024"],["承保日期>2020"]) :<list>\n
        :param clickStyle: 点击方式,"ctrl"|"shift" :<str>\n
        :param definePart: 要插入的区域,"row": "行", "mulRow": "嵌套的行", "col": "列", "mulCol": "嵌套的列", "context": "在上下文中过滤", "repRow": "替换_行", "repCol": "替换_列", "measure": "使用默认度量" :<str>\n
        :return: 无\n
        """
        clickStyle = clickStyle.lower()
        if isinstance(operationObjects, list) is not True:
            raise Exception(str(operationObjects) + "非列表, 格式错误")
        # menu = self.__cognosMenuList_2[level - 1]
        menu = "新页面-菜单控件-重名"

        # 控制变量num,用于点击第一次CTRL,
        num = 0
        for operationObject in operationObjects:
            args = operationObject.split(">")
            for i in range(10):
                if self.count(menu, arg=args) == 0:
                    sleep(0.3)
                else:
                    break
            if self.count(menu, arg=args) == 0:
                raise Exception(str(operationObject) + "不存在或层次定义错误")
            self.click(menu, arg=args)  ##
            num += 1
            if (num == 1):
                if clickStyle == 'ctrl':
                    ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                elif clickStyle == 'shift':
                    ActionChains(self.driver).key_down(Keys.SHIFT).perform()

            # 当操作对象是这种[2018-01, 2018-01]时，点击之后就结束循环
            if (num == len(operationObjects)) | (
                    (len(operationObjects) > 1) & (operationObjects[0] == operationObjects[len(operationObjects) - 1])):
                if clickStyle == 'ctrl':
                    ActionChains(self.driver).key_up(Keys.CONTROL).perform()
                elif clickStyle == 'shift':
                    ActionChains(self.driver).key_up(Keys.SHIFT).perform()
                self.click(menu, arg=args, type='right')

        if definePart == "context":
            self.wait_loaded("新页面-在上下文中过滤")
            self.click('新页面-在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("新页面-使用默认度量")
            self.click('新页面-使用默认度量')
        elif definePart.startswith("rep"):
            self.click('新页面-替换')
            self.click(self.__defineZoneDic_new_tab[definePart])
        else:
            self.click('新页面-插入')
            self.click(self.__defineZoneDic_new_tab[definePart])
        self.waitStatusLoaded_in_new_tab()

    def define_multi_in_new_tab(self, operationObjectsList, levelList, definePart):
        """ 插入控件到行|列|上下文|度量, 可以选取不同层次的控件, 点击方式固定为CTRL\n

        :param operationObjectsList: 要插入的多个控件名列表(例如: [["机构","日期"],["2020"]]) :<list>双重列表\n
        :param levelList: 控件层次列表(例如: [1, 2]) :<list>\n
        :param definePart: 要插入的区域,"row": "行", "mulRow": "嵌套的行", "col": "列", "mulCol": "嵌套的列", "context": "在上下文中过滤", "repRow": "替换_行", "repCol": "替换_列", "measure": "使用默认度量" :<str>\n
        :return: 无\n
        """
        if isinstance(operationObjectsList, list) is not True:
            raise Exception(str(operationObjectsList) + "非列表, 格式错误")
        for i in range(len(operationObjectsList)):
            menu = self.__cognosMenuList_new_tab[levelList[i] - 1]

            # 控制变量num,用于点击第一次CTRL,
            num = 0
            for operationObject in operationObjectsList[i]:
                for j in range(10):
                    if self.count(menu, arg=operationObject) == 0:
                        sleep(0.3)
                    else:
                        break
                if self.count(menu, arg=operationObject) == 0:
                    raise Exception(str(operationObject) + "不存在或层次定义错误")
                    # 有时候会点击失效，一次不行，再来一次
                for j in range(10):
                    try:
                        self.click(menu, arg=operationObject)
                        break
                    except Exception as e:
                        sleep(1)
                num += 1
                if num == 1:
                    ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                    # 当操作对象是这种[2018-01, 2018-01]时，点击之后就结束循环
                if (i == len(operationObjectsList) - 1) & (num == len(operationObjectsList[i])):
                    ActionChains(self.driver).key_up(Keys.CONTROL).perform()

                    # 有时候会点击失效，一次不行，再来一次
                    for j in range(10):
                        try:
                            self.click(menu, arg=operationObject, type="right")
                            break
                        except Exception as e:
                            sleep(1)

        if definePart == "context":
            self.wait_loaded("新页面-在上下文中过滤")
            self.click('新页面-在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("新页面-使用默认度量")
            self.click('新页面-使用默认度量')
        elif definePart.startswith("rep"):
            self.click('新页面-替换')
            self.click(self.__defineZoneDic_new_tab[definePart])
        else:
            self.click('新页面-插入')
            self.click(self.__defineZoneDic_new_tab[definePart])
        self.waitStatusLoaded_in_new_tab()

    def define_cross_in_new_tab(self, operationObjectsList, levelList, clickStyle:str, definePart):
        """跨级插入菜单，支持相同的父菜单下不同级子菜单的插入
        operationObjectsList: 要插入的多个控件名列表(例如: [["机构","日期"],["2020"]]) :<list>双重列表
        """
        if isinstance(operationObjectsList, list) is not True:
            raise Exception(str(operationObjectsList) + "非列表, 格式错误")

        for i in range(len(operationObjectsList)):
            menu = self.__cognosMenuList_new_tab[levelList[i] - 1]

            # 控制变量num,用于点击第一次CTRL,

            for  index,operationObject in enumerate(operationObjectsList[i]):
                for j in range(10):
                    if self.count(menu, arg=operationObject) == 0:
                        sleep(0.3)
                    else:
                        break
                if self.count(menu, arg=operationObject) == 0:
                    raise Exception(str(operationObject) + "不存在或层次定义错误")
                    # 有时候会点击失效，一次不行，再来一次
                for j in range(10):
                    try:
                        self.click(menu, arg=operationObject)
                        if index ==0:
                            ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                            sleep(1)
                            if clickStyle == "shift":
                                ActionChains(self.driver).key_down(Keys.SHIFT).perform()
                                sleep(1)
                        break
                    except Exception as e:
                        sleep(1)
                    # 当操作对象是这种[2018-01, 2018-01]时，点击之后就结束循环
            if clickStyle == "shift":
                ActionChains(self.driver).key_up(Keys.SHIFT).perform()

        ActionChains(self.driver).key_up(Keys.CONTROL).perform()

        menu_last = self.__cognosMenuList_new_tab[levelList[-1] - 1]
        object_last = operationObjectsList[-1][-1]
        self.click(menu_last, arg=object_last, type="right")



        if definePart == "context":
            self.wait_loaded("新页面-在上下文中过滤")
            self.click('新页面-在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("新页面-使用默认度量")
            self.click('新页面-使用默认度量')
        elif definePart.startswith("rep"):
            self.click('新页面-替换')
            self.click(self.__defineZoneDic_new_tab[definePart])
        else:
            self.click('新页面-插入')
            self.click(self.__defineZoneDic_new_tab[definePart])
        self.waitStatusLoaded_in_new_tab()

    def define_cross_with_duplicate_name_in_new_tab(self, operationObjectsList, clickStyle:str, definePart):
        """跨级插入菜单，支持相同的父菜单下不同级子菜单的插入,支持重名菜单
        operationObjectsList: 要插入的多个控件名列表(例如: [["机构","日期"],["2020"]]) :<list>双重列表
        """
        if isinstance(operationObjectsList, list) is not True:
            raise Exception(str(operationObjectsList) + "非列表, 格式错误")

        menu = "新页面-菜单控件-重名"

        for operationObjects in (operationObjectsList):


            # 控制变量num,用于点击第一次CTRL,

            for index, operationObject in enumerate(operationObjects):
                args = operationObject.split(">")

                for j in range(10):
                    if self.count(menu, arg=args) == 0:
                        sleep(0.3)
                    else:
                        break
                if self.count(menu, arg=args) == 0:
                    raise Exception(str(operationObject) + "不存在或层次定义错误")
                    # 有时候会点击失效，一次不行，再来一次
                for j in range(10):
                    try:
                        self.click(menu, arg=args)
                        if index ==0:
                            ActionChains(self.driver).key_down(Keys.CONTROL).perform()
                            sleep(1)
                            if clickStyle == "shift":
                                ActionChains(self.driver).key_down(Keys.SHIFT).perform()
                                sleep(1)
                        break
                    except Exception as e:
                        sleep(1)
            if clickStyle == "shift":
                ActionChains(self.driver).key_up(Keys.SHIFT).perform()

        ActionChains(self.driver).key_up(Keys.CONTROL).perform()

        args_last = operationObjectsList[-1][-1].split(">")
        self.click(menu, arg=args_last, type="right")

        if definePart == "context":
            self.wait_loaded("在上下文中过滤")
            self.click('新页面-在上下文中过滤')
        elif definePart == "measure":
            self.wait_loaded("新页面-使用默认度量")
            self.click('新页面-使用默认度量')
        elif definePart.startswith("rep"):
            self.click('新页面-替换')
            self.click(self.__defineZoneDic_new_tab[definePart])
        else:
            self.click('新页面-插入')
            self.click(self.__defineZoneDic_new_tab[definePart])
        self.waitStatusLoaded_in_new_tab()

    def downloadExcel_in_new_tab(self, downPath):
        """ 下载报表\n

        :param downPath: 下载路径(文件路径+文件名),例如: r"d:\A\101.xlsx" :<str>\n
        :return: 无\n
        """
        currnet_url = self.get_url()
        self.click('新页面-运行')
        self.down_by_element('新页面-以 Excel 2007格式', file_name=downPath)
        self.catch('IBM Cognos Viewer', mode='contain')
        print(self.get_title())
        print(self.get_url())
        self.close()
        self.catch(currnet_url, type='url')
        pass

    def showAllRows_in_new_tab(self):
        """
        显示所有行\n
        """
        # 点击更多
        sleep(3)
        self.click('新页面-更多-行', simulate=True)
        # 点击自定义
        self.click('新页面-自定义')
        # 点击可视项目下三角
        self.click('新页面-可视项目下三角')
        # 点击9999
        self.click('新页面-9999')
        # 点击确定
        self.click("新页面-显示所有行-确定")
        self.waitStatusLoaded_in_new_tab()
        sleep(2)

    def showAllCols_in_new_tab(self):
        """
        显示所有列\n
        """
        # 点击更多
        sleep(3)
        if self.count("新页面-横向滚动条") > 0:
            self.drag_then_drop("新页面-横向滚动条", "新页面-横向滚动条-最右")
        sleep(1)
        self.click('新页面-更多-列', simulate=True)
        # 点击自定义
        self.click('新页面-自定义')
        # 点击可视项目下三角
        self.click('新页面-可视项目下三角')
        # 点击9999
        self.click('新页面-9999')
        # 点击确定
        self.click("新页面-显示所有行-确定")
        self.waitStatusLoaded_in_new_tab()
        sleep(2)

    def delete_in_new_tab(self, deletePart):
        """ 删除行|列|上下文|度量的控件\n

        :param deletePart: 要删除的区域,"row": "行", "mulRow": "嵌套的行", "col": "列", "mulCol": "嵌套的列", "context": "上下文过滤器", "mulContext": "嵌套的上下文过滤器" :<str>\n
        :return: 无\n
        """
        delObj = self.__deleteZoneDic_new_tab[deletePart]
        self.click(delObj)
        self.click("新页面-删除")
        self.waitStatusLoaded_in_new_tab()

    def get_bas_table_data_in_new_tab(self):
        """
        获取多维分析表格页面数据-新页面内
        :return: 表格数据，二维数组
        """
        table = self.__find_element("新页面-数据表格")
        rows = table.find_elements_by_tag_name('tr')
        result = []
        for row in rows:
            # 在每行中获取所有单元格
            cells = row.find_elements_by_tag_name('td')
            rows_data = []
            if len(cells) == 0:
                continue
            for cell in cells:
                try:
                    cell_data = cell.get_attribute("title")
                except:
                    continue
                if not cell_data or cell_data == '':
                    continue
                if cell.get_attribute("rowspan") == "1" or cell.get_attribute("colspan") == "1":
                    cell_data = cell_data.split(".")[-1]
                rows_data.append(cell_data)
            if len(rows_data) > 0:
                result.append(rows_data)

        return result

    def login_province_analyze(self):
        """
        登录分省分析
        """
        self.loginMainSystem()
        self.enterSubSystem("分省分析")
        sleep(3)

    def get_bas_table_data(self):
        """
        获取多维分析表格页面数据
        :return: 表格数据，二维数组
        """
        table = self._Browser__find_element("数据表格")
        rows = table.find_elements_by_tag_name('tr')
        result = []
        for row in rows:
            # 在每行中获取所有单元格
            cells = row.find_elements_by_tag_name('td')
            rows_data = []
            if len(cells) == 0:
                continue
            for cell in cells:
                try:
                    cell_data = cell.get_attribute("title")
                except:
                    continue
                if not cell_data or cell_data == '':
                    continue
                if cell.get_attribute("rowspan") == "1" or cell.get_attribute("colspan") == "1":
                    cell_data = cell_data.split(".")[-1]
                rows_data.append(cell_data)
            if len(rows_data) > 0:
                result.append(rows_data)

        return result