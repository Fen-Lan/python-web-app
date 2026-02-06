# -*- coding: utf-8 -*- 
import traceback
from CLPC.tool import TOOL
from time import sleep, time
from CLPC.browser_visual import Browser
from CLPC.browser_visual import generate_new_driver
import sys, os
from CLPC.framework import *
from CLPC.office import excel
from CLPC.portal_visual import Portal
from CLPC.taoshubao_visual import Tao
from CLPC.scplus import SCPLUS
from CLPC.union_login import UNILOGIN
from CLPC.bdp import BDP
import take
import urllib.parse

@PRE_CHECK
def get_Id(drag_info_url):
    global_driver = generate_new_driver()
    # pp_url = "http://9.1.212.27:28080/main/"
    # inf_url = "http://9.1.212.27:28080/transactionmetadata.pinpoint"
    # drag_info_url = "http://9.1.212.27:28080/transactionList/qiye-order-service@BES/5m/2026-01-30-11-20-53?dragInfo=%7B%22x1%22:1769743023339,%22x2%22:1769743167407,%22y1%22:0,%22y2%22:2134,%22agentId%22:%22%22,%22dotStatus%22:%5B%22success%22,%22failed%22%5D%7D"
    try:

        browser = Browser(global_driver)  # 操作浏览器
        tsb = Tao(global_driver)  # 操作淘数宝
        scp = SCPLUS(global_driver)  # 操作闪查
        unilogin = UNILOGIN(global_driver)  # 操作内网门户
        portal = Portal(global_driver)  # 操作多维分析
        bdp = BDP(global_driver)  # 操作车险超多维
        base_url,service_name, time_interval,date_time, x1, x2, y1, y2 = take.get_url_info(drag_info_url)
        pp_url = base_url + "/main/"+service_name+ "@BES/"+time_interval + "/" + date_time
        # 比较 y1 和 y2，找到较大的值
        max_value = max(y1, y2)
        min_value = min(y1, y2)
        y_max = 10000
        #offset = 30
        oub_flag = False # 框选越界标志
        browser.create(pp_url)
        # sleep(10)

        if max_value > y_max:
            if max_value > 1000000:  # 防止越界框选数值过大
                max_value = min_value
                oub_flag = True
            # 计算较大值的最大位数
            max_digits = len(str(max_value))
            num_str = str(max_value).lstrip('0')
            # if not num_str:
            #     return 0,0,0
            y_digits = int (num_str[0])
            y_max = (y_digits+1) * (10 ** (max_digits - 1))
            print(f"y_max: {y_max}")
            if oub_flag:
                if y1 > y2:
                    y1 =y_max
                else:
                    y2 =y_max
            if max_value > 10000:
                sleep(10)
                browser.click("设置ms", simulate=False,type="left",arg="")
                sleep(2)
                browser.clear_input("输入y轴最大值",arg="")
                sleep(2)
                browser.input_text("输入y轴最大值",y_max,arg="")
                sleep(2)
                browser.click("应用", simulate=False,type="left",arg="")
        # location = browser.pos_ele_center("应用信息点位图", arg=None)
        # x = location['x']
        # y = location['y']
        # print(x, y)


        # 计算坐标
        location,size = browser.pos("点位图具体位置", arg=None)
        x0, y0 = location['x'], location['y']  # 点位图的左上角位置
        width, height = size['width'], size['height']  # 点位图的宽度和高度
        time_value = int(''.join(filter(str.isdigit, time_interval)))
        if "h" in time_interval:
            time_value = time_value * 60
        print(f"x0: {x0}, y0: {y0}, width: {width}, height: {height}, x1: {x1}, x2: {x2}, y1: {y1}, y2: {y2}, date_time: {date_time}, time_value: {time_value}")
        start_x, start_y, end_x, end_y = take.calculate_coordinates(y_max,x0, y0, width, height, x1, x2, y1, y2,date_time, time_value)
        print(f"start_x: {start_x}, start_y: {start_y}, end_x: {end_x}, end_y: {end_y}")

        # 使用计算出的坐标进行操作
        browser.simulate_area_select_and_switch(start_x, start_y, end_x, end_y,base_url)
        sleep(5)
        list_text = browser.text("应用信息", arg="")
        data = take.extract_data(list_text)
        return data

        # # 解析响应内容
        # if response_data:
        #     response_json = json.loads(response_data)
        #     metadata = response_json.get('metadata', [])
        #     if metadata:
        #         for item in metadata:
        #             traceId = item.get('traceId')
        #             agentId = item.get('agentId')
        #             print(f"traceId: {traceId}")
        #             print(f"agentId: {agentId}")
        #     else:
        #         print("metadata 为空")

    except:
        print(traceback.format_exc())
    finally:
        if global_driver:
            global_driver.quit()

