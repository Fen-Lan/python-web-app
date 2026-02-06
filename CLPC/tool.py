import time
from datetime import datetime
from time import sleep

import requests, os, traceback, json, sys, base64
import datetime
#from dateutil.relativedelta import relativedelta
from CLPC.framework import FRAME
from CLPC.framework import FUNC_USAGE_TRACKER, static_func_tracker, FRAME


@FUNC_USAGE_TRACKER
class TOOL:

    @staticmethod
    def get_my_asset(asset_name, timeout=60):
        exec_env = FRAME.get_env_arg("RPA_ENV", "")
        if exec_env == "prod" or exec_env == 'test': # 服务器端运行
            uuid = FRAME.get_env_arg("RPA_APP_CREATOR_UUID")

        else:  # 本地运行
            uuid = FRAME.get_user_uuid()

        if exec_env == "prod" or exec_env == 'test':  ## 服务器端运行
            if exec_env == "prod":  #生产服务器
                api_url = FRAME.BACKEND_ADDR_PROD + "api/base/getUserAssets"
            else:  ## 测试服务器
                api_url = 'http://zhida-dev:19090/' + "api/base/getUserAssets"
        else:  ## 本地运行
            local_env = FRAME.get_tool_env()
            if local_env == "prod":  ## 生产tool
                api_url = FRAME.BACKEND_ADDR_PROD + "api/base/getUserAssets"
            else:  # 测试tool
                api_url = 'http://zhida-dev:19090/' + "api/base/getUserAssets"

        data = {"uuid": uuid, "name": asset_name}
        json_data = json.dumps(data)
        try:
            response = requests.post(api_url, data=json_data, timeout=timeout)
        except:
            print("请求超时, 请稍后再试")
            raise Exception("请求超时")
        if response.status_code == 200:
            res_data = response.json()
            res_status = res_data.get("code")
            if res_status == 1:
                encode_data = res_data.get("data")
                if encode_data['type']=='string':
                    return encode_data['value']
                else:
                    key = TOOL.get_asset_key()
                    asset_value = TOOL.XXTEADecrypt(encode_data['value'],key)
                    return asset_value
            else:
                res_msg = res_data.get("msg")
                print(f"资产变量获取失败，原因: {res_msg}")
        else:
            print(f"资产变量获取失败，原因: {response.status_code}")

    @staticmethod
    def get_asset_key():
        """
        获取密钥
        """
        run_env = FRAME.get_env_arg("RPA_ENV", "")
        if run_env == "prod" or run_env == 'test':  ## 服务器端运行
            if run_env == "prod":
                api_url = FRAME.BACKEND_ADDR_PROD + "api/base/getXXTeaSecret"
            else:
                api_url = 'http://zhida-dev:19090/' + "api/base/getXXTeaSecret"
        else: ## 开发运行环境
            local_env = FRAME.get_tool_env()
            if local_env == "prod":  # 生产tool
                api_url = FRAME.BACKEND_ADDR_PROD + "api/base/getXXTeaSecret"
            else:  ## 测试tool
                api_url = 'http://zhida-dev:19090/' + "api/base/getXXTeaSecret"

        try:
            response = requests.get(api_url)
        except:
            print("请求超时, 请稍后再试")
            raise Exception("请求超时")
        return response.json().get("data")

    @staticmethod
    def decrypt(v, k):
        """XXTEA解密函数"""
        delta = 0x9E3779B9
        n = len(v)
        rounds = 6 + 52 // n

        sum = ((rounds * delta) & 0xFFFFFFFF & 0xFFFFFFFF)
        y = v[0]

        for i in range(rounds):
            e = (sum >> 2) & 3
            for p in range(n - 1, 0, -1):
                z = v[p - 1]
                v[p] = (v[p] - (
                            ((z >> 5 ^ y << 2) + (y >> 3 ^ z << 4)) ^ ((sum ^ y) + (k[(p & 3) ^ e] ^ z)))) & 0xFFFFFFFF
                y = v[p]
            z = v[n - 1]
            v[0] = (v[0] - (((z >> 5 ^ y << 2) + (y >> 3 ^ z << 4)) ^ ((sum ^ y) + (k[(0 & 3) ^ e] ^ z)))) & 0xFFFFFFFF
            y = v[0]
            sum = (sum - delta) & 0xFFFFFFFF

        return v

    @staticmethod
    def XXTEADecrypt(ciphertext, key):
        import struct
        """XXTEA解密主函数"""
        if len(key) < 16:
            # 密钥长度不足时进行填充
            key = key.ljust(16, '\x00')

        # 解码Base64并转换为uint32列表
        cipher_bytes = base64.b64decode(ciphertext)
        """将字符串转换为uint32列表"""
        n = len(cipher_bytes)
        # 填充到4的倍数
        if n % 4 != 0:
            cipher_bytes += b'\x00' * (4 - n % 4)
            n = len(cipher_bytes)

        v = [struct.unpack('<I', cipher_bytes[i:i + 4])[0] for i in range(0, n, 4)]
        # 处理密钥
        key_bytes = key.encode('utf-8')
        if len(key_bytes) < 16:
            key_bytes = key_bytes.ljust(16, b'\x00')
        n = len(key_bytes)
        # 填充到4的倍数
        if n % 4 != 0:
            key_bytes += b'\x00' * (4 - n % 4)
            n = len(key_bytes)

        k = [struct.unpack('<I', key_bytes[i:i + 4])[0] for i in range(0, n, 4)]
        if len(k) < 4:
            k.extend([0] * (4 - len(k)))

        # 解密
        decrypted = TOOL.decrypt(v, k)

        # 转换为字符串并去除填充
        result = (b''.join(struct.pack('<I', x) for x in decrypted)).decode('utf-8', errors='ignore')
        # 去除可能的空字符填充
        return result.rstrip('\x00')

    @staticmethod
    def join_str(*args):  # todo 如何在modules里定义可变参数
        """
        拼接输入变量为一个字符串
        """
        result = ''
        for s in args:
            result += str(s)
        print(result)
        return result

    @staticmethod
    def generate_time_str(year_type: str, year_flag: str, month_type: str, month_flag: str, day_type: str, day_flag: str = "", delta: str = "") -> str:
        """
        根据要求生成时间,delta在处理几月前时，不是判断大小月，需要自己保证不会出现2月30号等类似情况
        :param year_type: 年份格式，如yyyy yyyyy
        :param year_flag: 年份标记占位符，年，- 等
        :param month_type: 月份格式，如m,mm,mmm等
        :param month_flag: 月份标记占位符，月， - 等
        :param day_type: 日期格式，如d, dd, ddd等
        :param day_flag: 日志标记占位符，日， - 等
        :param delta: 时间偏移量，格式：xx天 几天前, xx月 几月前, xx年 几年前, 默认为空，输出当天日期
        """
        from datetime import datetime
        from datetime import timedelta
        today = datetime.today()
        year = datetime.today().year
        month = datetime.today().month
        day = datetime.today().day


        if "天" in delta:
            delta_day = int(delta.split("天")[0])
            target_date = today - timedelta(days=delta_day)
            target_year = target_date.year
            target_mon = target_date.month
            target_day = target_date.day

        elif "月" in delta:
            delta_mon = int(delta.split("月")[0])
            delta_year = round(delta_mon / 12)
            target_year = year - delta_year
            target_day = day
            target_mon = month - delta_mon
            while target_mon <= 0:
                target_mon += 12

        elif "年" in delta:
            delta_year = int(delta.split("年")[0])
            target_year = year - delta_year
            target_mon = month
            target_day = day

        elif delta == "":
            target_year = year
            target_mon = month
            target_day = day
        else:
            raise Exception("日期偏移量格式不正确")

        year_type_len = len(year_type)
        month_type_len = len(month_type)
        day_type_len = len(day_type)
        target_year_str = str(target_year).zfill(year_type_len)
        target_mon_str = str(target_mon).zfill(month_type_len)
        target_day_str = str(target_day).zfill(day_type_len)
        if day_flag == None:
            day_flag = ""
        result = f"{target_year_str}{year_flag}{target_mon_str}{month_flag}{target_day_str}{day_flag}"
        print(result)
        return result


    @staticmethod
    def get_verify_code(id_num):
        """
        通过身份账号获取企微动态验证码\n
        此方法只可在服务器端使用，账号需提前申请开通相关权限
        @id_num:身份证号
        """
        from CLPC.WeChatAuthRes import qyhAuthCode
        from CLPC.framework import FRAME
        if FRAME.get_env_arg('RPA_ENV') == 'prod' or FRAME.get_env_arg('RPA_ENV') == 'test':
            try:
                result = qyhAuthCode(id_num)
                if result:
                    return result
                else:
                    print("验证码获取失败，请确认网络环境，并提前申请账号权限")
                    return False
            except:
                print("验证码获取异常")
                return False

        else:
            print("本地开发，使用填入验证码")
            return False

    @staticmethod
    def robot_send_markdown_msg(webhook_id, msg):
        """
        使用群机器人发送markdown信息
        @rbt_id: 群机器人的账号
        @rbt_pwd: 群机器人的密码
        @msg: 待发送的markdown信息
        """
        from CLPC.WechatMsgSend import MsgSendByRobot

        try:
            wc = MsgSendByRobot(webhook_id)
            wc.sendMarkdownMsg(msg = msg)
        except Exception as e:
            print(e)
            print("群机器人markdown信息发送失败，msg:{}".format(msg))

    @staticmethod
    def robot_send_markdown_msg_intranet(rbt_id, rbt_pwd, msg):
        """
        使用群机器人发送markdown信息（内网接口）
        @rbt_id: 群机器人的账号
        @rbt_pwd: 群机器人的密码
        @msg: 待发送的markdown信息
        """
        from CLPC.WechatMsgSend import MsgSendByRobotIntranet

        try:
            print(msg)
            wc = MsgSendByRobotIntranet(rbt_id,rbt_pwd)
            wc.sendMarkdownMsg(msg = msg)
        except:
            print("群机器人markdown信息发送失败，msg:{}".format(msg))

    @staticmethod
    def robot_send_msg(rbt_id, msg, mention_list=[]):
        """
        使用群机器人发送文本信息
        @rbt_id: 群机器人的webhook
        @msg: 待发送的文本信息
        """
        from CLPC.WechatMsgSend import MsgSendByRobot

        try:
            wc = MsgSendByRobot(rbt_id)
            wc.sendMsg(msg = msg, mentioned_mobile_list=mention_list)
        except:
            print("群机器人信息发送失败，msg:{}".format(msg))

    @staticmethod
    def robot_send_file(rbt_id, files=[]):
        from CLPC.WechatMsgSend import MsgSendByRobot

        try:
            wc = MsgSendByRobot(rbt_id)
            wc.sendMsg(msg=None, file_path=files)
        except:
            print(traceback.format_exc())
            print("群机器人文件发送失败")

    @staticmethod
    def cp_file(source, result):
        from shutil import copyfile
        if not os.path.exists(source):
            print("{}不存在！".format(source))
            return
        copyfile(source, result)
        pass

    @staticmethod
    def del_file(path):
        """
        清空path下所有文件，但是不删除path
        """
        import os
        for i in os.listdir(path):
            path_file = os.path.join(path, i)
            if os.path.isfile(path_file):
                os.remove(path_file)
            else:
                TOOL.del_file(path_file)

    @staticmethod
    def send_email(receivers=[], cc=[], title='', content='', att=[], content_html=False):
        """
        发送电子邮件，发件人固定为 lcjqr
        :param receivers: 收件人，格式：['1@1.com','2@2.com','3@3.com']
        :param cc:
        :param title:
        :param content:
        :param att:
        :param content_html:
        """
        from CLPC.sendMail import send_mail
        if content_html:
            content_type = 'html'
        else:
            content_type = ''
        send_mail(to_mail=receivers, cc_mail=cc, to_title=title, to_content=content, to_att=att, content_type=content_type)


    @staticmethod
    def mov_file(source_path, target_path):
        """
        move file to target path
        """
        import shutil
        if not os.path.exists(source_path):
            print("{}不存在！".format(source_path))
            return
        shutil.copy2(source_path, target_path)
        sleep(2)
        # 删除原文件
        os.remove(source_path)
        pass



    @staticmethod
    def file_exist(filename):
        """
        判断文件是否存在
        """
        if os.path.exists(filename):
            return True
        return False


    @staticmethod
    def get_host_ip():
        """
        查询本机ip地址
        :return: ip
        """
        import socket

        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
        except:
            ip = ""
        finally:
            s.close()

        return ip

    @staticmethod
    def send_file_by_wx(receiver, info, files = []):
        """
        通过企业微信发送文件
        """
        from CLPC.WechatMsgSend import MsgSendByMobile
        try:
            wc = MsgSendByMobile()
            wc.sendMsg(userNum = receiver, msg = info, file_path = files)
        except:
            print(traceback.format_exc())
            print("发送文件错误")


    @staticmethod
    def send_markdown_info_by_wx(receiver, info, proj_name):
        """
        发送程序通知信息
        :params receiver 接收人手机号
        :params info 待发送信息
        :params proj_name 应用名称
        """
        from CLPC.WechatMsgSend import MsgSendByMobile
        try:
            msg = "应用：[{}]，{}".format(proj_name, info)
            wc = MsgSendByMobile()
            wc.sendMarkdownMsg(userNum = receiver, msg = msg)
        except:
            pass

    @staticmethod
    # @static_func_tracker
    def send_info_by_wx(receiver, info, proj_name):
        """
        发送程序通知信息
        :params receiver 接收人手机号
        :params info 待发送信息
        :params proj_name 应用名称
        """
        from CLPC.WechatMsgSend import MsgSendByMobile
        try:
            msg = "应用：[{}]，{}".format(proj_name, info)
            wc = MsgSendByMobile()
            wc.sendMsg(userNum = receiver, msg = msg)
        except:
            pass

    @staticmethod
    def send_err_by_wx(receiver : str, info : str, proj_name: str):
        """
        发送程序异常信息，自动拼接ip信息

        :param receiver: 接收人手机号
        :param info: 待发送信息
        :param proj_name: 应用名称
        """
        from CLPC.WechatMsgSend import MsgSendByMobile
        ip = TOOL.get_host_ip()
        msg = "应用：[{}]，IP: [{}]，{}".format(proj_name, ip, info)
        try:
            wc = MsgSendByMobile()
            wc.sendMsg(userNum = receiver, msg = msg)
        except:
            print(msg)



    @staticmethod
    def get_FileSize_kb(filePath):
        """
        返回指定文件的大小，单位kb

        :param filePath: 文件的完整路径
        """
        fsize = os.path.getsize(filePath)
        fsize = fsize / float(1024)
        return round(fsize)

    @staticmethod
    def convert_to_number(letter, columnA = 1):
        """
        sample：
        字母列号转数字
        columnA: 你希望A列是第几列(0 or 1)? 默认1
        return: int
        """
        ab = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        letter0 = letter.upper()
        w = 0
        for _ in letter0:
            w *= 26
            w += ab.find(_)
        return w - 1 + columnA

    @staticmethod
    def convert_to_letter(number, columnA = 1):
        """
        数字转字母列号

        :param columnA: 你希望A列是第几列(0 or 1)? 默认1
        :return: in upper case<str>
        """
        ab = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        n = number - columnA
        x = n % 26
        if n >= 26:
            n = int(n / 26)
            return TOOL.convert_to_letter(n, 1) + ab[x + 1]
        else:
            return ab[x + 1]

    @staticmethod
    def print_exception_msg(file_name):
        """
        传入参数固定为： __file__
        """
        import sys
        s = sys.exc_info()
        msg = "出错文件: {}, 出错行数: {}, 错误信息: {}".format(str(file_name).split("\\")[-1], s[2].tb_lineno, s[1])
        print(msg)
        return msg

    @staticmethod
    def handle_error(fun_name):
        from functools import wraps
        import traceback, inspect, sys
        @wraps(fun_name)
        def inner(*args, **kwargs):
            try:
                return fun_name(*args, **kwargs)
            except Exception as e:
                # print(traceback.format_exc())
                s = sys.exc_info()
                tb = s[2]
                err_line = traceback.extract_tb(tb)[1].lineno
                err_file = inspect.getfile(fun_name)
                err_info = s[1]
                print("出错文件: {}, 出错行数: {}, 错误信息: {}".format(err_file.split("\\")[-1], err_line, err_info))
        return inner

    @staticmethod
    def download_ftp_file(ip, usr, pwd, remote_file, local_file):
        from CLPC.FTP_SDK import ConnectFTP
        try:
            c = ConnectFTP(ip, usr, pwd)
            c.login()
        except:
            raise Exception("FTP连接失败")
        try:
            c.downLoadFile(local_file, remote_file)
            c.logout()
        except:
            raise Exception("文件下载失败")

    @staticmethod
    def download_ftp_folder(ip, usr, pwd, remote_folder, local_folder):
        from CLPC.FTP_SDK import ConnectFTP
        try:
            c = ConnectFTP(ip, usr, pwd)
            c.login()
        except:
            raise Exception("FTP连接失败")
        try:
            c.downLoadFiles(local_folder, remote_folder)
            c.logout()
        except:
            raise Exception("文件下载失败")

    @staticmethod
    def upload_ftp_file(ip, usr, pwd, local_file, remote_file):
        from CLPC.FTP_SDK import ConnectFTP
        try:
            c = ConnectFTP(ip, usr, pwd)
            c.login()
        except:
            raise Exception("FTP连接失败")
        try:
            c.upload_file(local_file, remote_file)
            c.logout()
        except:
            raise Exception("文件下载失败")

    @staticmethod
    def zip_folder(folder_path, output_path):
        """
        将文件夹压缩成zip格式
        """
        import zipfile
        import os
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, folder_path)
                    zipf.write(file_path, arcname)

    @staticmethod
    def compress_folder(folder_path, output_filename):
        """
        将文件夹压缩成tar.gz格式
        """
        import tarfile
        # 创建一个tar文件
        with tarfile.open(output_filename, 'w:gz') as tar:
            # 获取文件夹中的所有文件和文件夹
            for root, dirs, files in os.walk(folder_path):
                # 将文件夹中的文件添加到tar文件中
                for file in files:
                    path = os.path.join(root, file)
                    tar.add(path, arcname=os.path.relpath(path, folder_path))


    @staticmethod
    def wait_time(seconds:int):
        """
        :param seconds 等待时间，单位s
        """
        time.sleep(seconds)

    @staticmethod
    def get_year_of_today() -> str:
        """
        获取今天的年份
        """
        from datetime import datetime
        result = datetime.today().strftime("%Y")
        return result

    @staticmethod
    def get_month_of_today() -> str:
        """
        获取今天的月份
        """
        from datetime import datetime
        result = datetime.today().strftime("%m")
        return result

    @staticmethod
    def get_day_of_today() -> str:
        """
        获取今天的日
        """
        from datetime import datetime
        result = datetime.today().strftime("%d")
        return result

    @staticmethod
    def get_date_of_today() -> str:
        """
        获取今天的日期
        """
        from datetime import datetime
        result = datetime.today().strftime("%Y-%m-%d")
        return result

    @staticmethod
    def create_list() -> list:
        """
        创建一个空列表
        """

        return []

    @staticmethod
    def append_to_list(lst:list, item):
        """
        向列表中添加元素
        :param lst: 列表
        :param item: 待添加元素
        """
        lst.append(item)
        return lst

    @staticmethod
    def add(num1, num2):
        return eval(str(num1))+eval(str(num2))

    @staticmethod
    def getLenthOfList(list0):
        if not isinstance(list0,list):
            list0=eval(list0)
        return len(list0)

    @staticmethod
    def reverseList(list0):
        if not isinstance(list0,list):
            list0=eval(list0)
        return list(reversed(list0))

    @staticmethod
    def getElementInList(list0,numInt):
        if not isinstance(list0,list):
            list0=eval(list0)
        return list0[numInt]


    @staticmethod
    def getPartList(list0,fromIndex=0,endIndex=''):
        if not isinstance(list0,list):
            list0=eval(list0)
        if endIndex:
            return list0[fromIndex:endIndex]
        else:
            return list0[fromIndex:]

    @staticmethod
    def concatList(list0,list1):
        if not isinstance(list0,list):
            list0=eval(list0)
        if not isinstance(list1,list):
            list1=eval(list1)
        return list0+list1

    @staticmethod
    def assembleList(list0,list1):
        if not isinstance(list0,list):
            list0=eval(list0)
        if not isinstance(list1,list):
            list1=eval(list1)
        return [list0,list1]

    @staticmethod
    def html_table_2_png(html_file, png_file, width=1920):
        """
        将带table的html文件转换成png文件
        """
        from CLPC.browser_visual import generate_screenshot_driver
        from CLPC.framework import FRAME
        exec_env = FRAME.get_env_arg("RPA_ENV", "")

        try:
            t_d = generate_screenshot_driver(width=width)
            if not os.path.exists(html_file):
                raise Exception(f"文件{html_file}不存在，请确认")
            abs_path = os.path.abspath(html_file)
            t_d.get(f"file:///{abs_path}")
            element = t_d.find_element_by_xpath("//table")
            size = element.size
            tb_width = size['width']
            tb_height = size['height']
            print(f"计算宽度：{tb_width},计算高度{tb_height}")
            if exec_env == "prod" or exec_env == 'test':
                t_d.set_window_size(tb_width + 10, tb_height + 20)
            else:
                t_d.set_window_size(tb_width + 10, tb_height + 150)

            sleep(3)
            t_d.save_screenshot(png_file)
        except Exception as e:
            print("html转png失败", e)
        finally:
            if t_d:
                t_d.quit()

    @staticmethod
    def generate_BAS_duplicate_menu_list(fatherMenus,dateStrList,clickStyle='ctrl'):
        '''组装BAS同级同名日期菜单列表
        :param fatherMenus:父菜单列表，如['统计日期','承保年月']:<list>\n
        :param dateStrList:日期字符串列表，如['2023', '2024', '2025']、['2023-01', '2024-01', '2025-01']:<list>\n
        :param clickStyle: 点击方式,"ctrl"|"shift" :<str>\n
        '''
        fatherMenus=eval(str(fatherMenus))
        dateStrList=eval(str(dateStrList))
        
        returnList=[]
        for fm in fatherMenus:
            if clickStyle!='ctrl':
                dateStrList=[dateStrList[0],dateStrList[-1]]
            for dstr in dateStrList:
                returnList.append('{}>{}'.format(fm,dstr))
        return returnList

    @staticmethod
    def generate_BAS_cross_duplicate_menu_list(fatherMenus,dateStrList2d,clickStyle='ctrl'):
        '''组装BAS跨级同名日期菜单列表
        :param fatherMenus:父菜单列表，如['承保年月']:<list>\n
        :param dateStrList2d:日期字符串二维列表，如[['2023', '2024', '2025'], ['2023-01', '2024-01', '2025-01']]:<list>\n
        :param clickStyle: 点击方式,"ctrl"|"shift" :<str>\n
        '''
        fatherMenus=eval(str(fatherMenus))
        dateStrList2d=eval(str(dateStrList2d))
        
        returnList2d=[]
        if clickStyle!='ctrl':
            for dateStrListi in range(len(dateStrList2d)):
                dateStrList=dateStrList2d[dateStrListi]
                dateStrList2d[dateStrListi]=[dateStrList[0],dateStrList[-1]]
        for dateStrList in dateStrList2d:
            returnList=[]
            for fm in fatherMenus:
                for dstr in dateStrList:
                    returnList.append('{}>{}'.format(fm,dstr))
            returnList2d.append(returnList)
        return returnList2d
        

@FUNC_USAGE_TRACKER
class DATE_TOOL:

    @staticmethod
    def getDaysOfTheMonthBaseOneDay(baseDayStr,conn='2024-01-01'):
        """
        基于某日生成所在月份日期列表\n
        conn支持： 2024-01-01,2024-1-1,2024年01月01日,2024年1月1日\n
        传参举例：'2025-02-02','2024-01-01'\n
        返回举例：\n
            ['2025年02月01日', '2025年02月02日']
        """
        from datetime import datetime
        baseDayDate=datetime.strptime(baseDayStr,'%Y-%m-%d')
        year = baseDayDate.year

        if conn in ['2024-01-01',"2024-1-1"]:
            prototype = "{}-{}-{}"
        elif conn in ['2024年1月1日',"2024年01月01日"]:
            prototype = "{}年{}月{}日"        # 如2023年4月1日

        day=baseDayDate.day
        month=baseDayDate.month
        if conn not in ['2024年1月1日','2024-1-1']:
            month = str(month) if month >= 10 else "0" + str(month)
        else:
            month = str(month)
        days = []
        for i in range(1, day + 1):
            day = str(i)
            if conn not in ['2024年1月1日','2024-1-1']:
                day = str(i) if i >= 10 else "0" + str(i)
            else:
                day = str(i)
            formatStr = prototype.format(year, month,day)
            days.append(formatStr)
        # print(days)
        return days

    @staticmethod
    def getMonthsOfTheYearBaseOneDay(baseDayStr,conn='202401'):
        """
        基于某日生成所在年月份列表\n
        conn支持： 2024-01,2024001,2024年01月,2024年1月,2024/1月,202401\n
        传参举例：'2025-02-01','202401'\n
        返回举例：\n
            ['2025年01月', '2025年02月']\n
            ['2025/1月', '2025/2月']\n
            ['202501', '202502']
        """
        from datetime import datetime
        baseDayDate=datetime.strptime(baseDayStr,'%Y-%m-%d')
        year = baseDayDate.year
        cMonth = baseDayDate.month
        if conn == "2024年01月":
            prototype = "{0}年{1}月"
        elif conn == "2024/1月":
            prototype = "{0}/{1}月"
        else:
            prototype = "{0}" + conn[4:-2] + "{1}"
        months = []
        for i in range(1, cMonth + 1):
            if conn not in ['2024年1月,',"2024/1月"]:
                month = str(i) if i >= 10 else "0" + str(i)
            else:
                month = str(i)
            formatStr = str.format(prototype, year, month)
            months.append(formatStr)
        return months


    @staticmethod
    def getQuartersOfTheYearBaseOneDay(baseDayStr,conn='2024 年 1季'):
        """
        基于某日生成所在年季度列表\n
        conn支持： 2024Q1,2024年 1 季,2024年 1季,2024 年 1季\n
        传参举例：'2025-06-01','2024Q1'\n
        返回举例：\n
            ['2025Q1', '2025Q2']\n
            ['2025 年 1季', '2025 年 2季']
        """
        from datetime import datetime
        baseDayDate=datetime.strptime(baseDayStr,'%Y-%m-%d')
        year = baseDayDate.year
        cMonth = baseDayDate.month
        cQuarter = 0
        modulus = (cMonth - 1) // 3
        if modulus == 0:
            cQuarter = 1
        elif modulus == 1:
            cQuarter = 2
        elif modulus == 2:
            cQuarter = 3
        else:
            cQuarter = 4
        quarters = []
        for i in range(1, cQuarter+1):
            if conn=='2024Q1':
                prototype = str(year) + "Q{}"
            elif conn=='2024年 1 季':
                prototype = str(year) + "年 {} 季"
            elif conn=='2024年 1季':
                prototype = str(year) + "年 {}季"
            elif conn=='2024 年 1季':
                prototype = str(year) + " 年 {}季"
            formatStr = str.format(prototype, str(i))
            quarters.append(formatStr)
        return quarters

    @staticmethod
    def getMonthsOfYearOnYearBaseOneDay(baseDayStr,conn='202401',lenth=1,yoyType='年累计'):
        """
        基于某日生成所在年月份列表\n
        conn支持： 2024-01,2024001,2024年01月,2024年1月,2024/1月,202401\n
        yoyType支持： 年累计,当月\n
        传参举例：'2025-02-01','202401'\n
        返回举例：\n
            ['2024年01月', '2024年02月', '2025年01月', '2025年02月']\n
            ['2024/1月', '2024/2月', '2025/1月', '2025/2月']\n
            ['202401', '202402', '202501', '202502']
        """
        from datetime import datetime
        baseDayDate=datetime.strptime(baseDayStr,'%Y-%m-%d')
        year0 = baseDayDate.year
        yearList=[]
        for y in range(year0-lenth+1,year0+1):
            yearList.append(y)
            
        cMonth = baseDayDate.month
        if conn == "2024年01月":
            prototype = "{0}年{1}月"
        elif conn == "2024/1月":
            prototype = "{0}/{1}月"
        else:
            prototype = "{0}" + conn[4:-2] + "{1}"
        months = []
        
        for year in yearList:
            if yoyType=='年累计':
                beginM=1
            elif yoyType=='当月':
                beginM=cMonth
                
            for i in range(beginM, cMonth + 1):
                if conn not in ['2024年1月,',"2024/1月"]:
                    month = str(i) if i >= 10 else "0" + str(i)
                else:
                    month = str(i)
                formatStr = str.format(prototype, year, month)
                months.append(formatStr)
        return months


    @staticmethod
    def getQuartersOfYearOnYearBaseOneDay(baseDayStr,conn='2024 年 1季',lenth=1,yoyType='年累计'):
        """
        基于某日生成所在年季度列表\n
        conn支持： 2024Q1,2024年 1 季,2024年 1季,2024 年 1季\n
        yoyType支持： 年累计,当季\n
        传参举例：'2025-06-01','2024Q1',2\n
        返回举例：\n
            ['2024Q1', '2024Q2', '2025Q1', '2025Q2']\n
            ['2024 年 1季', '2024 年 2季', '2025 年 1季', '2025 年 2季']
        """
        from datetime import datetime
        baseDayDate=datetime.strptime(baseDayStr,'%Y-%m-%d')
        year0 = baseDayDate.year
        yearList=[]
        for y in range(year0-lenth+1,year0+1):
            yearList.append(y)
            
        cMonth = baseDayDate.month
        cQuarter = 0
        modulus = (cMonth - 1) // 3
        if modulus == 0:
            cQuarter = 1
        elif modulus == 1:
            cQuarter = 2
        elif modulus == 2:
            cQuarter = 3
        else:
            cQuarter = 4
        quarters = []
        for year in yearList:
            if yoyType=='年累计':
                beginQ=1
            elif yoyType=='当季':
                beginQ=cQuarter
                
            for i in range(beginQ, cQuarter+1):
                if conn=='2024Q1':
                    prototype = str(year) + "Q{}"
                elif conn=='2024年 1 季':
                    prototype = str(year) + "年 {} 季"
                elif conn=='2024年 1季':
                    prototype = str(year) + "年 {}季"
                elif conn=='2024 年 1季':
                    prototype = str(year) + " 年 {}季"
                formatStr = str.format(prototype, str(i))
                quarters.append(formatStr)
                
        return quarters

    @staticmethod
    def getYearsBaseOneDay(baseDayStr:str,lenth:int,fmt='2024年'):
        """
        基于某日生成近N年年份列表\n
        fmt支持：'2024','2024年'\n
        传参举例：'2025-01-01',4,'年'\n
        返回举例：\n
            ['2022', '2023', '2024', '2025']\n
            ['2022年', '2023年', '2024年', '2025年'] \n
        """
        from datetime import datetime
        suffix=''
        if fmt=='2024':
            suffix=''
        elif fmt=='2024年':
            suffix='年'
        baseDayDate=datetime.strptime(baseDayStr,'%Y-%m-%d')
        baseYear = baseDayDate.year
        yearsList = []
        # 倒序输出
        for i in range(lenth-1, -1, -1):
            year = baseYear - i
            year = str(year) + suffix
            yearsList.append(year)
        return yearsList


    @staticmethod
    def generate_time_str(year_type: str, year_flag: str, month_type: str, month_flag: str, day_type: str,
                          day_flag: str = "", delta: str = "") -> str:
        """
        根据要求生成时间,delta在处理几月前时，不是判断大小月，需要自己保证不会出现2月30号等类似情况
        :param year_type: 年份格式，如yyyy yyyyy
        :param year_flag: 年份标记占位符，年，- 等
        :param month_type: 月份格式，如m,mm,mmm等
        :param month_flag: 月份标记占位符，月， - 等
        :param day_type: 日期格式，如d, dd, ddd等
        :param day_flag: 日志标记占位符，日， - 等
        :param delta: 时间偏移量，格式：xx天 几天前, xx月 几月前, xx年 几年前, 默认为空，输出当天日期
        :return: 返回日期字符串
        """
        from datetime import datetime
        from datetime import timedelta
        today = datetime.today()
        year = datetime.today().year
        month = datetime.today().month
        day = datetime.today().day

        if "天" in delta:
            delta_day = int(delta.split("天")[0])
            target_date = today - timedelta(days=delta_day)
            target_year = target_date.year
            target_mon = target_date.month
            target_day = target_date.day

        elif "月" in delta:
            delta_mon = int(delta.split("月")[0])
            delta_year = round(delta_mon / 12)
            target_year = year - delta_year
            target_day = day
            target_mon = month - delta_mon
            while target_mon <= 0:
                target_mon += 12

        elif "年" in delta:
            delta_year = int(delta.split("年")[0])
            target_year = year - delta_year
            target_mon = month
            target_day = day

        elif delta == "":
            target_year = year
            target_mon = month
            target_day = day
        else:
            raise Exception("日期偏移量格式不正确")
        year_type_len = len(year_type)
        month_type_len = len(month_type)
        day_type_len = len(day_type)
        target_year_str = str(target_year).zfill(year_type_len)
        target_mon_str = str(target_mon).zfill(month_type_len)
        target_day_str = str(target_day).zfill(day_type_len)
        result = f"{target_year_str}{year_flag}{target_mon_str}{month_flag}{target_day_str}{day_flag}"
        print(result)
        return result

    @staticmethod
    def get_year_of_today() -> str:
        """
        获取今天的年份
        """
        from datetime import datetime
        result = datetime.today().strftime("%Y")
        return result

    @staticmethod
    def get_month_of_today() -> str:
        """
        获取今天的月份
        """
        from datetime import datetime
        result = datetime.today().strftime("%m")
        return result

    @staticmethod
    def get_day_of_today() -> str:
        """
        获取今天的日
        """
        from datetime import datetime
        result = datetime.today().strftime("%d")
        return result

    @staticmethod
    def get_date_of_today() -> str:
        """
        获取今天的日期
        """
        from datetime import datetime
        result = datetime.today().strftime("%Y-%m-%d")
        return result

    @staticmethod
    def get_easy_date(need_date_str):
        """
        快速生成报表计算中常用日期
        :param need_date_str: 需要的日期，如：今天，昨天，上月首日，上月末日，本月首日，去年今天，去年昨天，上周一，上周五，上周日，本周一
        :return: datetime.date
        """
        TODAY = datetime.datetime.today()
        result = ''
        if need_date_str == "今天":
            result = TODAY
            return result
        elif need_date_str == "昨天":
            result = TODAY - datetime.timedelta(days=1)
            return result
        elif need_date_str == "上月首日":
            result = (TODAY - datetime.timedelta(days=TODAY.day)).replace(day=1)
            return result
        elif need_date_str == "上月末日":
            result = TODAY - datetime.timedelta(days=TODAY.day)  # 上月最后一天
            return result
        elif need_date_str == '上月今天':
            result = TODAY - relativedelta(months=1)  # 上月今天
            return result
        elif need_date_str == "本月首日":
            result = TODAY.replace(day=1)
            return result
        elif need_date_str == "今年首日":
            result = TODAY.replace(month=1, day=1)
            return result
        elif need_date_str == "去年首日":
            result = TODAY.replace(year=TODAY.year - 1, month=1, day=1)
            return result
        elif need_date_str == "去年今天":
            result = TODAY - relativedelta(years=1)  # 去年今日
            return result
        elif need_date_str == "去年昨天":
            result = TODAY - relativedelta(years=1) - datetime.timedelta(days=1)  # 去年昨日
            return result
        elif need_date_str == "上周一":
            # 计算上周周一的日期
            today_weekday = TODAY.weekday()
            result = TODAY - datetime.timedelta(days=(today_weekday + 7))
            return result
        elif need_date_str == "上周五":
            # 计算上周周一的日期
            today_weekday = TODAY.weekday()
            result = TODAY - datetime.timedelta(days=(today_weekday + 3))
            return result
        elif need_date_str == "上周日":
            # 计算上周周一的日期
            today_weekday = TODAY.weekday()
            result = TODAY - datetime.timedelta(days=(today_weekday + 1))
            return result
        elif need_date_str == "本周一":
            # 计算本周周一的日期
            today_weekday = TODAY.weekday()
            result = TODAY - datetime.timedelta(days=today_weekday)
            return result
        elif need_date_str == "本月末":
            # 计算本月最后一天的日期
            next_month = TODAY.replace(day=28) + datetime.timedelta(days=4)  # 先加4天，确保下个月的日期在下个月
            result = next_month - datetime.timedelta(days=next_month.day)
            return result
        elif need_date_str == "本年末":
            # 计算本年最后一天的日期
            result = TODAY.replace(month=12, day=31)
            return result

    @staticmethod
    def gnrt_date_str(date: datetime.date, str_format: str) -> str:
        """
        生成日期字符串
        :param date: 日期
        :param str_format: 需要生成的日期格式
        :return: 日期字符串str
        """
        year_str = str(date.year)
        month_str = str(date.month)
        day_str = str(date.day)

        if str_format == "2024-01-01":
            year_str = year_str.zfill(4)
            month_str = month_str.zfill(2)
            day_str = day_str.zfill(2)
            result = f"{year_str}-{month_str}-{day_str}"

        elif str_format == "2024-1-1":
            year_str = year_str.zfill(4)
            result = f"{year_str}-{month_str}-{day_str}"

        elif str_format == "2024年01月01日":
            year_str = year_str.zfill(4)
            month_str = month_str.zfill(2)
            day_str = day_str.zfill(2)
            result = f"{year_str}年{month_str}月{day_str}日"

        elif str_format == "2024年1月1日":
            year_str = year_str.zfill(4)
            result = f"{year_str}年{month_str}月{day_str}日"

        elif str_format == "2024年 1 季":
            year_str = year_str.zfill(4)
            qurtar_str = DATE_TOOL.get_quarter_list(date.month)
            result = f"{year_str}年 {qurtar_str} 季"

        elif str_format == "2024年 1季":
            year_str = year_str.zfill(4)
            qurtar_str = DATE_TOOL.get_quarter_list(date.month)
            result = f"{year_str}年 {qurtar_str}季"

        elif str_format == "2024 年 1季":
            year_str = year_str.zfill(4)
            qurtar_str = DATE_TOOL.get_quarter_list(date.month)
            result = f"{year_str} 年 {qurtar_str}季"

        elif str_format == "2024Q1":
            year_str = year_str.zfill(4)
            qurtar_str = DATE_TOOL.get_quarter_list(date.month)
            result = f"{year_str}Q{qurtar_str}"

        elif str_format == "2024/1月":
            year_str = year_str.zfill(4)
            result = f"{year_str}/{month_str}月"

        elif str_format == "2024年01月":
            year_str = year_str.zfill(4)
            month_str = month_str.zfill(2)
            result = f"{year_str}年{month_str}月"

        elif str_format == "2024年1月":
            year_str = year_str.zfill(4)
            result = f"{year_str}年{month_str}月"

        elif str_format == "2024001":
            year_str = year_str.zfill(4)
            month_str = month_str.zfill(2)
            result = f"{year_str}0{month_str}"

        elif str_format == '2024-01':
            year_str = year_str.zfill(4)
            month_str = month_str.zfill(2)
            result = f"{year_str}-{month_str}"

        elif str_format == '202401':
            year_str = year_str.zfill(4)
            month_str = month_str.zfill(2)
            result = f"{year_str}{month_str}"

        elif str_format == "2024":
            year_str = year_str.zfill(4)
            result = f"{year_str}"

        elif str_format == "2024年":
            year_str = year_str.zfill(4)
            result = f"{year_str}年"

        elif str_format == "1月":
            result = f"{month_str}月"
        elif str_format == '01月':
            month_str = month_str.zfill(2)
            result = f"{month_str}月"

        else:
            raise Exception("输入日期格式有误，请检查")

        return result

    @staticmethod
    def gnrt_easy_date_str(need_date_str: str, str_format: str) -> str:
        """
        生成报表中常用日期字符串
        :param need_date_str: 需要的日期，如：今天，昨天，上月首日，上月末日，本月首日，去年今天，去年昨天，上周一，上周五，上周日，本周一
        :param str_format: 需要生成的日期格式
        :return: 日期字符串str
        """
        date = DATE_TOOL.get_easy_date(need_date_str)
        result = DATE_TOOL.gnrt_date_str(date, str_format)
        return result

    @staticmethod
    def get_quarter_list(month):
        """
        返回从触发所在月份的去年的季度到今年的季度
        """
        if month <= 3:
            this_quarter = 1
        elif month <= 6:
            this_quarter = 2
        elif month <= 9:
            this_quarter = 3
        else:
            this_quarter = 4
        return this_quarter

    @staticmethod
    def gnrt_easy_date_V2(reference_day:str, direction:str, offset:str, day:str, offset_num: int=0, str_format:str="2024-01-01") -> str:
        """
        更加个性化的生成日期
        :param reference_day: 参考日期，今天、昨天
        :param direction: 方向，过去、未来
        :param offset: 偏移量，几天、几月、几年
        :param day: 随需日期，当天、周初、月初、年初、季度初, 周末、月末、年末、季度末等
        :param offset_num: 偏移天数，针对X天的偏移量
        :param str_format: 输出的日期格式
        """
        if reference_day == "今天":
            ref = datetime.datetime.today()
        elif reference_day == "昨天":
            ref = datetime.datetime.today() - datetime.timedelta(days=1)
        else:
            raise Exception("参考日期格式不正确")
        if offset == "无偏移":
            result = ref
        elif offset == '一天':
            result = ref - relativedelta(days=1) if direction == '过去' else ref + relativedelta(days=1)
        elif offset == '一个月':
            result = ref - relativedelta(months=1) if direction == '过去' else ref + relativedelta(months=1)
        elif offset == '一年':
            result = ref - relativedelta(years=1) if direction == '过去' else ref + relativedelta(years=1)
        elif offset == '一季度':
            result = ref - relativedelta(months=3) if direction == '过去' else ref + relativedelta(months=3)
        elif offset == '一周':
            result = ref - relativedelta(days=7) if direction == '过去' else ref + relativedelta(days=7)
        elif offset == 'X天':
            result = ref - relativedelta(days=offset_num) if direction == '过去' else ref + relativedelta(days=offset_num)
        elif offset == 'X月':
            result = ref - relativedelta(months=offset_num) if direction == '过去' else ref + relativedelta(months=offset_num)
        elif offset == 'X年':
            result = ref - relativedelta(years=offset_num) if direction == '过去' else ref + relativedelta(years=offset_num)
        else:
            raise Exception("偏移量格式不正确")

        if day == '当天':
            result = result
        elif day == '周一':
            result = result - relativedelta(days=result.weekday())  # 周一
        elif day == '周五':
            result = result - relativedelta(days=(result.weekday() - 4))  # 周五
        elif day == '周日':
            result = result - relativedelta(days=(result.weekday() - 6))  # 周日
        elif day == '周六':
            result = result - relativedelta(days=(result.weekday() - 5))
        elif day == '月初':
            result = result.replace(day=1)
        elif day == '月末':
            next_month = result.replace(day=28) + relativedelta(days=4)    # 先加4天，确保下个月的日期在下个月
            result = next_month - relativedelta(days=next_month.day)
        elif day == '年初':
            result = result.replace(month=1, day=1)
        elif day == '年末':
            result = result.replace(month=12, day=31)
        elif day == '季度初':
            month = result.month
            if month <= 3:
                result = result.replace(month=1, day=1)
            elif month <= 6:
                result = result.replace(month=4, day=1)
            elif month <= 9:
                result = result.replace(month=7, day=1)
            else:
                result = result.replace(month=10, day=1)
        elif day == '季度末':
            month = result.month
            if month <= 3:
                next_month = result.replace(month=4, day=1)
                result = next_month - relativedelta(days=1)
            elif month <= 6:
                next_month = result.replace(month=7, day=1)
                result = next_month - relativedelta(days=1)
            elif month <= 9:
                next_month = result.replace(month=10, day=1)
                result = next_month - relativedelta(days=1)
            else:
                next_month = result.replace(month=1, day=1, year=result.year + 1)
                result = next_month - relativedelta(days=1)
        else:
            raise Exception("日期类型不正确")

        return DATE_TOOL.gnrt_date_str(result, str_format)


