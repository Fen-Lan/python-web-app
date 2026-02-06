import datetime
from time import sleep

import requests, os, traceback, json, sys, base64


ARG_MAP = {}  # 保存启动参数
ACTIVE_CHROME_DRIVERS = [] # 保存创建的chromedriver

# 全局集合，用于存储被调用的方法信息
CALLED_METHODS = set()

# 删除历史执行日志
if os.path.exists('flow_log.log'):
    with open('flow_log.log', 'w', encoding='UTF-8') as file:
        file.truncate(0)

# 读取启动变量
if len(sys.argv) == 2:  # 暂时只接受一个启动参数，BASE64编码
    """
    # 参数样例：
    {"triggerUrls":[],
    "triggerFileInfos":[],
    "triggerParams":[
    {"triggerParamKey":"userName","triggerParamValue":"220204199100000011"},
    {"triggerParamKey":"userPwd","triggerParamValue":"Clpctysf2020@"},
    {"triggerParamKey":"loginURL","triggerParamValue":"http://9.23.28.27:8080/workbench/workbench/login.html"}],
    "triggerUsers":[]}
    """

    """
    {"triggerUrls":["http://zhida-dev:19090/upload/600ad05817522d06d9b6959bb6901cf0_20250327111249.xlsx","http://zhida-dev:19090/upload/f20b05988b5ceec75733421b8025a361_20250327160619.xlsx"],
    "triggerFileInfos":[
        {"triggerFileUrl":"http://zhida-dev:19090/upload/600ad05817522d06d9b6959bb6901cf0_20250327111249.xlsx",
         "triggerFileName":"结果.xlsx",
         "triggerFileId":855},
        {"triggerFileUrl":"http://zhida-dev:19090/upload/f20b05988b5ceec75733421b8025a361_20250327160619.xlsx",
         "triggerFileName":"aa.xlsx",
         "triggerFileId":858}],
    "triggerParams":[],
    "triggerUsers":[{"triggerUserName":"韩丹","triggerUserCode":"360424199108306737"}]}
    """

    start_args = sys.argv[1]  # 获取启动参数
    if start_args and len(start_args) > 0:
        try:
            arg_string = base64.b64decode(start_args).decode()  # base64解码为字符串
            arg_json = json.loads(arg_string)  # 转为json对象
            # print("启动参数：",arg_string)
            try:
                trigger_params = arg_json.get('triggerParams')  # 获取triggerParams对象
                for param in trigger_params:
                    key = param.get('triggerParamKey')
                    value = param.get('triggerParamValue')
                    ARG_MAP.update({key: value})
            except:
                print('未获取到参数')
                print('当前转码前参数为：', start_args)
            try:
                files_info = arg_json.get('triggerFileInfos')  # 获取文件信息
                for f_info in files_info:
                    try:
                        file_url = f_info.get('triggerFileUrl')
                        file_name = f_info.get('triggerFileName')
                        cmd = 'wget -O {} {}'.format(file_name, file_url)
                        print('开始下载文件：', file_name) # 下载任务中上传的文件
                        os.system(cmd)
                    except:
                        print('文件下载失败')
            except:
                print('未获取到文件信息')

        except:
            print('启动参数读取失败')
            print('当前转码前参数为：', start_args)



class FRAME:
    # MAIN_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

    BACKEND_ADDR_PROD = "http://9.2.255.77:8080/"  # 后端F5地址

    @staticmethod
    def clear_log():
        if os.path.exists('flow_log.log'):
            with open('flow_log.log', 'w', encoding='UTF-8') as file:
                file.truncate(0)

    @staticmethod
    def check_log(status: str, step_id: str, err=None):
        """
        记录每行拖拽的结果
        """
        import sys
        if err:
            print("**错误信息:**", err)
        with open('flow_log.log', 'a', encoding='UTF-8') as f:
            if err:
                err_msg = sys.exc_info()[1]
                print(f'status:{status}, step_id:{step_id}, err: {err_msg}', file=f)
                # print('finish !', file=f)
            else:
                print(f'status:{status}, step_id:{step_id}', file=f)
        pass

    @staticmethod
    def finish_log():
        """
        标记执行结束
        """
        with open('flow_log.log', 'a', encoding='UTF-8') as f:
            print('finish !', file=f)

    @staticmethod
    def business_start_log():
        """
        业务开始执行标记,24h任务使用
        """
        with open('flow_log.log', 'a', encoding='UTF-8') as f:
            print('business start !', file=f)
        print("##########BUSINESS START##########")


    @staticmethod
    def business_end_log():
        """
        业务结束执行标记,24h任务使用
        """
        with open('flow_log.log', 'a', encoding='UTF-8') as f:
            print('business end !', file=f)
        print("##########BUSINESS END##########")

    @staticmethod
    def get_param(param_name):
        """
        @param_name: 参数名称
        根据参数名，获取启动参数
        """
        return ARG_MAP.get(param_name, '')

    @staticmethod
    def get_env_arg(arg_name, default_value=None):
        """
        获取指定环境变量
        @arg_name: 变量名称
        @default_value: 未找到时返回的默认值
        """
        try:
            env_var = os.environ.get(arg_name)
            print("环境变量: {} : {}".format(arg_name, env_var))
            return env_var
        except:
            print("未获取到环境变量[{}]".format(arg_name))
            return default_value

    @staticmethod
    def get_yiyun_ids(default_value=[]):
        """
        获取易云触发任务的data_id列表
        """
        try:
            id_list_str = os.environ.get("RPA_DATA_IDS")
            print(f"原始id_str:{id_list_str}")
            try:
                id_list = json.loads(id_list_str)
            except:
                print("json 解析失败")
                return default_value
            print("id_list的类型：", type(id_list))
            try:
                if len(id_list) > 0:
                    for id in id_list:
                        print(f"id的类型:{type(id)}, id的值:{id}")
            except:
                pass

            return id_list
        except:
            print("未获取到易云传入的数据id列表")
            return default_value

    @staticmethod
    def report_local_task_log(app_id, exec_ip, exec_time, log_msg):
        """
        上报本地执行任务的信息日志
        """
        # todo 上报任务执行信息

    @staticmethod
    def get_task_id():
        """
        获取任务ID
        """
        try:
            env_var = os.environ.get('RPA_TASKID')
            print(env_var)
            return env_var
        except:
            raise Exception("未查询到任务ID")

    # 获取port_range范围内的可用端口
    @staticmethod
    def get_available_port():
        """
        查询可用端口
        """
        import socket
        import random

        port_ = FRAME.get_available_port_by_api()
        if not port_:
            print("通过api未获取到可用端口，直接尝试本地随机获取可用端口")
            PORT_RANGE = (29200, 29500)
            ports = list(range(PORT_RANGE[0], PORT_RANGE[1]))
            random.shuffle(ports)

            for port in ports:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                try:
                    sock.bind(('127.0.0.1', port))
                    sock.close()
                    print("本地获取可用端口：", port)
                    return port
                except:
                    pass
            else:
                raise Exception('未找到可用端口')

        else:
            return port_

    @staticmethod
    def get_available_port_by_api():
        """
        通过api获取本地可用的端口
        """
        url = "http://127.0.0.1:18091/rim/get_port"
        try:
            response = requests.get(url=url)
            rj = response.json()
            port = rj['data']
            print("api获取可用端口为：", port)
            return port
        except:
            print("触发视频录制失败")
            return None

    @staticmethod
    def trigger_rim(debugging_port):
        """
        触发录制
        """
        print('trigger_rim')
        url = 'http://127.0.0.1:18091/rim/trigger'
        headers = {'Content-Type': 'application/json'}

        task_id = FRAME.get_task_id()
        app_id = FRAME.get_env_arg('RPA_APPID', 0)
        if_cut_gif = FRAME.get_env_arg('RPA_IF_CUT_GIF', 'True')

        print(if_cut_gif)
        print(if_cut_gif == "True")
        if if_cut_gif == "True":
            if_cut_gif = True
        else:
            if_cut_gif = False
        print(if_cut_gif)
        data = {
            "task_id": int(task_id),
            "port": int(debugging_port),
            "if_cut_gif": if_cut_gif,
            "app_id": int(app_id)
        }
        print(data)
        try:
            response = requests.post(url=url, headers=headers, json=data)
            print("Trigger rim: resp=%s" % str(response.text))
        except:
            print("触发视频录制失败")


    @staticmethod
    def get_operating_system():
        """
        获取操作系统
        """
        import platform
        system = platform.system()

        if system == 'Windows':
            return 'Windows'
        elif system == 'Darwin':
            return 'MacOS'
        elif system == 'Linux':
            return 'Linux'
        else:
            return 'Unknown'

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
    def get_app_id():
        """
        获取应用id
        """
        try:
            with open('proj.conf', 'r') as f:
                app_info = f.read()
                app_id = json.loads(app_info).get('appInfo').get("appId")
                return app_id
        except:
            raise Exception("未查询到应用ID")

    @staticmethod
    def get_app_version():
        """
        获取应用版本
        """
        try:
            with open('proj.conf', 'r') as f:
                app_info = f.read()
                app_version = json.loads(app_info).get('appInfo').get("version")
                return app_version
        except:
            print("未查询到应用版本")
            return ""

    @staticmethod
    def get_user_uuid():
        """
        本地获取登录用户的uuid
        """
        URL = 'http://127.0.0.1:8080/rpa/localInfo'
        try:
            response = requests.request("GET", URL)
            res_text = response.text
            code = json.loads(res_text).get('code')
            if code == 1:
                uuid = json.loads(res_text).get('data').get("uuid")
                return uuid
            else:
                print("未查询到用户uuid")
        except:
            print("接口调用失败")
            return ""

    @staticmethod
    def get_tool_env():
        """
        获取 rpa_tool 的环境，是prod 还是 test
        """
        URL = 'http://127.0.0.1:8080/rpa/localInfo'
        try:
            response = requests.request("GET", URL)
            res_text = response.text
            code = json.loads(res_text).get('code')
            if code == 1:
                env = json.loads(res_text).get('data').get("environment")
                return env
            else:
                print("未获取到tool运行环境")
        except:
            print("接口调用失败")
            return ""

def CLEAN_DRIVER(func):
    """
    程序执行结束后，清理还存在的driver
    """
    # import psutil
    #
    # def get_chromedriver_process_count():
    #     count = 0
    #     for proc in psutil.process_iter(['name']):
    #         # 检查进程名是否为"chromedriver"
    #         if proc.info['name'] == 'chromedriver':
    #             count += 1
    #     return count

    def wrapper(*args, **kwargs):
        try:
            # Execute the decorated function
            return func(*args, **kwargs)
        finally:
            # Quit all active drivers
            exec_env = FRAME.get_env_arg("RPA_ENV", "")
            if exec_env == "prod" or exec_env == 'test':
                for driver in ACTIVE_CHROME_DRIVERS:
                    try:
                        driver.quit()
                    except Exception as e:
                        print(f"Error quitting driver: {e}")
                ACTIVE_CHROME_DRIVERS.clear()
                print("所有driver清理完成")
            else:
                print("本地执行，不清理driver")

    return wrapper


def LOCAL_EXEC(func):
    """
    本地执行时，默认发送任务执行信息
    """
    def wrapper(*args, **kwargs):
        print('开始本地执行任务')
        BACKEND_ADDR_PROD = "http://9.2.255.77:8080/"  # 后端F5地址
        BACKEND_ADDR_TEST = "http://zhida-dev:19090/"  # 后端F5地址
        API = 'api/base/exeInfoTrack'
        TASK_ID = ''
        SYSTEM = 'windows'
        # todo 区分本地和服务器，这个直接嵌入所有的应用， 因为拖拽也要用

        try:
            try:
                CONTENT = "任务执行成功"
                EXEC_IP = FRAME.get_host_ip()  # 获取执行机器ip
                APP_ID = FRAME.get_app_id()  # 获取应用id
                UUID = FRAME.get_user_uuid()  # 用户标识
                VERSION = FRAME.get_app_version()  # 获取应用版本
                LOCAL_LOG_URL = BACKEND_ADDR_PROD + API  # PROD
                LOCAL_LOG_URL = BACKEND_ADDR_TEST + API  # TEST
                # todo 接口调用记录结束日志信息，记录信息： uuid用户标识、ip、appId应用标识、taskId任务标识（如果有）、content日志内容
                payload = {"appId": int(APP_ID),
                           "ip": EXEC_IP,
                           "content": CONTENT,
                           "uuid":UUID,
                           "systemType": SYSTEM,
                           "version": VERSION}

                headers = {
                    "source":       "RpaApplication",
                    "content-type": "application/json"
                }
                print(payload)
                try:
                    response = requests.request("POST", LOCAL_LOG_URL, json=payload, headers=headers)
                    # print(response.text)
                    code = json.loads(response.text).get('code')
                    if code == 1:
                        print("日志上传成功")
                    else:
                        print("日志上传失败")

                except:
                    print(traceback.format_exc())
                    print("日志上传失败")
            except:
                print("日志记录失败")

            result = func(*args, **kwargs)
            print('本次任务执行完成')
            return result
        except SystemExit:
            CONTENT = "程序异常退出，异常信息：" + traceback.format_exc()
            print("程序异常结束")
            print(traceback.format_exc())


    return wrapper

def PRE_CHECK(func):
    """
    预检逻辑
    """

    def wrapper(*args, **kwargs):
        exec_env = FRAME.get_env_arg("RPA_ENV", "")  # 获取执行环境
        is_pre_check = FRAME.get_env_arg('RPA_IS_PRE_CHECK', 'False')  # 获取是否为预检任务
        is_run_logic = FRAME.get_env_arg('RPA_IS_EXECUTE_MAIN', 'True')  # 获取预检时是否执行业务逻辑（录入型应用不执行业务）

        if exec_env == "test":
            PRE_CHECK_URL = "http://zhida-dev:19090/api/base/rpaCallback"  # test 预检状态反馈接口
        elif exec_env == "prod":
            PRE_CHECK_URL = FRAME.BACKEND_ADDR_PROD + 'api/base/rpaCallback'  # prod
        else:
            print("本地执行")
            PRE_CHECK_URL = ""

        if func.__name__ == 'main':  # 只对main有效
            if exec_env == 'prod' or exec_env == "test":  # 检查是否为服务器执行环境
            # 服务器环境，执行预检逻辑

                if is_pre_check == 'True':  # 只在预检才调用数据
                    # 获取当前应用appId
                    app_id = FRAME.get_env_arg('RPA_APPID', 0)
                    # 调用后端预检成功接口
                    payload = {"id": int(app_id)}
                    headers = {
                        "source":       "RpaApplication",
                        "content-type": "application/json"
                    }
                    print(payload)
                    try:
                        response = requests.request("POST", PRE_CHECK_URL, json=payload, headers=headers)

                        print(response.text)
                    except:
                        print("预检结果反馈失败")

                    print("预检完成")  ## 预检结束，更新标志位，然后继续执行应用

                    if is_run_logic == 'True':
                        print('预检执行业务逻辑，执行main')
                        try:
                            return func(*args, **kwargs)
                        except SystemExit:
                            print("程序异常退出")
                        finally:
                            upload_func_usage()  # 上报组件使用情况， 录入类应用因为不执行具体的预检任务，所以埋点不会统计

                    else:
                        return

                else:
                #  非预检任务（普通任务），直接执行逻辑
                    fun_name = func.__name__
                    print("执行", fun_name)
                    try:
                        func(*args, **kwargs)
                    except:
                        print("详细错误信息：")
                        print(traceback.format_exc())
                        sys.exit(1)

            else:
                #  本地执行时，直接执行逻辑
                fun_name = func.__name__
                print("执行", fun_name)

                try:
                    return func(*args, **kwargs)
                except SystemExit:
                    print("程序异常退出")
                    print("完整报错信息：")
                    print(traceback.format_exc())

                except:
                    print("程序错误信息：")
                    print(traceback.format_exc())

        else:
            #  非main，直接执行逻辑
            print('执行业务逻辑')
            try:
                return func(*args, **kwargs)
            except SystemExit:
                print("程序错误信息：")
                print(traceback.format_exc())
            except:
                print("详细错误信息：")
                print(traceback.format_exc())
                print("程序异常退出")

    return wrapper





def FUNC_USAGE_TRACKER(cls):
    """
    类装饰器：记录类中所有方法的调用情况,并反馈到CALLED_METHODS集合中
    """
    # 遍历类的所有属性
    for name, attr in cls.__dict__.items():
        # 排除特殊方法和私有方法
        if (name.startswith('__') and name.endswith('__')) or name.startswith('_'):
            continue

        # 处理普通方法
        if callable(attr) and not isinstance(attr, (staticmethod, classmethod)):
            original_method = attr
            setattr(cls, name, static_func_tracker(cls, original_method))

        # 处理类方法
        elif isinstance(attr, classmethod):
            original_method = attr.__func__
            setattr(cls, name, classmethod(class_func_tracker(original_method)))

        # 处理静态方法
        elif isinstance(attr, staticmethod):
            original_method = attr.__func__
            setattr(cls, name, static_func_tracker(cls, original_method))
    return cls

def static_func_tracker(cls, func):
    """
    静态方法、普通方法埋点
    """
    def wrapper(*args, **kwargs):
        method_name = f"{cls.__name__}.{func.__name__}"
        CALLED_METHODS.add(method_name)
        return func(*args, **kwargs)
    return wrapper


def class_func_tracker(func):
    """
    类方法埋点
    """
    import  functools
    @functools.wraps(func)
    def wrapper(cls, *args, **kwargs):
        class_name = cls.__name__
        method_name = func.__name__
        full_method_name = f"{class_name}.{method_name}"

        CALLED_METHODS.add(full_method_name)
        return func(cls, *args, **kwargs)

    return wrapper

def upload_func_usage():
    """
    上传
    """
    BACKEND_ADDR_PROD = "http://9.2.255.77:8080/"  # 后端F5地址
    BACKEND_ADDR_TEST = "http://zhida-dev:19090/"  # 后端F5地址
    API = 'api/base/appFunctionRecord'

    LOCAL_URL = BACKEND_ADDR_TEST + API # 测试环境地址
    PROD_URL = BACKEND_ADDR_PROD + API # 生产环境地址

    app_id = FRAME.get_env_arg("RPA_APPID")
    uuid = FRAME.get_env_arg("RPA_APP_CREATOR_UUID")

    # app_id = FRAME.get_app_id() # 用于本地测试
    # uuid = FRAME.get_user_uuid() # 用于本地测试

    func_list = list(CALLED_METHODS)
    payload = {"appId":        int(app_id),
               "uuid":         uuid,
               "functionNames": func_list,
               "moduleId":     ""}

    headers = {
        "source":       "RpaApplication",
        "content-type": "application/json"
    }
    print(payload)
    try:
        response = requests.request("POST", PROD_URL, json=payload, headers=headers)  # todo 后续dev更新成 只有test环境才请求local
        code = json.loads(response.text).get('code')
        if code == 1:
            print("组件使用情况上传成功")
        else:
            print(json.loads(response.text).get('msg'))
            print("组件使用情况上传失败")

    except:
        print(traceback.format_exc())
        print("组件使用情况上传失败")


