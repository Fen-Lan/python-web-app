import time
import urllib.parse
import json
from datetime import datetime

def extract_data(list_text):
    lines = list_text.splitlines()
    
    # 初始化结果列表
    data = []
    
    # 定义行数偏移
    agent_id_offset = 5  # 第6行
    trace_id_offset = 7  # 第8行
    
    # 定义步长
    step = 8  # 每8行一个周期
    
    # 遍历所有行
    for i in range(0, len(lines), step):
        if i + agent_id_offset < len(lines) and i + trace_id_offset < len(lines):
            agent_id = lines[i + agent_id_offset]
            trace_id = lines[i + trace_id_offset]
            data.append((agent_id, trace_id))
    
    return data


def calculate_coordinates(y_max,x0, y0, width, height, x1, x2, y1, y2, date_time, time_range_minutes):
    # 将 date_time 转换为时间戳
    date_time_obj = datetime.strptime(date_time, "%Y-%m-%d-%H-%M-%S")
    current_time = int(date_time_obj.timestamp() * 1000)
    offset = 15 # 添加偏移量
    # 计算 x 轴的时间范围
    start_time = current_time - time_range_minutes * 60 * 1000
    end_time = current_time
    
    # 计算 x 轴的像素位置
    start_x = int(x0 + (x1 - start_time)  / (end_time - start_time)  * width)
    end_x = int(x0 + (x2 - start_time)  / (end_time - start_time) * width)
    
    # 计算 y 轴的像素位置
    start_y = int(y0 + (1 - (y1 / y_max)) * height)
    end_y = int(y0 + (1 - (y2 / y_max)) * height)
    if start_x > end_x:
        start_x, end_x = end_x, start_x
    if start_y > end_y:
        start_y, end_y = end_y, start_y

    return start_x, start_y, end_x, end_y



def get_url_info(inf_url):
    # 给定的 URL
    drag_info_url = inf_url

    # 解析 URL
    parsed_url = urllib.parse.urlparse(drag_info_url)

    # 获取 base_url
    base_url = f"{parsed_url.scheme}://{parsed_url.netloc}"

    # 获取路径部分
    path = parsed_url.path

    # 分割路径部分
    path_parts = path.split('/')

    # 提取 service_name 和 time_interval
    service_name = path_parts[2].split('@')[0]
    time_interval = path_parts[3]
    date_time = path_parts[4]

    # 获取查询参数部分
    query = parsed_url.query

    # 解析查询参数
    query_params = urllib.parse.parse_qs(query)

    # 提取 dragInfo
    drag_info = query_params.get('dragInfo', [''])[0]

    # 解析 dragInfo 为 JSON
    drag_info_dict = json.loads(urllib.parse.unquote(drag_info))

    # 提取 x1, x2, y1, y2
    x1 = drag_info_dict.get('x1', None)
    x2 = drag_info_dict.get('x2', None)
    y1 = drag_info_dict.get('y1', None)
    y2 = drag_info_dict.get('y2', None)
    
    return base_url,service_name, time_interval,date_time, x1, x2, y1, y2

    # 打印结果
    # print(f"Service Name: {service_name}")
    # print(f"Time Interval: {time_interval}")
    # print(f"x1: {x1}")
    # print(f"x2: {x2}")
    # print(f"y1: {y1}")
    # print(f"y2: {y2}")