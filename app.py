import json
import subprocess
import uuid
import getData
from flask import Flask, request, jsonify, send_from_directory,render_template,make_response
from flask_cors import CORS
import requests
import logging
import os
from data__update import push_data
from waitress import serve
from flask import Response, stream_with_context
# 新增：导入并发相关模块
import concurrent.futures
from urllib.parse import urlparse, parse_qs, unquote

app = Flask(__name__,template_folder='templates')
# 详细配置CORS，允许所有来源、所有方法和所有头部
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)

# 配置日志
logging.basicConfig(level=logging.INFO)

# ---------------------- 原有函数保持不变 ----------------------
def find_call_stack(data,path=None):
    if path is None:
        path = []
    if isinstance(data, dict):
        if 'callStack' in data:
            print(f"找到 callStack，长度: {len(data['callStack'])}")
            return data['callStack']
        # 递归查找字典中的所有值
        for key,value in data.items():  # 修复原代码bug：data.values() -> data.items()
            result = find_call_stack(value,path + [key])
            if result:
                return result
    elif isinstance(data, list):
        for index,item in enumerate(data):
            result = find_call_stack(item, path + [str(index)])
            if result:
                return result
    print("未找到 callStack 数据")
    logging.info(f"未在路径 {' -> '.join(path)} 找到 callStack 数据")
    return None

# 提供静态文件服务
@app.route('/<path:filename>')
def serve_static(filename):
    if filename.endswith('.py') or filename in ['health','console','example.json'] or filename.endswith('.html') or filename.endswith('.css') or filename.endswith('.json'):
        os.abort(404)
    return send_from_directory(os.path.dirname(os.path.abspath(__file__)), filename)

@app.route('/console')
def console():
    # 返回 404 错误
    return "Not Found", 404

@app.before_request
def before_request():
    if request.path == '/console':
        # 返回 404 错误
        return "Not Found", 404

@app.route('/')
def landing():
    # 使用服务器端渲染生成首页
    return render_template('main.html')

# 解析 callStack 数据的函数
def parse_call_stack(call_stack_data):
    result = []
    if isinstance(call_stack_data, list):
        for item in call_stack_data:
            if not isinstance(item, list):  # 确保有足够的字段
                continue
                
            parsed_item = {
                'index': len(result),  # 索引
                'app_name': item[4] if len(item) > 4 else '',  # 第5个字段：应用名称
                'level': item[5] if len(item) > 5 else '',  # 第6个字段：层级
                'method_name': item[10] if len(item) > 10 else '',  # 第11个字段：方法名称
                'params': item[11] if len(item) > 11 else '',  # 第12个字段：入参
                'call_time': item[12] if len(item) > 12 else '',  # 第13个字段：调用时间
                'gap_time': item[13] if len(item) > 13 else '',  # 第14个字段：Gap时间
                'exec_time': item[14] if len(item) > 14 else '',  # 第15个字段：Exec时间
                'self_time': item[16] if len(item) > 16 else '',  # 第17个字段：Self时间
                'call_class': item[17] if len(item) > 17 else '',  # 第18个字段：调用类
                'call_api': item[19] if len(item) > 19 else '',  # 第20个字段：调用API
                'app_ip': item[18] if len(item) > 18 else '',  # 第19个字段：应用ip
                'agent_id': item[20] if len(item) > 20 else '',  # 第21个字段：Agent ID
                'is_error': item[22] if len(item) > 22 else '',  # 第22个字段：是否为错误请求
                'all_fields': item  # 保存所有字段
            }
            result.append(parsed_item)
    print(f"解析出 {len(result)} 个callStack项")
    return result

# 构建树状结构
def build_tree_structure(parsed_items):
    # 根据层级信息构建树状结构
    tree = []
    node_map = {}
    
    try:
        # 首先创建所有节点的映射
        for item in parsed_items:
            node_id = str(item['index'])
            # 安全获取字段值，避免None值
            call_class = item.get('call_class') or 'UnknownClass'
            method_name = item.get('method_name') or 'unknownMethod'
            gap_time = item.get('gap_time') or '0'
            exec_time = item.get('exec_time') or '0'
            self_time = item.get('self_time') or '0'
            
            node = {
                'id': node_id,
                'name': f"{call_class}:{method_name}:{gap_time}:{exec_time}:{self_time}",
                'data': item,
                'children': [],
                'level': item.get('level', '')
            }
            node_map[node_id] = node
        
        # 根据层级关系构建树
        # 层级格式假设为 "1", "1.1", "1.2", "2" 等
        for node_id, node in node_map.items():
            # 确保level是字符串类型
            level = str(node.get('level', ''))
            if not level:
                # 没有层级信息，直接加入树
                tree.append(node)
                continue
            
            # 查找父节点
            if '.' in level:
                try:
                    parent_level = '.'.join(level.split('.')[:-1])
                    parent_found = False
                    for potential_parent in node_map.values():
                        # 比较时也确保转换为字符串
                        if str(potential_parent.get('level', '')) == parent_level:
                            potential_parent['children'].append(node)
                            parent_found = True
                            break
                    if not parent_found:
                        tree.append(node)
                except Exception as e:
                    print(f"Error processing level {level}: {str(e)}")
                    tree.append(node)
            else:
                # 顶级节点
                tree.append(node)
        return tree
    except Exception as e:
        print(f"Error building tree structure: {str(e)}")
        # 返回原始节点列表作为备用
        return [{
            'id': str(item['index']),
            'name': f"{item.get('call_class', 'UnknownClass')}:{item.get('method_name', 'unknownMethod')}",
            'data': item,
            'children': []
        } for item in parsed_items]

# 验证函数
def validate_agent_transaction(agentId, transactionId):
    # 这里可以添加具体的验证逻辑，例如查询数据库或调用其他API
    valid_pairs = {
        "transactionId": "agentId"
    }
    
    if agentId in valid_pairs and transactionId in valid_pairs[agentId]:
        return True
    return False

# 修改parse_url路由，确保传递agent_id到build_tree_structure
@app.route('/api/parse-url', methods=['POST'])
def parse_url():
    try:
        print("Received parse-url request")
        data = request.get_json()
        print(f"Request data: {data}")
        
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        transactionId = data.get('transactionId')
        agentId = data.get('agentId')  # 获取 AgentId
        
        if not transactionId:
            return jsonify({'error': 'TransactionId is required'}), 400

        if not agentId:
            return jsonify({'error': 'AgentId is required'}), 400
        
        print(f"Processing TransactionId: {transactionId}, AgentId: {agentId}")
        
        # 定义三个固定的URL前缀，不在URL中直接拼接agentId
        url_prefixes = [
            "http://9.2.99.4:28080/transactionInfo.pinpoint?traceId=",
            "http://9.1.212.27:28080/transactionInfo.pinpoint?traceId=",
            "http://9.1.210.86:28082/transactionInfo.pinpoint?traceId="
        ]
        
        # 构建完整的URL，不包含agentId参数
        urls = [f"{prefix}{transactionId}&agentId={agentId}" for prefix in url_prefixes]
        
        # 打印拼接完的完整URL
        for i, url in enumerate(urls):
            print(f"拼接后的第{i+1}个URL: {url}")
        
        # 同时请求三个URL，取第一个成功的响应
        responses = []
        content_fetched = False
        content = {}  # 初始化 content 为一个空字典
        for url in urls:
            try:
                print(f"Fetching URL: {url}")
                response = requests.get(url, timeout=10)
                response.raise_for_status()
                print(f"URL request successful: {url}")
                content = response.json()
                call_stack= find_call_stack(content)
                if call_stack:
                    parsed_items = parse_call_stack(call_stack)
                    tree_structure = build_tree_structure(parsed_items)
                     # 获取 transaction_id, path, application_name
                    transaction_id = content.get('transactionId', '')
                    path = content.get('applicationName', '')
                    application_name = content.get('applicationId', '')
                    agent_id = content.get('agentId', '')
                    # 筛选 agentId 匹配的项
                    filtered_items = []
                    for item in parsed_items:
                        # 核心：为每条筛选后的项注入原始入参的transactionId和agentId（精准关联）
                        item['transactionId'] = transactionId  # 注入请求传入的transactionId
                        item['agentId'] = agentId              # 注入请求传入的agentId
                        if agent_id == agentId:
                            filtered_items.append(item)
                    print(f"过滤条件 agentId: {agentId}")
                    if filtered_items:
                        responses.append((url, filtered_items))
                        print(f"Parsed items count: {len(parsed_items)}")
                        print(f"traceid: {transaction_id}, 路径: {path}, 应用名称: {application_name}, 唯一标识符: {agent_id}")
                    else:
                        print(f"No matching items found for the provided agentId in URL: {url}")
                        continue
                    # 调用 push_data 方法将数据插入数据库
                    # push_data(transaction_id, path, application_name)
                    content_fetched = True
            except requests.RequestException as e:
                print(f"Failed to fetch URL {url}: {str(e)}")
                continue

        # 选择包含最多匹配项的响应
        if responses:
            best_response = max(responses, key=lambda x: len(x[1]))
            url, filtered_items = best_response
            tree_structure = build_tree_structure(filtered_items)
            print(f"Tree structure built successfully with {len(tree_structure)} nodes")
            return jsonify({
                'success': True,
                'parsed_items': filtered_items,
                'tree_structure': tree_structure,
                'used_agent_id': agentId,  # 返回使用的agentId供前端验证
                'Content-Type': 'application/json; charset=utf-8'
                # 'used_url': url
            })
        else:
            # 如果所有 URL 都不匹配，返回一个适当的响应
            return jsonify({
                'success': False,
                'error': 'No matching items found for the given agentId in any URL',
                'used_agent_id': agentId  # 返回使用的agentId供前端验证
            }), 400
        
    # 异常处理部分保持不变...
    except Exception as e:
        import traceback
        print(f"Unexpected error in parse-file: {str(e)}")
        print(f"Error stack trace: {traceback.format_exc()}")
        return jsonify({'error': f'处理文件时出错: {str(e)}', 'details': '请检查文件格式是否正确'}), 500

# ---------------------- 优化后的 parse_link 路由 ----------------------
# 新增：封装单条 (agentId, transactionId) 的请求逻辑
def fetch_parse_url_data(agentId, transactionId):
    """
    单条 agentId/transactionId 调用 parse-url 接口，返回解析后的parsed_items
    异常时返回空列表，避免单个请求失败影响整体
    """
    try:
        data = {
            'transactionId': transactionId,
            'agentId': agentId
        }
        # 调用本地 parse-url 接口
        response = requests.post(
            f"http://localhost:62000/api/parse-url",
            json=data,
            timeout=15  # 单个请求超时时间（可调整）
        )
        # 非200响应视为失败，返回空列表
        if response.status_code != 200:
            logging.warning(f"parse-url请求失败: agentId={agentId}, transactionId={transactionId}, 状态码={response.status_code}")
            return []
        
        parsed_data = response.json()
        parsed_items = parsed_data.get('parsed_items', [])
        
        # 为每个item补充transactionId和agentId
        # for item in parsed_items:
        #     item['transactionId'] = transactionId
        #     item['agentId'] = agentId
        
        return parsed_items
    except Exception as e:
        logging.error(f"处理agentId={agentId}, transactionId={transactionId}时出错: {str(e)}")
        return []

@app.route('/parse-link', methods=['GET'])
def parse_link():
    try:
         # 新增：生成唯一请求ID，用于请求隔离和追踪
        request_id = str(uuid.uuid4())
        logging.info(f"【解析请求-{request_id}】开始执行，接收URL: {request.args.get('url')}")

        url = request.args.get('url')
        print(f"Received parse-link request with URL: {url}")
        if not url:
            return jsonify({"error": "URL 不能为空"}), 400
        
        # 从URL中提取多组 agentId/transactionId
        data = getData.get_Id(url)
        if not data:
            return jsonify({"error": "未从URL中提取到有效的agentId/transactionId"}), 400
        
        # ---------------------- 并行请求核心逻辑 ----------------------
        # 初始化线程池（max_workers建议根据服务器性能调整，默认5）
        max_workers = min(10, len(data))  # 避免线程数超过待处理的组数
        all_parsed_items = []  # 收集所有并发请求的结果
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            # 提交所有任务到线程池，返回 future 对象列表
            future_to_params = {
                executor.submit(fetch_parse_url_data, agentId, transactionId): (agentId, transactionId)
                for agentId, transactionId in data
            }
            
            # 遍历已完成的future，收集结果
            for future in concurrent.futures.as_completed(future_to_params):
                agentId, transactionId = future_to_params[future]
                try:
                    parsed_items = future.result()
                    if parsed_items:
                        all_parsed_items.extend(parsed_items)
                        logging.info(f"成功获取 agentId={agentId}, transactionId={transactionId} 的数据，共{len(parsed_items)}条")
                except Exception as e:
                    logging.error(f"获取 agentId={agentId}, transactionId={transactionId} 数据时异常: {str(e)}")
        
        # 如果所有请求都无结果，返回提示
        if not all_parsed_items:
            return jsonify({"error": "所有agentId/transactionId组合均未查询到有效数据"}), 404
        
        # ---------------------- 原有结果汇总逻辑（复用） ----------------------
        summary = {}
        param_stats = {}

        for item in all_parsed_items:
            # 从item中获取已注入的核心标识（parse-url已注入，直接取用）
            agentId = str(item.get('agentId', 'null'))  # 原始传入的agentId
            transactionId = str(item.get('transactionId', 'null'))  # 原始传入的transactionId
            agent_id_str = str(item.get('agent_id','null'))  # 接口返回的agent_id（保留原有逻辑）
            method_name = str(item.get('method_name','null'))
            params = str(item.get('params','null'))
            
            # 优化summary：携带transactionId和原始agentId
            if agentId not in summary:
                summary[agentId] = {
                    'agentId': agentId,  # 替换原有agent_id为原始传入的agentId
                    'transactionIds': set(),  # 收集该agentId对应的所有transactionId
                    'methods': []
                }
            # 将当前transactionId加入集合（自动去重）
            summary[agentId]['transactionIds'].add(transactionId)
            # 方法信息中携带transactionId，关联到具体请求
            summary[agentId]['methods'].append({
                'transactionId': transactionId,
                'method_name': method_name,
                'params': params
            })

            # 优化param_stats：携带transactionId/agentId，便于详情追溯
            if params and params.lower() not in ['null', 'none', '']:
                if method_name == 'SQL-BindValue':
                    # 按逗号拆分 params
                    params_list = params.split(',')
                    for param in params_list:
                        param = param.strip()
                        if param not in param_stats:
                            param_stats[param] = {'count': 0, 'details': []}
                        param_stats[param]['count'] += 1
                        # 详情中携带完整标识，前端可直接展示
                        param_stats[param]['details'].append({
                            'transactionId': transactionId,
                            'agentId': agentId,
                            'method_name': method_name,
                            'agent_id': agent_id_str,
                            'params': params,
                            'call_time': item.get('call_time',''),
                            'exec_time': item.get('exec_time',''),
                            'self_time': item.get('self_time',''),
                            'gap_time': item.get('gap_time',''),
                            'call_api': item.get('call_api', ''),
                            'call_class': item.get('call_class', ''),
                            'is_error': item.get('is_error', '')
                        })
                else:
                    if params not in param_stats:
                        param_stats[params] = {'count': 0, 'details': []}
                    param_stats[params]['count'] += 1
                    # 详情中携带完整标识
                    param_stats[params]['details'].append({
                        'transactionId': transactionId,
                        'agentId': agentId,
                        'method_name': method_name,
                        'agent_id': agent_id_str,
                        'params': params,
                        'call_time': item.get('call_time',''),
                        'exec_time': item.get('exec_time',''),
                        'self_time': item.get('self_time',''),
                        'gap_time': item.get('gap_time',''),
                        'call_api': item.get('call_api', ''),
                        'call_class': item.get('call_class', ''),
                        'is_error': item.get('is_error', '')
                    })

        # 集合转列表，便于前端JSON解析（set不可序列化）
        for agentId in summary:
            summary[agentId]['transactionIds'] = list(summary[agentId]['transactionIds'])

        return jsonify({
            'request_id': request_id,
            'summary': summary,
            'param_stats': param_stats,
            'total_items': len(all_parsed_items)  # 新增：返回总数据条数，便于前端验证
        })
    except Exception as e:
        logging.error(f"Error in parse_link: {str(e)}", exc_info=True)
        return jsonify({'error': f"Internal server error: {str(e)}"}), 500

# ---------------------- 原有中间件/启动逻辑保持不变 ----------------------
@app.after_request
def add_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['Cache-Control'] = 'no-cache'
    return response

class CustomServerHeaderMiddleware():
    def __init__(self,app):
        self.app = app
    def __call__(self,environ,start_response):
        def custom_start_response(status,headers,exc_info=None):
            # 移除Server标头
            headers = [(name,value) for name,value in headers if name.lower() != 'server']
            # 添加自定义的Server头
            headers.append(('Server','MyServer'))
            return start_response(status,headers,exc_info)
        return self.app(environ,custom_start_response)
app.wsgi_app = CustomServerHeaderMiddleware(app.wsgi_app)

if __name__ == "__main__":
    serve(app,host='0.0.0.0',port=62000)