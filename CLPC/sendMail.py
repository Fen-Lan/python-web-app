
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import smtplib
import poplib
import os
from functools import wraps
import base64
from time import sleep


FROM_MAIL = "lcr@chinalife-p.com.cn"  # 发件人
SMTP_SERVER = '9.0.1'  # 163邮箱服务器
SSL_PORT = '46'  # 加密端口
USER_PWD = "ABCa"  # 163邮箱授权码

# 接受：收件人，主题，内容
# 返回：邮件发送结果
def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read())
        return encoded_string.decode("utf-8")

def tst():

    # 用于编码的图像文件路径
    image_path = "xx.png"

    # 编码图像
    encoded_image = encode_image(image_path)

    # 打印编码后的图像数据
    # print(encoded_image)

    content="""
<p>
    这是一段富文本示例，包含各种样式功能。
</p>

<p>
    <strong>这是加粗文本。</strong>
</p>

<p>
    <em>这是斜体文本。</em>
</p>

<p>
    <u>这是带有下划线的文本。</u>
</p>

<p>
    这是<span style="color: red;">红色字体</span>和<span style="color: blue;">蓝色字体</span>的例子。
</p>

<p>
    这是带有<span style="background-color: yellow;">黄色背景</span>的文本。
</p>

<p style="text-align: right;">
    这是右对齐的文本。
</p>
<table style="border-collapse: collapse;">
    <tr>
        <th style="border: 1px solid black; padding: 8px;">姓名</th>
        <th style="border: 1px solid black; padding: 8px;">年龄</th>
        <th style="border: 1px solid black; padding: 8px;">性别</th>
    </tr>
    <tr>
        <td style="border: 1px solid black; padding: 8px;">张三</td>
        <td style="border: 1px solid black; padding: 8px;">25</td>
        <td style="border: 1px solid black; padding: 8px;">男</td>
    </tr>
    <tr>
        <td style="border: 1px solid black; padding: 8px;">李四</td>
        <td style="border: 1px solid black; padding: 8px;">30</td>
        <td style="border: 1px solid black; padding: 8px;">女</td>
    </tr>
    <tr>
        <td style="border: 1px solid black; padding: 8px;">王五</td>
        <td style="border: 1px solid black; padding: 8px;">28</td>
        <td style="border: 1px solid black; padding: 8px;">男</td>
    </tr>
</table>
<img src="data:image/jpeg;base64,{}" alt="Image Description" style="width: 200px; height: 150px;">>
<a href="http://9.16.120.100/">Click here(新版门户)</a> to visit the website.

""".format(encoded_image)
    send_mail(to_mail=['lichangli@chinalife-p.com.cn'],to_content=content,content_type='html')



def retry_send(send_func):
    """装饰器：邮件推送 报错重试"""
    @wraps(send_func)
    def decorated(*args, **kwargs):
        for i in range(5):
            print("第 {} 次邮件推送尝试".format(i+1))
            result = send_func(*args, **kwargs)
            if result:
                return True
            sleep(5)
        return False
 
    return decorated


@retry_send
def send_mail(to_mail=[], cc_mail=[], to_title='【RPA服务通知】', to_content='老师好', to_att=[],content_type=''):
    # @to_mail 接收人 list  如： ['164696@qq.com','wei.zhang78@pactera.com']
    # @cc_mail 抄送人 list  如： ['164696@qq.com','wei.zhang78@pactera.com']    
    # @to_title 邮件标题 string 如：'【RPA测试】财务报表'
    # @to_content 邮件正文 string 如：'这是您本月的财务报表'
    # @to_att 邮件附件 list 如: ['D:\警报阈值.xlsx','D:\xxx.xlsx']
    # @content_type 正文类型 str 如 'html',默认空字符串，纯文本发送
    to_mail = to_mail   # 收件人邮箱，可以使列表
    title = to_title  # 邮件标题
    content = to_content  # 邮件内容

    ret = True
    FROM_MAIL = "lcjqr@chinalife-p.com.cn"  # 发件人
    TO_MAIL = to_mail          # 收件人
    CC_MAIL = cc_mail          # 收件人
    

    SMTP_SERVER = '9.0.16.46'  # 163邮箱服务器
    SSL_PORT = '465'  # 加密端口
    USER_NAME = FROM_MAIL  # 163邮箱用户名
    USER_PWD = "ABCabc123"  # 163邮箱授权码
    msg = MIMEMultipart()  # 实例化email对象
    h=Header('小财机器人', 'utf-8')
    h.append('<'+FROM_MAIL+'>', 'ascii')
    msg['from'] = h # 对应发件人邮箱昵称、发件人邮箱账号
    msg['to'] = ';'.join(TO_MAIL)  # 对应收件人邮箱昵称、收件人邮箱账号
    msg['Cc'] = ';'.join(CC_MAIL)     # 对应收件人邮箱昵称、收件人邮箱账号
    msg['subject'] = title  # 邮件的主题
    RECIVER = TO_MAIL + CC_MAIL
    if content_type=='html':
        txt = MIMEText(content,'html')
    else:
        txt = MIMEText(content)
    msg.attach(txt)
    if to_att:
        for i in to_att:
            basename = os.path.basename(i)
            # print(basename)
            att = MIMEText(open(i, "rb").read(), "base64", 'utf-8')
            att.add_header("Content-Disposition",
                           'attachment', filename=dd_b64(basename))
            att["Content-Type"] = "application/octet-stream"
            msg.attach(att)
    try:
        # 纯粹的ssl加密方式
        smtp = smtplib.SMTP_SSL(SMTP_SERVER, SSL_PORT)  # 邮件服务器地址和端口
        # smtp.set_debuglevel(1)
        smtp.ehlo()  # 用户认证
        smtp.login(USER_NAME, USER_PWD)  # 括号中对应的是发件人邮箱账号、邮箱密码
        smtp.sendmail(FROM_MAIL, RECIVER, msg.as_string())  # 收件人邮箱账号、发送邮件
        smtp.quit()  # 等同 smtp.close()  ,关闭连接
        print(">>>>>>> 邮件发送成功")
    except Exception as e:
        ret = False
        print(">>>>>>>:%s" % e)
    return ret


def dd_b64(param):
    """
    对邮件header及附件的文件名进行两次base64编码，防止outlook中乱码。
    email库源码中先对邮件进行一次base64解码然后组装邮件
    :param param: 需要防止乱码的参数
    :return:
    """
    param = '=?utf-8?b?' + base64.b64encode(param.encode('UTF-8')).decode() + '?='
    # param = '=?utf-8?b?' + base64.b64encode(param.encode('UTF-8')).decode() + '?='
    return param


def start():

    # 调用
    # 参数  1.接收人 []
    #      2. 邮件标题
    send_mail(["lichangli@chinalife-p.com.cn"], [], '【邮件服务通知】test','老师好： \n附件，请查收', [r'C:\Users\10190\Desktop\RPAdev\0先看说明.txt'])

    # send_mail(['sunxiaoli@gz.chinalife-p.com.cn'], [], '【RPA服务通知】2020年1月查勘员下载超时清单','老师好： \n附件是上月的查勘员下载超时清单，请查收', [r'E:\虚拟机共享\2020-01.xlsx'])
    

    # f = r'C:\Users\16469\Desktop\截止12月末总公司各部门预算\投资管理部.xlsx'
    # msg = '投资管理部\n现发送贵部截至1月6日部门费用预算执行情况表。\n如有问题, 请联系财务会计部申淑冰（6355）。'
    # title = '截止12月末投资管理部预算使用情况'
    # send_mail(['zhaolu@chinalife-p.com.cn'], ['zhaoyinglin@chinalife-p.com.cn'], title, msg, [f])

    # f = r'C:\Users\16469\Desktop\截止12月末总公司各部门预算\投资管理部.xlsx'
    # msg = '投资管理部\n现发送贵部截至1月6日部门费用预算执行情况表。\n如有问题, 请联系财务会计部申淑冰（6355）。'
    # title = '截止12月末投资管理部预算使用情况'
    # send_mail(['zhaolu@chinalife-p.com.cn'], ['zhaoyinglin@chinalife-p.com.cn'], title, msg, [f])




# 调用说明
# from sendMail import *
# start()