
from ftplib import FTP
import socket
import os

# 链接FTP
class ConnectFTP(object):
    # 初始化
    def __init__(self, hostaddr, username, password, remotedir="/", port=21):
        self.hostaddr = hostaddr
        self.port = port
        self.username = username
        self.password = password
        self.remotedir = remotedir
        self.ftp = FTP()
        self.timeout = 120
        self.bIsDir = False
        self.path = ""
        self.encoding = "GBK"

    # 登出
    def logout(self):
        try:
            self.ftp.close()
            # self.ftp.set_debuglevel(0)
        except Exception as e:
            print(e)

    # 登录
    def login(self):
        ftp = self.ftp
        try:
            socket.setdefaulttimeout(self.timeout)
            ftp.set_pasv(True)
            # print("开始连接到: %s" %(self.hostaddr))
            ftp.encoding = self.encoding
            ftp.connect(host=self.hostaddr, port=self.port)
            print("成功连接到: %s" %(self.hostaddr))
            # print("开始登录到: %s" %(self.hostaddr))
            ftp.login(user=self.username, passwd=self.password)
            print("成功登录到: %s" %(self.hostaddr))
            # rpa.logger.info("welcome: %s" %(ftp.getwelcome()))
            self.cwdDir(self.remotedir)
        except Exception as e:
            print("登录FTP异常: %s" %e)
            self.logout()
            ftp = None
        return ftp

    # 切换目录
    def cwdDir(self, remotedir):
        try:
            self.ftp.cwd(remotedir)
            # rpa.logger.info("当前目录: %s" %self.ftp.pwd())
        except Exception as e:
            print("切换目录失败: %s" %e)

    # 创建目录
    def createFtpDir(self, remotedir):
        try:
            self.ftp.mkd(remotedir)
        except:
            print("目录已存在 %s" %remotedir)

    # 判断remote path isDir or isFile
    def isDir(self,path):
        try:
            self.ftp.cwd(path)
            self.ftp.cwd("..")
            return True
        except Exception as e:
            return False

    # 上传文件
    def upload_file(self, localfile, remotefile):
        try:
            if not os.path.isfile(localfile):
                return False
            file_handler = open(localfile, "rb")
            self.ftp.storbinary("STOR %s" %remotefile, file_handler)
            self.ftp.set_debuglevel(0)
            file_handler.close()
            # print("已传送: %s" %localfile)
            return True
        except Exception as e:
            print("文件上传异常: %s" %e)
            return False

    # 批量上传文件
    def upload_files(self, localdir="./", remotedir = "./"):
        if not os.path.isdir(localdir):
            return False
        localnames = os.listdir(localdir)
        self.ftp.cwd(remotedir)
        for item in localnames:
            localSrc = os.path.join(localdir, item)
            if os.path.isdir(localSrc):
                self.createFtpDir(item)
                self.upload_files(localdir=localSrc, remotedir=item)
            else:
                self.upload_file(localSrc, item)
        self.ftp.cwd("..")
       
    # 下载文件
    def downLoadFile(self, localFile, remoteFile):
        try:
            file_handler = open(localFile, "wb")
            # 接收服务器上文件并写入本地文件
            self.ftp.retrbinary("RETR %s" % (remoteFile), file_handler.write)
            file_handler.close()
            return True
        except Exception as e:
            print("下载文件异常: %s" %e)
            return False

    # 下载整个目录下的文件
    def downLoadFiles(self, localDir="./", remoteDir="./"):
        if not os.path.exists(localDir):
            os.makedirs(localDir)
        self.ftp.cwd(remoteDir)
        # print(self.ftp.pwd())
        remoteNames = self.ftp.nlst()
        for item in remoteNames:
            localSrc = os.path.join(localDir, item)
            if self.isDir(item):
                self.downLoadFiles(localDir=localSrc, remoteDir=item)
            else:
                self.downLoadFile(localSrc, item)
        self.ftp.cwd("..")
                
    # 删除文件
    def deleteFile(self, remoteFile):
        self.ftp.delete(remoteFile)
    
    # 备份文件
    def rename(self,fromname, toname):
        self.ftp.rename(fromname, toname)
        self.createFtpDir(fromname)



class Tool:

    def exceptionLogger(pyName,e):
        """
        打印异常日志；
        参数pyName为py文件名，无需后缀，str；
        参数e为异常，Exception
        """
        print(pyName,'.py: ','错误：',type(e),e.__traceback__.tb_lineno,'行',e)


# 调用示例
'''
from FTP_SDK import *
class PATH:
    
    REPORT_PATH = r'C:\\Users\\10190\\Downloads\\RPA_Reports'
    PROJ_NAME = 'xx'
    PROJ_PATH = REPORT_PATH+'\\'+PROJ_NAME

    TODAYSTR = '2024-01-19'
    LAST_MONTH = '2023-12'
    FTP_HOST = '9.23.23.171'
    FTP_USER = 'ftp1'
    FTP_PWD = ''
    FTP_PORT = 21
    FTP_FOLDER='/1010/xx/{}'.format(LAST_MONTH)
    FTP_ABS_PATH = '/home/{}/{}'.format(FTP_USER,FTP_FOLDER)

    FOLDER_PREST_FULL = r'{}\{}'.format(PROJ_PATH,LAST_MONTH)

def start(local_dir, remote_dir):
    """下载整个目录下的文件"""
    ftp = None
    try:
        ftp = ConnectFTP(PATH.FTP_HOST, PATH.FTP_USER, PATH.FTP_PWD)
        ftp.login()
        
        ftp.upload_files(local_dir, remote_dir)
        
        # 下载文件
        ftp.downLoadFiles(local_dir, remote_dir)
        
        ftp.deleteFile(remote_dir+'/111.docx')
        
    except Exception as e:
        Tool.exceptionLogger(pyName,e)
        raise e
    finally:
        if ftp:
            ftp.logout()

start(PATH.FOLDER_PREST_FULL,PATH.FTP_ABS_PATH)
'''