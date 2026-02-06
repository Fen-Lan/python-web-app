import base64
import hashlib
from Crypto.Cipher import PKCS1_v1_5
from Crypto.PublicKey import RSA
import requests
import json
import datetime
import random
import time




def parse(s):
    res = Base64Decode.base64decode(s)
    print("用户名：{}，身份证号：{}".format(res[1], res[0]))
    return res


def qyhAuthCode(user):
    """
    获取双因素认证验证码
    :param user: 用户名/身份证号
    :return: 验证码
    """
    id_card = user

    # 要发给企业微信的消息，用企业微信的公钥加密
    # 收到企业微信的响应，用自己系统的私钥解密

    # 企业微信公钥
    qy_publicKey="MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQCa0EyUR93S3gBdio1uPdoJpdmHcfFokOd+PFlp4AYtDSQeHmHcQeDDyZYxrP792xEvFCizOAB1Oidrp7By+2PPGEQxMienTZ3ygMK6DdTiYj0Ix5CWnvTkvmhcM+DsZlK7pV9VDxGxwLaxqJ+UNX/xVKrseQkpg8CLxwZjCf968QIDAQAB"
    # 测试系统
    # publicKey="MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDjolHtuMcoZaQjOGc64H2oNGd2sqoIEW4NsMs4r74+uZ0pFPUlW1VaC1ajKklYvHouDcAmO9XQ89YHBuNiPQABc3ne59MyaId0Qs9zAlLwKVnPwkcyZnZ/u0hunDR/klYRimgbJ0l9S6y3PFE4T13xXnuBmmrGt2oeRiNLiINoEQIDAQAB";
    # privateKey="MIICdwIBADANBgkqhkiG9w0BAQEFAASCAmEwggJdAgEAAoGBAOOiUe24xyhlpCM4Zzrgfag0Z3ayqggRbg2wyzivvj65nSkU9SVbVVoLVqMqSVi8ei4NwCY71dDz1gcG42I9AAFzed7n0zJoh3RCz3MCUvApWc/CRzJmdn+7SG6cNH+SVhGKaBsnSX1LrLc8UThPXfFee4Gaasa3ah5GI0uIg2gRAgMBAAECgYEA4NWjrHSUqY7y9yKvu5SeKHRSOQgxLzTgCb+0ifHzq4qz8y7TD6nNfNm0IgcTWQUYfMJyJpF1GCSvIlOoZZCwneEVU9fJ9Nh1FpMqn0rV9QuGfQIQd6ZztjO7it6J/n5TtGTWqDF18HPjwNzx8Im5aDSwxWR0+yM/3uA06gA2DQ0CQQD6t8S7F9XF8oGTv10n0sPg7VtSMAmZr1bERWMN3ZhRfHDSAc3zRprc6hYuSbZhaJMrZGlotPu/soup9jc6gK1HAkEA6G4M8gPtLrOgHi1KOr2NaV2B5uA1SiH65bTiLCA0N4sqRWOYqi9aWgU0p3Ytc6bTSi++n9OvXrRpi/lXZ+ML5wJAAhTEbUklXR9GNBPCkjINrjBKMcR0T/JEphxtVhAg04xU42lgbESJxIus43V5LhXQIuwSc+wMquqwfhitHK80wQJAJnVvNuxnZn7aU6Py0F1k9LZANE+NAcM1nKSdd+esPDSOvgSI0kAblyGdYMgxJR3JgFD+HbwNHIKFpF+RkuMCqQJBALYEe1hiFRq5pAhHhAeSYBJickaJp//eG2w7QNZ5/5iFypc1PpVyT5hCvZ7L2c1+qvuSM+E6Jj2WdMNkcZVvL/c="
    # url = "http://wxqytest.chinalife-p.com.cn:30013/dynamic_pass_web/TwoFactor/qyhAuthCode"
    # url = "http://9.1.188.42:18080/dynamic_pass_web/TwoFactor/qyhAuthCode"

    # 正式系统
    publicKey = "MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDnGeoxgSgBNWFMBH+mrxls7D9/oOz20oWZeWM5owTDEBE/g3NW6qrVP/tLP/oOaoO8ZiFNS3AxuRPAqG2KHjACMbpbi6QGzQe0zb9E3QUSMW9fc+8SnNi15yzbdG+NiVK2WIojysjrlYoHv0f1qpL6NAv5B8mtfcYAEls99NUk4wIDAQAB"
    privateKey = "MIICdwIBADANBgkqhkiG9w0BAQEFAASCAmEwggJdAgEAAoGBAOcZ6jGBKAE1YUwEf6avGWzsP3+g7PbShZl5YzmjBMMQET+Dc1bqqtU/+0s/+g5qg7xmIU1LcDG5E8CobYoeMAIxuluLpAbNB7TNv0TdBRIxb19z7xKc2LXnLNt0b42JUrZYiiPKyOuVige/R/Wqkvo0C/kHya19xgASWz301STjAgMBAAECgYBh/O78BpN7z7JtlQq7Fktlj9Zsu0M+dI0JQhr8eU6vlsR5dbcWB3Jf8T0P7xSLwEYTQAqmx3HO43aoncG0apBXz48IiHWwFQxki+Dv6Wz/K04NpX2Gbq5wulD2wz7ARUeubwuZL+zRofbPEYgnjR2X2uYlT+JgX7GnXA2YrFfUYQJBAPDAXC7TN8ka8Zk/jh3Vh11vHnsm2eRSnQhbyV+hwSdNaLkV12tLMZSM8PE2+0D7RolA9uOUql+LCm7RCYQ2/70CQQD1vRU4M+T9ZeAOdJ+xiu0IIhLgSMGxJUvmuVQ50xv6xCd+kKJZd0zdM+m71R+CWHUgn01Cwktc230fRltDLTEfAkEA3RoMjwR0S0FveKqhvkyIUQroF3oKymIEzdReEHHhjlLNRo4ElQkts5vs+9rezUL3+L2tAD1cavqqzjM0ZjSMkQJBANzEIxUb4gQYiwLRmUoKckoVOnoOQxyfUiIUq2tLkl5l7MlSrNfNStuLMNfTbvxN9eP52ZI5NTVV5oG7Vm/yYKcCQEH+d/DLvZLyXWoswj+quUSL0n5as7AtXrh9yXJtAdcBQ3BVGo6xVudIx2HSR4cXiWp3LKnMFwnaeTasBvJh0JM="
    # url = "http://10.20.102.59:8084/dynamic_pass_web/TwoFactor/qyhAuthCode"
    url = "http://9.1.208.78:32211/dynamic_pass_web/TwoFactor/qyhAuthCode"

    headers = {'Content-Type': 'application/json;charset=utf8'}

    sys_id = "rpa"
    random_code = ''.join(random.sample('1234567890', 6))
    timestamp = str(int(round(time.time() * 1000)))

    sign = hashlib.md5(bytes(id_card+random_code+timestamp, encoding="utf8")).hexdigest().upper()
    source_param = {
        "id": id_card,
        "random_code": random_code,
        "timestamp": timestamp
        }

    param = RSACrypto.encrypt(json.dumps(source_param, separators=(',',':')), qy_publicKey)

    params = {
        "sys_id": sys_id,
        "sign": sign,
        "param": param
    }

    
    res = requests.post(url, headers=headers, json=params)

    if res.status_code == 200:
        data = res.json()
        print(data)
        if data.get("errCode") == "000000":
            code = RSACrypto.decrypt(data.get("res"),privateKey)
            print("获取验证码成功：",code)
            return code
        else:
            print("获取验证码失败，错误代码：{}，错误信息：{}".format(data.get("errCode"),data.get("res")))
            return False
    else:
        print("获取验证码失败，服务器返回码：{}".format(res.status_code))
        return False


class RSACrypto(object):

    @staticmethod
    def encrypt(plain_text: str, public_key: str) -> str:
        """
        加密
        :param plain_text: 明文
        :param public_key: 公钥
        :return: 密文
        """
        plain_bytes = plain_text.encode()
        public_key = RSACrypto._format_public_key(public_key)
        key = RSA.import_key(public_key)
        crypto = PKCS1_v1_5.new(key)
        cipher_array = crypto.encrypt(plain_bytes)
        cipher_b64 = base64.b64encode(cipher_array)
        return cipher_b64.decode()

    @staticmethod
    def decrypt(cipher_text: str, private_key: str) -> str:
        """
        解密
        :param cipher_text: 密文
        :param private_key: 密钥
        :return: 明文
        """
        private_key = RSACrypto._format_private_key(private_key)
        key = RSA.import_key(private_key)
        crypto = PKCS1_v1_5.new(key)
        cipher_bytes = base64.b64decode(cipher_text)
        plain_bytes = crypto.decrypt(cipher_bytes,None)
        return plain_bytes.decode()


    @staticmethod
    def _format_public_key(public_key: str) -> str:
        """
        将公钥字符串处理成可导入的格式
        :param public_key: 公钥
        :return: str
        """
        start = '-----BEGIN RSA PUBLIC KEY-----\n'
        end = '\n-----END RSA PUBLIC KEY-----'
        key = public_key
        if not key.startswith(start):
            key = start + key
        if not key.endswith(end):
            key = key + end
        return key

    @staticmethod
    def _format_private_key(private_key: str) -> str:
        """
        将私钥字符串处理成可导入的格式
        :param private_key: 私钥
        :return: str
        """
        start = '-----BEGIN RSA PRIVATE KEY-----\n'
        end = '\n-----END RSA PRIVATE KEY-----'
        key = private_key
        if not key.startswith(start):
            key = start + key
        if not key.endswith(end):
            key = key + end
        return key


class Base64Decode(object):

    @staticmethod
    def generateDecoder():
        Base64ByteToStr = [
            'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
            'U', 'V', 'W', 'X', 'Y', 'Z', 'a', 'b', 'c', 'd',
            'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x',
            'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n',
            'y', 'z', '0', '1', '2', '3', '4', '5', '6', '7',
            '8', '9', '+', '/'
        ]
        StrToBase64Byte = bytearray(128)
        for i in range(len(StrToBase64Byte) - 1):
            StrToBase64Byte[i] = 0
        for i in range(len(Base64ByteToStr) - 1):
            StrToBase64Byte[ord(Base64ByteToStr[i])] = i
        return StrToBase64Byte

    @staticmethod
    def base64decode(encVal):
        RANGE = 0xff
        StrToBase64Byte = Base64Decode.generateDecoder()
        srcBytes = bytearray(encVal.encode("utf-8"))
        base64bytes = bytearray(len(srcBytes))
        bos = bytearray()
        for i in range(len(srcBytes)):
            ind = int(srcBytes[i])
            base64bytes[i] = StrToBase64Byte[ind]
        for i in range(0, len(base64bytes), 4):
            deBytes = bytearray(3)
            delen = 0
            for k in range(3):
                if (i + k + 1) <= len(base64bytes) - 1 and base64bytes[i + k + 1] >= 0:
                    tmp = Base64Decode.unsigned_right_shitf(int(base64bytes[i + k + 1]) & RANGE, 2 + 2 * (2 - (k + 1)))
                    deBytes[k] = ((int(base64bytes[i + k]) & RANGE) << (2 + 2 * k) & RANGE) | int(tmp)
                    delen += 1
            for k in range(delen):
                bos.append(deBytes[k])
        strs = bos.decode("utf-8")
        return strs.split("|")[0], strs.split("|")[1]


    @staticmethod
    def int_overflow(val):
        maxint = 2147483647
        if not -maxint - 1 <= val <= maxint:
            val = (val + (maxint + 1)) % (2 * (maxint + 1)) - maxint - 1
        return val


    @staticmethod
    def unsigned_right_shitf(n, i):
        # 数字小于0，则转为32位无符号uint
        if n < 0:
            n = ctypes.c_uint32(n).value
        # 正常位移位数是为正数，但是为了兼容js之类的，负数就右移变成左移好了
        if i < 0:
            return -Base64Decode.int_overflow(n << abs(i))
        return Base64Decode.int_overflow(n >> i)