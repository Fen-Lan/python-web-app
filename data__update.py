import pymysql
from datetime import datetime


def push_data(transaction_id,path,application_name):
    # 连接 MySQL 数据库
    connection = pymysql.connect(
        host='9.1.202.247',  # 数据库地址
        user='yunwei',  # 数据库用户名
        password='mG!2pauyS*0sC@qXl',  # 数据库
        database='monitor' , # 数据库名
        port=3309
    )

    # 将 DataFrame 中的数据插入到数据库中
    call_type = "1" #61000的工具为1
    try:
        with connection.cursor() as cursor:
            sql = "INSERT INTO application_call(transactionId, path, applicationName, call_time,call_type) VALUES (%s,%s,%s,%s,%s)"
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute(sql, (
                transaction_id,
                path,
                application_name,
                current_time,
                call_type
            ))
            connection.commit()
        print("数据插入成功")
    except Exception as e:
        print(f"插入数据时发生错误: {e}")
    finally:
        connection.close()

