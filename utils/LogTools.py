import json
import logging
import logging.handlers
import os
import datetime
import asyncio
import time
from utils.sqlite_tools import get_SqlLite, WriteLog


def setWebsocket(new_value):
    global websocket_info
    websocket_info = new_value


def getWebsocket():
    global websocket_info
    return websocket_info


class LoggerTools():
    send_msg = {}
    time_span = 0.1
    level_relations = {
        'debug': logging.DEBUG,
        'info': logging.INFO,
        'warning': logging.WARNING,
        'error': logging.ERROR,
        'crit': logging.CRITICAL
    }

    def __init__(self, logger_name, logger_level='debug'):
        """
        :param logger_name: logger记录笔者
        :param logger_level: 日志级别
        """
        self.logger_name = logger_name
        self.logger = logging.getLogger(logger_name)
        self.logger.setLevel(self.level_relations.get(logger_level))
        # 统一输出格式
        self.logger_formatter = logging.Formatter("%(filename)s|%(lineno)d|%(name)s|%(levelname)8s|%(message)s")

        # 导入sqlite对象
        # self.sqlite_tool = SqliteTools()
        self.sqlite_tool = get_SqlLite()

    def add_file_handler(self):
        cur_path = os.path.dirname(os.path.realpath(__file__))  # log_path是存放日志的路径
        father_path = os.path.abspath(os.path.dirname(cur_path) + os.path.sep + ".") + '\\log'
        try:
            if not os.path.exists(father_path):
                os.mkdir(father_path)
        except Exception as e:
            os.mkdir(father_path)
        self.file_handler = logging.handlers.RotatingFileHandler(
            filename=father_path + '\\' + str(self.logger_name) + '.log', backupCount=7, encoding="utf-8")
        self.file_handler.setFormatter(self.logger_formatter)

    def add_handler(self):
        self.logger.addHandler(self.file_handler)

    def control_console_handler(self):
        self.ch = logging.StreamHandler()
        self.ch.setFormatter(self.logger_formatter)

    def get_client_username(self):
        result = self.sqlite_tool.start(table="user_login", operation_type=9, settings=None, language=None)
        if result.get("result") == "success":
            user_list = result.get("user_list")
            if len(user_list):
                self.username = user_list.get("username")
            else:
                self.username = None
        else:
            self.username = None

    def set_client_username(self, username):
        self.username = username

    def connect_sqlite(self, level, log_message, module_code):
        """
            module_name: 模块名称
            create_time: 创建时间
            log_level: 日志级别
            username: 用户名称
            send_code: 同步状态
            log_message: log详细信息
        """
        # 数据分割处理
        try:
            self.table = "log_record"
            sql = 'insert into %s values (null, "%s", "%s", "%s", %d, "%s", "%s", "%s")' % (
                self.table, self.logger_name, level, self.username, 0, log_message, module_code,
                datetime.datetime.now())
            self.sqlite_tool.cursor.execute(sql)
            self.sqlite_tool.conn.commit()
        except Exception as e:
            print(e)

    def start(self):
        self.control_console_handler()
        self.add_file_handler()
        self.add_handler()
        self.get_client_username()

    def info(self, msg, module_code, level='info'):
        time.sleep(self.time_span)
        self.connect_sqlite(level, msg, module_code)
        self.logger.info(msg)

    def debug(self, msg, module_code, level='debug'):
        time.sleep(self.time_span)
        self.connect_sqlite(level, msg, module_code)
        self.logger.debug(msg)

    def error(self, msg, module_code, level='error'):
        time.sleep(self.time_span)
        self.connect_sqlite(level, msg, module_code)
        self.logger.error(msg)

    def exception(self, msg, module_code, level='exception'):
        time.sleep(self.time_span)
        self.connect_sqlite(level, msg, module_code)
        self.logger.exception(msg)

    def warning(self, msg, module_code, level='warning'):
        time.sleep(self.time_span)
        self.connect_sqlite(level, msg, module_code)
        self.logger.warning(msg)


g_logs_client = LoggerTools("Client")


def get_logs_client():
    return g_logs_client


async def log_thread_tools():
    """
        线程单独进行向服务端上传log记录
        1. 获取本地sqlite中的日志记录
        2. 上传服务端并获取正确反馈
        3. 收到正确反馈,则认为上传完毕
        4. 上传完毕的日志记录在本地sqlite中删除
    """
    global websocket_info
    sqlite_obj = get_SqlLite()
    while True:
        sql = "SELECT * FROM %s" % "log_record"
        log_record_objs = sqlite_obj.cursor.execute(sql)
        num = 0
        log_record_list = []
        for item in log_record_objs:
            num += 1
            log_record_dict = {}
            log_record_dict['id'] = item[0]
            log_record_dict['module_name'] = item[1]
            log_record_dict['log_level'] = item[2]
            log_record_dict['username'] = item[3]
            log_record_dict['send_code'] = item[4]
            log_record_dict['log_message'] = item[5]
            log_record_dict['module_code'] = item[6]
            log_record_dict['create_time'] = item[7]
            log_record_list.append(log_record_dict)
        if num != 0 and log_record_list:
            send_message = {"action": "client_log", "log_record_list": log_record_list}
            WriteLog(json.dumps(send_message))
            await websocket_info.send(json.dumps(send_message))
        time.sleep(10)


def run_log_thread_tools():
    try:
        loop_log = asyncio.new_event_loop()
        asyncio.set_event_loop(loop_log)
        loops = asyncio.get_event_loop()
        loops.run_until_complete(log_thread_tools())
    except Exception as e:
        print("run_log_thread_tools--exception", e)


def delete_sqlite_log(id_list):
    sqlite_tool = get_SqlLite()
    for item in id_list:
        try:
            sql = 'delete from %s where id = "%d"' % ("log_record", item)
            sqlite_tool.cursor.execute(sql)
            sqlite_tool.commit()
            WriteLog(sql)
        except Exception as e:
            print(e)
