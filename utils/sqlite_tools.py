"""
    文件说明： 客户端启动连接Sqlite
     创建表和插入默认配置信息是服务内部调用
     修改配置信息  对外开放
     更改系统语言 对外开放
     代码 调用start()
        传参：table, operation_type, settings, language
        table 和 operation_type为必传字段
"""

import json
import os
import sqlite3
import datetime
import sys


def WriteLog(logMsg):
    logfile = open("debuglog.txt", "a", encoding='utf-8')
    logfile.write(logMsg)
    logfile.write("\r\n")
    logfile.close()


class SqliteTools():
    def __init__(self):
        WriteLog("Init client DB")
        db_file_path = os.getcwd() + '\\client.db'
        WriteLog(
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " " + db_file_path + " " + __file__ + " " + str(
                sys._getframe().f_lineno))
        self.conn = sqlite3.connect(db_file_path, check_same_thread=False)
        self.cursor = self.conn.cursor()

    def create_setting_table(self):
        try:
            sql = 'create table %s(id int primary key,server_ip varchar(128),port varchar(20) ,pid varchar(20) , language_set varchar(20))' % self.table
            self.cursor.execute(sql)
            # 同时加入配置信息
            result = self.insert_data()
            if result.get("result") == "success":
                return {'result': 'success', 'msg': 'create table success and insert data success'}
            return {'result': 'success', 'msg': 'create table success but insert data failed'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def insert_data(self):
        sql = 'insert into %s (id, server_ip, port, pid, language_set) values (1, \'ws://127.0.0.1:8000\', \'9222\', \'3664\', \'en_US\')' % self.table
        try:
            self.cursor.execute(sql)
            self.commit()
            return {'result': 'success', 'msg': 'insert data success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def select_data(self):
        sql = "SELECT *  from %s" % self.table
        try:
            data_json = {}
            data = self.cursor.execute(sql)
            num = 0
            for item in data:
                data_json['server'] = item[1]
                data_json['port'] = item[2]
                data_json['pid'] = item[3]
                data_json['system_language'] = item[4]
                num = num + 1
            self.commit()
            return {'result': 'success', 'msg': num, 'data': data_json}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def update_language(self, language):
        try:
            sql = 'UPDATE "{}" set language_set="{}" where ID=1'.format(self.table, language)
            self.cursor.execute(sql)
            self.commit()
            return {'result': 'success', 'msg': 'update system language success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def update_settings(self, setting):
        try:
            server = json.dumps(setting['server'])
            sql = "UPDATE %s set server_ip=%s where ID=1" % (
                self.table, server)
            self.cursor.execute(sql)
            self.commit()
            return {'result': 'success', 'msg': 'update system settings success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def create_password_manage(self):
        try:
            sql = 'create table %s(id integer primary key, name varchar(128), login_name varchar(256), password varchar(256))' % self.table
            self.cursor.execute(sql)
            return {'result': 'success', 'msg': 'create passwordManage success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def create_password_node(self):
        try:
            sql = 'create table %s(id integer primary key, divId varchar(128), password_id int)' % self.table
            self.cursor.execute(sql)
            return {'result': 'success', 'msg': 'create passwordNode success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def insert_password_manage(self, data):
        try:
            sql = 'insert into %s values (null , "%s", "%s", "%s")' % (
                self.table, data['name'], data['login_name'], data['password'])
            self.cursor.execute(sql)
            self.commit()
            return {'result': 'success', 'msg': 'insert data success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def update_password_data(self, new_data):
        try:
            sql = "UPDATE %s set name='%s', password='%s', login_name='%s' where ID=%d" % (
                self.table, new_data['name'], new_data['password'], new_data['login_name'], new_data['id'])
            self.cursor.execute(sql)
            self.commit()
            return {'result': 'success', 'msg': 'update system settings success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def delete_password_manage(self, user_id):
        try:
            sql = 'delete from %s where id = "%d"' % (self.table, user_id)
            self.cursor.execute(sql)
            self.commit()
            return {'result': 'success', 'msg': 'delete data success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def select_pwd_manage(self):
        try:
            data_json = []
            sql = "SELECT *  from %s" % self.table
            data = self.cursor.execute(sql)
            for item in data:
                li = []
                li.append(item[0])
                li.append(item[1])
                li.append(item[2])
                li.append("**********")
                li.append(item[3])
                data_json.append(li)
            self.commit()
            return data_json
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def create_user_login(self):
        try:
            sql = 'create table %s(id integer primary key, username varchar(64),password varchar(256) ,mac varchar(256),version varchar(256), expiration_time TIMESTAMP, local_pwd varchar(256))' % self.table
            self.cursor.execute(sql)
            return {'result': 'success', 'msg': 'create userModel success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def insert_user_info(self, user_json):
        try:
            sql = 'insert into %s values (null, "%s", "%s", "%s", "%s", "%s", "%s")' % (
                self.table, user_json['username'], user_json['new_password'], user_json['mac'], user_json['version'],
                datetime.datetime.now(), user_json['local_pwd'])
            self.cursor.execute(sql)
            self.commit()
            return {'result': 'success', 'msg': 'insert userModel success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def update_user_data(self, user_json):
        try:
            sql = "UPDATE %s set username='%s', password='%s', mac='%s', version='%s', expiration_time='%s', local_pwd='%s' where ID=%d" % (
                self.table, user_json['username'], user_json['new_password'], user_json['mac'], user_json['version'],
                datetime.datetime.now(), user_json['local_pwd'], user_json['id'])
            self.cursor.execute(sql)
            self.commit()
            return {'result': 'success', 'msg': 'update system settings success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def select_user_data(self):
        try:
            sql = "SELECT * FROM %s" % self.table
            data = self.cursor.execute(sql)
            num = 0
            user_list = {}
            for item in data:
                num += 1
                user_list['id'] = item[0]
                user_list['username'] = item[1]
                user_list['password'] = item[2]
                user_list['mac'] = item[3]
                user_list['version'] = item[4]
                user_list['expiration_time'] = item[5]
                user_list['local_pwd'] = item[6]
            self.commit()
            return {'result': 'success', 'msg': num, "user_list": user_list}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def create_log_record(self):
        try:
            sql = 'create table %s(id integer primary key, module_name varchar(64),log_level varchar(32),username varchar(32), send_code int, log_message varchar(256), module_code varchar(64), create_time TIMESTAMP)' % self.table
            self.cursor.execute(sql)
            return {'result': 'success', 'msg': 'create LogRecord success'}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def select_log_record(self):
        try:
            data_json = []
            sql = "SELECT *  from %s limit 1" % self.table
            data = self.cursor.execute(sql)
            num = 0
            for item in data:
                li = []
                li.append(item[0])
                li.append(item[1])
                li.append(item[2])
                li.append(item[3])
                data_json.append(li)
                num = num + 1
            self.commit()
            return {'result': 'success', 'msg': num, "data": data_json}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def select_psw_node(self):
        try:
            data_json = []
            sql = "SELECT *  from %s limit 1" % self.table
            data = self.cursor.execute(sql)
            num = 0
            for item in data:
                li = []
                li.append(item[0])
                li.append(item[1])
                li.append(item[2])
                data_json.append(li)
                num = num + 1
            self.commit()
            return {'result': 'success', 'msg': num, "data": data_json}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def select_pwd_manage_limit(self):
        try:
            data_json = []
            sql = "SELECT *  from %s limit 1" % self.table
            data = self.cursor.execute(sql)
            num = 0
            for item in data:
                li = []
                li.append(item[0])
                li.append(item[1])
                li.append(item[2])
                data_json.append(li)
                num = num + 1
            self.commit()
            return {'result': 'success', 'msg': num, "data": data_json}
        except Exception as e:
            return {'result': 'failed', 'msg': str(e)}

    def commit(self):
        self.conn.commit()

    def close(self):
        self.conn.close()

    def start(self, table, operation_type, settings, language):
        """
        type, 0:创建表  1:插入数据(暂时不对外)  2: 查询数据  3: 修改数据

        :param table:
        :param operation_type:
        :return:
        """

        result = None
        self.table = table
        if operation_type == 0:
            result = self.create_setting_table()
        elif operation_type == 1:
            result = self.create_password_manage()
        elif operation_type == 2:
            result = self.select_data()
        elif operation_type == 3:
            if language != None:
                result = self.update_language(language)
            elif settings:
                result = self.update_settings(settings)
        elif operation_type == 4:
            result = self.select_pwd_manage()
        elif operation_type == 5:
            result = self.insert_password_manage(settings)
        elif operation_type == 6:
            result = self.delete_password_manage(language)
        elif operation_type == 7:
            result = self.update_password_data(settings)
        elif operation_type == 8:
            result = self.create_user_login()
        elif operation_type == 9:
            result = self.select_user_data()
        elif operation_type == 10:
            result = self.insert_user_info(settings)
        elif operation_type == 11:
            result = self.update_user_data(settings)
        elif operation_type == 12:
            result = self.create_log_record()
        elif operation_type == 13:
            result = self.create_password_node()
        elif operation_type == 14:
            result = self.select_log_record()
        elif operation_type == 15:
            result = self.select_psw_node()
        elif operation_type == 16:
            result = self.select_pwd_manage_limit()

        # 提交事务：
        # self.conn.commit()
        # self.conn.close()
        WriteLog(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " start table:" + table + " " + str(
            operation_type) + " " + str(settings) + " " + str(language) + " " + " result:" + str(
            result) + " " + __file__ + " " + str(sys._getframe().f_lineno))
        return result


g_SqlLiteTool = SqliteTools()


def get_SqlLite():
    return g_SqlLiteTool
