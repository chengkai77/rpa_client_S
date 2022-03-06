import json
import datetime
import sys
import os
import win32con
import win32api
from dateutil.relativedelta import relativedelta
from utils.UserManager import UserManger
from utils.login_encryption import login_encrypt
from utils.sqlite_tools import get_SqlLite


def setUserLogin_key(key):
    global login_user_key
    login_user_key = key


def getUserLogin_key():
    global login_user_key
    return login_user_key


# 登录加密key
def setLoginSecret_key(key):
    global secret_key
    secret_key = key


def getLoginSecret_key():
    global secret_key
    return secret_key


# 强制修改密码
async def check_pwd_strategy(websocket, pwd_length, rule, error_num, secret_key, hostname, mac):
    # 用户需要强制修改密码
    user_login_manager = UserManger()
    user_json = user_login_manager.create_user_win(pwd_length, rule, error_num)
    if not user_json.get("type"):
        sys.exit()
    # 密码加密
    old_password = login_encrypt(user_json.get("old_password"), secret_key)
    new_password = login_encrypt(user_json.get("new_password"), secret_key)
    confirm_password = login_encrypt(user_json.get("confirm_password"), secret_key)
    user_json['old_password'] = old_password
    user_json['new_password'] = new_password
    user_json['confirm_password'] = confirm_password
    user_json['mac'] = mac
    user_json['type'] = 0
    user_json['action'] = "change_password"
    await websocket.send(json.dumps(user_json))
    try:
        change_pwd_result = await websocket.recv()
        change_pwd_result = json.loads(change_pwd_result)
        if change_pwd_result.get("code") == 0:
            user_json['code'] = 0
            user_json['local_pwd'] = user_json["new_password"]
            user_json['new_password'] = change_pwd_result.get("new_password")
            return user_json
        else:
            return change_pwd_result
    except Exception as e:
        print(e)
        return {}


async def check_change_pwd_websocket(websocket, sqlite_user, login_dict, hostname, mac, operation_type):
    websocket = websocket
    local_password = sqlite_user.get("user_list").get("password", "")
    login_dict["is_check_pwd"] = operation_type
    sqlite_user = sqlite_user
    await websocket.send(json.dumps(login_dict))
    try:
        login_encryption_result = await websocket.recv()
        login_encryption_result = json.loads(login_encryption_result)
        version = login_encryption_result.get("version")
        secret_key = login_encryption_result.get("key")
        setUserLogin_key(secret_key)
        sqlite_tool = get_SqlLite()
        if operation_type == 1:
            # 过期时间和策略版本号验证
            local_version = sqlite_user.get("user_list").get("version", None)
            local_expiration_time = sqlite_user.get("user_list").get("expiration_time")
            # expiration_time 1, 6, 12三种情况，单位为月
            expiration_cycle = login_encryption_result.get("expiration_time")
            ciphertext_key = login_encryption_result.get("ciphertext_key")
            local_expiration_time = datetime.datetime.strptime(local_expiration_time, "%Y-%m-%d %H:%M:%S.%f")
            expiration_time = local_expiration_time + relativedelta(months=int(expiration_cycle))
            current_time = datetime.datetime.now()
            if ciphertext_key != "" and local_password != ciphertext_key:
                sql = "UPDATE %s set password='%s' where ID=%d" % (
                    "user_login", ciphertext_key, sqlite_user.get("user_list").get("id"))
                sqlite_tool.cursor.execute(sql)
                sqlite_tool.commit()
                return
            if local_version == str(version):
                return
            elif expiration_time < current_time:
                return
            else:
                operation_type = 2

        if login_encryption_result.get("code") == 0 and operation_type in [0, 2]:
            error_num = int(login_encryption_result.get("error_number"))
            error_num_list = list(range(error_num + 1)[1:])[::-1]
            for item in error_num_list:
                change_pwd_result = await check_pwd_strategy(websocket, login_encryption_result.get("pwd_length"),
                                                             login_encryption_result.get("rule"), item,
                                                             secret_key, hostname, mac)

                # 写入本地数据库中
                if change_pwd_result.get("code") == 0:
                    change_pwd_result['version'] = version
                    # 插入用户数据 或者 修改用户数据
                    if operation_type == 0:
                        sqlite_result = sqlite_tool.start(table="user_login", operation_type=10,
                                                          settings=change_pwd_result,
                                                          language=None)
                    else:
                        change_pwd_result['id'] = sqlite_user.get("user_list").get("id")
                        sqlite_result = sqlite_tool.start(table="user_login", operation_type=11,
                                                          settings=change_pwd_result,
                                                          language=None)
                    if sqlite_result.get("result") == "failed":
                        win32api.MessageBox(0, _("Operation failed, try again!"), "ChocLead",
                                            win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                    else:
                        win32api.MessageBox(0, _("Data has been changed, please login again"), "ChocLead",
                                            win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                    os._exit(0)
                else:
                    if item == 1:
                        win32api.MessageBox(0, _("The maximum number of changes reached!"), "ChocLead",
                                            win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                        sys.exit()
                    else:
                        win32api.MessageBox(0, change_pwd_result.get("message"), "ChocLead",
                                            win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
        else:
            win32api.MessageBox(0, _("Unable to get key file!"), "ChocLead",
                                win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
            sys.exit()
    except Exception as e:
        win32api.MessageBox(0, _("Unable to get key file!"), "ChocLead",
                            win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
        sys.exit()
