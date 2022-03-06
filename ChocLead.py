from utils.LogTools import get_logs_client, setWebsocket, run_log_thread_tools, delete_sqlite_log
from utils.sqlite_tools import WriteLog

# 加载logger配置

logger = get_logs_client()
logger.start()

try:
    # 优先加载autogui库，否则跟win32会冲突
    import pyautogui
    # 右下角加载提醒
    # pyinstall打包版——必须在ChocLead.py引入win32timezone否则将报错 No module named 'win32timezone'，参考：https://blog.csdn.net/zhuangd/article/details/113929605
    import win32timezone
    import loading
    import threading
    import netifaces
    import signal

    chocLeadLoading = loading.chocLeadLoading()


    def loadingIcon():
        chocLeadLoading.createWindow()


    t = threading.Thread(target=loadingIcon)
    t.setDaemon(True)
    t.start()
    import logging
    import tkinter as tk
    import subprocess
    import base64
    import socket
    import psutil
    import web
    import time
    import os
    import json
    import sys
    import importlib
    import win32gui_struct
    import win32print
    import sap_target
    import ie_target
    import requests
    import websockets
    import shutil
    import win32com.client
    import win32api, win32gui, win32con, win32ui
    import pythoncom
    from os import stat
    import win32process
    from win32com.client import constants
    from smtplib import SMTP
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email.mime.image import MIMEImage
    from email.mime.multipart import MIMEMultipart
    from email import encoders
    from email.header import Header
    import subprocess
    from PIL import ImageGrab
    from imp import reload
    import traceback
    import pypyodbc
    import datetime
    import asyncio
    import websockets
    import json
    import uuid
    import login_window
    import socket
    import icon_transfer
    import ui_spy
    import icon_target
    import icon_coordinate_target
    import icon_position
    import mouse_target
    import screen_ratio
    import chrome_socket
    import fuzzyMatch
    import ctypes
    from tornado import ioloop
    from tornado.web import Application
    from tornado.websocket import WebSocketHandler
    import Lbt.SAP.Actions as SAP_Module
    import Lbt.EXCEL.Actions as EXCEL_Module
    import Lbt.FILE.Actions as FILE_Module
    import Lbt.ACCESS.Actions as ACCESS_Module
    import Lbt.MAIL.Actions as MAIL_Module
    import Lbt.MainFunction.Actions as MainFunction_Module
    import Lbt.MouseKeyboard.Actions as MouseKeyboard_Module
    import Lbt.SelfDevelopment.Actions as SelfDevelopment_Module
    import Lbt.WEB.Actions as WEB_Module
    import Lbt.General.Actions as General_Module
    import Lbt.AI.Actions as AI_Module
    import Lbt.HCI.Actions as HCI_Module
    import Lbt.General.encryption as encryption
    from Lbt.ChocLeadException.Error import CustomError
    import Lbt.ChocLeadException as lbtExc
    import winreg
    import ie_crawl
    import inspect
    from pynput import keyboard
    from utils.sqlite_tools import SqliteTools, get_SqlLite
    from utils.i18n_language import i18n_main
    from utils.UserManager import UserManger
    from utils.login_encryption import login_encrypt
    from dateutil.relativedelta import relativedelta
    from utils.UserPasswordManager import check_change_pwd_websocket, setLoginSecret_key
    from utils.ServerAddr_tools import ServerAddressTool
    from utils.pwd_management import setCode
    import tray

    current_path = os.path.dirname(os.path.realpath(__file__))
    sys.path.append(current_path)

    robot_status = tray.robot_status
    trayStart = "no"
    client_port = "1234"
    app = ""
    reconnect = "no"
    chromeStart = "no"
except Exception as e:
    print("init import error")
    error_info = "init import error\n" + str(e)
    logger.error(error_info, 10000)
    exit(1)


def _async_raise(tid, exctype):
    """raises the exception, performs cleanup if needed"""
    tid = ctypes.c_long(tid)
    if not inspect.isclass(exctype):
        exctype = type(exctype)
    res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(exctype))
    if res == 0:
        raise ValueError("invalid thread id")
    elif res != 1:
        # """if it returns a number greater than one, you're in trouble,
        # and you should call it again with exc=NULL to revert the effect"""
        ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)
        raise SystemError("PyThreadState_SetAsyncExc failed")


def stop_thread(thread):
    _async_raise(thread.ident, SystemExit)


# 根据进程名字获取pid
def getPidByName(Str):
    pids = psutil.process_iter()
    pidList = []
    for pid in pids:
        if pid.name() == Str:
            return pid.pid
    return None


# 根据pid获取进程
def getNameByPid(id):
    pids = psutil.process_iter()
    pidList = []
    for pid in pids:
        if pid.pid == id:
            return pid.name()
    return None


# 向谷歌插件发送结束通知
def sendEndMsg():
    client = chrome_socket.getClient()
    if client:
        chrome_socket.sendMessage({"type": "end"})


# 设置启动端口
class WebServer(web.auto_application):
    def run(self, port, *middleware):
        func = self.wsgifunc(*middleware)
        return web.httpserver.runsimple(func, ('0.0.0.0', port))


class start:
    def POST(self):
        try:
            web.header("Access-Control-Allow-Origin", "*")
            mac = get_mac_address()
            # 获取计算机名称
            hostname = socket.gethostname()
            # # 获取本机UUID
            uuid = hostname
            result = {"mac": mac, "uuid": uuid, "result": "success"}
            return json.dumps(result)
        except Exception as e:
            result = {"result": "error", "msg": str(e)}
            return json.dumps(result)

    def GET(self):
        try:
            web.header("Access-Control-Allow-Origin", "*")
            request_params = web.input()
            sqlite_tool = SqliteTools()
            username = ""
            result = sqlite_tool.start(table="user_login", operation_type=9, settings=None, language=None)
            if result.get("result") == "success" and result.get("msg") > 0:
                username = result.get("user_list").get("username")
            data = {"code": 0}
            if request_params.get("username") == username:
                sql = "SELECT *  from %s" % "password_manage"
                objs = sqlite_tool.cursor.execute(sql)
                name_li = []
                for item in objs:
                    name_li.append({"name": item[1]}, )
                data["pwd_manage_name_li"] = name_li
            return json.dumps(data)
        except Exception as e:
            logger.error(str(e), 10000)
            data = {"code": -1}
            return json.dumps(data)


def startServer(port):
    try:
        global app
        urls = ('/', 'start')
        app = web.application(urls, globals())
        os.environ["PORT"] = port
        app.run()
    except Exception as e:
        print(e)
        logger.error(str(e), 10000)


# 获取计算机mac地址
def get_mac_address():
    # mac = uuid.UUID(int=uuid.getnode()).hex[-12:]
    # return ":".join([mac[e:e + 2] for e in range(0, 11, 2)])
    mac_amount = 0
    mac_all = ''
    routingNicMacAddr_list = []
    routingNicMacAddr_dict = {}
    interface_name = netifaces.interfaces()
    for i in interface_name:
        routingNicMacAddr = netifaces.ifaddresses(i)[netifaces.AF_LINK][0]['addr']
        if routingNicMacAddr == '':
            continue
        mac_amount = mac_amount + 1
        routingNicMacAddr_list.append(routingNicMacAddr)
        routingNicMacAddr_dict[i] = routingNicMacAddr
    routingNicMacAddr_list_set = list(set(routingNicMacAddr_list))
    routingNicMacAddr_list_set.sort(reverse=True)
    for i in routingNicMacAddr_list_set:
        mac_all = mac_all + "_" + i
    mac_all = "mac_amount:%d" % mac_amount + mac_all
    mac_return = ""
    try:
        nic_key = 'SYSTEM\ControlSet001\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}'
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, nic_key)
        index = 0
        nic_information = dict()
        while True:
            try:
                sub_key = winreg.EnumKey(key, index)
                sub_key = winreg.OpenKey(key, sub_key)
                nic_name, _ = winreg.QueryValueEx(sub_key, 'DriverDesc')
                nic_id, _ = winreg.QueryValueEx(sub_key, 'NetCfgInstanceId')
                nic_information[nic_id] = nic_name
                index = index + 1
            except Exception:
                break
        for key, value in nic_information.items():
            try:
                mac_value = routingNicMacAddr_dict[key]
            except Exception:
                mac_value = None
            if "ethernet" in value.lower() and "virtual" not in value.lower():
                if mac_value:
                    # print(key, value, sep=":")
                    # print(routingNicMacAddr_dict[key])
                    mac_return = routingNicMacAddr_dict[key]
                    break
            elif ("wlan" in value.lower() or "wireless" in value.lower()) and "virtual" not in value.lower():
                if mac_value:
                    # print(key, value, sep=":")
                    # print(routingNicMacAddr_dict[key])
                    mac_return = routingNicMacAddr_dict[key]
                    break
            elif "bluetooth" in value.lower() and "virtual" not in value.lower():
                if mac_value:
                    # print(key, value, sep=":")
                    # print(routingNicMacAddr_dict[key])
                    mac_return = routingNicMacAddr_dict[key]
                    break
        if len(mac_return) >= 1:
            return mac_return
        else:
            return mac_all
    except Exception:
        return mac_all


def copytemptoerror():
    try:
        os.remove("error.py")
    except Exception:
        pass
    try:
        os.rename("temporary.py", "error.py")
    except Exception:
        pass
    try:
        os.remove("temporary.py")
    except Exception:
        pass


def switch_icon(hwnd, icon, title):
    try:
        # win32gui.SetForegroundWindow(hwnd)
        hinst = win32gui.GetModuleHandle(None)
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        hicon = win32gui.LoadImage(hinst,
                                   icon,
                                   win32con.IMAGE_ICON,
                                   0,
                                   0,
                                   icon_flags)
        message = win32gui.NIM_MODIFY
        notify_id = (hwnd,
                     0,
                     win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                     win32con.WM_USER + 20,
                     hicon,
                     title)
        win32gui.Shell_NotifyIcon(message, notify_id)
    except Exception as e:
        print(e)


# 接收谷歌插件回复的消息
class sendMessage():
    global websocket

    def __init__(self):
        self.loop = asyncio.get_event_loop()
        self.message = ""

    async def send_msg(self):
        await websocket.send(json.dumps(self.message))

    def start(self, message):
        self.message = message
        self.loop = asyncio.get_event_loop()
        self.loop.set_debug(True)
        asyncio.set_event_loop(self.loop)
        self.loop.run_until_complete(self.send_msg())


# 向服务器端认证，用户名密码通过才能退出循环
# 20210318：由于计算机的双网卡，增加uuid验证
async def auth_mac(websocket):
    """
        1. 用户验证
        2. 创建机器人
    """
    global user_dict, check_time, trayStart, client_port, loginCode
    try:
        version_path = os.path.abspath(sys.executable)
        info = win32api.GetFileVersionInfo(version_path, os.sep)
        ms = info['FileVersionMS']
        ls = info['FileVersionLS']
        version = '%d%d%d%d' % (win32api.HIWORD(ms), win32api.LOWORD(ms), win32api.HIWORD(ls), win32api.LOWORD(ls))
        custom_versions = int(version)
    except Exception as e:
        custom_versions = 1019
    check_time = 1
    mac = str(get_mac_address())
    # 获取计算机名称
    hostname = socket.gethostname()
    # 获取本机IP
    robotip = socket.gethostbyname(hostname)
    uuid = hostname
    # 验证机器人密码本地是否存在
    # sqlite_tool = SqliteTools()
    sqlite_tool = get_SqlLite()
    try:
        sqlite_user = sqlite_tool.start(table="user_login", operation_type=9, settings=None, language=None)
        # 向服务器获取密钥
        login_dict = {
            "mac": mac, "username": "", "action": "Login_Encryption", "ip": robotip
        }
        # msg == 0 表示客户端初次登录
        if sqlite_user.get("msg") == 0:
            await check_change_pwd_websocket(websocket, sqlite_user, login_dict, hostname, mac, operation_type=0)
        # 否 表示非初次登录 需要验证过期时间和策略版本号
        else:
            login_dict['username'] = sqlite_user.get("user_list").get("username")
            await check_change_pwd_websocket(websocket, sqlite_user, login_dict, hostname, mac, operation_type=1)
    except Exception as e:
        print(e)
    cred_text = {"mac": mac, "hostname": hostname, "action": "check_mac", "ip": robotip, "port": client_port,
                 "uuid": uuid, "custom_versions": custom_versions}
    await websocket.send(json.dumps(cred_text))
    response_str = ""
    try:
        response_str = await websocket.recv()
        response_str_dict = json.loads(response_str)
        response_str_dict_result = response_str_dict.get("result")
        response_str_dict_secret = response_str_dict.get("secret", "")
        response_str_dict_version = response_str_dict.get("versions", "")
    except Exception as e:
        response_str_dict = response_str
        response_str_dict_result = response_str
        response_str_dict_secret = response_str
        response_str_dict_version = response_str

    if os.path.isfile("upgrade.bat"):
        os.remove("upgrade.bat")
    if isinstance(response_str_dict, dict):
        if custom_versions < response_str_dict["versions"]:
            user_selection = win32api.MessageBox(0, _("version too low"), "ChocLead",
                                                 win32con.MB_ICONINFORMATION | win32con.MB_OKCANCEL | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
            if user_selection == 2:
                await websocket.close(reason="客户端版本不是最新版本，需要更新！")
                sys.exit()
            elif user_selection == 1:
                import struct
                phone = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                address = ""
                result = sqlite_tool.start(table="client_settings", operation_type=2, settings=None,
                                           language=None)
                if result.get("result") != "failed" and result:
                    address = result["server"]
                    if address.split("//", 1)[0] == "ws:":
                        address = address.split("//", 1)[1].split(":", 1)[0]
                    else:
                        address = address.split("//", 1)[1]
                phone.connect((address, 9000))
                cmd = "1"
                phone.send(cmd.encode('utf-8'))
                obj = phone.recv(4)
                header_size = struct.unpack('i', obj)[0]
                '''
                        header_dic = {
                        'filename': filename,
                        'file_size': os.path.getsize(filename)
                    }
                '''
                header_bytes = phone.recv(header_size)
                header_json = header_bytes.decode('utf-8')
                header_dic = json.loads(header_json)
                total_size = header_dic['file_size']
                file_path = os.path.dirname(sys.executable)
                index3 = version_path.rfind("\\")
                old_name = version_path[index3 + 1:]
                filename = "NewChocLead.exe"
                with open('%s/%s' % (file_path, filename), 'wb') as f:
                    recv_size = 0
                    window = tk.Tk()
                    window.title(_("更新ChocLead"))
                    window.geometry("630x150")
                    screenwidth = window.winfo_screenwidth()
                    screenheight = window.winfo_screenheight()
                    window.geometry('%dx%d+%d+%d' % (630, 150, (screenwidth - 630) / 2, (screenheight - 150) / 2))
                    window.wm_attributes('-topmost', 1)
                    tk.Label(window, text=_("Download progress")).place(x=20, y=30)
                    canvas = tk.Canvas(window, width=465, height=22, bg="white")
                    canvas.place(x=110, y=60)
                    fill_line = canvas.create_rectangle(1.5, 1.5, 0, 23, width=0, fill="green")
                    while recv_size < total_size:
                        line = phone.recv(1024)
                        f.write(line)
                        recv_size += len(line)
                        n = recv_size / total_size
                        n = n * 465
                        canvas.coords(fill_line, (0, 0, n, 60))
                        window.update()
                        # print('总大小：%s     已下载：%s' % (total_size, recv_size))
                        # print('\r' + '[下载进度]:%s%.2f%%' % ('>' * int(recv_size * 50 / total_size), float(recv_size / total_size * 100)), end=' ')
                    print("下载已完成!")
                if os.path.isfile("upgrade.bat"):
                    os.remove("upgrade.bat")
                b = open("upgrade.bat", 'w')
                TempList = "@echo off\n"
                TempList += "if not exist " + filename + " exit \n"
                TempList += "ping " + address + " -n 3 >nul \n"
                TempList += "taskkill /f /im " + "\"" + old_name + "\" " + "\n"
                TempList += "taskkill /f /im " + "\"" + old_name + "\" " + "\n"
                TempList += "ping " + address + " -n 3 >nul \n"
                TempList += "del " + "/f/s/q/a " + "\"" + old_name + "\" " + "\n"
                TempList += "del " + "/f/s/q/a " + "\"" + old_name + "\" " + "\n"
                TempList += "ren " + "NewChocLead.exe" + " ChocLead.exe" + "\n"
                TempList += "start " + "ChocLead.exe"
                b.write(TempList)
                b.close()
                subprocess.Popen("upgrade.bat")
                # os.rename(file_path + filename, file_path + old_name)
                await websocket.close(reason="no reson！")
                os.kill(os.getpid(), signal.SIGINT)  # 进行升级，退出此程序
                return
    if response_str_dict_result == "no":
        sqlite_tool.close()
        text_path = os.getcwd() + '\\client.db'
        with open(text_path, 'w') as f1:
            f1.seek(0)
            f1.truncate()
        win32api.MessageBox(0, "机器人server端信息异常", "ChocLead",
                            win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
        sys.exit()
    if "duplicateMac" in response_str_dict_result:
        login_window.error_remainder(_('duplicateMacInfo'))
        await websocket.close(reason="There are duplicate mac bindings in the database")
        sys.exit()
    elif "similarMac" in response_str_dict_result:
        login_window.error_remainder(_("similarMacInfo"))
        await websocket.close(reason="There are similar mac bindings in the database")
        sys.exit()
    else:
        loginCode = response_str_dict_result
        setLoginSecret_key(response_str_dict_secret)
        tray.setUser(loginCode)
        tray.setVersion(response_str_dict_version)
        if trayStart == "no":
            trayStart = "yes"
            trayMain = tray.main()
            disconnect = breakConnect()
            threads = [threading.Thread(target=trayMain.start), threading.Thread(target=disconnect.start),
                       threading.Thread(target=startServer, args=(client_port,))]
            for t in threads:
                t.start()
            win32api.MessageBox(0, _("connectInfo"), "ChocLead",
                                win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)


class breakConnect():
    global websocket, loginCode

    def __init__(self):
        self.loop = asyncio.get_event_loop()

    async def send_msg(self):
        await websocket.send(json.dumps({"action": "exit", "username": loginCode}))

    def start(self):
        robot_status = tray.getValue()
        while robot_status != "exit":
            time.sleep(1)
            robot_status = tray.getValue()
            if robot_status == "break":
                hwnd = tray.getHwnd()
                switch_icon(hwnd, "icons\\error.ico", _("robot_title"))
                tray.setValue("disconnect")
                tray.setConnect("disconnect")
                policy = asyncio.get_event_loop_policy()
                policy.set_event_loop(policy.new_event_loop())
                self.loop = asyncio.get_event_loop()
                self.loop.set_debug(True)
                asyncio.set_event_loop(self.loop)
                self.loop.run_until_complete(self.send_msg())
                win32api.MessageBox(0, _("breakInfo"), "ChocLead",
                                    win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
        else:
            global app
            app.stop()
            files = ["icons\\robot.ico", "icons\\error.ico", "icons\\connect.ico", "icons\\disconnect.ico",
                     "icons\\help.ico", "icons\\about.ico", "icons\\logout.ico", "icons\\loading.ico",
                     "icons\\stting.ico"]

            for file in files:
                if os.path.exists(file):
                    try:
                        os.remove(file)
                    except Exception:
                        pass
            connect_status = tray.getConnect()
            if connect_status == "connect":
                policy = asyncio.get_event_loop_policy()
                policy.set_event_loop(policy.new_event_loop())
                self.loop = asyncio.get_event_loop()
                self.loop.set_debug(True)
                asyncio.set_event_loop(self.loop)
                self.loop.run_until_complete(self.send_msg())


# 向服务器端认证，用户名密码通过才能退出循环
async def auth_system(websocket):
    global user_dict, check_time, trayStart, client_port, loginCode
    if check_time == 3:
        login_window.error_remainder(_("loginInfo"))
        await websocket.close(reason="login 3 times with error")
        sys.exit()
    username = user_dict['username']
    loginCode = username
    password = user_dict['password']
    # 获取计算机名称
    hostname = socket.gethostname()
    # 获取本机IP
    robotip = socket.gethostbyname(hostname)
    mac = str(get_mac_address())
    uuid = hostname
    cred_text = {"mac": mac, "hostname": hostname, "username": username, "password": password, "action": "check_user",
                 "time": check_time, "port": client_port, "ip": robotip, "uuid": uuid}
    try:
        await websocket.send(json.dumps(cred_text))
        response_str = await websocket.recv()
        if "success" in response_str:
            sys.exit()

        elif "different" in response_str:
            login_window.error_remainder(_("hostInfo"))
            await websocket.close(reason="user already has a mac address")
            sys.exit()
        elif "exit" in response_str:
            login_window.error_remainder(_("loginWarning"))
            await websocket.close(reason="login 3 times with error")
            sys.exit()
        else:
            if "2" in response_str:
                check_time = 2
            elif "1" in response_str:
                check_time = 3
            login = login_window.loginWindow(response_str)
            user_dict = login.start_window()
            username = user_dict['username']
            if username == "" or username == None:
                await websocket.close(reason="no user")
                return False
            else:
                await auth_system(websocket)
    except Exception as e:
        print(e)


# web target function
async def web_targeting(websocket, username, browser, delay):
    global exit_force
    result = {}
    pythoncom.CoInitialize()
    try:
        if browser == "ie":
            web_hwnd = win32gui.FindWindow("IEFrame", None)
            if web_hwnd:
                ie_target.setClickEvent(0)
                target = ie_target.Ie_Target()
                result = target.start(delay)
            else:
                result["status"] = "error"
                result["msg"] = "Client_Error: No ie browser"
            result['action'] = 'web_target'
            result['username'] = username
            await websocket.send(json.dumps(result))
        elif browser == "chrome":
            chrome_socket.setUserName(username)
            chrome_json = {"type": "catch", "delay": delay}
            chrome_socket.sendMessage(chrome_json)
        lbtExc.setExit(1)
        pythoncom.CoUninitialize()
    except Exception as e:
        result['msg'] = str(e)
        result['action'] = 'web_target'
        result['username'] = username
        await websocket.send(json.dumps(result))
        lbtExc.setExit(1)
        ie_target.setClickEvent(0)
        pythoncom.CoUninitialize()


# sap target function
async def sap_targeting(websocket, username):
    result = {}
    result['username'] = username
    pythoncom.CoInitialize()
    try:
        target = sap_target.targeting()
        sap_id = target.main()
        if "!" in sap_id:
            result['result'] = "error"
            result['msg'] = sap_id
        else:
            sap_ids = sap_id.split("/")
            id = ""
            for i in range(len(sap_ids)):
                if i > 3:
                    if i == 4:
                        id = sap_ids[i]
                    else:
                        id = id + "/" + sap_ids[i]
            sap_id = id
            result['result'] = "success"
            result['sap_id'] = "\"" + sap_id + "\""
    except Exception as e:
        result['result'] = "error"
        result['msg'] = str(e)
    result['action'] = 'sap_target'
    await websocket.send(json.dumps(result))
    lbtExc.setExit(1)
    pythoncom.CoUninitialize()


# chrome recording
async def chrome_recording(websocket, username, record, auto):
    result = {}
    result["auto"] = auto
    try:
        chrome_socket.setUserName(username)
        chrome_socket.setAuto(auto)
        if record == 1:
            chrome_json = {"type": "record"}
            chrome_socket.sendMessage(chrome_json)
        else:
            chrome_json = {"type": "endRecord"}
            chrome_socket.sendMessage(chrome_json)
    except Exception as e:
        result['msg'] = str(e)
        result['action'] = 'chrome_record'
        result['username'] = username
        await websocket.send(json.dumps(result))


# web crawl
async def web_crawl(websocket, username, browser):
    result = {}
    try:
        if browser == "chrome":
            chrome_socket.setUserName(username)
            chrome_json = {"type": "crawlConfig"}
            chrome_socket.sendMessage(chrome_json)
        elif browser == "ie":
            pythoncom.CoInitialize()
            ieTarget = ie_crawl.Ie_Crawl()
            result = json.loads(ieTarget.start())
            result["action"] = "web_crawl"
            result['username'] = username
            await websocket.send(json.dumps(result))
            lbtExc.setExit(1)
            pythoncom.CoUninitialize()
    except Exception as e:
        if browser == "ie":
            lbtExc.setExit(1)
            pythoncom.CoUninitialize()
        result['msg'] = str(e)
        result['action'] = 'web_crawl'
        result['username'] = username
        await websocket.send(json.dumps(result))


# sap recording
async def sap_recording(websocket, username, record, auto):
    global record_path, session
    result = {}
    result['username'] = username
    result['auto'] = auto
    try:
        connection_num = 0
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection_num = application.Children.Count
        except Exception:
            result['result'] = "error"
            result['msg'] = "No sap connection!"
            result['action'] = 'sap_record'
            await websocket.send(json.dumps(result))
        if connection_num == 0:
            result['result'] = "error"
            result['msg'] = "No sap connection!"
            result['action'] = 'sap_record'
            await websocket.send(json.dumps(result))
        else:
            connection = application.Children(0)
            flag = 0
            wait = 0
            while flag == 0 and wait < 10:
                try:
                    session = connection.Children(0)
                    flag = 1
                except Exception:
                    wait += 1
                    time.sleep(0.5)
            if record == 1:
                now = str(int(time.time()))
                record_result = session.Record
                if record_result == True:
                    session.record = 0
                session.recordFile = now + ".vbs"
                session.record = 1
                record_path = session.RecordFile
                logger.info(str(username) + " sap录制文件地址: " + record_path, 10000)
                if os.path.exists(record_path) == False:
                    with open(record_path, "w"):
                        pass
                sap_window = win32gui.FindWindow("SAP_FRONTEND_SESSION", None)
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys('^')
                win32gui.SetForegroundWindow(sap_window)
            else:
                if record_path == "":
                    result['result'] = "error"
                    result['msg'] = "SAP recording is not enabled!"
                    result['action'] = 'sap_record'
                    await websocket.send(json.dumps(result))
                result['record_file'] = record_path
                code = ""
                with open(record_path, mode='r', encoding='utf-8', errors='ignore') as f:
                    data = f.readlines()
                    for i in range(28, len(data)):
                        line = data[i]
                        if line != '\n' and line != '' and len(line) != 2:
                            code = code + line
                session.record = 0
                try:
                    os.remove(record_path)
                except Exception:
                    pass
                result['action'] = 'sap_record'
                result['record_code'] = code
                result['result'] = "success"
                await websocket.send(json.dumps(result))
    except Exception as e:
        try:
            os.remove(record_path)
        except Exception:
            pass
        result['result'] = "error"
        result['msg'] = str(e)
        result['action'] = 'sap_record'
        await websocket.send(json.dumps(result))


# handle target function
async def handle_targeting(websocket, username):
    pythoncom.CoInitialize()
    platform = win32gui.GetForegroundWindow()
    try:
        spy = ui_spy.uiSpy()
        result = spy.main()
    except Exception as e:
        result = {}
        result['msg'] = str(e).replace('\n', '<br>')
        result['result'] = "error"
    # 检查一开始获取的句柄号是否为choclead句柄号
    hwnd = win32gui.FindWindow("Chrome_WidgetWin_1", "ChocLead - Google Chrome")
    if hwnd != platform and hwnd:
        platform = hwnd
    returnRPA(platform)
    result['action'] = 'handle_target'
    result['username'] = username
    await websocket.send(json.dumps(result))
    lbtExc.setExit(1)
    pythoncom.CoUninitialize()


# icon target function
async def icon_targeting(websocket, username, position, delay):
    pythoncom.CoInitialize()
    platform = win32gui.GetForegroundWindow()
    try:
        result = icon_target.start(position, delay)
    except Exception as e:
        result = {}
        result['msg'] = str(e).replace('\n', '<br>')
        result['result'] = "error"
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('^')
        win32gui.SetForegroundWindow(platform)
    except Exception:
        lbtExc.setExit(1)
        pythoncom.CoUninitialize()
        pass
    result['action'] = 'icon_target'
    result['username'] = username
    await websocket.send(json.dumps(result))
    lbtExc.setExit(1)
    pythoncom.CoUninitialize()


# icon target function
async def icon_coordinate_targeting(websocket, username, delay):
    pythoncom.CoInitialize()
    platform = win32gui.GetForegroundWindow()
    try:
        result = icon_coordinate_target.start(delay)
    except Exception as e:
        result = {}
        result['msg'] = str(e).replace('\n', '<br>')
        result['result'] = "error"
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('^')
        win32gui.SetForegroundWindow(platform)
    except Exception:
        lbtExc.setExit(1)
        pythoncom.CoUninitialize()
        pass
    result['action'] = 'icon_coordinate_target'
    result['username'] = username
    await websocket.send(json.dumps(result))
    lbtExc.setExit(1)
    pythoncom.CoUninitialize()


# mouse target function
async def mouse_targeting(websocket, username, method):
    pythoncom.CoInitialize()
    platform = win32gui.GetForegroundWindow()
    ratio = screen_ratio.get_ratio()
    try:
        """
        sun13 与前台配合起来，我必须传文本 测试
        """
        method = method.replace("\"", "")
        target = mouse_target.targeting()
        result = target.main(method)
        result['ratio'] = "\"" + str(int(ratio * 100)) + "%" + "\""
    except Exception as e:
        result = {}
        result['msg'] = str(e).replace('\n', '<br>')
        result['result'] = "error"
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('^')
        win32gui.SetForegroundWindow(platform)
    except Exception:
        pass
    result['action'] = 'mouse_target'
    result['username'] = username
    await websocket.send(json.dumps(result))
    lbtExc.setExit(1)
    pythoncom.CoUninitialize()


async def web_dragging_mouse_targeting(websocket, username, method):
    pythoncom.CoInitialize()
    platform = win32gui.GetForegroundWindow()
    ratio = screen_ratio.get_ratio()
    try:
        """
        sun13 与前台配合起来，我必须传文本 测试
        """
        method = method.replace("\"", "")
        target = mouse_target.targeting()
        result = target.main(method)
        result['ratio'] = "\"" + str(int(ratio * 100)) + "%" + "\""
    except Exception as e:
        result = {}
        result['msg'] = str(e).replace('\n', '<br>')
        result['result'] = "error"
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('^')
        win32gui.SetForegroundWindow(platform)
    except Exception:
        pass
    result['action'] = 'web_dragging_mouse_target'
    result['username'] = username
    await websocket.send(json.dumps(result))
    lbtExc.setExit(1)
    pythoncom.CoUninitialize()


async def runtask_temporary(websocket, recv_str):
    print('enter', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    # global chrome, chromedriver
    try:
        os.remove("temporary.py")
    except Exception:
        pass
    result = {}
    code = ""
    try:
        codes = recv_str['code']
        if isinstance(codes, str):
            codes = codes.split("\n")
            codes_list = []
            for i in range(len(codes) - 1):
                code_json = {}
                code_json["code"] = codes[i] + "\n"
                codes_list.append(code_json)
            codes = codes_list
        for single in codes:
            code = code + single["code"]
        username = recv_str['username']
        taskname = recv_str['taskname']
    except Exception as e:
        code = ''
        logger.error(username + "get robot code failed!" + str(recv_str), 10000)
    filePath = "temporary.py"
    try:
        os.remove(filePath)
    except Exception:
        time.sleep(0.2)
    with open(filePath, mode='w', encoding='utf-8') as ff:
        ff.write(code)
    print('write', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    wait = 0
    sttime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    time.sleep(2)
    while wait < 5:
        if os.path.exists(filePath):
            wait = 5
            lbtExc.setId("start", "", "")
            # 判断是否有谷歌流程，如果有，运行成功或者失败都通知插件结束
            informChrome = False
            if 'browser="chrome"' in code:
                informChrome = True
            # global keyboardListener
            try:
                exec(code)
                run_result = "success"
            except Exception as err:
                divId, groupId, action = lbtExc.getId()
                error_list = traceback.format_exception(*sys.exc_info(), limit=None, chain=True)
                # 计算 报错line的number数
                for error in error_list:
                    if "<string>" in error:
                        line = error.split(",")
                        # 提取机器人执行的报错行数
                        line_number = int(line[1].split(" ")[2])
                        print("temporary.py error_line_number:%s" % line_number)
                        break
                line_code = codes[line_number - 1]["code"]
                try:
                    line = codes[line_number - 2]["code"].split("# node_no: ")[1].split(" ")[0]
                except:
                    line = ""
                code_error = {}
                code_error["code"] = str(line_code)
                code_error["conso"] = str(err)
                code_error["line"] = str(line)
                result["conso"] = json.dumps(code_error)
                if not action:
                    result["msg"] = "Generate Codes Error!"
                else:
                    result["msg"] = str(action) + " Failed!"
                run_result = "failed"
                result["result"] = run_result
            if informChrome:
                sendEndMsg()
            if run_result == "success":
                result["result"] = run_result
                continue
            else:
                conso = result["conso"].replace("\n", "<br>")
                result["conso"] = conso
                conso_list = conso.split("<br>")
                logger.error("Some Process Error!" + str(result["msg"]), 10000)
                logger.error("Error:" + str(result["conso"]), 10000)
                copytemptoerror()
        else:
            time.sleep(1)
            wait += 1
    print('run', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    result['end_time'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    result['action'] = 'run_result'  # 运行结果返回到服务器
    result['username'] = username
    result['taskname'] = taskname
    result['start_time'] = sttime
    hostname = socket.gethostname()
    ip = socket.gethostbyname(hostname)
    result['runtaskoi'] = ip
    # result, start_time, end_time, msg
    await websocket.send(json.dumps(result))


async def runCycletask_temporary(websocket, recv_str):
    logger.info("runCycletask_temporary: enter %s" % datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 10000)
    try:
        os.remove("temporary.py")
    except Exception:
        pass
    result = {}
    code = ""
    try:
        codes = recv_str['code']
        if isinstance(codes, str):
            codes = codes.split("\n")
            codes_list = []
            for i in range(len(codes) - 1):
                code_json = {}
                code_json["code"] = codes[i] + "\n"
                codes_list.append(code_json)
            codes = codes_list
        for single in codes:
            code = code + single["code"]
        username = recv_str['username']
        taskname = recv_str['taskname']
    except Exception as e:
        code = ''
        logger.error(username + "get robot code failed!" + str(recv_str), 10000)
    filePath = "temporary.py"
    try:
        os.remove(filePath)
    except Exception:
        time.sleep(0.2)
    with open(filePath, mode='w', encoding='utf-8') as ff:
        ff.write(code)

    logger.info("runCycletask_temporary: write %s" % datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 10000)
    wait = 0
    sttime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    time.sleep(2)
    while wait < 5:
        if os.path.exists(filePath):
            wait = 5
            lbtExc.setId("start", "", "")
            # 判断是否有谷歌流程，如果有，运行成功或者失败都通知插件结束
            informChrome = False
            if 'browser="chrome"' in code:
                informChrome = True
            # global keyboardListener
            try:
                exec(code)
                run_result = "success"
            except Exception as err:
                divId, groupId, action = lbtExc.getId()
                error_list = traceback.format_exception(*sys.exc_info(), limit=None, chain=True)
                # 计算 报错line的number数
                for error in error_list:
                    if "<string>" in error:
                        line = error.split(",")
                        # 提取机器人执行的报错行数
                        line_number = int(line[1].split(" ")[2])
                        print("temporary.py error_line_number:%s" % line_number)
                        break
                line_code = codes[line_number - 1]["code"]
                try:
                    line = codes[line_number - 2]["code"].split("# node_no: ")[1].split(" ")[0]
                except:
                    line = ""
                code_error = {}
                code_error["code"] = str(line_code)
                code_error["conso"] = str(err)
                code_error["line"] = str(line)
                result["conso"] = json.dumps(code_error)
                if not action:
                    result["msg"] = "Generate Codes Error!"
                else:
                    result["msg"] = str(action) + " Failed!"
                run_result = "failed"
                result["result"] = run_result
            if informChrome:
                sendEndMsg()
            if run_result == "success":
                result["result"] = run_result
                continue
            else:
                conso = result["conso"].replace("\n", "<br>")
                result["conso"] = conso
                conso_list = conso.split("<br>")
                logger.error("Some Process Error!" + str(result["msg"]), 10000)
                logger.error("Error:" + str(result["conso"]), 10000)
                copytemptoerror()
        else:
            time.sleep(1)
            wait += 1
    logger.info("runCycletask_temporary: run %s" % datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 10000)
    result['end_time'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    result['action'] = 'cycle_result'  # 运行结果返回到服务器
    result['username'] = username
    result['taskname'] = taskname
    result['start_time'] = sttime
    hostname = socket.gethostname()
    ip = socket.gethostbyname(hostname)
    result['runtaskoi'] = ip
    logger.info(result, 10000)
    # result, start_time, end_time, msg
    await websocket.send(json.dumps(result))


# 键盘点击事件
# 监听按压
def on_press(key):
    global on_press_status, first_press_start
    on_press_start = time.time()
    if format(key) not in on_press_status.keys():
        first_press_start[format(key)] = on_press_start
    on_press_status[format(key)] = True


# 监听释放
def on_release(key):
    global keyboardListener, exitStatus, exit_force, first_press_start, on_press_status
    try:
        time_interval = time.time() - first_press_start[format(key)]
    except Exception:
        # print("遇到错误:time_interval被置为0")
        time_interval = 0
    # ###按下一个键盘到释放的时间（监控pyautogui按键，time_interval < 0.01）
    # print("the interval between first press and release:",time_interval)
    # print("已经释放:%s" %key)
    if format(key) == "Key.esc" and time_interval > 0.01:
        lbtExc.setExit(1)
        keyboardListener.stop()
        exit_force = 1
        ui_spy.last_java_hwnd = None
        ui_spy.java_hwnd = None
        print("Key.esc has been input")
        return False
    elif format(key) == "Key.pause" and time_interval > 0.01:
        lbtExc.setExit(1)
        exit_force = 0
        keyboardListener.stop()
        print("Key.pause has been input")
        return False
    try:
        del on_press_status[format(key)]
    except Exception:
        pass
    try:
        del first_press_start[format(key)]
    except Exception:
        pass


# 监听esc退出事件
def listen():
    global keyboardListener, first_press_start, on_press_status
    # 按键第一次press时间
    first_press_start = {}
    # 按键press状态
    on_press_status = {}
    with keyboard.Listener(on_press=on_press, on_release=on_release) as keyboardListener:
        keyboardListener.join()


def returnRPA(hwnd):
    try:
        # 若最小化，则将其最大化显示
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('^')
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


async def run_temporary(websocket, recv_str):
    """
    sun13 2021-01-04
    重写立即执行代码的功能
    execute temporary.py
    """
    global loginCode
    pythoncom.CoInitialize()
    platform = win32gui.GetForegroundWindow()
    logger.info('process enter %s' % datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 10000)
    file_path = "temporary.py"
    try:
        os.remove(file_path)
    except Exception as e:
        logger.error("os.remove(temporary.py) " + str(e), 10000)
    result = {}
    code = ""
    if "choclead_secret_key" in recv_str:
        General_Module.set_c_k(recv_str['choclead_secret_key'])
    try:
        log_lines_num = recv_str['log_lines_num']
    except:
        log_lines_num = ""
    try:
        log_run_time1 = recv_str['log_run_time1']
        log_run_time = recv_str['log_run_time']
    except:
        log_run_time1 = ""
        log_run_time = ""
    result["log_lines_num"] = log_lines_num
    result["log_run_time1"] = log_run_time1
    result["log_run_time"] = log_run_time
    try:
        codes = recv_str['code']
        if isinstance(codes, str):
            codes = codes.split("\n")
            codes_list = []
            for i in range(len(codes) - 1):
                code_json = {}
                code_json["code"] = codes[i] + "\n"
                codes_list.append(code_json)
            codes = codes_list
        for single in codes:
            code = code + single["code"]
        username = recv_str['username']
    except Exception as e:
        logger.exception(str(e), 10000)
        code = ''
        logger.error(username + "get robot code failed!" + str(recv_str), 10000)
        result["conso"] = str(e)
        result["msg"] = "Get Code Failed!"
        logger.error("Get Code Failed! " + str(e), 10000)
        copytemptoerror()
        result['action'] = 'run'
        result['username'] = loginCode
        result['commander'] = ""
        result['booked_time'] = ""
        result['uuid'] = uuid
        returnRPA(platform)
        await websocket.send(json.dumps(result))
    if "commander" in recv_str:
        commander = recv_str['commander']
    else:
        commander = ""
    if "booked_time" in recv_str:
        booked_time = recv_str['booked_time']
    else:
        booked_time = ""
    # time.sleep(0.2)
    # TODO 考虑之后不生成文件的话，直接注释掉
    with open(file_path, mode='w', encoding='utf-8') as ff:
        ff.write(code)
    logger.info('process write %s' % datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 10000)
    wait = 0
    # time.sleep(2)
    while wait < 5:
        if os.path.exists(file_path):
            wait = 5
            # 运行之前清空全局变量divId,groupId,action值 传"" 无法情况id，因为def setId(val1,val2,val3):
            #     if val1:
            lbtExc.setId("start", "", "")
            # 判断是否有谷歌流程，如果有，运行成功或者失败都通知插件结束
            informChrome = False
            if 'browser="chrome"' in code:
                informChrome = True
            # global keyboardListener
            try:
                # 清空打印数组
                lbtExc.setListBlank()
                exec(code)
                # #设置默认esc退出状态
                # lbtExc.setExit(0)
                # # 退出监听线程
                # wait_seconds = 0
                # while wait_seconds < 3:
                #     try:
                #         if keyboardListener:
                #             keyboardListener.stop()
                #             wait_seconds = 3
                #     except Exception:
                #         wait_seconds+=1
                #         time.sleep(1)
                # 运行完代码后，向谷歌插件发送结束通知
                if informChrome:
                    sendEndMsg()
                try:
                    result["result"] = "success"
                    run_result = result["result"]
                except Exception as e:
                    exstr = str(traceback.format_exc()).replace('\n', '<br>')
                    result["conso"] = exstr
                    run_result = "failed"
                if run_result == "success":
                    continue
                else:
                    conso = result["conso"].replace("\n", "<br>")
                    result["conso"] = conso
                    logger.error("Some Process Error!" + str(result["msg"]), 10000)
                    logger.error("Error:" + str(result["conso"]), 10000)
                    copytemptoerror()
            except Exception as err:
                """
                报错改进
                """
                # 设置默认esc退出状态
                lbtExc.setExit(0)
                # 运行完代码后，向谷歌插件发送结束通知
                if informChrome:
                    sendEndMsg()
                result["result"] = "error"
                # 判断是否是esc中断流程引起的错误
                # if err == "Client_Error: The process was forced to pause and exit!":
                if "Client_Error:" in str(err):
                    divId, groupId, action = lbtExc.getId()
                    for i, line_code in enumerate(codes):
                        if divId in line_code["code"] and groupId in line_code["code"]:
                            try:
                                node = codes[i - 1]["code"].split("# node_no: ")[1].split(" ")[0]
                            except:
                                node = ""
                            break
                else:  # 非中断流程引起的错误
                    # 根据行号做报错的兜底方案
                    error_list = traceback.format_exception(*sys.exc_info(), limit=None, chain=True)
                    # 计算 报错line的number数
                    for error in error_list:
                        if "<string>" in error:
                            line = error.split(",")
                            # 提取机器人执行的报错行数
                            line_number = int(line[1].split(" ")[2])
                            print("temporary.py error_line_number:%s" % line_number)
                            break
                    try:
                        node = codes[line_number - 2]["code"].split("# node_no: ")[1].split(" ")[0]
                    except:
                        node = ""
                    action = codes[line_number - 1]["function"]
                    divId = codes[line_number - 1]["div_id"]
                    groupId = codes[line_number - 1]["group_id"]
                    # # 流程group的报错
                    # if codes[line_number]["group_id"] is not None:
                    #     result["step"] = codes[line_number]["group_id"]
                    # else:
                    #     result["step"] = codes[line_number]["div_id"]
                result["node"] = str(node)
                exstr = str(traceback.format_exc()).replace('\n', '<br>')
                result["msg"] = str(action) + " Failed!"
                if groupId:
                    result["step"] = groupId
                    result["groupId"] = divId
                else:
                    result["step"] = divId
                copytemptoerror()
                printContent = lbtExc.getList()
                if printContent:
                    result["conso"] = printContent + "<br>" + str(err)
                else:
                    result["conso"] = str(err)
                # 详细报错
                # result["conso"] = str(exstr)
                logger.error("Details_Error or exit:\n" + str(exstr), 10000)
                result['action'] = 'run'
                result['username'] = username
                result['commander'] = commander
                result['booked_time'] = booked_time
        else:
            time.sleep(1)
            wait += 1
    logger.info('process run %s' % datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 10000)
    result['action'] = 'run'
    result['username'] = username
    result['commander'] = commander
    result['booked_time'] = booked_time
    printContent = lbtExc.getList()
    result['print'] = printContent
    returnRPA(platform)
    logger.info("process execute_result:" + result.get("result") + " %s" % datetime.datetime.now().strftime(
        '%Y-%m-%d %H:%M:%S'), 10000)
    await websocket.send(json.dumps(result))
    pythoncom.CoUninitialize()
    lbtExc.setExit(1)


async def check_action_thread(action, recv_str, username):
    if action == 'web_target':
        browser = recv_str['browser']
        delay = recv_str['delay']
        await web_targeting(websocket, username, browser.lower(), delay)
    elif action == 'sap_target':
        await sap_targeting(websocket, username)
    elif action == 'handle_target':
        await handle_targeting(websocket, username)
    elif action == 'icon_target':
        position = recv_str['position']
        delay = recv_str['delay']
        await icon_targeting(websocket, username, position, delay)
    elif action == 'icon_coordinate_target':
        delay = recv_str['delay']
        await icon_coordinate_targeting(websocket, username, delay)
    elif action == 'mouse_target':
        method = recv_str['method']
        await mouse_targeting(websocket, username, method)
    elif action == 'web_dragging_mouse_target':
        method = recv_str['method']
        await web_dragging_mouse_targeting(websocket, username, method)
    elif action == 'web_crawl':
        browser = recv_str['browser']
        await web_crawl(websocket, username, browser.lower())
    elif action == 'run':
        try:
            await run_temporary(websocket, recv_str)
        except Exception as e:
            logger.exception(str(e), 10000)


# 检查服务端传输的动作类型
async def check_action(websocket, recv_str):
    global exitStatus, exit_force, keyboardListener
    # print(recv_str, "++++++++++++++++++++++++")
    lbtExc.setExit(0)
    result_check_action = {}
    platform = win32gui.GetForegroundWindow()
    exit_force = 0
    try:
        if recv_str != '':
            recv_str = json.loads(recv_str)
            action = recv_str.get('action')
            username = recv_str.get('username')
            try:
                browser = recv_str.get('browser')
            except Exception:
                browser = ""
            if "commander" in recv_str:
                commander = recv_str.get('commander')
            else:
                commander = ""
            if "booked_time" in recv_str:
                booked_time = recv_str.get('booked_time')
            else:
                booked_time = ""
            # 获取秘钥
            if action == 'get_secret_key':
                try:
                    secret_key_path = recv_str.get('choclead_secret_key')
                    secret_key_path = secret_key_path.replace("\"", "")
                    secret_key_path = secret_key_path.replace("\\", "\\\\")
                    # print(secret_key_path)
                    secret_key = encryption.get_secret_key(secret_key_path)
                    password_result = {"secret_key": secret_key, "action": "secret_key_result", "result": "success"}
                    await websocket.send(json.dumps(password_result))
                except Exception as e:
                    password_result = {"secret_key": "", "action": "secret_key_result", "result": "error"}
                    await websocket.send(json.dumps(password_result))
            if action == 'runtask':
                await runtask_temporary(websocket, recv_str)
            elif action == 'runCycletask':
                await runCycletask_temporary(websocket, recv_str)
            elif action == 'sap_record':
                record = recv_str.get('record')
                auto = recv_str.get('auto')
                await sap_recording(websocket, username, record, auto)
            elif action == 'chrome_record':
                record = recv_str.get('record')
                auto = recv_str.get('auto')
                await chrome_recording(websocket, username, record, auto)
            elif action == 'web_crawl' and browser.lower() == "chrome":
                await web_crawl(websocket, username, browser.lower())
            # run和target使用线程方法,监听ESC按键 Forced退出
            elif action in ["web_target", "sap_target", "handle_target", "icon_target", "icon_coordinate_target",
                            "mouse_target", "web_dragging_mouse_target", "web_crawl", "run"]:
                # 分线程执行键盘监听事件
                tlisten = threading.Thread(target=listen)
                tlisten.setDaemon(True)
                tlisten.start()
                # run和target使用线程方法,监听ESC按键 Forced退出
                tcheck_action_thread = threading.Thread(target=asyncio.run,
                                                        args=(check_action_thread(action, recv_str, username),))
                tcheck_action_thread.setDaemon(True)
                tcheck_action_thread.start()
                while True:
                    time.sleep(1)
                    exitStatus = lbtExc.getExit()
                    if exitStatus == 1:
                        # 强制退出
                        if exit_force == 1:
                            error = "Client_Error: The process was forced to exit immediately!"
                            if action == "run":
                                stop_thread(tcheck_action_thread)
                                divId, groupId, action = lbtExc.getId()
                                result_check_action["conso"] = str(error)
                                result_check_action["msg"] = str(action) + " Failed!"
                                result_check_action["result"] = "failed"
                                if groupId:
                                    result_check_action["step"] = groupId
                                    result_check_action["groupId"] = divId
                                else:
                                    result_check_action["step"] = divId
                                result_check_action['action'] = 'run'
                                result_check_action['username'] = username
                                result_check_action['commander'] = commander
                                result_check_action['booked_time'] = booked_time
                                returnRPA(platform)
                                await websocket.send(json.dumps(result_check_action))
                                # print("check_action_thread exitStatus is 1 and exit_force is 1,exit while")
                                break
                            elif action == "web_target":
                                browser = recv_str.get('browser')
                                if browser == "ie":
                                    hm = ie_target.getHm()
                                    if hm:
                                        hm.UnhookMouse()  # 取消鼠标监听
                                        ie_target.setHm("")
                                    ie_target.setClickEvent(1)
                                    returnRPA(platform)
                                    break
                                # elif browser == "chrome":
                                #     chrome_json = {"type": "endCatch"}
                                #     chrome_socket.sendMessage(chrome_json)
                                #     exit_force = 0
                                #     break
                            elif action == "sap_target":
                                sap_target.setClickEvent(1)
                                returnRPA(platform)
                                break
                            elif action == "handle_target":
                                stop_thread(tcheck_action_thread)
                                # ui_spy.setClickEvent(1)
                                hm = ui_spy.getHm()
                                if hm:
                                    hm.UnhookMouse()  # 取消鼠标监听
                                    ui_spy.setHm("")
                                hwnd = win32gui.GetForegroundWindow()
                                win32gui.InvalidateRect(hwnd, None, True)
                                win32gui.UpdateWindow(hwnd)
                                win32gui.RedrawWindow(hwnd, None, None,
                                                      win32con.RDW_FRAME | win32con.RDW_INVALIDATE | win32con.RDW_UPDATENOW | win32con.RDW_ALLCHILDREN)
                                ui_spy.refresh_all()
                                result_check_action['action'] = 'handle_target'
                                result_check_action['username'] = username
                                await websocket.send(json.dumps(result_check_action))
                                returnRPA(platform)
                                break
                            elif action == "icon_target":
                                ui_spy.setClickEvent(1)
                                returnRPA(platform)
                                break
                            elif action == "icon_coordinate_target":
                                ui_spy.setClickEvent(1)
                                returnRPA(platform)
                                break
                            elif action == "mouse_target":
                                stop_thread(tcheck_action_thread)
                                result_check_action['action'] = 'mouse_target'
                                result_check_action['username'] = username
                                await websocket.send(json.dumps(result_check_action))
                                try:
                                    shell = win32com.client.Dispatch("WScript.Shell")
                                    shell.SendKeys('^')
                                    win32gui.SetForegroundWindow(platform)
                                except Exception:
                                    pass
                                break
                            elif action == "web_dragging_mouse_target":
                                stop_thread(tcheck_action_thread)
                                result_check_action['action'] = 'web_dragging_mouse_target'
                                result_check_action['username'] = username
                                await websocket.send(json.dumps(result_check_action))
                                try:
                                    shell = win32com.client.Dispatch("WScript.Shell")
                                    shell.SendKeys('^')
                                    win32gui.SetForegroundWindow(platform)
                                except Exception:
                                    pass
                                break
                            elif action == "web_crawl" and browser.lower() == "ie":
                                stop_thread(tcheck_action_thread)
                                # 获取所有ie窗口
                                ieTarget_new = ie_crawl.Ie_Crawl()
                                ieTarget_new.getWindows()
                                # 遍历ie窗口清除插入的js文件
                                for ieTarget_new.ie in ieTarget_new.window_list:
                                    ieTarget_new.doc = ieTarget_new.ie.Document
                                    ieTarget_new.pw = ieTarget_new.doc.parentWindow
                                    body = ieTarget_new.doc.body
                                    ieTarget_new.removeCode(body, ieTarget_new.pw)
                                    ieTarget_new.pw.execScript(
                                        "rpa_finish();if (document.getElementById('rpa_bot_window')){document.getElementById('rpa_bot_window').parentElement.removeChild(document.getElementById('rpa_bot_window'))};delete document.divJson;delete document.div1;delete document.div2;delete document.rpa_div;delete document.rpa_status;delete document.chocleadStatus;")
                                # 发送结束result
                                result_check_action['action'] = 'web_crawl'
                                result_check_action['username'] = username
                                result_check_action[
                                    'msg'] = "the JSON object must be str, bytes or bytearray, not NoneType"
                                await websocket.send(json.dumps(result_check_action))
                                # 回到ChocLead主页
                                try:
                                    shell = win32com.client.Dispatch("WScript.Shell")
                                    shell.SendKeys('^')
                                    win32gui.SetForegroundWindow(platform)
                                except Exception:
                                    pass
                                break
                        # 暂停退出/正常运行退出exit_force = 0
                        else:
                            break
                # 退出监听线程
                wait_seconds = 0
                while wait_seconds < 3:
                    try:
                        if keyboardListener:
                            keyboardListener.stop()
                            wait_seconds = 3
                    except Exception:
                        wait_seconds += 1
                        time.sleep(1)

            # 日志处理action
            elif action == "client_log":
                if recv_str.get("code") == 1:
                    delete_sqlite_log(recv_str.get("log_id_list"))

            elif action == "check_user":
                if recv_str.get("code") == 0:
                    setCode(True)
                else:
                    setCode(False)
                    # logger.error("用户: %s进入到密码管理界面失败!" % self.username.get(), 10000)
                    win32api.MessageBox(0, _("Operation failed, please try again!"), "ChocLead",
                                        win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
            elif action == "keepalive_ack":
                global websocket_flag
                websocket_flag = True
                logger.debug("websocket keepalive_ack", 10000)
            else:
                return True
        else:
            return True
    except Exception as e:
        logger.error(str(e), 10000)
        return True


async def check_websocket_status(websocket):
    global websocket_flag
    websocket_flag = False
    flag = 0
    while True:
        time.sleep(3)
        if websocket_flag or flag == 0:
            cred_text = {"action": "keepalive"}
            logger.debug("websocket keepalive", 10000)
            await websocket.send(json.dumps(cred_text))
            flag = 1
        else:
            # logger.error("Websocket reconnect", 10000)
            tray.setValue("reconnect")


def check_websocket_thread_tools(websocket):
    try:
        loop_log = asyncio.new_event_loop()
        asyncio.set_event_loop(loop_log)
        loops = asyncio.get_event_loop()
        loops.run_until_complete(check_websocket_status(websocket))
    except Exception as e:
        print(e)
        # logger.exception("Websocket Tools Exception: %s" % e, 10000)


# 客户端主逻辑
async def main_logic():
    try:
        global port, websocket, robot_status, reconnect, chromeStart
        hld = win32gui.FindWindow(None, u"TecleadRobotIcon")
        if hld != 0 and robot_status != "reconnect":
            win32api.MessageBox(0, _("Choclead cannot be started repeatedly"), "ChocLead",
                                win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
            return
        icon_transfer.createIcon()
        # 创建sqlite服务并插入配置信息
        sqlite_tool = get_SqlLite()

        WriteLog(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "Loop: " + " " + __file__ + " " + str(
            sys._getframe().f_lineno))

        result = sqlite_tool.start(table="client_settings", operation_type=2, settings=None, language=None)
        if len(result) and result.get("result") == "failed":
            result = sqlite_tool.start(table="client_settings", operation_type=0, settings=None, language=None)
            if result.get("result") == "failed" and result.get("msg") != "table client_settings already exists":
                win32api.MessageBox(0, "client_settings init error", "ChocLead",
                                    win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                return
            else:
                pass
        else:
            pass

        result = sqlite_tool.start(table="password_manage", operation_type=16, settings=None, language=None)
        if result.get("result") == "failed":
            result = sqlite_tool.start(table="password_manage", operation_type=1, settings=None, language=None)
            if result.get("result") == "failed" and result.get("msg") != "table password_manage already exists":
                win32api.MessageBox(0, "password_manage init error", "ChocLead",
                                    win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                return
            else:
                pass
        else:
            pass

        result = sqlite_tool.start(table="user_login", operation_type=9, settings=None, language=None)
        if len(result) and result.get("result") == "failed":
            result = sqlite_tool.start(table="user_login", operation_type=8, settings=None, language=None)
            if result.get("result") == "failed" and result.get("msg") != "table user_login already exists":
                win32api.MessageBox(0, "user_login init error", "ChocLead",
                                    win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                return
            else:
                pass
        else:
            pass

        result = sqlite_tool.start(table="log_record", operation_type=14, settings=None, language=None)
        if len(result) and result.get("result") == "failed":
            result = sqlite_tool.start(table="log_record", operation_type=12, settings=None, language=None)
            if result.get("result") == "failed" and result.get("msg") != "table log_record already exists":
                win32api.MessageBox(0, "log_record init error", "ChocLead",
                                    win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                return
            else:
                pass
        else:
            pass

        result = sqlite_tool.start(table="password_manage_node", operation_type=15, settings=None, language=None)
        if len(result) and result.get("result") == "failed":
            result = sqlite_tool.start(table="password_manage_node", operation_type=13, settings=None, language=None)
            if result.get("result") == "failed" and result.get("msg") != "table password_manage_node already exists":
                win32api.MessageBox(0, "password_manage_node init error", "ChocLead",
                                    win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                return
            else:
                pass
        else:
            pass
        try:
            WriteLog(datetime.datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S") + "  This is client setting:" + " " + __file__ + " " + str(
                sys._getframe().f_lineno))
            result = sqlite_tool.start(table="client_settings", operation_type=2, settings=None, language=None)
            WriteLog(datetime.datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S") + "  client_settings:" + " " + __file__ + " " + str(sys._getframe().f_lineno))
            if result.get("result") == "success":
                WriteLog(datetime.datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S") + "  client_settings:" + " " + __file__ + " " + str(sys._getframe().f_lineno))
                data = result.get("data")
                WriteLog(datetime.datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S") + "  client_settings:" + " " + __file__ + " " + str(sys._getframe().f_lineno))
                tray.setLanguage(data.get("system_language"))
                WriteLog(datetime.datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S") + "  client_settings:" + " " + __file__ + " " + str(sys._getframe().f_lineno))
            tray.setLanguage('en_US')
        except Exception as e:
            WriteLog(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "  Exception:" + str(
                e) + " " + __file__ + " " + str(sys._getframe().f_lineno))
            tray.setLanguage('en_US')
        i18n_main()

        WriteLog(datetime.datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S") + "  This is ServerAddressTool setting:" + " " + __file__ + " " + str(
            sys._getframe().f_lineno))
        addre_tool = ServerAddressTool(sqlite_tool)
        code, res = addre_tool.getServerInfo()
        if code == 0:
            return
            # 获取系统运行的配置文件settings
        WriteLog(datetime.datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S") + " Have set ServerAddressTool setting:" + " " + __file__ + " " + str(
            sys._getframe().f_lineno))
        chrome_port = "5678"
        client_port = "1234"
        server_address = res[0]
        pid = res[2]
        try:
            # 判断是否需要启动chrome服务监听
            if chromeStart == "no":
                chromeStart = "yes"
                chrome_socket.setChromePort(chrome_port)
                t = threading.Thread(target=chrome_socket.main)
                t.setDaemon(True)
                t.start()
        except Exception as e:
            print("chrome_socket error")
            error_info = "chrome_socket error\n" + str(e)
            logger.error(error_info, 10000)
            return False
        # 判断是否需要创建连接(根据program_name)
        if reconnect == "no":
            try:
                filepath = os.path.abspath(__file__)
                filenames = filepath.split("\\")
                program_name = filenames[len(filenames) - 1]
            except Exception:
                program_name = "ChocLead.exe"
            try:
                tray.setProgramName(program_name)
                create_connect = False
                if pid == "" and pid is not None:
                    current_pid = getPidByName(program_name)
                else:
                    current_name = getNameByPid(pid)
                    if current_name == program_name:
                        # 取消加载窗口
                        loading.setStatus(1)
                        login_window.error_remainder(_("repeatInfo"))
                        sys.exit()
                    else:
                        if current_name:
                            current_name = current_name.split(".")[0]
                        if program_name:
                            program_name = program_name.split(".")[0]
                        if current_name == program_name:
                            # 取消加载窗口
                            loading.setStatus(1)
                            login_window.error_remainder(_("repeatInfo"))
                            sys.exit()
                        else:
                            create_connect = True
                            old_str = "pid = " + str(pid)
                            program_name = "ChocLead.exe"
                            new_pid = getPidByName(program_name)
                            if new_pid != None:
                                new_str = "pid = " + str(new_pid)
            except Exception as e:
                print("reconnect is no tray.setProgramName error")
                error_info = "reconnect is no tray.setProgramName error\n" + str(e)
                logger.error(error_info, 10000)
                return False
        else:
            create_connect = True

        WriteLog(datetime.datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S") + " Begin Connect to Server:" + server_address + " " + pid + " " + reconnect + " " + str(
            create_connect) + " " + __file__ + " " + str(sys._getframe().f_lineno))
        ResetServer = int(3)
        while (create_connect == True) and (ResetServer > 0):
            code, res = addre_tool.getServerInfo()
            if code == 0:
                WriteLog(datetime.datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S") + " Have getServerInfo fail:" + " " + __file__ + " " + str(
                    sys._getframe().f_lineno))
                return
            server_address = res[0]

            try:
                # 创建和服务端的消息通讯
                if create_connect:
                    WriteLog(datetime.datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S") + " Conncet to server:" + " " + server_address + " " + __file__ + " " + str(
                        sys._getframe().f_lineno))
                    async with websockets.connect(server_address + '/websocketlink/', ping_timeout=60) as websocket:
                        # 取消加载窗口
                        setWebsocket(websocket)
                        loading.setStatus(1)
                        sendMsg = sendMessage()
                        chrome_socket.setServer(sendMsg)
                        # icon_transfer.createIcon()
                        WriteLog(datetime.datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S") + " Connect server success:" + server_address + " " + __file__ + " " + str(
                            sys._getframe().f_lineno))
                        await auth_mac(websocket)
                        hwnd = tray.getHwnd()
                        switch_icon(hwnd, "icons\\robot.ico", _("robot_title"))
                        logger.info("机器人状态: %s" % robot_status, 10000)
                        if robot_status == "reconnect":
                            tray.setConnect("connect")
                            win32api.MessageBox(0, _("reconnectInfo"), "ChocLead",
                                                win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                        tray.setValue("connect")
                        show_error = "no"
                        # 启动上传日志线程
                        # await check_websocket_status(websocket)
                        log_tt = threading.Thread(target=(run_log_thread_tools))
                        log_tt.start()
                        # 检测websocket是否活跃
                        websocket_check = threading.Thread(target=(check_websocket_thread_tools), args=(websocket,))
                        websocket_check.start()

                        while True:
                            global websocket_flag
                            websocket_flag = True
                            try:
                                recv_str = await websocket.recv()
                                await check_action(websocket, recv_str)
                            except Exception as e:
                                robot_status = tray.getValue()
                                if robot_status == "connect":
                                    hwnd = tray.getHwnd()
                                    switch_icon(hwnd, "icons\\error.ico", _("robot_title"))
                                    tray.setValue("disconnect")
                                elif robot_status == "reconnect":
                                    reconnect = "yes"
                                    loading.setStatus(0)
                                    t = threading.Thread(target=loadingIcon)
                                    t.setDaemon(True)
                                    t.start()
                                    await main_logic()
                                elif robot_status == "exit" or robot_status == "button_exit":
                                    chrome_socket.quit()
                                    os.kill(os.getpid(), signal.SIGINT)
                                if "1006" in str(e) and show_error == "no" and websocket_flag:
                                    time.sleep(3)
                                    tray.setValue("reconnect")
                                    show_error = "yes"
            except Exception as e:
                # 取消加载窗口
                WriteLog(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " Client error:" + str(
                    e) + " " + __file__ + " " + str(sys._getframe().f_lineno))
                loading.setStatus(1)
                logger.error(str(e), 10000)
                hwnd = tray.getHwnd()
                switch_icon(hwnd, "icons\\error.ico", _("robot_title"))
                tray.setValue("button_exit")
                if "10061" in str(e) or "11001" in str(e):
                    win32api.MessageBox(0, _("serverWarning"), "ChocLead",
                                        win32con.MB_OK | win32con.MB_ICONERROR | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)

                    addre_tool = ServerAddressTool(sqlite_tool)
                    code, res = addre_tool.setting()
                    if code == 1:
                        return False
                        # 获取系统运行的配置文件settings
                    ResetServer = ResetServer - 1
                    pass
                elif "invalid literal" in str(e) or "HTTP 404" in str(e):
                    win32api.MessageBox(0, _("serverWarning"), "ChocLead",
                                        win32con.MB_OK | win32con.MB_ICONERROR | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)

                    addre_tool = ServerAddressTool(sqlite_tool)
                    code, res = addre_tool.setting()
                    if code == 1:
                        return False
                        # 获取系统运行的配置文件settings
                    ResetServer = ResetServer - 1
                    pass
                else:
                    if "local variable 'r' referenced before assignment" not in str(e):
                        win32api.MessageBox(0, str(e), "ChocLead",
                                            win32con.MB_OK | win32con.MB_ICONERROR | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                        return False

    except Exception as e:
        print("init main_logic error")
        error_info = "init main_logic error\n" + str(e)
        logger.error(error_info, 10000)
        return False


loops = asyncio.get_event_loop()
loops.run_until_complete(main_logic())
# asyncio.get_event_loop().run_until_complete(main_logic())
