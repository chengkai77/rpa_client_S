import asyncio
import json
import win32gui
from tornado import ioloop
from tornado.web import Application
from tornado.websocket import WebSocketHandler
from Lbt import logger

client = None


def setAuto(val):
    global auto
    auto = val


def getAuto():
    global auto
    return auto


def setUserName(val):
    global username
    username = val


def getUserName():
    global username
    return username


def setResult(obj):
    global result
    result = obj


def getResult():
    global result
    return result


def setChromePort(val):
    global chrome_port
    chrome_port = val


def getChromePort():
    global chrome_port
    return chrome_port


def setServer(obj):
    global server
    server = obj


def getServer():
    global server
    return server


def setClient(obj):
    global client
    client = obj


def getClient():
    global client
    return client


def sendMessage(message):
    client = getClient()
    client.write_message(json.dumps(message))


def quit():
    global app
    try:
        app.stop()
    except Exception as e:
        print(e)


class EchoWebSocket(WebSocketHandler):
    global server, client

    def open(self):
        logger.info("WebSocket server of client has started successfully!", 10000)

    # 处理client发送的数据
    def on_message(self, message):
        if message != "ping":
            print(message)
        if message == "ping":
            self.write_message("pong")
        elif message == "close":
            ioloop.IOLoop.current().stop()
        try:
            message_json = json.loads(message)
            if "from" in message_json:
                # 将数据发送给连接服务器的websocket
                if "type" in message_json:
                    type = message_json["type"]
                    if type == "target":
                        # 聚焦chocLead窗口
                        hwnd = win32gui.FindWindow(None, "ChocLead - Google Chrome")
                        if hwnd:
                            win32gui.SetForegroundWindow(hwnd)
                        # 传输target值
                        message_json["browser"] = "chrome"
                        message_json["action"] = "web_target"
                        message_json["status"] = "success"
                        username = getUserName()
                        message_json["username"] = username
                        server = getServer()
                        server.start(message_json)
                    elif type == "records":
                        records_js = {}
                        record_list = []
                        records = message_json["records"]
                        for record in records:
                            command = record["command"]
                            if auto == 1:
                                record_json = {}
                                parameters = {}
                                if command == "open":
                                    element = record["target"][0]
                                    record_json["action"] = "Web_StartBrowser"
                                    parameters["login_url"] = '"' + element[0] + '"'
                                    parameters["browser"] = "\"chrome\""
                                    parameters["web_active"] = "\"yes\""
                                    parameters["size"] = "\"2\""
                                    record_json["parameters"] = parameters
                                    record_list.append(record_json)
                                elif command == "type":
                                    element = record["target"][0]
                                    method = element[1]
                                    if "xpath" in method:
                                        method = "xpath"
                                    elif method == "link":
                                        method = "linkText"
                                    record_json["action"] = "Web_SendKeys"
                                    parameters["method"] = '"' + method + '"'
                                    parameters["tag_name"] = '"' + record["tagName"] + '"'
                                    parameters["iframe"] = '"' + record["frameLocation"] + '"'
                                    parameters["element_name"] = '"' + element[0].replace("id=", "").replace("link=",
                                                                                                             "") + '"'
                                    parameters["content"] = '"' + record["value"] + '"'
                                    parameters["web_active"] = "\"yes\""
                                    record_json["parameters"] = parameters
                                    record_list.append(record_json)
                                elif command == "click":
                                    element = record["target"][0]
                                    method = element[1]
                                    if "xpath" in method:
                                        method = "xpath"
                                    elif method == "link":
                                        method = "linkText"
                                    record_json["action"] = "Web_ClickElement"
                                    parameters["method"] = '"' + method + '"'
                                    parameters["tag_name"] = '"' + record["tagName"] + '"'
                                    parameters["iframe"] = '"' + record["frameLocation"] + '"'
                                    parameters["click_option"] = "\"left\""
                                    parameters["element_name"] = '"' + element[0].replace("id=", "").replace("link=",
                                                                                                             "") + '"'
                                    parameters["web_active"] = "\"yes\""
                                    record_json["parameters"] = parameters
                                    record_list.append(record_json)
                                elif command == "selectWindow":
                                    window_target = str(record["target"][0][0].replace("win_ser_", ""))
                                    if window_target == "local":
                                        window_target = "0"
                                    record_json["action"] = "Web_SwitchToWindow"
                                    parameters["sequence"] = window_target
                                    record_json["parameters"] = parameters
                                    record_list.append(record_json)
                                elif command == "close":
                                    record_json["action"] = "Web_Close"
                                    record_list.append(record_json)
                                elif command == "select":
                                    element = record["target"][0]
                                    method = element[1]
                                    if "xpath" in method:
                                        method = "xpath"
                                    elif method == "link":
                                        method = "linkText"
                                    record_json["action"] = "Web_Select"
                                    parameters["method"] = method
                                    parameters["tag_name"] = '"' + record["tagName"] + '"'
                                    parameters["iframe"] = '"' + record["frameLocation"] + '"'
                                    parameters["element_name"] = '"' + element[0].replace("id=", "").replace("link=",
                                                                                                             "") + '"'
                                    parameters["content"] = '"' + record["value"] + '"'
                                    parameters["web_active"] = "\"yes\""
                                    record_json["parameters"] = parameters
                                    record_list.append(record_json)
                            else:
                                record_str = []
                                if command == "open":
                                    element = record["target"][0]
                                    record_str.append("Client_Trans: Web_StartBrowser")
                                    record_str.append("(")
                                    record_str.append("Client_Trans: Url")
                                    record_str.append("=<span style='color:red'>" + '"' + element[0] + '"' + "</span>)")
                                    record_list.append(record_str)
                                elif command == "type":
                                    element = record["target"][0]
                                    method = element[1]
                                    if "xpath" in method:
                                        method = "xpath"
                                    elif method == "link":
                                        method = "link_text"
                                    record_str.append("Client_Trans: Web_SendKeys")
                                    record_str.append("(")
                                    record_str.append("Client_Trans: Find Element Method")
                                    record_str.append("=<span style='color:red'>" + '"' + method + '"' + "</span>, ")
                                    record_str.append("Client_Trans: Element Property")
                                    if method == "id":
                                        ele_pro = element[0].replace("id=", "").replace("link=", "")
                                    else:
                                        ele_pro = element[0].replace("link=", "")
                                    record_str.append("=<span style='color:red'>" + '"' + ele_pro + '"' + "</span>, ")
                                    record_str.append("Client_Trans: Tag Name")
                                    record_str.append(
                                        "=<span style='color:red'>" + '"' + record["tagName"] + '"' + "</span>,")
                                    record_str.append("Client_Trans: Iframe")
                                    record_str.append(
                                        "=<span style='color:red'>" + '"' + record["frameLocation"] + '"' + "</span>,")
                                    record_str.append("Client_Trans: Content")
                                    record_str.append(
                                        "=<span style='color:red'>" + '"' + record["value"] + '"' + "</span>)")
                                    record_list.append(record_str)
                                elif command == "click":
                                    element = record["target"][0]
                                    method = element[1]
                                    if "xpath" in method:
                                        method = "xpath"
                                    elif method == "link":
                                        method = "link_text"
                                    record_str.append("Client_Trans: Web_ClickElement")
                                    record_str.append("(")
                                    record_str.append("Client_Trans: Find Element Method")
                                    record_str.append("=<span style='color:red'>" + '"' + method + '"' + "</span>, ")
                                    record_str.append("Client_Trans: Element Property")
                                    if method == "id":
                                        ele_pro = element[0].replace("id=", "").replace("link=", "")
                                    else:
                                        ele_pro = element[0].replace("link=", "")
                                    record_str.append("=<span style='color:red'>" + '"' + ele_pro + '"' + "</span>,")
                                    record_str.append("Client_Trans: Tag Name")
                                    record_str.append(
                                        "=<span style='color:red'>" + '"' + record["tagName"] + '"' + "</span>,")
                                    record_str.append("Client_Trans: Iframe")
                                    record_str.append(
                                        "=<span style='color:red'>" + '"' + record["frameLocation"] + '"' + "</span>)")
                                    record_list.append(record_str)
                                elif command == "selectWindow":
                                    window_target = str(record["target"][0][0].replace("win_ser_", ""))
                                    if window_target == "local":
                                        window_target = "0"
                                    record_str.append("Client_Trans: Web_SwitchToWindow")
                                    record_str.append("(")
                                    record_str.append("Client_Trans: Element Sequence")
                                    record_str.append("=<span style='color:red'>" + window_target + "</span>)")
                                    record_list.append(record_str)
                                elif command == "close":
                                    record_str.append("Client_Trans: Web_Close")
                                    record_list.append(record_str)
                                elif command == "select":
                                    element = record["target"][0]
                                    method = element[1]
                                    if "xpath" in method:
                                        method = "xpath"
                                    elif method == "link":
                                        method = "link_text"
                                    record_str.append("Client_Trans: Web_Select")
                                    record_str.append("(")
                                    record_str.append("Client_Trans: Find Element Method")
                                    record_str.append("=<span style='color:red'>" + '"' + method + '"' + "</span>, ")
                                    record_str.append("Client_Trans: Element Property")
                                    if method == "id":
                                        ele_pro = element[0].replace("id=", "").replace("link=", "")
                                    else:
                                        ele_pro = element[0].replace("link=", "")
                                    record_str.append("=<span style='color:red'>" + '"' + ele_pro + '"' + "</span>, ")
                                    record_str.append("Client_Trans: Tag Name")
                                    record_str.append(
                                        "=<span style='color:red'>" + '"' + record["tagName"] + '"' + "</span>,")
                                    record_str.append("Client_Trans: Iframe")
                                    record_str.append(
                                        "=<span style='color:red'>" + '"' + record["frameLocation"] + '"' + "</span>,")
                                    record_str.append("Client_Trans: Content")
                                    record_str.append(
                                        "=<span style='color:red'>" + '"' + record["value"] + '"' + "</span>)")
                                    record_list.append(record_str)
                        records_js["action"] = "chrome_record"
                        records_js["auto"] = auto
                        records_js["result"] = "success"
                        records_js["username"] = getUserName()
                        records_js["record_code"] = record_list
                        server = getServer()
                        server.start(records_js)
                    elif type == "sendCrawlConfig":
                        message_json["action"] = "web_crawl"
                        message_json["type"] = "doCrawl"
                        message_json["username"] = getUserName()
                        server = getServer()
                        server.start(message_json)
                    else:
                        setResult(message_json)
            else:
                if "type" in message_json:
                    if message_json["type"] == "wsConnect":
                        if message_json["msg"] == "success":
                            setClient(self)
                # 将数据发送给插件
                self.write_message(message)
        except Exception:
            pass

    def on_close(self):
        print("WebSocket closed, reason: " + str(self.close_reason))

    def ping(self, data):
        if self.ws_connection is None:
            print("Chrome socket has been disconnected!")
        else:
            self.ws_connection.write_ping(data)

    def on_pong(self, data):
        pass

    # 允许所有跨域通讯，解决403问题
    def check_origin(self, origin):
        return True


def mywatcher():
    try:
        client = getClient()
        client.ping(b"ping")
    except Exception:
        pass


def main():
    global chrome_port, app
    application = Application([
        (r"/", EchoWebSocket),
    ])
    try:
        chrome_port = getChromePort()
    except Exception:
        chrome_port = "5678"
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    app = application.listen(chrome_port)
    # ioloop.PeriodicCallback(mywatcher, 5000).start()  # start scheduler 每隔2s执行一次
    ioloop.IOLoop.current().start()

# main()
