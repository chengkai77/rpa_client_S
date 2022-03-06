# -*- coding:utf-8 -*-
import json
import threading
import win32api
import win32com
import win32con
import win32gui
from win32com.client import Dispatch
from pynput.mouse import Listener, Controller
import time
import PyHook3
import pythoncom

# 获取xpath函数
js_code = """
            function readXPath(element) {
                if (element.id !== "") {//判断id属性，如果这个元素有id，则显 示//*[@id="xPath"]  形式内容
                    return "//*[#tagName=\'" + element.tagName + "\'@id=\'" + element.id + "\']";
                }
                //这里需要需要主要字符串转译问题，可参考js 动态生成html时字符串和变量转译（注意引号的作用）
                if (element == document.body) {//递归到body处，结束递归
                    return '/html/' + element.tagName.toLowerCase();
                }
                var ix = 1,//在nodelist中的位置，且每次点击初始化
                siblings = element.parentNode.childNodes;//同级的子元素
                for (var i = 0, l = siblings.length; i < l; i++) {
                    var sibling = siblings[i];
                    //如果这个元素是siblings数组中的元素，则执行递归操作
                    if (sibling == element) {
                        return arguments.callee(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix) + ']';
                    //如果不符合，判断是否是element元素，并且是否是相同元素，如果是相同的就开始累加
                    } else if (sibling.nodeType == 1 && sibling.tagName == element.tagName) {
                        ix++;
                    }
                }
            }

            function mouseMove(e){            
                e = e || window.event;
                var scrollX = document.documentElement.scrollLeft || document.body.scrollLeft;
                var scrollY = document.documentElement.scrollTop || document.body.scrollTop;
                var x = e.pageX - scrollX || e.clientX;
                var y = e.pageY - scrollY || e.clientY;
                ele = document.elementFromPoint(x,y);
                //消除其他所有之前选中元素样式
                if(document.querySelectorAll){
                    var elements = document.querySelectorAll(".select-target");
                }else{
                    var elements = [];//定义一个数组用来存class相同的节点
                    //1、查找node所有的子节点
                    var nodes = document.getElementsByTagName("*");
                    /*node.getElementsByTagName("*") 的意思是通过标签名查找所以node节点下所有的节点*为通配符*/
                    for(var i = 0; i < nodes.length; i++){//遍历每一个节点
                        if (nodes[i].classList != null && typeof(nodes[i].classList) != 'undefined'){
                            if(nodes[i].classList.indexOf("select-target") != -1){//判断每一个节点的class属性名是否等于 传入的class名
                                elements.push(nodes[i]);
                            }
                        }
                    }
                }
                for (var i=0;i<elements.length;i++){
                    var el = elements[i];
                    if (el != ele){
                        try{
                            el.classList.remove('select-target');
                        }catch(e){
                            el.className = el.className.replace(/select-target/g, "");
                        }
                    }
                }
                if(ele){
                    var draw = "",classFound = "";
                    var classList = ele.classList || ele.className;
                    //判断元素是否有class
                    if (typeof(classList) == "undefined"){
                        draw = "yes";
                    }else{
                        try{
                            //ie8及以上版本
                            if (classList.contains('select-target') == false){
                                draw = "yes";
                                classFound = "yes";
                            }
                        }catch(err){
                            //ie5、6、7
                           if (classList.indexOf('select-target') == -1){
                                draw = "yes";
                                classFound = "yes";
                            } 
                        }
                    }
                    if (draw == "yes"){
                        var div_json = {};
                        var frame_list = [];
                        var tagName = ele.tagName;
                        if (tagName){
                            if (tagName == "HTML" || tagName == "BODY" || tagName == "IFRAME"){
                                return false;
                            }
                            div_json.tagName = tagName;
                            if (tagName == "A"){
                                div_json.linkText = ele.innerText;
                                var divs = document.getElementsByTagName(tagName);
                                var n = 0;
                                for (var i=0;i<divs.length;i++){
                                    var div = divs[i];
                                    var linkText = div.innerText;
                                    if (linkText == div_json.linkText){
                                        if (div == ele){
                                            break;
                                        }else{
                                            n = n + 1;
                                        }
                                    }
                                }
                                div_json.link_sequence = n;
                            }
                        }
                        var divId = ele.id;
                        if (divId){
                            div_json.divId = divId;
                            div_json.id_sequence = 0;
                        }
                        var divName = ele.getAttribute('name');
                        if (divName){
                            var divs = document.getElementsByName(divName);
                            var n = 0;
                            for (var i=0;i<divs.length;i++){
                                var div = divs[i];
                                if (div == ele && div.tagName == tagName){
                                    break;
                                }else if(div.tagName == tagName){
                                    n = n + 1;
                                }
                            }
                            div_json.Name = divName;
                            div_json.name_sequence = n;
                        }
                        var className = ele.getAttribute('class') || ele.getAttribute('className') || ele.className;
                        if (className){
                            var divs = new Array;
                            try{
                                //ie8及以上
                                divs = document.getElementsByClassName(className);
                            }catch(err){
                                //ie5、6、7
                                var all_divs = document.getElementsByTagName('*');
                                for (var i=0; i<all_divs.length; i++){
                                    var child = all_divs[i];
                                    if (child.className.indexOf(className) != -1){
                                        divs.push(child);
                                    }
                                }
                            }
                            var n = 0;
                            for (var i=0;i<divs.length;i++){
                                var div = divs[i];
                                if (div == ele && div.tagName == tagName){
                                    break;
                                }else if(div.tagName == tagName){
                                    n = n + 1;
                                }
                            }
                            div_json.className = className;
                            div_json.class_sequence = n;
                        }
                        var xpath = readXPath(ele);
                        div_json.className = className;
                        div_json.class_sequence = n;
                        div_json.xpath = xpath;
                        var current_window = window.self;
                        while (current_window != window.top){
                            var all_wins = current_window.parent.frames;
                            for (var i=0;i<all_wins.length;i++){
                                var win = all_wins[i];
                                if (win == current_window){
                                    frame_list.unshift(i);
                                }
                            }
                            current_window = current_window.parent;
                        }
                        div_json.frame = frame_list;
                        if (typeof(JSON) == 'undefined' || typeof(JSON.stringify) == 'undefined'){
                            var div = "{";
                            for (var key in div_json){
                                div = div + '\"' + key + '\"' + ":" + '\"' + div_json[key] + '\"' + ",";
                            }
                            if (div != "{"){
                                div = div.substr(0,div.length-1);
                            }
                            div = div + "}";
                            window.top.document.divJson = div;
                        }else{
                            window.top.document.divJson = JSON.stringify(div_json);
                        }
                        if (classFound == "yes"){
                            try{
                                //ie8及以上
                                ele.classList.add("select-target");
                            }catch(err){
                                //ie5、6、7
                                ele.className = ele.className + " select-target";
                            }
                        }else{
                            ele.className = "select-target";
                        }
                    }
                }
            }

            function clearChocLeadStyle(e){
                if (document.documentElement.className.indexOf('select-target') != -1){
                    try{
                        //ie8及以上
                        document.documentElement.classList.remove('select-target');
                    }catch(err){
                        //ie5、6、7
                        document.documentElement.className.replace(/select-target/g,"");
                    }
                }
                if (document.body.className.indexOf('select-target') != -1){
                    try{
                        //ie8及以上
                        document.body.classList.remove('select-target');
                    }catch(err){
                        //ie5、6、7
                        document.body.className.replace(/select-target/g,"");
                    }
                }
                if(document.querySelectorAll){
                    //ie8及以上
                    var elements = document.querySelectorAll(".select-target");
                }else{
                    //ie5、6、7
                    var elements = [];//定义一个数组用来存class相同的节点
                    //1、查找node所有的子节点
                    var nodes = document.getElementsByTagName("*");
                    /*node.getElementsByTagName("*") 的意思是通过标签名查找所以node节点下所有的节点*为通配符*/
                    for(var i = 0; i < nodes.length; i++){//遍历每一个节点
                        if (nodes[i].className != null && typeof(nodes[i].className) != 'undefined'){
                            if(nodes[i].className.indexOf("select-target") != -1){//判断每一个节点的class属性名是否等于 传入的class名
                                elements.push(nodes[i]);
                            }
                        }
                    }
                }
                for (var i=0;i<elements.length;i++){
                    var el = elements[i];
                    try{
                        el.classList.remove('select-target');
                    }catch(err){
                        el.className = el.className.replace(/select-target/g,"");
                    }
                }
            }
            
            function rpa_mouse_click(e){
                e = e || window.event;
                var target = e.target || e.srcElement;
                if (target.tagName.toLowerCase() === 'a'){
                    if ( e && e.preventDefault ) {
                        //阻止默认浏览器动作(W3C) 
                        e.preventDefault(); 
                    }else{
                        //IE中阻止函数器默认动作的方式 
                        window.event.returnValue = false; 
                        return false;
                    }    
                }
            }
            
            if (document.body){
                try{
                    document.body.addEventListener("mousemove",mouseMove);
                }catch(err){
                    document.body.onmousemove = mouseMove;
                }
                try{
                    document.body.addEventListener("mouseout",clearChocLeadStyle);
                }catch(err){
                    document.body.onmouseout = clearChocLeadStyle; 
                }
                try{
                    document.body.addEventListener("mouseclick",rpa_mouse_click);
                }catch(err){
                    document.body.onmouseclick = rpa_mouse_click;
                }
            }            
    """

# 页面css属性
css = """
            if (document.getElementsByTagName("head")){
                var style = document.createElement('style');
                style.type = 'text/css'; 
                style.id = "chocLead_css";
                var outline_support = true;
                try{
                    document.body.style.outline = "double red !important";
                    document.body.style.outline = "";
                }catch(e){
                    outline_support = false;
                }
                if (outline_support){
                    try{
                        style.innerHTML=".select-target{background:rgba(0,0,0,.3);opacity:0.5;outline:double red !important;pointer-events:none;}";
                    }catch(err){
                        style.styleSheet.cssText = ".select-target{background:rgba(0,0,0,.3);opacity:0.5;outline:double red !important;pointer-events:none;}";
                    }  
                }else{
                    try{
                        style.innerHTML=".select-target{background:rgba(0,0,0,.3);opacity:0.5;border:2px dashed red !important;pointer-events:none;}";
                    }catch(err){
                        style.styleSheet.cssText = ".select-target{background:rgba(0,0,0,.3);opacity:0.5;border:2px dashed red !important;pointer-events:none;}";
                    }    
                }
                document.getElementsByTagName('HEAD').item(0).appendChild(style);
            }
    """

# 页面js函数
js = """
            if (document.getElementsByTagName("head")){
                var scriptObj = document.createElement("script");
                scriptObj.type = "text/javascript";
                scriptObj.id = "chocLead_js";
                document.getElementsByTagName("head")[0].appendChild(scriptObj);
                try{
                    scriptObj.innerHTML = """ + js_code + """;
                }catch(e){
                    scriptObj.text = """ + js_code + """;
                }
            }
    """

# js添加meta
meta = """
        function addMeta(name,content){//手动添加mate标签
            let meta = document.createElement('meta');
            meta.content = "IE=Edge";
            meta.id = "chocLead_meta";
            document.getElementsByTagName('head')[0].appendChild(meta);
        }

    """

# 移除chocLead插入进网页的js跟css代码
remove_js = """
            var deleteCss = document.getElementById('chocLead_css');
            if (deleteCss){
                deleteCss.parentNode.removeChild(deleteCss);
            }
            var deleteJs = document.getElementById('chocLead_js');
            if (deleteJs){
                deleteJs.parentNode.removeChild(deleteJs);
            }
            if (document.body){
                try{
                    document.body.removeEventListener("mousemove",mouseMove);
                }catch(err){
                    document.body.onmousemove = null;
                }
                try{
                    document.body.removeEventListener("mouseout",clearChocLeadStyle);
                }catch(err){
                    document.body.onmouseout = null; 
                } 
                try{
                    document.body.removeEventListener("mouseclick",rpa_mouse_click);
                }catch(err){
                    document.body.onmouseclick = null;
                }
            }            
        """

# 监听鼠标点击事件
click_event = 0


def setClickEvent(value):
    global click_event
    click_event = value


def getClickEvent():
    global click_event
    return click_event


def setHm(value):
    global hm
    hm = value


def getHm():
    global hm
    return hm


# def on_click(x, y, button, pressed):
#     if "left" in str(button):
#         if not pressed:
#             setClickEvent(1)
#             # Stop listener
#             return False
#     elif "right" in str(button):
#         if not pressed:
#             setClickEvent(1)
#             # Stop listener
#             return False


# def mouse():
# with Listener(on_click=on_click) as listener:
#     listener.join()


class Ie_Target():
    def __init__(self):
        self.ie = ""
        self.doc = ""
        self.pw = ""
        self.url = ""
        self.window_json = {}
        self.window_list = []

    def onMouseEvent(self, event):
        setClickEvent(1)
        return False

    def listen(self):
        self.hm = PyHook3.HookManager()
        self.hm.MouseLeftDown = self.onMouseEvent
        self.hm.HookMouse()
        setHm(self.hm)
        pythoncom.PumpMessages()

    # 获取单个Ie浏览器
    def getIe(self):
        ShellWindowsCLSID = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'
        ShellWindows = Dispatch(ShellWindowsCLSID)
        for shellwindow in ShellWindows:
            if shellwindow.LocationURL != "":
                if "Internet Explorer" in str(shellwindow):
                    self.ie = shellwindow
                    self.doc = self.ie.Document
                    self.pw = self.doc.parentWindow
                    break

    # 获取所有ie窗口
    def getWindows(self):
        """
        获取所有IE窗口
        :return:            窗口的数据集
        """
        ShellWindowsCLSID = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'
        ShellWindows = Dispatch(ShellWindowsCLSID)
        window_list = []
        for shellwindow in ShellWindows:
            if shellwindow.LocationURL != "":
                if "Internet Explorer" in str(shellwindow):
                    if shellwindow not in window_list:
                        window_list.append(shellwindow)
                        hwnd = shellwindow.HWND
                        try:
                            hwnd_window = self.window_json[hwnd]
                            hwnd_window.append(shellwindow)
                        except Exception:
                            hwnd_window = []
                            hwnd_window.append(shellwindow)
                        self.window_json.update({hwnd: hwnd_window})
                        self.window_list.append(shellwindow)

    # 给页面插入js、css文件，同时监听鼠标移出主体事件
    def insertCode(self, body, pw):
        # 给html页面插入css值
        pw.execScript(css)
        # 给html页面js函数
        pw.execScript(js)  # BW网站测试失败
        # 给html手工添加meta
        # pw.execScript(meta)
        # 遍寻html的iframe框架，同时给框架中也插入js跟css
        try:
            childNode = body.getElementsByTagName('iframe')
            for child in childNode:
                pw = child.contentWindow
                body = child.contentWindow.document.body
                self.insertCode(body, pw)
            childNode = body.getElementsByTagName('frame')
            for child in childNode:
                pw = child.contentWindow
                body = child.contentWindow.document.body
                self.insertCode(body, pw)
        except Exception:
            pass

    # 移出给页面插入的js、css文件，同时取消监听移出主体事件
    def removeCode(self, body, pw):
        pw.execScript(remove_js)
        # 遍寻html的iframe框架，同时删除插入框架中的js跟css
        try:
            childNode = body.getElementsByTagName('iframe')
            for child in childNode:
                pw = child.contentWindow
                body = child.contentWindow.document.body
                self.removeCode(body, pw)
        except Exception:
            pass
        try:
            childNode = body.getElementsByTagName('frame')
            for child in childNode:
                pw = child.contentWindow
                body = child.contentWindow.document.body
                self.removeCode(body, pw)
        except Exception:
            pass

    # 鼠标移动循环函数
    def start(self, delay):
        result = {}
        platform = win32gui.GetForegroundWindow()
        try:
            t = threading.Thread(target=self.listen)
            t.start()
            # 获取当前ie窗口句柄将其聚焦
            try:
                web_hwnd = win32gui.FindWindow("IEFrame", None)
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys('^')
                win32gui.SetForegroundWindow(web_hwnd)
            except Exception:
                pass
            time.sleep(int(delay))
            # 获取所有ie窗口
            self.getWindows()
            # 给每个ie窗口插入js文件
            for self.ie in self.window_list:
                self.doc = self.ie.Document
                self.pw = self.doc.parentWindow
                body = self.doc.body
                self.insertCode(body, self.pw)
            while True:
                clickEvent = getClickEvent()
                if clickEvent == 1:
                    self.hm.UnhookMouse()  # 取消鼠标监听
                    setHm("")
                    setClickEvent(0)
                    # 遍历ie窗口清除插入的js文件
                    for self.ie in self.window_list:
                        self.doc = self.ie.Document
                        self.pw = self.doc.parentWindow
                        body = self.doc.body
                        self.removeCode(body, self.pw)
                    # 获取当前聚焦窗口，找到其对应ie浏览器
                    current_hwnd = win32gui.GetForegroundWindow()
                    # 判断当前点击窗口是否为ie窗口
                    try:
                        self.ie_windows = self.window_json[current_hwnd]
                        for ie in self.ie_windows:
                            doc = ie.Document
                            hidden = doc.hidden
                            if not hidden:
                                self.ie = ie
                                self.doc = self.ie.Document
                                self.pw = self.doc.parentWindow
                                break
                    except Exception:
                        pass
                    break
            try:
                try:
                    result = json.loads(self.doc.divJson)
                    # 整理ie7之前版本获取的iframe数据格式
                    try:
                        frame = result["frame"]
                        if isinstance(frame, str):
                            frame_list = frame.split(",")
                            frames = []
                            for i in frame_list:
                                frames.append(int(i))
                            result["frame"] = frames
                    except Exception:
                        pass
                    # 清除保存的变量
                    self.pw.execScript("delete document.divJson")
                except Exception:
                    pass
                self.ie = ""
                result["status"] = "success"
            except Exception:
                result["status"] = "error"
                result["msg"] = "No selected element"
        except Exception as e:
            self.hm.UnhookMouse()  # 取消鼠标监听
            setHm("")
            print(e)
            result["status"] = "error"
            result["msg"] = str(e)
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('^')
            win32gui.SetForegroundWindow(platform)
        except Exception:
            pass
        return result

# if __name__ == '__main__':
#     ieTarget = Ie_Target()
#     ieResult = ieTarget.start(0)
#     print(ieResult)
