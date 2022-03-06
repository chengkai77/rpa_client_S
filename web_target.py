import os
import tkinter
import tkinter.messagebox
import traceback
import pythoncom
import win32gui
import win32print
import win32con
import win32api
from selenium import webdriver
from win32.lib import win32con
from win32api import GetSystemMetrics

def get_real_resolution():
    """获取真实的分辨率"""
    hDC = win32gui.GetDC(0)
    # 横向分辨率
    w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
    # 纵向分辨率
    h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
    return w, h


def get_screen_size():
    """获取缩放后的分辨率"""
    w = GetSystemMetrics (0)
    h = GetSystemMetrics (1)
    return w, h

class targeting():
    def __init__(self):
        self.frame = ""
        self.root = ""
        self.adj_x = 0
        self.adj_y = 0
        self.iframe = []
        self.clear_js = ''
        self.result = {}
        self.hwnd_title = dict()
        self.title = ''
        self.hwnd = ''
        self.all_handles = []

    def get_all_hwnd(self, hwnd, mouse):
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            self.hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})

    def get_hwnd(self, title):
        win32gui.EnumWindows(self.get_all_hwnd, 0)
        for h, t in self.hwnd_title.items():
            if t is not '':
                if title.lower() in t.lower():
                    try:
                        if t.lower()[:len(title.lower())] == title.lower():
                            left, top, right, bottom = win32gui.GetWindowRect(h)
                            return h
                    except Exception:
                        continue

    def get_hwnd_title(self, hwnd):
        win32gui.EnumWindows(self.get_all_hwnd, 0)
        for h, t in self.hwnd_title.items():
            if h == hwnd:
                return t

    def clickMouse(self,event):
        x = event.x_root - self.width - self.adj_x
        y = event.y_root - self.height - self.adj_y
        print(x,y,self.height)
        try:
            self.root.withdraw()
            self.driver.execute_script(self.clear_js)
        except Exception:
            try:
                js = """var elements = document.querySelectorAll(".select-target");
                      for (var i=0;i<elements.length;i++){
                         var el = elements[i];
                         var className = el.className;
                         el.className = className.replace(" select-target","");
                         el.className = className.replace("select-target","");
                      }"""
                self.driver.execute_script(js)
            except Exception:
                pass
        try:
            self.driver.switch_to.window(self.all_handles[0])
        except Exception:
            pass
        self.result['iframe'] = self.iframe
        self.root.destroy()


    def callBack(self,event):
        try:
            hwnd = self.get_hwnd(self.title)
            current_hwnd = win32gui.GetForegroundWindow()
            print(hwnd,current_hwnd)
            x = event.x_root / self.ratio - self.width - self.adj_x
            y = event.y_root / self.ratio - self.height - self.adj_y
            if current_hwnd != hwnd:
                try:
                    self.driver.execute_script(self.clear_js)
                except Exception:
                    js = """var elements = document.querySelectorAll(".select-target");
                        for (var i=0;i<elements.length;i++){
                            var el = elements[i];
                            var className = el.className;
                            el.className = className.replace(" select-target","");
                            el.className = className.replace("select-target","");
                        }"""
                    self.driver.execute_script(js)
                current_title = self.get_hwnd_title(current_hwnd).replace(' - Google Chrome','')
                print(current_title)
                self.title = current_title
                self.hwnd = current_hwnd
                for handle in self.all_handles:
                    self.driver.switch_to.window(handle)
                    if self.driver.title == self.title:
                        break
            else:
                if hwnd == None:
                    try:
                        self.driver.execute_script(self.clear_js)
                    except Exception:
                        js = """var elements = document.querySelectorAll(".select-target");
                            for (var i=0;i<elements.length;i++){
                                var el = elements[i];
                                var className = el.className;
                                el.className = className.replace(" select-target","");
                                el.className = className.replace("select-target","");
                            }"""
                        self.driver.execute_script(js)
                    self.title = self.get_hwnd_title(self.hwnd).replace(' - Google Chrome','')
                    for handle in self.all_handles:
                        self.driver.switch_to.window(handle)
                        if self.driver.title == self.title:
                            break
            self.clear_js = '$(".select-target").removeClass("select-target");'
            if x < 0 or y < 0:
                try:
                    self.driver.execute_script(self.clear_js)
                except Exception:
                    js = """var elements = document.querySelectorAll(".select-target");
                        for (var i=0;i<elements.length;i++){
                            var el = elements[i];
                            var className = el.className;
                            el.className = className.replace(" select-target","");
                            el.className = className.replace("select-target","");
                        }"""
                    self.driver.execute_script(js)
                self.driver.switch_to.default_content()
                self.adj_x = 0
                self.adj_y = 0
                self.iframe = []
            js = """
                $(".select-target").removeClass("select-target");
                var style = document.createElement('style');
                style.type = 'text/css';
                style.rel = 'stylesheet';
                //for Chrome Firefox Opera Safari
                style.appendChild(document.createTextNode('.select-target {border:2px solid #ef4300 !important}'));
                //for WEB
                //style.styleSheet.cssText = code;
                var head = document.getElementsByTagName('head')[0];
                head.appendChild(style);
                //获取xpath
                function readXPath(element) {
                    if (element.id !== "") {//判断id属性，如果这个元素有id，则显 示//*[@id="xPath"]  形式内容
                        return '//*[@id=\"' + element.id + '\"]';
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
                };
                var el_json = {};
                var element = document.elementFromPoint("""+str(x)+","+str(y)+");""""
                if (element){
                    try{
                        el_json.x = element.getBoundingClientRect().left;
                    }catch{
                        var offset = element.offsetTop;
                        while(element.offsetParent != null){
                            offset+=element.offsetParent.offsetTop;
                        }
                    }
                    el_json.y = element.getBoundingClientRect().top;
                    el_json.height = element.offsetHeight;
                    el_json.width = element.offsetWidth;
                    el_json.id = element.id;
                    el_json.name = element.getAttribute('name');
                    if (el_json.name != null && el_json.name != "" && typeof(el_json.name) != "undefined"){
                        eles = document.getElementsByName(el_json.name);
                        for (var i=0;i<eles.length;i++){
                            ele = eles[i];
                            if (ele == element){
                                el_json.name_item = i;
                                break;
                            }
                        }
                    }
                    el_json.className = element.className;
                    if (el_json.className != null && el_json.className != "" && typeof(el_json.className) != "undefined"){
                        eles = document.getElementsByClassName(el_json.name);
                        for (var i=0;i<eles.length;i++){
                            ele = eles[i];
                            if (ele == element){
                                el_json.className_item = i;
                                break;
                            }
                        }
                    }
                    el_json.tagName = element.tagName;
                    if (el_json.tagName != null && el_json.tagName != "" && typeof(el_json.tagName) != "undefined"){
                        eles = document.getElementsByTagName(el_json.tagName);
                        for (var i=0;i<eles.length;i++){
                            ele = eles[i];
                            if (ele == element){
                                el_json.tagName_item = i;
                                break;
                            }
                        }
                    }
                    el_json.text = element.textContent;
                    if (el_json.tagName != "A"){
                        el_json.text = "";
                    }
                    el_json.xpath = readXPath(element);
                    try{
                        element.classList.add("select-target");
                    }catch{
                        element.className = "select-target";
                    }
                    return el_json;  
                }else{
                    return false;
                }
            """
            try:
                self.result = self.driver.execute_script(js)
            except Exception:
                exstr = str(traceback.format_exc())
                jquery = """var script = document.createElement("script");
                      script.type = "text/javascript";
                      script.src = "http://10.190.11.142:8000/static/javascript/jvisio-flow/lib/jquery/jquery.min.js";
                      document.getElementsByTagName('head')[0].appendChild(script);"""
                self.driver.execute_script(jquery)
                self.result = self.driver.execute_script(js)
            if self.result != False:
                iframes = self.driver.find_elements_by_tag_name('iframe')
                if self.result['tagName'] == 'IFRAME':
                    # self.root.withdraw()
                    for i in range(len(iframes)):
                        iframe = iframes[i]
                        iframe_x = int(iframe.location['x'])
                        iframe_y = int(iframe.location['y'])
                        if abs(iframe_x - int(self.result['x'])) <= 1 and abs(iframe_y - int(self.result['y'])) <= 1:
                            try:
                                self.driver.execute_script(self.clear_js)
                            except Exception:
                                js = """var elements = document.querySelectorAll(".select-target");
                                    for (var i=0;i<elements.length;i++){
                                        var el = elements[i];
                                        var className = el.className;
                                        el.className = className.replace(" select-target","");
                                        el.className = className.replace("select-target","");
                                    }"""
                                self.driver.execute_script(js)
                            self.driver.switch_to_frame(i)
                            self.iframe.append(i)
                            if len(self.iframe) > 0:
                                self.adj_x+=iframe_x
                                self.adj_y+=iframe_y
                            else:
                                self.adj_x = iframe_x
                                self.adj_y = iframe_y
                            break
        except Exception:
            exstr = traceback.format_exc()
            try:
                js = """
                    var elements = document.querySelectorAll(".select-target");
                    for (var i=0;i<elements.length;i++){
                        var el = elements[i];
                        var className = el.className;
                        el.className = className.replace(" select-target","");
                        el.className = className.replace("select-target","");
                    }
                    var style = document.createElement('style');
                    style.type = 'text/css';
                    style.rel = 'stylesheet';
                    //for Chrome Firefox Opera Safari
                    style.appendChild(document.createTextNode('.select-target {border:2px solid #ef4300 !important}'));
                    //for WEB
                    //style.styleSheet.cssText = code;
                    var head = document.getElementsByTagName('head')[0];
                    head.appendChild(style);
                    //获取xpath
                    function readXPath(element) {
                        if (element.id !== "") {//判断id属性，如果这个元素有id，则显 示//*[@id="xPath"]  形式内容
                            return '//*[@id=\"' + element.id + '\"]';
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
                    };
                    var el_json = {};
                    var element = document.elementFromPoint(""" + str(x) + "," + str(y) + ");""""
                    if (element){
                        try{
                            el_json.x = element.getBoundingClientRect().left;
                        }catch{
                            var offset_top = element.offsetTop;
                            while(element.offsetParent != null){
                                offset_top+=element.offsetParent.offsetTop;
                            }
                            el_json.x = offset_top;
                        }
                        el_json.y = element.getBoundingClientRect().top;
                        el_json.height = element.offsetHeight;
                        el_json.width = element.offsetWidth;
                        el_json.id = element.id;
                        el_json.name = element.getAttribute('name');
                        if (el_json.name != null && el_json.name != "" && typeof(el_json.name) != "undefined"){
                            eles = document.getElementsByName(el_json.name);
                            for (var i=0;i<eles.length;i++){
                                ele = eles[i];
                                if (ele == element){
                                    el_json.name_item = i;
                                    break;
                                }
                            }
                        }
                        el_json.className = element.className;
                        if (el_json.className != null && el_json.className != "" && typeof(el_json.className) != "undefined"){
                            eles = document.getElementsByClassName(el_json.name);
                            for (var i=0;i<eles.length;i++){
                                ele = eles[i];
                                if (ele == element){
                                    el_json.className_item = i;
                                    break;
                                }
                            }
                        }
                        el_json.tagName = element.tagName;
                        if (el_json.tagName != null && el_json.tagName != "" && typeof(el_json.tagName) != "undefined"){
                            eles = document.getElementsByTagName(el_json.tagName);
                            for (var i=0;i<eles.length;i++){
                                ele = eles[i];
                                if (ele == element){
                                    el_json.tagName_item = i;
                                    break;
                                }
                            }
                        }
                        el_json.text = element.textContent;
                        if (el_json.tagName != "A"){
                            el_json.text = "";
                        }
                        el_json.xpath = readXPath(element);
                        try{
                            element.classList.add("select-target");
                        }catch{
                            element.className = "select-target";
                        }
                        return el_json;  
                    }else{
                        return false;
                    }
                """
                self.result = self.driver.execute_script(js)
                if self.result != False:
                    iframes = self.driver.find_elements_by_tag_name('iframe')
                    if self.result['tagName'] == 'IFRAME':
                        for i in range(len(iframes)):
                            iframe = iframes[i]
                            iframe_x = int(iframe.location['x'])
                            iframe_y = int(iframe.location['y'])
                            if abs(iframe_x - int(self.result['x'])) <= 1 and abs(
                                    iframe_y - int(self.result['y'])) <= 1:
                                try:
                                    self.driver.execute_script(self.clear_js)
                                except Exception:
                                    js2 = """var elements = document.querySelectorAll(".select-target");
                                        for (var i=0;i<elements.length;i++){
                                            var el = elements[i];
                                            var className = el.className;
                                            el.className = className.replace(" select-target","");
                                            el.className = className.replace("select-target","");
                                        }"""
                                    self.driver.execute_script(js2)
                                self.driver.switch_to_frame(i)
                                if len(self.iframe) > 0:
                                    self.adj_x += iframe_x
                                    self.adj_y += iframe_y
                                else:
                                    self.adj_x = iframe_x
                                    self.adj_y = iframe_y
                                self.iframe.append(i)
                                break
            except Exception:
                exstr = traceback.format_exc()
                print(exstr)

    def xFunc1(self,event):
        self.root.destroy()

    def xFunc(self,event):
        try:
            if event.keycode == 27:
                self.root.destroy()
        except Exception:
            pass

    def main(self, port, chromedriver):
        try:
            pythoncom.CoInitialize()
            option = webdriver.ChromeOptions()
            option.add_argument("disable-infobars")
            option.add_argument("--no-sandbox")
            option.HideCommandPromptWindow = True
            option.debugger_address = "127.0.0.1:" + port
            try:
                self.driver = webdriver.Chrome(chrome_options=option)
            except Exception as e:
                try:
                    self.driver = webdriver.Chrome(executable_path=chromedriver,chrome_options=option)
                except Exception:
                    exstr = str(traceback.format_exc())
                    print(exstr)
                    self.result['status'] = 'failed'
                    self.result['msg'] = str(exstr).replace('\n', '<br>')
                    return self.result
            self.title = self.driver.title
            self.hwnd = self.get_hwnd(self.title)
            self.all_handles = self.driver.window_handles
            outerHeight = self.driver.execute_script('return window.outerHeight')
            outerWidth = self.driver.execute_script('return window.outerWidth')
            innerHeight = self.driver.execute_script('return window.innerHeight')
            innerWidth = self.driver.execute_script('return window.innerWidth')
            screen_left = self.driver.execute_script('return window.screenLeft')
            screen_top = self.driver.execute_script('return window.screenTop')
            self.height = outerHeight - innerHeight + screen_top
            self.width = outerWidth - innerWidth + screen_left
            ratio_js = '''
                                    var ratio = 0;
                                    var screen = window.screen;
                                    var ua = navigator.userAgent.toLowerCase();
                                    if (window.devicePixelRatio !== undefined){
                                        ratio=window.devicePixelRatio;
                                    }else if (ua.indexOf('msie') != - 1){
                                        if (screen.deviceXDPI && screen.logicalXDPI){
                                            ratio=screen.deviceXDPI / screen.logicalXDPI;
                                        }
                                    }else if (window.outerWidth !== undefined && window.innerWidth !== undefined){
                                        ratio=window.outerWidth / window.innerWidth;
                                    }
                                    if (ratio){
                                        ratio=Math.floor(ratio * 100) / 100;
                                    }
                                    return ratio;
                        '''
            self.ratio = self.driver.execute_script(ratio_js)
            self.root = tkinter.Tk()
            # self.x = self.root.winfo_screenwidth()
            self.x = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            # 获取当前屏幕的宽
            # self.y = self.root.winfo_screenheight()
            self.y = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            hDC = win32gui.GetDC(0)
            # 横向分辨率
            HORZRES = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
            # 纵向分辨率
            VERTRES = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
            self.scalex = HORZRES / self.x
            print(self.scalex)
            self.scaley = VERTRES / self.y
            print(self.scaley)
            self.root.attributes('-toolwindow', True,
                                '-alpha', 0.3,
                                '-fullscreen', True,
                                '-topmost', True)
            self.root.overrideredirect(True)
            self.frame = tkinter.Frame(self.root, width=self.root.winfo_screenwidth(), height=self.root.winfo_screenheight())
            self.frame.bind("<Motion>", self.callBack)
            self.frame.bind("<Key>", self.xFunc)
            # self.frame.bind("<FocusOut>", self.xFunc1)
            self.frame.focus_set()
            self.frame.pack()
            self.root.bind("<Button-1>", self.clickMouse)
            self.root.mainloop()
            self.root.quit()
            self.result['status'] = 'success'
            self.driver = ''
            return self.result
        except Exception:
            exstr = str(traceback.format_exc())
            self.result['status'] = 'failed'
            self.result['msg'] = str(exstr).replace('\n','<br>')
            try:
                self.root.destroy()
            except Exception:
                pass
            return self.result

# if __name__ == '__main__':
#     target = targeting()
#     target.main()

