import threading
import traceback
import pythoncom
import win32com.client, win32gui
import time
import win32print
import win32con
from pynput import mouse
from pynput.mouse import Controller
import win32api
from Lbt import logger

click_event = 0
exit = 0


def setClickEvent(value):
    global click_event
    click_event = value


def getClickEvent():
    global click_event
    return click_event


def on_click(x, y, button, pressed):
    if "left" in str(button):
        if not pressed:
            setClickEvent(1)
            # Stop listener
            return False
    elif "right" in str(button):
        if not pressed:
            setClickEvent(1)
            # Stop listener
            return False


def on_press(key):
    global exit
    try:
        if format(key) == "Key.esc":
            exit = 1
            return False
    except Exception:
        pass


class targeting():
    def __init__(self):
        self.x = ""
        self.y = ""
        self.ratio = ""
        self.session = ""
        self.target_id = ""
        self.sap_hwnd = ""
        self.sap_array = []
        self.platform_hwnd = ""
        self.session_list = []
        self.numIdx = 10
        self.json = {}

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

    def get_screen_ratio(self):
        hDC = win32gui.GetDC(0)
        dpiA = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES) / win32print.GetDeviceCaps(hDC, win32con.HORZRES)
        dpiB = win32print.GetDeviceCaps(hDC, win32con.LOGPIXELSX) / 0.96 / 100
        if dpiA == 1:
            return dpiB
        elif dpiB == 1:
            return dpiA
        elif dpiA == dpiB:
            return dpiA

    def get_real_resolution(self):
        """获取真实的分辨率"""
        hDC = win32gui.GetDC(0)
        # 横向分辨率
        w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
        # 纵向分辨率
        h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
        return w, h

    def get_screen_size(self):
        """获取缩放后的分辨率"""
        w = win32api.GetSystemMetrics(0)
        h = win32api.GetSystemMetrics(1)
        return w, h

    def findSapCompent(self, sapId):
        try:
            gui = self.session.findById(sapId)
            sap_children_count = gui.Children.Count
            for i in range(0, sap_children_count):
                gui_component = gui.Children[i]
                gui_component_id = gui_component.Id
                gui_component_screenLeft = gui_component.ScreenLeft
                gui_component_screenTop = gui_component.ScreenTop
                gui_component_height = gui_component.Height
                gui_component_width = gui_component.Width
                if gui_component_screenLeft < self.x * self.ratio and gui_component_screenTop < self.y * self.ratio and gui_component_screenLeft + gui_component_width > self.x * self.ratio and gui_component_screenTop + gui_component_height > self.y * self.ratio:
                    try:
                        sub_gui = self.session.findById(gui_component_id)
                        sap_sub_children_count = sub_gui.Children.Count
                        if sap_sub_children_count > 0:
                            sapId = self.findSapCompent(gui_component_id)
                    except Exception:
                        sapId = gui_component_id
                    break
        except Exception:
            pass
        return sapId

    def listen(self):
        with mouse.Listener(on_click=on_click) as mouseListener:
            mouseListener.join()
        # with keyboard.Listener(on_press=on_press) as keyboardListener:
        #     keyboardListener.join()

    def main(self):
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection_num = application.Children.Count
            if connection_num == 0:
                return "No connection!"
            for i in range(connection_num):
                connection = application.Children(i)
                flag = 0
                wait = 0
                while flag == 0 and wait < 10:
                    try:
                        session_num = connection.Children.Count
                        for j in range(session_num):
                            session = connection.Children(j)
                            self.session_list.append(session)
                            self.json[session.Id + "-connection"] = i
                            self.json[session.Id + "-session"] = j
                        flag = 1
                    except Exception:
                        wait += 1
                        time.sleep(0.5)
            self.platform_hwnd = win32gui.GetForegroundWindow()
            sap_window = win32gui.FindWindow("SAP_FRONTEND_SESSION", None)
            sap_log = win32gui.FindWindow("SAP_FRONTEND_SESSION", "SAP")
            sap_window2 = win32gui.FindWindowEx(0, sap_log, "SAP_FRONTEND_SESSION", None)
            next_sap_window = 1
            if sap_window != 0:
                self.sap_array.append(sap_window)
                current_sap_window = sap_window
                while next_sap_window != 0:
                    try:
                        next_sap_window = win32gui.FindWindowEx(0, current_sap_window, "SAP_FRONTEND_SESSION", None)
                        if next_sap_window != 0:
                            self.sap_array.append(next_sap_window)
                            current_sap_window = next_sap_window
                    except Exception:
                        break
            if sap_window != 0 and sap_window != sap_log:
                self.sap_hwnd = sap_window
            elif sap_window != 0 and sap_window2 != 0:
                self.sap_hwnd = sap_window2
            else:
                if sap_log != 0:
                    self.sap_hwnd = sap_log
                else:
                    return "No existing SAP window!"
            if len(self.sap_array) > 1:
                win32gui.SetWindowPos(self.platform_hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
                for sap_hwnd in self.sap_array:
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shell.SendKeys('^')
                    win32gui.SetForegroundWindow(sap_hwnd)
                    for session in self.session_list:
                        if session.isActive == True:
                            self.json[str(sap_hwnd) + "-connection"] = self.json[session.Id + "-connection"]
                            self.json[str(sap_hwnd) + "-session"] = self.json[session.Id + "-session"]
                            self.json[sap_hwnd] = session
                            if sap_hwnd == self.sap_hwnd:
                                self.session = session
                            break
                win32gui.SetWindowPos(self.platform_hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
                win32gui.SetWindowPos(self.sap_hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
                win32gui.SetWindowPos(self.sap_hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
            else:
                self.session = session
                self.json[str(self.sap_hwnd) + "-connection"] = self.json[self.session.Id + "-connection"]
                self.json[str(self.sap_hwnd) + "-session"] = self.json[self.session.Id + "-session"]
                self.json[self.sap_hwnd] = self.session
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('^')
            win32gui.SetForegroundWindow(self.sap_hwnd)
            try:
                self.ratio = round(self.get_screen_ratio(), 2)
            except Exception:
                real_resolution = self.get_real_resolution()
                screen_size = self.get_screen_size()
                self.ratio = round(real_resolution[0] / screen_size[0], 2)
            # print(self.ratio)
            self.session.lockSessionUI()
            mouse = Controller()
            global click_event, exit
            click_event = 0
            t = threading.Thread(target=self.listen)
            t.start()
            # state_left = win32api.GetKeyState(0x01)  # Left button down = 0 or 1. Button up = -127 or -128
            while True:
                if exit == 0:
                    current_hwnd = win32gui.GetForegroundWindow()
                    try:
                        ele = session.findById("wnd[1]")
                    except Exception:
                        ele = ""
                    if current_hwnd == self.sap_hwnd or ele != "":
                        # a = win32api.GetKeyState(0x01)
                        a = getClickEvent()
                        if a == 1:  # Button state changed
                            try:
                                self.session.unlockSessionUI()
                                self.session.findById(self.target_id).Visualize(False)
                            except Exception:
                                exstr = str(traceback.format_exc())
                                logger.warning("sap target click error: " + exstr, 10010)
                            break
                        self.x = mouse.position[0] / self.ratio
                        self.y = mouse.position[1] / self.ratio
                        eles = self.session.findByPosition(self.x, self.y, False)
                        if eles != None:
                            sap_id = eles[0]
                            try:
                                sap_control_type = self.session.findById(sap_id).Type
                            except Exception:
                                sap_control_type = ""
                            # print(sap_control_type)
                            if sap_control_type == "GuiTableControl" or sap_control_type == "GuiShell":
                                sap_id = self.findSapCompent(sap_id)
                            #     col = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").CurrentCellColumn
                            #     row = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").CurrentCellRow
                            #     print(col, row)
                            if self.target_id != sap_id:
                                if self.target_id != "":
                                    self.session.findById(self.target_id).Visualize(False)
                                self.session.findById(sap_id).Visualize(True)
                                self.target_id = sap_id
                    else:
                        try:
                            self.session.unlockSessionUI()
                            if self.target_id != "":
                                self.session.findById(self.target_id).Visualize(False)
                        except Exception:
                            exstr = str(traceback.format_exc())
                            logger.warning("sap target click error: " + exstr, 10010)
                            pass
                        if current_hwnd in self.sap_array:
                            self.sap_hwnd = current_hwnd
                            self.session = self.json[self.sap_hwnd]
                            self.session.lockSessionUI()
                            self.target_id = ""
                else:
                    print("esc exit")
                    self.session.unlockSessionUI()
                    self.target_id = ""
                    break
            try:
                win32gui.SetForegroundWindow(self.platform_hwnd)
            except Exception:
                exstr = traceback.format_exc()
                logger.error("set focus of robot page error: " + exstr, 10010)
            return self.target_id
        except Exception:
            exstr = str(traceback.format_exc())
            logger.error("sap connection error: " + exstr, 10010)
            return "No connection!"
# target = targeting()
# target.main()
