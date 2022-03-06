import tray
from Lbt import logger
import tkinter as tk
import win32con
import win32gui
import win32print
import os


class ServerAddressTool():
    dialog_opened = False
    code = 1

    def __init__(self, sqlite_tool):
        self.sqlite_tool = sqlite_tool

    def usr_txt(self):
        try:
            settings = {
                "server": self.server_addr.get()
            }
            result = self.sqlite_tool.start(table="client_settings", operation_type=3, settings=settings, language=None)
            if result.get("result") == "success":
                logger.info("设置服务地址成功", 10000)
                self.root.destroy()
                self.dialog_opened = False
                self.code = 0
            else:
                logger.error("设置服务地址失败", 10000)
                self.dialog_opened = False

        except Exception as e:
            logger.exception(str(e), 10000)

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

    def setting(self):
        if self.dialog_opened == False:
            self.dialog_opened = True
            self.ratio = round(self.get_screen_ratio(), 2)
            self.root = tk.Tk()
            # 获取当前屏幕大小(分辨率)
            screenwidth = self.root.winfo_screenwidth()
            screenheight = self.root.winfo_screenheight()
            self.root.title("ChocLead")
            self.root.iconbitmap(default=os.getcwd() + '\\icons\\robot.ico')
            self.root.configure(bg='white')
            self.root.attributes('-topmost', True)
            # 设置界面高宽
            width = 300 * self.ratio
            heigh = 190 * self.ratio
            # ‘窗口宽 x 窗口高 + 窗口位于x轴位置 + 窗口位于y轴位置
            self.root.geometry('%dx%d+%d+%d' % (width, heigh, (screenwidth - width) / 2, (screenheight - heigh) / 2))
            self.root.resizable(False, False)
            self.setting_list = self.read_setting_sql()
            self.server_addr = tk.StringVar(value=f'{self.setting_list[0]}')
            tk.Label(self.root, text=_("Server Address"), bg='white').place(x=10 * self.ratio, y=40 * self.ratio)

            entry_ip = tk.Entry(self.root, textvariable=self.server_addr, bd=2, width=20, relief="groove")
            entry_ip.place(x=120 * self.ratio, y=40 * self.ratio)
            # 点击按钮,执行回调函数
            bt_save = tk.Button(self.root, text=_("confirm"), command=self.usr_txt, width=10)
            bt_save.place(x=50 * self.ratio, y=130 * self.ratio)
            bt_quit = tk.Button(self.root, text=_("quit"), command=self.usr_sign_quit, width=10)
            bt_quit.place(x=170 * self.ratio, y=130 * self.ratio)
            self.root.mainloop()
            if self.code == 0:
                self.setting_list[0] = self.server_addr.get()
            return self.code, self.setting_list

    def getServerInfo(self):
        self.setting_list = self.read_setting_sql()
        retCode = 0
        if len(self.setting_list) > 0:
            retCode = 1
        return retCode, self.setting_list

    def read_setting_sql(self):
        obj_data_value = []
        obj_data = self.sqlite_tool.start(table="client_settings", operation_type=2, settings=None, language=None)
        if obj_data.get("result") != "Fail":
            data = obj_data.get("data")
            del data["system_language"]
            for key, value in data.items():
                obj_data_value.append(value)
        return obj_data_value

    def usr_sign_quit(self):
        tray.setValue("exit")
        self.root.destroy()
