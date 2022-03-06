from Lbt import logger
import win32api
import win32con
import win32gui_struct
import win32print
import win32ui
import tkinter as tk
from utils.set_win_center import set_win_center, set_system_language
from utils.sqlite_tools import SqliteTools, get_SqlLite
from utils.LogTools import LoggerTools, get_logs_client
import os
import webbrowser

try:
    import winxpgui as win32gui
except ImportError:
    import win32gui

robot_status = "connect"
connect_status = "connect"
hwnd = ""
login_user = ""
userInfo = ""
program_name = "ChocLead.exe"
robot_title = "ChocLead"


def setLanguage(newValue):
    global system_language
    system_language = newValue


def getLanguage():
    global system_language
    return system_language


def setProgramName(newValue):
    global program_name
    program_name = newValue


def getProgramName():
    global program_name
    return program_name


def getUser():
    global login_user
    return login_user


def setUser(newValue):
    global login_user
    login_user = newValue
    try:
        import json
        login_user = json.loads(login_user)["result"]
    except:
        login_user = login_user
    userInfo = "Welcome " + str(login_user)
    get_logs_client().set_client_username( str(login_user) )


def setConnect(newValue):
    global connect_status
    connect_status = newValue


def getConnect():
    global connect_status
    return connect_status


def setValue(newValue):
    global robot_status
    robot_status = newValue


def getValue():
    global robot_status
    return robot_status


def setHwnd(newValue):
    global hwnd
    hwnd = newValue


def getHwnd():
    global hwnd
    return hwnd


def setVersion(value):
    last_number = str(value)[-1]
    version = str(value)[:-1]
    li = [str(i) + "." for i in str(version)]
    value = ''.join(li) + last_number
    global choclead_version
    choclead_version = value


class SysTrayIcon(object):
    '''TODO'''
    QUIT = 'QUIT'
    SPECIAL_ACTIONS = [QUIT]
    connect = "connect"
    FIRST_ID = 1023

    def __init__(self,
                 icon,
                 hover_text,
                 menu_options,
                 on_quit,
                 default_menu_index=None,
                 window_class_name=None, ):

        self.icon = icon
        self.hover_text = hover_text
        self.on_quit = on_quit

        menu_options = menu_options + ((_("quit_des"), self.QUIT),)
        self._next_action_id = self.FIRST_ID
        self.menu_actions_by_id = set()
        self.menu_options = self._add_ids_to_menu_options(list(menu_options))
        self.menu_actions_by_id = dict(self.menu_actions_by_id)
        del self._next_action_id

        self.default_menu_index = (default_menu_index or 0)
        self.window_class_name = window_class_name or "TecleadRobotIcon"
        message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): self.restart,
                       win32con.WM_DESTROY: self.destroy,
                       win32con.WM_COMMAND: self.command,
                       win32con.WM_USER + 20: self.notify, }
        # Register the Window class.
        window_class = win32gui.WNDCLASS()
        hinst = window_class.hInstance = win32gui.GetModuleHandle(None)
        window_class.lpszClassName = self.window_class_name
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        window_class.hbrBackground = win32con.COLOR_WINDOW
        window_class.lpfnWndProc = message_map  # could also specify a wndproc.
        classAtom = win32gui.RegisterClass(window_class)
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(classAtom,
                                          self.window_class_name,
                                          style,
                                          0,
                                          0,
                                          win32con.CW_USEDEFAULT,
                                          win32con.CW_USEDEFAULT,
                                          0,
                                          0,
                                          hinst,
                                          None)
        setHwnd(self.hwnd)
        win32gui.UpdateWindow(self.hwnd)
        self.notify_id = None
        self.refresh_icon()

        win32gui.PumpMessages()

    def _add_ids_to_menu_options(self, menu_options):
        """
            function: 依次累加 FIRST_ID + id
            params： [('断开连接', 'icons\\disconnect.ico', <bound method main.disconnect of <tray.main object at 0x000001C9ECBB4A88>>)]
            result： [('断开连接', 'icons\\disconnect.ico', <bound method main.disconnect of <tray.main object at 0x000001C9ECBB4A88>>, 1023)]
        """
        result = []
        for menu_option in menu_options:
            option_text, option_action = menu_option
            if callable(option_action) or option_action in self.SPECIAL_ACTIONS:
                self.menu_actions_by_id.add((self._next_action_id, option_action))
                result.append(menu_option + (self._next_action_id,))
            else:
                result.append((option_text,
                               self._add_ids_to_menu_options(option_action),
                               self._next_action_id))
            self._next_action_id += 1
        return result

    def refresh_icon(self):
        # Try and find a custom icon
        hinst = win32gui.GetModuleHandle(None)  # 获取win32gui模块链接库的句柄
        if os.path.isfile(self.icon):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst,
                                       self.icon,
                                       win32con.IMAGE_ICON,
                                       0,
                                       0,
                                       icon_flags)
        else:
            print("Can't find icon file - using default.")
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        if self.notify_id:
            message = win32gui.NIM_MODIFY
        else:
            message = win32gui.NIM_ADD
        self.notify_id = (self.hwnd,
                          0,
                          win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                          win32con.WM_USER + 20,
                          hicon,
                          self.hover_text)
        win32gui.Shell_NotifyIcon(message, self.notify_id)
        # win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, self.notify_id)

    def restart(self, hwnd, msg, wparam, lparam):
        self.refresh_icon()

    def destroy(self, hwnd, msg, wparam, lparam):
        if self.on_quit: self.on_quit(self)
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)  # Terminate the app.

    def notify(self, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONDBLCLK:
            pass
        elif lparam == win32con.WM_RBUTTONUP:
            self.show_menu()
        elif lparam == win32con.WM_LBUTTONUP:
            self.show_menu()
        return True

    def show_menu(self):
        menu = win32gui.CreatePopupMenu()

        self.connect = getValue()

        if self.connect == "connect":
            menu_options_list = []
            for i in range(len(self.menu_options)):
                if i == 1:
                    continue
                menu_options_list.append(self.menu_options[i])
        else:
            menu_options_list = []
            for i in range(1, len(self.menu_options)):
                menu_options_list.append(self.menu_options[i])
        self.create_menu(menu, menu_options_list)
        # win32gui.SetMenuDefaultItem(menu, 1000, 0)

        pos = win32gui.GetCursorPos()
        # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
        try:
            win32gui.SetForegroundWindow(self.hwnd)
        except Exception:
            pass
        win32gui.TrackPopupMenu(menu,
                                win32con.TPM_LEFTALIGN,
                                pos[0],
                                pos[1],
                                0,
                                self.hwnd,
                                None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

    def create_menu(self, menu, menu_options):
        maxId = 1
        for option_text, option_action, option_id in menu_options[::-1]:
            if option_id in self.menu_actions_by_id:
                if maxId < option_id + 1:
                    maxId = option_id + 1
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                wID=option_id)
                win32gui.InsertMenuItem(menu, 0, 1, item)
            else:
                submenu = win32gui.CreatePopupMenu()
                self.create_menu(submenu, option_action)
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hSubMenu=submenu)
                win32gui.InsertMenuItem(menu, 0, 1, item)

        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, maxId + 1, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING | win32con.MF_GRAYED | win32con.MF_DISABLED, maxId + 2, login_user)

    def prep_menu_icon(self, icon):
        # First load the icon.
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)
        hIcon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)

        hwndDC = win32gui.GetWindowDC(self.hwnd)
        dc = win32ui.CreateDCFromHandle(hwndDC)
        memDC = dc.CreateCompatibleDC()
        iconBitmap = win32ui.CreateBitmap()
        iconBitmap.CreateCompatibleBitmap(dc, ico_x, ico_y)
        oldBmp = memDC.SelectObject(iconBitmap)
        brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)

        win32gui.FillRect(memDC.GetSafeHdc(), (0, 0, ico_x, ico_y), brush)
        win32gui.DrawIconEx(memDC.GetSafeHdc(), 0, 0, hIcon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)

        memDC.SelectObject(oldBmp)
        memDC.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, hwndDC)

        return iconBitmap.GetHandle()

    def command(self, hwnd, msg, wparam, lparam):
        id = win32gui.LOWORD(wparam)
        self.execute_menu_option(id)

    def execute_menu_option(self, id):
        menu_action = self.menu_actions_by_id[id]
        if menu_action == self.QUIT:
            win32gui.DestroyWindow(self.hwnd)
        else:
            menu_action(self)


# Minimal self test. You'll need a bunch of ICO files in the current working
# directory in order for this to work...
class main():
    """
        action:
            break: 断开链接
            exit: 推出
            connect：重新链接
            set_up：设置
    """
    global hwnd
    dialog_opened = False

    global system_language

    is_language_win_show = False

    def __init__(self):
        # self.sqlite_tool = SqliteTools()
        self.sqlite_tool = get_SqlLite()
        self.hover_text = robot_title
        self.menu_options = (
            (_("disconnect_des"), self.disconnect),
            (_("connect_des"), self.connect),
            (_("Server_address"), self.server_address),
            (_("MyPassWord"), self.password_manage),
            (_("set_lang"), self.set_lang),
            (_("setting"), self.setting),
            (_("help_des"), self.help),
            (_("about_des"), self.about)
        )

    def start(self):
        """
        params:
            icon 机器人图标
            hover_text 机器人名称
            menu_options 菜单项
            on_quit=None, 退出按键
            default_menu_index=None,
            window_class_name=None,
        """
        self.sysTrayIcon = SysTrayIcon('icons\\robot.ico', self.hover_text, self.menu_options, on_quit=self.quit,
                                       default_menu_index=1)

    def info(self, sysTrayIcon):
        win32api.MessageBox(0, userInfo, "ChocLead",
                            win32con.MB_OK | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)

    def connect(self, sysTrayIcon):
        setValue("reconnect")

    def server_address(self, sysTrayIcon):
        try:
            # 获取本地server address
            address_ip = ""
            sql = "SELECT server_ip  from 'client_settings'"
            objs = self.sqlite_tool.cursor.execute(sql)
            self.sqlite_tool.commit()
            for item in objs:
                if str(item[0]).split("//", 1)[0] == "ws:":
                    address_ip = "http://" + str(item[0]).split('//', 1)[1]
                else:
                    address_ip = "https://" + str(item[0]).split('//', 1)[1]
            webbrowser.open(address_ip, autoraise=True)
        except Exception as e:
            logger.exception("setting: server_address" + str(e), 10000)

    def set_lang(self, sysTrayIcon):
        if self.is_language_win_show == False:
            zh_text = '中文'
            en_text = 'English'
            self.is_language_win_show = True
            window = tk.Tk()
            window.resizable(False, False)  # 窗口不可调整大小
            # 第2步，给窗口的可视化起名字
            window.title('ChocLead')
            window.update()
            window.iconbitmap(default=os.getcwd() + '\icons\\robot.ico')
            window.configure(bg='white')
            set_win_center(window, 350, 200)

            zh_button = tk.Button(window, text=zh_text, bg='white', font=('Arial', 12), width=12, height=1,
                                  command=lambda: set_system_language(window, 'zh_CN'), relief="groove")
            zh_button.grid(column=5, row=2)
            en_button = tk.Button(window, text=en_text, bg='white', font=('Arial', 12), width=12, height=1,
                                  command=lambda: set_system_language(window, 'en_US'), relief="groove")
            en_button.grid(column=5, row=4)
            col_count, row_count = window.grid_size()
            for col in range(col_count):
                window.grid_columnconfigure(col, minsize=13)
            for row in range(row_count):
                window.grid_rowconfigure(row, minsize=13)

            lable = tk.Label(window,
                             text=_("Robot must reboot to affect after changing language!"),
                             font=('楷书', 8), width=28, height=1, padx=10, pady=10, bg='white')
            lable.grid(column=5, sticky=tk.E)
            window.mainloop()
            self.is_language_win_show = False

    def password_manage(self, sysTrayIcon):
        from utils.pwd_management import UserPassManger
        tt = UserPassManger()
        tt.index_show()

    def disconnect(self, sysTrayIcon):
        setValue("break")

    def about(self, sysTrayIcon):
        global choclead_version
        test_i18n = _('aboutInfo') % choclead_version
        win32api.MessageBox(0, test_i18n, "ChocLead",
                            win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)

    def help(self, sysTrayIcon):
        # login_window.info_window("ChocLead","\t使用指导\n1.请将机器人程序置于启动状态\n2.点击重新连接可以重新连至服务器\n3.退出后请重新启动ChocLead.exe")
        win32api.MessageBox(0, _('helpInfo'), "ChocLead",
                            win32con.MB_OK | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)

    def quit(self, sysTrayIcon):
        setValue("exit")

    def usr_txt(self):
        try:
            settings = {
                "server": self.server_addr.get()
            }
            result = self.sqlite_tool.start(table="client_settings", operation_type=3, settings=settings, language=None)
            if result.get("result") == "success":
                logger.info("设置服务地址成功", 10000)
                self.root.destroy()
                win32api.MessageBox(0, _("Saved successfully, please restart the robot" + '!'), "ChocLead",
                                    win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)

                win32gui.DestroyWindow(getHwnd())
                self.quit(None)
                self.dialog_opened = False
            else:
                logger.error("设置服务地址失败", 10000)
                win32api.MessageBox(0, _("Saved failed, please restart the robot") + '!', "ChocLead",
                                    win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
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

    def setting(self, sysTrayIcon):
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
            # serverAddr = self.setting_list.get
            self.server_addr = tk.StringVar(value=f'{self.setting_list[0]}')
            tk.Label(self.root, text=_("Server Address"), bg='white').place(x=10 * self.ratio, y=40 * self.ratio)

            entry_ip = tk.Entry(self.root, textvariable=self.server_addr, bd=2, width=20, relief="groove")
            entry_ip.place(x=120 * self.ratio, y=40 * self.ratio)
            # 点击按钮,执行回调函数
            bt_save = tk.Button(self.root, text=_("save"), command=self.usr_txt, width=10)
            bt_save.place(x=50 * self.ratio, y=130 * self.ratio)
            bt_quit = tk.Button(self.root, text=_("quit"), command=self.usr_sign_quit, width=10)
            bt_quit.place(x=170 * self.ratio, y=130 * self.ratio)
            self.root.mainloop()

    def read_setting_sql(self):
        obj_data_value = []
        obj_data = self.sqlite_tool.start(table="client_settings", operation_type=2, settings=None, language=None)
        if obj_data.get("result") != "Fail":
            data = obj_data.get("data")
            del data["system_language"]
            for key, value in data.items():
                obj_data_value.append(value)
        return obj_data_value

    # setting退出
    def usr_sign_quit(self):
        self.root.destroy()
