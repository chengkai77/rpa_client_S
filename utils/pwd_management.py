import time
from tkinter import ttk, Scrollbar, VERTICAL, RIGHT, Button, CENTER, BOTH, Y, GROOVE, END, Label
import tkinter as tk
import tray
from utils.UserPasswordManager import getLoginSecret_key
from utils.login_encryption import login_encrypt
from utils.sqlite_tools import SqliteTools, get_SqlLite, WriteLog
import os
import win32api
import win32con
from utils.LogTools import getWebsocket
import json
import asyncio

win_index_code = False
is_check_user = False


def setCode(value):
    global win_index_code
    win_index_code = value


def getCode():
    global win_index_code
    return win_index_code


class verification_window(tk.Frame):
    # 调用时初始化
    def __init__(self):
        global root
        root = tk.Tk()
        root.title('ChocLead')  # 标题
        root.iconbitmap(default=os.getcwd() + '\\icons\\robot.ico')
        root.configure(bg='white')
        root.resizable(False, False)
        # 窗口大小设置为150x150

        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        width = 450
        height = 250
        x = int((screenwidth - width) / 2)
        y = int((screenheight - height) / 2)
        root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        super().__init__()
        username = tray.getUser()
        self.username = tk.StringVar(value=username)
        self.password = tk.StringVar()
        self.pack()
        self.main_window()

        root.mainloop()

    # 窗口布局
    def main_window(self):
        global root

        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        width = 450
        height = 250
        x = int((screenwidth - width) / 2)
        y = int((screenheight - height) / 2)
        root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        tk.Label(root, text=_("username_des"), bg='white').place(x=60, y=50)
        tk.Label(root, text=_("Password: "), bg='white').place(x=60, y=90)
        entry_username = tk.Entry(root, textvariable=self.username, bd=2, relief=GROOVE, width=20,
                                  state="disabled")
        entry_username.place(x=160, y=50)
        entry_user_password = tk.Entry(root, textvariable=self.password, bd=2, relief=GROOVE,
                                       width=25, show='*')
        entry_user_password.place(x=160, y=90)
        bt_login = tk.Button(root, text=_("confirm"), command=self.user_confirm_button, width=10)
        bt_login.place(x=100, y=180)
        bt_logquit = tk.Button(root, text=_("cancel"), command=self.cancel_button, width=10)
        bt_logquit.place(x=250, y=180)

    # 验证函数
    def user_confirm_button(self):
        WriteLog("user_confirm_button ")
        if self.username.get() == "" or self.password.get() == "":
            win32api.MessageBox(0, _("Enter username or password!"), "ChocLead",
                                win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
        else:
            # 用户验证 --> 用户名和密码
            user_data = {
                "username": self.username.get(),
                "password": self.password.get()
            }
            self.send_message(user_data)
            root.destroy()

    def send_message(self, user_data):
        WriteLog("send_message ")
        loop_log = asyncio.new_event_loop()
        asyncio.set_event_loop(loop_log)
        loops = asyncio.get_event_loop()
        loops.run_until_complete(self.check_user_pwd_manager(user_data))

    async def check_user_pwd_manager(self, user_data):
        WriteLog("check_user_pwd_manager ")
        cred_text = user_data
        cred_text['action'] = "check_user"
        websocket_server = getWebsocket()
        try:
            await websocket_server.send(json.dumps(cred_text))
        except Exception as e:
            print(e)

    def cancel_button(self):
        root.destroy()


# 用户密码管理
class UserPassManger():
    cell_msg = None
    is_index_show = False
    is_edit_box_show = False
    is_add_user_show = False

    def __init__(self):
        self.sqlite_tool = get_SqlLite()

    def check_operate_user(self):
        setCode(False)
        global is_check_user
        if is_check_user == False:
            is_check_user = True
            self.login_check_user = verification_window()
            is_check_user = False

    def index_show(self, internal_call=0):
        if internal_call == 0:
            self.check_operate_user()
            time.sleep(2)
            code = getCode()
        else:
            code = True
        if self.is_index_show == False and code == True:
            self.win = tk.Tk()  # 窗口
            self.win.title('ChocLead')  # 标题
            self.win.iconbitmap(default=os.getcwd() + '\\icons\\robot.ico')
            screenwidth = self.win.winfo_screenwidth()  # 屏幕宽度
            screenheight = self.win.winfo_screenheight()  # 屏幕高度
            width = 550
            height = 340
            x = int((screenwidth - width) / 2)
            y = int((screenheight - height) / 2)
            self.win.geometry('{}x{}+{}+{}'.format(width, height, x, y))  # 大小以及位置
            self.win.resizable(False, False)
            tabel_frame = tk.Frame(self.win, bd=16)
            tabel_frame.grid(column=5, row=2)
            for row in range(4):
                tabel_frame.grid_rowconfigure(row)

            for column in range(4):
                tabel_frame.grid_columnconfigure(column)

            # xscroll = Scrollbar(tabel_frame, orient=HORIZONTAL)
            yscroll = Scrollbar(tabel_frame, orient=VERTICAL)

            columns = [_("#"), _('Node name'), _('LoginAccount'), str(_("Password: ")).replace(": ", "")]
            self.table = ttk.Treeview(
                master=tabel_frame,  # 父容器
                height=10,  # 表格显示的行数,height行
                columns=columns,  # 显示的列
                show='headings',  # 隐藏首列
                # xscrollcommand=xscroll.set,  # x轴滚动条
                yscrollcommand=yscroll.set,  # y轴滚动条
            )

            for column in columns:
                if column == "#":
                    self.table.heading(column=column, text=column, anchor=CENTER)  # 定义表头
                    self.table.column(column=column, width=60, anchor=CENTER)  # 定义列
                else:
                    self.table.heading(column=column, text=column, anchor=CENTER)  # 定义表头
                    self.table.column(column=column, width=150, anchor=CENTER)  # 定义列

            # xscroll.config(command=table.xview)
            # xscroll.pack(side=BOTTOM, fill=X)
            yscroll.config(command=self.table.yview)
            yscroll.pack(side=RIGHT, fill=Y)
            self.table.pack(fill=BOTH, expand=True)

            self.index_data_insert()

            Label(self.win, text=_("Click row to edit or delete!")).grid(column=5)

            btn_frame = tk.Frame(self.win)
            btn_frame.grid(column=5, row=10)
            Button(btn_frame, text=_("New growth"), width=12, command=self.add_user_show).pack()
            self.table.bind("<<TreeviewSelect>>", self.mouse_click_event)
            self.win.mainloop()
            global is_check_user
            is_check_user = False

    def edit_box_show(self, cell_msg):
        self.win.destroy()
        self.cell_msg = cell_msg.get("values")
        if self.is_edit_box_show == False:
            self.is_edit_box_show = True
            self.edit_window = tk.Tk()
            self.edit_window.title('ChocLead')
            # self.edit_window.update()
            self.edit_window.iconbitmap(default=os.getcwd() + '\\icons\\robot.ico')
            self.edit_window.configure(bg='white')
            self.edit_window.resizable(False, False)
            self.edit_name = tk.StringVar(value=self.cell_msg[1])
            self.edit_login_name = tk.StringVar(value=self.cell_msg[2])
            self.edit_password = tk.StringVar(value=self.cell_msg[3])

            screenwidth = self.edit_window.winfo_screenwidth()
            screenheight = self.edit_window.winfo_screenheight()
            width = 550
            height = 230
            x = int((screenwidth - width) / 2)
            y = int((screenheight - height) / 2)
            self.edit_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
            tk.Label(self.edit_window, text=_("Node name") + ":", bg='white').place(x=90, y=30)
            tk.Label(self.edit_window, text=_("LoginAccount") + ":", bg='white').place(x=90, y=60)
            tk.Label(self.edit_window, text=_("Password: "), bg='white').place(x=90, y=90)
            # 密码名称
            entry_name = tk.Entry(self.edit_window, textvariable=self.edit_name, bd=2, relief=GROOVE, width=25)
            entry_name.place(x=200, y=30)
            # login
            entry_login_name = tk.Entry(self.edit_window, textvariable=self.edit_login_name, bd=2, relief=GROOVE,
                                        width=25)
            entry_login_name.place(x=200, y=60)
            # 密码
            entry_pwd = tk.Entry(self.edit_window, textvariable=self.edit_password, bd=2, relief=GROOVE, width=25,
                                 show='*')
            entry_pwd.place(x=200, y=90)
            # 数据赋值
            # # 确认 和  删除按钮
            bt_login = tk.Button(self.edit_window, text=_("strike out"), fg="red", command=self.delete_button,
                                 width=10)
            bt_login.place(x=150, y=160)
            bt_logquit = tk.Button(self.edit_window, text=_("renew"), command=self.confirm_button, width=10)
            bt_logquit.place(x=300, y=160)
            # 主循环
            self.edit_window.mainloop()

    def mouse_click_event(self, event):
        cell_xid = event.widget.selection()
        cell_msg = self.table.item(cell_xid[0])
        self.edit_box_show(cell_msg)

    def index_data_insert(self):
        result = self.sqlite_tool.start(table="password_manage", operation_type=4, settings=None, language=None)
        if result:
            for index, data in enumerate(result):
                self.table.insert('', END, values=data)  # 添加数据到末尾

    def confirm_button(self):
        self.edit_window.destroy()
        self.is_edit_box_show = False
        user_msg = self.cell_msg
        new_data = {
            "id": user_msg[0],
            "name": self.edit_name.get(),
            "login_name": self.edit_login_name.get()
        }
        if self.edit_password.get() == "**********":
            new_data['password'] = user_msg[4]
        else:
            password = login_encrypt(self.edit_password.get(), getLoginSecret_key())
            new_data['password'] = password
        # 修改sqlite数据库对应的信息
        result = self.sqlite_tool.start(table="password_manage", operation_type=7, settings=new_data, language=None)
        if result.get("result") == "failed":
            win32api.MessageBox(0, _("Operation failed, try again!"), "ChocLead",
                                win32con.MB_ICONINFORMATION | win32con.MB_OKCANCEL | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
        self.index_show(internal_call=1)

    def delete_button(self):
        self.edit_window.destroy()
        self.is_edit_box_show = False
        user_id = self.cell_msg[0]
        # 删除sqlite数据库对应的信息
        result = self.sqlite_tool.start(table="password_manage", operation_type=6, settings=None, language=user_id)
        if result.get("result") == "failed":
            win32api.MessageBox(0, _("Failed to delete, please try again!"), "ChocLead",
                                win32con.MB_ICONINFORMATION | win32con.MB_OKCANCEL | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
        else:
            # 删除关联子表信息
            del_sql = 'delete from %s where password_id = "%s"' % ("password_manage_node", user_id)
            self.sqlite_tool.cursor.execute(del_sql)
            self.sqlite_tool.commit()
            self.index_show(internal_call=1)

    def add_user_show(self, operation_type=0, add_name=""):
        self.add_button_success = False
        if operation_type == 0:
            self.win.destroy()
        if self.is_add_user_show == False:
            try:
                self.is_add_user_show = True
                self.add_window = tk.Tk()
                self.add_window.title('ChocLead')
                # self.add_window.update()
                self.add_window.iconbitmap(default=os.getcwd() + '\\icons\\robot.ico')
                self.add_window.configure(bg='white')
                self.add_window.resizable(False, False)
                self.add_name = tk.StringVar()
                if operation_type == 1:
                    self.add_name = tk.StringVar(value=add_name)
                self.add_login_name = tk.StringVar()
                self.add_password = tk.StringVar()
                self.add_login_type = tk.StringVar()

                screenwidth = self.add_window.winfo_screenwidth()
                screenheight = self.add_window.winfo_screenheight()
                width = 550
                height = 230
                x = int((screenwidth - width) / 2)
                y = int((screenheight - height) / 2)
                self.add_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
                tk.Label(self.add_window, text=_("Node name") + ":", bg='white').place(x=90, y=30)
                tk.Label(self.add_window, text=_("LoginAccount") + ":", bg='white').place(x=90, y=60)
                tk.Label(self.add_window, text=_("Password: "), bg='white').place(x=90, y=90)
                # 密码名称
                entry_name = tk.Entry(self.add_window, textvariable=self.add_name, bd=2, relief=GROOVE, width=25)
                entry_name.place(x=200, y=30)
                # 密码
                entry_pwd = tk.Entry(self.add_window, textvariable=self.add_login_name, bd=2, relief=GROOVE, width=25)
                entry_pwd.place(x=200, y=60)

                entry_pwd = tk.Entry(self.add_window, textvariable=self.add_password, bd=2, relief=GROOVE, width=25,
                                     show='*')

                entry_pwd.place(x=200, y=90)

                # 确认按钮
                bt_logquit = tk.Button(self.add_window, text=_("save"), command=self.add_button_event, width=10)
                bt_logquit.place(x=300, y=160)

                # 主循环
                self.add_window.mainloop()
                self.is_add_user_show = False
                data = {}
                if self.add_button_success:
                    data = {
                        "name": self.add_name.get(),
                        "login_name": self.add_login_name.get(),
                        "password": login_encrypt(self.add_password.get(), getLoginSecret_key())
                    }
                return data
            except Exception as e:
                print("eeee", e)

    def add_button_event(self):
        data = {
            "name": self.add_name.get(),
            "login_name": self.add_login_name.get(),
            "password": login_encrypt(self.add_password.get(), getLoginSecret_key())
        }
        is_insert_pwd_sqlite = False
        for item in data.items():
            if item[1] == "":
                is_insert_pwd_sqlite = True
                msg = _("Please enter %s !") % item[0]
                win32api.MessageBox(0, msg, "ChocLead",
                                    win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
                break
        if is_insert_pwd_sqlite == False:
            self.add_window.destroy()
            # 密码加密
            # 判断该类型是否以存在
            sql = 'SELECT *  from %s where name="%s"' % ("password_manage", data['name'])
            res = self.sqlite_tool.cursor.execute(sql)
            num = 0
            for item in res:
                num += 1
            if num < 1:
                result = self.sqlite_tool.start(table="password_manage", operation_type=5, settings=data, language=None)
                if result.get("result") == "failed":
                    win32api.MessageBox(0, _("Failed to add, please try again!"), "ChocLead",
                                        win32con.MB_ICONINFORMATION | win32con.MB_OKCANCEL | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)

                else:
                    self.add_button_success = True
                    self.is_add_user_show = False
                    self.index_show(internal_call=1)
            else:
                win32api.MessageBox(0, _("Data already exists!"), "ChocLead",
                                    win32con.MB_ICONINFORMATION | win32con.MB_OKCANCEL | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)

    def close(self):
        self.win.destroy()

    def cancel_button(self):
        self.login_check_user.destroy()
