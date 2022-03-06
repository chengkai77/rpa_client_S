import tkinter as tk
from tkinter import *
from utils.sqlite_tools import get_SqlLite
import os
import win32api
import win32con

tis_msg_menu_zh = {
    0: "无限制",
    2: "小写，大写，数字，符号中任意2种",
    3: "小写，大写，数字，符号中任意3种",
    4: "小写，大写，数字，符号全部",
}

tis_msg_menu_en = {
    0: "No limit",
    2: "2 out of lower lowercase, uppercase, number and symbol",
    3: "3 out of lower lowercase, uppercase, number and symbol",
    4: "all of lower lowercase, uppercase, number and symbol",
}


# 用户密码管理
class UserManger():
    is_user_show = False

    def __init__(self):
        self.button_type = False
        self.sqlite_tool = get_SqlLite()

    def create_user_win(self, pwd_length, rule, error_num):
        if self.is_user_show == False:
            self.is_user_show = True
            self.reset_password_window = tk.Tk()
            self.reset_password_window.title('ChocLead')
            self.reset_password_window.update()
            self.reset_password_window.iconbitmap(default=os.getcwd() + '\\icons\\robot.ico')
            self.reset_password_window.configure(bg='white')
            self.reset_password_window.attributes('-topmost', True)
            self.reset_password_window.resizable(False, False)
            self.username = tk.StringVar()
            self.old_password = tk.StringVar()
            self.new_password = tk.StringVar()
            self.confirm_password = tk.StringVar()

            screenwidth = self.reset_password_window.winfo_screenwidth()
            screenheight = self.reset_password_window.winfo_screenheight()
            width = 550
            height = 300
            x = int((screenwidth - width) / 2)
            y = int((screenheight - height) / 2)
            self.reset_password_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
            tk.Label(self.reset_password_window, text=_("username_des"), bg='white').place(x=85, y=20)
            tk.Label(self.reset_password_window, text=_("Old Password :"), bg='white').place(x=85, y=60)
            tk.Label(self.reset_password_window, text=_("New password :"), bg='white').place(x=85, y=100)
            tk.Label(self.reset_password_window, text=_("Confirm Password :"), bg='white').place(x=85, y=140)
            text_i18n = _("Password length cannot be less than %s, %s can try %s times!")
            if "密码长度不可小于" in text_i18n:
                tk.Label(self.reset_password_window,
                         text=text_i18n % (pwd_length, tis_msg_menu_zh[rule], error_num), bg='yellow',
                         font=("楷书", 8)).place(x=85, y=180)
            else:
                test_i18n = text_i18n % (pwd_length, tis_msg_menu_en[rule], error_num)
                if len(test_i18n) > 62:
                    test_i18n = test_i18n[:59] + "\n" + test_i18n[59:]

                tk.Label(self.reset_password_window,
                         text=test_i18n, bg='yellow',
                         font=("楷书", 8)).place(x=85, y=180)
            # 用户名称
            entry_name = tk.Entry(self.reset_password_window, textvariable=self.username, bd=2, relief=GROOVE,
                                  width=25)
            entry_name.place(x=230, y=20)
            # # 密码名称
            entry_name = tk.Entry(self.reset_password_window, textvariable=self.old_password, bd=2, relief=GROOVE,
                                  width=25, show='*')
            entry_name.place(x=230, y=60)
            # # login
            entry_login_name = tk.Entry(self.reset_password_window, textvariable=self.new_password, bd=2, relief=GROOVE,
                                        width=25, show='*')
            entry_login_name.place(x=230, y=100)
            # # 密码
            entry_pwd = tk.Entry(self.reset_password_window, textvariable=self.confirm_password, bd=2, relief=GROOVE,
                                 width=25, show='*')
            entry_pwd.place(x=230, y=140)
            # 数据赋值
            # # 确认 和  删除按钮
            bt_login = tk.Button(self.reset_password_window, text=_("save"), command=self.check_label, width=10)
            bt_login.place(x=120, y=240)
            bt_log_quit = tk.Button(self.reset_password_window, text=_("cancel"), command=self.close, width=10)
            bt_log_quit.place(x=300, y=240)
            # 主循环
            self.reset_password_window.mainloop()
            self.is_user_show = False
            user_dict = {
                "type": self.button_type,
                "username": self.username.get(),
                "old_password": self.old_password.get(),
                "new_password": self.new_password.get(),
                "confirm_password": self.confirm_password.get(),
            }
            return user_dict

    def check_label(self):
        if self.username.get() == '':
            win32api.MessageBox(0, _("Please fill in the user name!"), "ChocLead",
                                win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
        if self.old_password.get() == '':
            win32api.MessageBox(0, _("Please fill in the old password!"), "ChocLead",
                                win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)

        elif self.new_password.get() == '':
            win32api.MessageBox(0, _("Please fill in the new password!"), "ChocLead",
                                win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
        elif self.confirm_password.get() == '':
            win32api.MessageBox(0, _("Please confirm the password again!"), "ChocLead",
                                win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST | win32con.MB_DEFAULT_DESKTOP_ONLY)
        else:
            self.button_type = True
            self.reset_password_window.destroy()

    def close(self):
        self.reset_password_window.destroy()
