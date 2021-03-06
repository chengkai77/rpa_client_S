import base64
import os
import win32api
import win32con
import win32gui
import win32print
import time

status = 0


def setStatus(val):
    global status
    status = val


def getStatus():
    global status
    return status


class chocLeadLoading():
    def __init__(self):
        self.createPic()
        wc = win32gui.WNDCLASS()
        wc.hbrBackground = win32con.COLOR_BTNFACE + 1
        wc.hCursor = win32gui.LoadCursor(0, win32con.IDI_APPLICATION)
        wc.lpszClassName = "ChocLead Window"
        wc.lpfnWndProc = self.WndProc
        self.hinst = win32gui.GetModuleHandle(None)
        iconPathName = "icons\\loading.ico"
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        wc.hIcon = win32gui.LoadImage(self.hinst, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
        self.reg = win32gui.RegisterClass(wc)

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

    def WndProc(self, hwnd, msg, wParam, lParam):
        if msg == win32con.WM_PAINT:
            hdc, ps = win32gui.BeginPaint(hwnd)
            rect = win32gui.GetClientRect(hwnd)
            win32gui.DrawText(hdc, 'ChocLead', len('ChocLead'), rect,
                              win32con.DT_SINGLELINE | win32con.DT_CENTER | win32con.DT_VCENTER)
            win32gui.EndPaint(hwnd, ps)
        if msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
        return win32gui.DefWindowProc(hwnd, msg, wParam, lParam)

    def createPic(self):
        # 根据base64编码生成ico图标
        icon = b'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'
        if not os.path.exists("icons"):
            os.mkdir("icons")
        self.ico_path = "icons\\loading.ico"
        if not os.path.exists(self.ico_path):
            img_data = base64.b64decode(icon)
            # 注意：如果是"data:image/jpg:base64,"，那你保存的就要以png格式，如果是"data:image/png:base64,"那你保存的时候就以jpg格式。
            with open(self.ico_path, 'wb') as f:
                f.write(img_data)

    def loadingString(self, hwnd):
        global status
        while status == 0:
            status = getStatus()
            animation = ["■□□□□□□□□□", "■■□□□□□□□□", "■■■□□□□□□□", "■■■■□□□□□□", "■■■■■□□□□□", "■■■■■■□□□□",
                         "■■■■■■■□□□", "■■■■■■■■□□", "■■■■■■■■■□", "■■■■■■■■■■"]
            for i in range(len(animation)):
                time.sleep(0.1)
                win32gui.SetWindowText(hwnd, animation[i % len(animation)])

    def createWindow(self):
        global status
        if status == 0:
            dwStyle = win32con.WS_POPUP | win32con.WS_DLGFRAME | win32con.WS_VISIBLE | win32con.WS_CAPTION | win32con.WS_SYSMENU | win32con.MB_TOPMOST
            try:
                ratio = round(self.get_screen_ratio(), 2)
            except Exception:
                ratio = 1
            width = 160
            height = 30
            w = int((win32api.GetSystemMetrics(0) - width * ratio) / 2)
            h = int((win32api.GetSystemMetrics(1) - height * ratio) / 2)
            hwnd = win32gui.CreateWindow(self.reg, 'ChocLead', dwStyle, w, h, width, height, 0, 0, self.hinst, None)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                                   win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
            win32gui.SetLayeredWindowAttributes(hwnd, win32api.RGB(0, 0, 0), 80, win32con.LWA_ALPHA)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
            win32gui.UpdateWindow(hwnd)
            self.loadingString(hwnd)

# chocLeadLoading().createWindow()
