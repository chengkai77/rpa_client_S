import win32api
import win32con
import win32gui
import win32print
import ctypes


def get_real_resolution():
    """获取真实的分辨率"""
    hwnd = 0
    hDC = win32gui.GetWindowDC(hwnd)
    # 横向分辨率
    w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
    # 纵向分辨率
    h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
    win32gui.ReleaseDC(hwnd, hDC)
    return w, h


def get_screen_size():
    """获取缩放后的分辨率"""
    # user32 = ctypes.windll.user32
    # user32.SetProcessDPIAware()
    # w = user32.GetSystemMetrics(0)
    # h = user32.GetSystemMetrics(1)
    w = win32api.GetSystemMetrics(0)
    h = win32api.GetSystemMetrics(1)
    return w, h


def get_ratio():
    try:
        w1, h1 = get_screen_size()
        w2, h2 = get_real_resolution()
        ratio = w2 / w1
    except Exception as e:
        print(e)
        ratio = 1
    return ratio
