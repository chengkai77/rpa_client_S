import win32gui
import sys
import tray
from utils.sqlite_tools import SqliteTools


def set_win_center(root, curWidth, curHight):
    '''
    设置窗口大小，并居中显示
    :param root:主窗体实例
   :param curWidth:窗口宽度，非必填，默认200
   :param curHight:窗口高度，非必填，默认200
   :return:无
   '''
    if not curWidth:
        '''获取窗口宽度，默认200'''
        curWidth = root.winfo_width()
    if not curHight:
        '''获取窗口高度，默认200'''
        curHight = root.winfo_height()

        # 获取屏幕宽度和高度
    scn_w, scn_h = root.maxsize()

    # 计算中心坐标
    cen_x = (scn_w - curWidth) / 2
    cen_y = (scn_h - curHight) / 2

    # 设置窗口初始大小和位置
    size_xy = '%dx%d+%d+%d' % (curWidth, curHight, cen_x, cen_y)
    root.geometry(size_xy)


def set_system_language(window, language):
    """
        设置全局语言, 退出机器人
    """
    global system_language
    if language != '' or language != None:
        sqlite_tool = SqliteTools()
        system_language = language
        window.destroy()
        sqlite_tool.start(table="client_settings", operation_type=3, settings=None, language=system_language)
        win32gui.DestroyWindow(tray.getHwnd())
        win_main = tray.main()
        win_main.quit(None)
        sys.exit(0)
