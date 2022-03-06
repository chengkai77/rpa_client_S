# encoding:utf-8
import threading
import traceback

import win32api
import win32com.client
import win32con
import win32gui
import uiautomation as auto
import win32print
from pynput import mouse, keyboard
import tkinter as tk
import time

from Lbt.EXCEL.gen import getting_location_range, column_to_name
from Lbt.JavaBridge import javaaccessbridge
import PyHook3
import pythoncom

click_event = 0
exit = 0
last_java_hwnd = None
java_hwnd = None


def refresh_all():
    hwnd_title = dict()

    def get_all_hwnd(hwnd, mouse):
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})

    win32gui.EnumWindows(get_all_hwnd, 0)
    for hwnd, title in hwnd_title.items():
        try:
            win32gui.InvalidateRect(hwnd, None, True)
            win32gui.RedrawWindow(hwnd, None, None,
                                  win32con.RDW_FRAME | win32con.RDW_INVALIDATE | win32con.RDW_UPDATENOW | win32con.RDW_ALLCHILDREN)
        except Exception as e:
            continue


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


class uiSpy():
    def __init__(self):
        self.top_hwnd = []
        self.position = []
        self.ratio = 1
        self.x = ""
        self.y = ""
        self.platform_hwnd = win32gui.GetForegroundWindow()
        self.rectanglePen = win32gui.CreatePen(win32con.PS_SOLID, 2, win32api.RGB(255, 0, 0))
        self.hwnd = win32gui.GetForegroundWindow()
        self.root = ""
        self.label = ""
        self.content = ""
        self.left = 0
        self.right = 0
        self.top = 0
        self.bottom = 0
        self.control = ""
        self.left = 0
        self.top = 0

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

    def highlightWindow(self):
        hwnd = win32gui.GetDesktopWindow()
        hwndDC = win32gui.GetWindowDC(hwnd)
        if hwndDC:
            hPen = win32gui.CreatePen(win32con.PS_SOLID, 2, win32api.RGB(255, 0, 0))
            win32gui.SelectObject(hwndDC, hPen)
            hBrush = win32gui.GetStockObject(win32con.HOLLOW_BRUSH)
            win32gui.SelectObject(hwndDC, win32gui.GetStockObject(win32con.HOLLOW_BRUSH))
            win32gui.Rectangle(hwndDC, self.position[0], self.position[1], self.position[2], self.position[3])
            win32gui.DeleteObject(hPen)
            win32gui.DeleteObject(hBrush)
            win32gui.DeleteObject(hwndDC)
            win32gui.ReleaseDC(self.hwnd, hwndDC)

    def refreshWindow(self):
        hwnd = win32gui.GetForegroundWindow()
        win32gui.InvalidateRect(hwnd, None, True)
        win32gui.UpdateWindow(hwnd)
        win32gui.RedrawWindow(hwnd, None, None,
                              win32con.RDW_FRAME | win32con.RDW_INVALIDATE | win32con.RDW_UPDATENOW | win32con.RDW_ALLCHILDREN)
        refresh_all()

    def find_subHandle(self, control, x, y):
        children = control.GetChildren()
        for child in children:
            try:
                rect = child.BoundingRectangle
                left = rect.left
                top = rect.top
                right = rect.right
                bottom = rect.bottom
                if x < right and x > left and y < bottom and y > top:
                    son_children = child.GetChildren()
                    if len(son_children) > 0:
                        son_child = self.find_subHandle(child, x, y)
                        return son_child
                    else:
                        return child
            except Exception:
                pass
        return control

    def find_idxSubHandle(self, pHandle, winClass, sonHandle):
        index = 1
        handle = win32gui.FindWindowEx(pHandle, 0, winClass, None)
        if handle:
            while handle != sonHandle:
                try:
                    handle = win32gui.FindWindowEx(pHandle, handle, winClass, None)
                    index += 1
                except Exception:
                    index = 0
                    break
        else:
            index = 0
        return index

    def getNum(self):
        i = 0
        element = self.control
        while element:
            element = element.GetPreviousSiblingControl()
            if element:
                i += 1
        return i

    def onMouseEvent(self, event):
        setClickEvent(1)
        return False

    def onMouseMove(self, event):
        return False

    def listen(self):
        # with mouse.Listener(on_click=on_click) as mouseListener:
        #     mouseListener.join()
        # with keyboard.Listener(on_press=on_press) as keyboardListener:
        #     keyboardListener.join()
        self.hm = PyHook3.HookManager()
        self.hm.MouseLeftDown = self.onMouseEvent
        self.hm.HookMouse()
        setHm(self.hm)
        pythoncom.PumpMessages()

    def main(self):
        global click_event, exit
        click_event = 0
        exit = 0
        t = threading.Thread(target=self.listen)
        t.start()
        platform = win32gui.GetForegroundWindow()
        # state_left = win32api.GetKeyState(0x01)
        # try:
        #     self.ratio = round(self.get_screen_ratio(), 2)
        # except Exception:
        #     real_resolution = self.get_real_resolution()
        #     screen_size = self.get_screen_size()
        #     self.ratio = round(real_resolution[0] / screen_size[0], 2)
        try:
            javaaccessbridge.start_java_application_connection()
        except Exception as e:
            print(str(e))
        while True:
            if exit == 0:
                """
                增加java应用的支持，判断是否是java window,如果有是走新代码，否则老代码
                """
                x, y = win32api.GetCursorPos()
                hwnd = win32gui.WindowFromPoint((x, y))
                # print(javaaccessbridge.get_window_info(hwnd))
                # try:
                #     hwnd_title = javaaccessbridge.get_window_info(hwnd)["title"]
                # except Exception as e:
                #     #print(str(e))
                #     continue
                global last_java_hwnd, java_hwnd
                # if hwnd_title != "rpa_ui_shadow":
                #     is_java_window = javaaccessbridge.is_java_window_by_hwnd(hwnd)
                #     if is_java_window:
                #         java_hwnd = hwnd
                #     else:
                #         java_hwnd = None
                # 遮罩被去掉了
                try:
                    is_java_window = javaaccessbridge.is_java_window_by_hwnd(hwnd)
                    print("is_java_window:" + str(is_java_window))
                    if is_java_window:
                        java_hwnd = hwnd
                    else:
                        java_hwnd = None
                except Exception:
                    is_java_window = ""
                # java 路径
                if is_java_window:
                    try:
                        # 减少刷新数据的次数
                        if last_java_hwnd != java_hwnd:
                            last_java_hwnd = java_hwnd
                            javaaccessbridge.save_properties = None
                            # javaaccessbridge.ui_root = None
                            javaaccessbridge.refreshWindow(java_hwnd)
                            # javaaccessbridge.set_ui_shadow(self.root)
                            javaaccessbridge.ac_child_list = []
                            javaaccessbridge.last_rect = None
                            jab = javaaccessbridge.JABContext(hwnd=java_hwnd, vmID=None, accContext=None)
                            javaaccessbridge.get_all_ac_infos(jab)
                        else:
                            mouse_x, mouse_y = win32gui.GetCursorPos()
                            javaaccessbridge.find_target(mouse_x, mouse_y)
                            if click_event == 1:
                                self.hm.UnhookMouse()  # 取消鼠标监听
                                setHm("")
                                ac_properties = jab.get_target_accessible_context_properties(
                                    javaaccessbridge.jab_ac_target)
                                javaaccessbridge.hwnd_dict["java_app"] = 1
                                save_properties = {"hwnd": javaaccessbridge.hwnd_dict, "Nodes": ac_properties}
                                element_information = javaaccessbridge.get_element_information(
                                    javaaccessbridge.jab_ac_target)
                                offsetX = mouse_x - element_information["bounds"]["x"]
                                offsetY = mouse_y - element_information["bounds"]["y"]
                                image_width = element_information["bounds"]["width"]
                                image_height = element_information["bounds"]["height"]
                                print("target_parents: " + str(save_properties))
                                result = {'list': str(save_properties), 'offsetX': offsetX, 'offsetY': offsetY,
                                          'width': image_width, 'height': image_height}
                                # get text 专用属性
                                result["name"] = javaaccessbridge.accessilbe_context_get_text_java_api(
                                    javaaccessbridge.jab_ac_target)
                                # self.result = result
                                print("完成target目标:" + str(result))
                                javaaccessbridge.save_properties = None
                                # javaaccessbridge.ui_root = None
                                last_java_hwnd = None
                                refresh_all()
                                # self.root.destroy()
                                return result
                                # break
                    except Exception as e:
                        print(str(traceback.format_exc()))
                else:
                    """
                    句柄
                    """
                    try:
                        h = win32gui.GetForegroundWindow()
                        # if h != self.hwnd:
                        #     self.hwnd = h
                        #     self.root.geometry("0x0")
                        #     self.root.update()
                        #     self.control = ""
                        #     try:
                        #         self.refreshWindow()
                        #     except Exception as e:
                        #         print(e)
                        # if self.left:
                        #     if x < self.left or x > self.right or y < self.top or y > self.bottom:
                        #         self.root.geometry("0x0")
                        #         self.root.update()
                        x, y = win32api.GetCursorPos()
                        control = auto.ControlFromCursor()
                        # if len(control.GetChildren()) == 0:
                        #     if x > self.right or x < self.left or y > self.bottom or y < self.top or not self.left:
                        #         position_rect = control.BoundingRectangle
                        #         self.left = int(position_rect.left) - 1
                        #         self.top = int(position_rect.top) - 1
                        #         self.right = int(position_rect.right) + 1
                        #         self.bottom = int(position_rect.bottom) + 1
                        #         self.root.attributes('-toolwindow', True, '-alpha', 0.3, '-topmost', True)
                        #         alignstr = '%dx%d+%d+%d' % (self.right - self.left, self.bottom - self.top, self.left, self.top)
                        #         self.root.geometry(alignstr)
                        #         self.root.update()
                        position_rect = control.BoundingRectangle
                        left = int(position_rect.left) - 1
                        top = int(position_rect.top) - 1
                        right = int(position_rect.right) + 1
                        bottom = int(position_rect.bottom) + 1
                        position = []
                        position.append(left)
                        position.append(top)
                        position.append(right)
                        position.append(bottom)
                        # if left > self.left and left < self.right and top > self.top and top < self.bottom:
                        #     position.append(self.left)
                        #     position.append(self.top)
                        #     position.append(self.right)
                        #     position.append(self.bottom)
                        # else:
                        #     position.append(left)
                        #     position.append(top)
                        #     position.append(right)
                        #     position.append(bottom)
                        if position != self.position:
                            if len(self.position) > 0:
                                try:
                                    self.refreshWindow()
                                except Exception as e:
                                    print(e)
                            self.position = position
                            try:
                                self.highlightWindow()
                            except Exception as e:
                                print(e)
                        if self.x != x or self.y != y:
                            self.x = x
                            self.y = y
                            try:
                                self.highlightWindow()
                            except Exception as e:
                                print(e)
                            # try:
                            #     self.label["text"] = str(x) + "," + str(y)
                            #     self.label.place(x=x - self.left, y=y - self.top)
                            #     self.root.update()
                            # except Exception:
                            #     pass
                        clickEvent = getClickEvent()
                        if clickEvent == 1:
                            self.hm.MouseMove = self.onMouseMove  # 禁用鼠标悬浮
                            time.sleep(0.2)
                            # 重新获取元素，防止金蝶遮罩情况
                            screenWidth, screenHeight = auto.GetScreenSize()
                            auto.MoveTo(screenWidth, screenHeight, moveSpeed=0)
                            control = auto.ControlFromPoint(x, y)
                            control = self.find_subHandle(control, x, y)
                            # auto.MoveTo(x, y, moveSpeed=0)
                            self.hm.UnhookMouse()  # 取消鼠标监听
                            setHm("")
                            # print(control)
                            # 获取句柄元素所有文本
                            result = {}
                            active_hwnd = win32gui.GetForegroundWindow()
                            if win32gui.GetClassName(
                                    active_hwnd) == "XLMAIN" and control.ControlTypeName == "CustomControl":
                                # 如果control为Excel中的图表对象，则返回图表中系列项名称
                                try:
                                    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
                                except Exception as e:
                                    excel = win32com.client.Dispatch('Excel.Application')
                                result["series_name"] = []
                                result["selected_cells"] = []
                                try:
                                    active_workbook = excel.ActiveWorkbook
                                    workbook = active_workbook
                                    sheet = workbook.ActiveSheet
                                    chart_name = control.Name
                                    count = sheet.ChartObjects(chart_name).Chart.FullSeriesCollection().Count
                                    for i in range(1, count + 1):
                                        result["series_name"].append(
                                            sheet.ChartObjects(chart_name).Chart.FullSeriesCollection(i).Name)
                                        A = sheet.ChartObjects(chart_name).Chart.FullSeriesCollection(i).Formula
                                        index1 = A.rfind("!")
                                        index2 = A.rfind(",")
                                        B = A[index1 + 1:index2]
                                        B = B.replace("$", "")
                                        result["selected_cells"].append(B)
                                    excel = None
                                except Exception as e:
                                    excel = None
                                try:
                                    result["sheet_name"] = sheet.Name
                                except Exception:
                                    result["sheet_name"] = ""
                            try:
                                Text = control.GetTextPattern().DocumentRange.GetText(-1)
                                result['text'] = Text
                            except Exception:
                                result['text'] = ""
                            try:
                                Value = control.GetValuePattern().Value
                                result['value'] = Value
                            except Exception:
                                result['value'] = ""
                            finally:
                                if not result['value']:
                                    try:
                                        Value = control.GetPropertyValue(30045)
                                        result['value'] = Value
                                    except Exception:
                                        result['value'] = ""
                            try:
                                Name = control.Name
                                result['name'] = Name
                            except Exception:
                                result['name'] = ""
                            try:
                                HelpText = control.GetPropertyValue(30013)
                                result['helptext'] = HelpText
                            except Exception:
                                result['helptext'] = ""
                                # 获取数据透视表的数据源的Sheet名字并返回
                            # try:
                            # 方法一：
                            # sheet_name_data = control.GetParentControl().GetParentControl().GetParentControl().Name
                            # sheet_name = sheet_name_data.replace('工作表 ', '')
                            # 方法二：
                            #     try:
                            #         excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
                            #     except Exception as e:
                            #         excel = win32com.client.Dispatch('Excel.Application')
                            #     sheet = excel.ActiveWorkbook.ActiveSheet
                            #     sheet_name = sheet.Name
                            #     result['pivot_sheet'] = sheet_name
                            #     print(sheet_name)
                            # except:
                            #     result['pivot_sheet'] = ''
                            # 根据Excel数据透视表获取透视表名
                            # try:
                            #     pivot_table_name = control.GetParentControl().Name
                            #     result['pivot_table_name'] = pivot_table_name
                            # except:
                            #     result['pivot_table_name'] = ''
                            if win32gui.GetClassName(
                                    active_hwnd) == "XLMAIN" and control.ControlTypeName == "DataItemControl":
                                try:
                                    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
                                except Exception as e:
                                    excel = win32com.client.Dispatch('Excel.Application')
                                try:
                                    # 获取工作表中的原数据的范围
                                    ws = excel.ActiveWorkbook.ActiveSheet
                                    sheet_name = ws.Name
                                    pivot_table_name = control.GetParentControl().Name
                                    # 判断target之后是否存在数据透视表
                                    if not ws.PivotTables():
                                        cols = {}  # Dictionary holding named range objects
                                        # Determine the number of columns used in the top row
                                        xlToLeft = -4159
                                        col_count = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                                        column_names = [ws.Cells(1, col).Value for col in range(1, col_count + 1)]
                                        result['pivot_row_fields'] = []
                                        result['pivot_column_fields'] = []
                                        result['pivot_page_fields'] = []
                                        result['pivot_table_filters_values'] = column_names
                                        result['chosen_sheet1'] = ws.Name
                                        info = ws.UsedRange
                                        # print(info[0])
                                        nrows = info.Rows.Count
                                        ncols = info.Columns.Count
                                        column_name = column_to_name(ncols)
                                        data_range = f'A1:{column_name}{nrows}'
                                        result['pivot_range'] = data_range
                                        result['pivot_table_name'] = ""

                                        # Print each column heading
                                        # worksheet(行，列)
                                        # for col in range(1, col_count + 1):
                                        #     print("Col {}, {}".format(col, ws.Cells(2, col).Value))
                                    # 更新
                                    else:
                                        pt = ws.PivotTables(pivot_table_name)
                                        result['pivot_table_name'] = sheet_name + "!" + pivot_table_name
                                        row_fields = [i.Name for i in pt.RowFields]
                                        column_fields = [i.Name for i in pt.ColumnFields]
                                        # 去掉值
                                        if "值" in column_fields:
                                            column_fields = column_fields[:-1]
                                        page_fields = [i.Name for i in pt.PageFields]
                                        result['pivot_row_fields'] = row_fields
                                        result['pivot_column_fields'] = column_fields
                                        result['pivot_page_fields'] = page_fields
                                        data_list = [i.Name for i in pt.PivotFields()]
                                        if "值" in data_list:
                                            result['pivot_table_filters_values'] = data_list[:-1]
                                        else:
                                            result['pivot_table_filters_values'] = data_list
                                        # result['pivot_table_name'] = sheet_name + "!" + pivot_table_name
                                        pivot_range1 = pt.PivotCache().SourceData
                                        pivot_range2 = pt.SourceData
                                        # 数据处理。
                                        chosen_sheet1, data_range = getting_location_range(pivot_range2)
                                        result['chosen_sheet1'] = chosen_sheet1
                                        result['pivot_range'] = data_range
                                except:
                                    result['pivot_row_fields'] = []
                                    result['pivot_column_fields'] = []
                                    result['pivot_page_fields'] = []
                                    result['pivot_table_filters_values'] = []
                                    result['pivot_table_name'] = ''
                                    result['chosen_sheet1'] = ''
                                    result['pivot_range'] = ''
                            control_json = {}
                            controlList = []
                            depth = 0
                            control_name = ""
                            control_className = ""
                            while control:
                                depth += 1
                                controlJson = {}
                                self.control = control
                                num = self.getNum()
                                # numList.insert(0, num)
                                controlJson["num"] = num
                                if control.AutomationId:
                                    controlJson["id"] = control.AutomationId
                                else:
                                    controlJson["id"] = ""
                                if control.ControlTypeName:
                                    controlJson["type"] = control.ControlTypeName
                                if control.Name:
                                    controlJson["name"] = control.Name
                                    if depth == 1:
                                        control_json["controlName"] = control.Name
                                        control_name = control.Name
                                else:
                                    controlJson["name"] = ""
                                    if depth == 1:
                                        control_json["controlName"] = ""
                                        control_name = ""
                                if control.ClassName:
                                    controlJson["class"] = control.ClassName
                                    if depth == 1:
                                        control_json["controlClass"] = control.ClassName
                                        control_className = control.ClassName
                                else:
                                    controlJson["class"] = ""
                                    if depth == 1:
                                        control_json["controlClass"] = ""
                                        control_className = ""
                                controlList.insert(0, controlJson)
                                control = control.GetParentControl()
                                # print(control)
                            try:
                                ControlTypeName_json = controlList[0]
                                ControlTypeName = ControlTypeName_json["type"]
                                ControlClass = ControlTypeName_json["class"]
                                if ControlTypeName == "PaneControl" and ControlClass == "#32769":
                                    hwnd_json = controlList[1]
                                    controlList.remove(controlList[0])
                                else:
                                    hwnd_json = controlList[0]
                            except Exception:
                                pass
                            try:
                                self.refreshWindow()
                            except Exception as e:
                                print(e)
                            # 移回原来坐标
                            auto.MoveTo(x, y, moveSpeed=0)
                            hwnd_json['controlName'] = control_json["controlName"]
                            hwnd_json['controlClass'] = control_json["controlClass"]
                            # depth去除主窗体跟背景窗体
                            depth = depth - 2
                            result['depth'] = depth
                            result['list'] = str(controlList)
                            result['name'] = control_name
                            result['class'] = control_className
                            if x < right and x > left and y < bottom and y > top:
                                result['width'] = right - left - 2
                                result['height'] = bottom - top - 2
                                result['offsetX'] = x - left - 1
                                result['offsetY'] = y - top - 1
                            try:
                                result['window_name'] = hwnd_json['name']
                            except Exception:
                                result['window_name'] = ""
                            try:
                                result['window_class'] = hwnd_json['class']
                            except Exception:
                                result['window_class'] = ""
                            # 返回平台桌面
                            try:
                                shell = win32com.client.Dispatch("WScript.Shell")
                                shell.SendKeys('^')
                                win32gui.SetForegroundWindow(platform)
                            except Exception:
                                pass
                            # self.result = result
                            # self.root.destroy()
                            # break
                            return result
                    except Exception:
                        pass
            else:
                # self.refreshWindow()
                # self.result = {}
                # self.root.destroy()
                # break
                if self.hm:
                    self.hm.UnhookMouse()  # 取消鼠标监听
                    setHm("")
                return {}

    def tkWindow(self):
        self.root = tk.Tk()
        self.root.title("rpa_ui_shadow")
        hDC = win32gui.GetDC(0)
        # 横向分辨率
        HORZRES = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
        # 纵向分辨率
        VERTRES = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
        # self.root.geometry(str(HORZRES)+"x"+str(VERTRES))
        self.root.attributes('-toolwindow', True, '-alpha', 0.3, '-topmost', True)
        self.root.overrideredirect(True)
        label_border = tk.Frame(self.root, background="red", borderwidth=1)
        label = tk.Label(label_border, text="", bd=0)
        label_border.pack(side=tk.LEFT, fill="both", expand=True)
        label.pack(fill="both", expand=True, padx=1, pady=1)
        self.label = tk.Label(self.root, text="")
        self.label.place(x=0, y=0)
        self.root.after(1000, self.main)
        self.root.mainloop()
        self.root.quit()
        return self.result

# uiSpy().main()
