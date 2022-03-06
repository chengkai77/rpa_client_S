import numpy as np
import cv2
from PIL import ImageGrab
import base64
import os
import win32gui, win32con
from math import floor
import time


def isInsideRect(p, r):
    return p[0] >= r[0] and p[0] <= r[2] + r[0] and p[1] >= r[1] and p[1] <= r[3] + r[1]


mainWindowChooseRect = (2, 2, 382, 122)
bChooseTemplate = False
bChooseTemplateDone = False
bChoosingTemplate = False
bMatchTemplate = False
bMatchTemplateDone = False
reMatch = False
bTemplate = -1
phase = 1
chooseTemplate = np.zeros((600, 600, 3), np.uint8)
result = np.zeros((600, 600, 3), np.uint8)
mainWindow = np.zeros((600, 1000, 3), np.uint8)
template = np.zeros((60, 60, 3), np.uint8)
point1 = (0, 0)
point2 = (0, 0)
newPoint = (0, 0)


def cancelSetPost(hwnd):
    try:
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
    except Exception:
        pass


def setPost(hwnd):
    try:
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
    except Exception:
        pass


def mainWindowCallback(event, x, y, flags, param):
    global mainWindowChooseRect, bChooseTemplate, bTemplate, bMatchTemplate
    if event == cv2.EVENT_LBUTTONDOWN:
        if isInsideRect((x, y), mainWindowChooseRect):
            bChooseTemplate = True
    if event == cv2.EVENT_RBUTTONUP:
        if bTemplate > -1:
            bTemplate = -1
            bMatchTemplate = True


def chooseTemplateCallback(event, x, y, flags, param):
    global bChooseTemplateDone, point1, point2, bChoosingTemplate
    if event == cv2.EVENT_LBUTTONDOWN:
        point1 = (x, y)
    if event == cv2.EVENT_MOUSEMOVE and flags == cv2.EVENT_FLAG_LBUTTON:
        point2 = (x, y)
        bChoosingTemplate = True
    if event == cv2.EVENT_LBUTTONUP:
        point2 = (x, y)
        bChooseTemplateDone = True
        bChoosingTemplate = False


def locate(tp, im):
    grayTp = cv2.cvtColor(tp, cv2.COLOR_BGR2GRAY)
    grayTp = cv2.blur(grayTp, (3, 3))
    gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
    gray = cv2.blur(gray, (3, 3))
    (tH, tW) = gray.shape
    found = None
    scales = [0.5, 0.6, 0.7, 0.8, 0.9, 1., 1.1, 1.2, 1.25, 1.3, 1.4, 1.5, 1.6, 1.7, 1.75, 1.8, 1.9, 2.0]
    for scale in scales:
        resized = cv2.resize(grayTp, (0, 0), fx=scale, fy=scale)
        r = gray.shape[1] / float(resized.shape[1])
        if resized.shape[0] > tH or resized.shape[1] > tW:
            break
        res = cv2.matchTemplate(gray, resized, cv2.TM_CCOEFF_NORMED)
        (_, maxVal, _, maxLoc) = cv2.minMaxLoc(res)

        clone = np.dstack([resized, resized, resized])
        cv2.rectangle(clone, (maxLoc[0], maxLoc[1]),
                      (maxLoc[0] + tW, maxLoc[1] + tH), (0, 0, 255), 2)
        if found is None or maxVal > found[0]:
            found = (maxVal, maxLoc, scale)

    (_, maxLoc, r) = found
    (startX, startY) = (int(maxLoc[0]), int(maxLoc[1]))
    (endX, endY) = (int((maxLoc[0] + grayTp.shape[1] * r)), int((maxLoc[1] + grayTp.shape[0] * r)))
    return (startX, startY), (endX, endY), found[0]


def start(delay):
    global phase, bTemplate, bChooseTemplate, bMatchTemplate, bChooseTemplateDone, bMatchTemplateDone, reMatch
    phase = 1
    icon_result = {}
    icon_result['status'] = "success"
    icon_result['img_data'] = ""
    width = ""
    height = ""
    while True:
        try:
            if phase == 1:
                click_image1_path = "icons\\click.jpg"
                mainWin = '  '
                cv2.namedWindow(mainWin)
                cv2.setMouseCallback(mainWin, mainWindowCallback, mainWin)
                clickme1 = cv2.imread(click_image1_path, flags=cv2.IMREAD_COLOR)
                if bTemplate == 1:
                    bTemplate = 0
                    try:
                        scale = min(mainWindowChooseRect[3] * 1. / template.shape[0],
                                    mainWindowChooseRect[2] * 1. / template.shape[1])
                        tmp = cv2.resize(template, (0, 0), fx=scale, fy=scale)
                        mainWindow[mainWindowChooseRect[1]:mainWindowChooseRect[1] + tmp.shape[0],
                        mainWindowChooseRect[0]:mainWindowChooseRect[0] + tmp.shape[1]] = tmp
                    except Exception:
                        pass
                phase = 2
                cv2.imshow(mainWin, clickme1)
            elif phase == 2:
                autosize = cv2.getWindowProperty(mainWin, cv2.WND_PROP_AUTOSIZE)
                hwnd = win32gui.FindWindow("Main HighGUI class", mainWin)
                setPost(hwnd)
                if autosize < 1:
                    cv2.destroyAllWindows()
                    cancelSetPost(hwnd)
                    break
                if bChooseTemplate:
                    cv2.destroyWindow(mainWin)
                    cv2.waitKey(500)
                    bChooseTemplate = False
                    phase = 3
                if bMatchTemplate:
                    cv2.destroyWindow(mainWin)
                    cv2.waitKey(500)
                    bMatchTemplate = False
                    phase = 5
            elif phase == 3:
                time.sleep(delay)
                ii = ImageGrab.grab()
                rgb = np.array(ii)
                chooseTemplate = cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)
                veil = np.ones_like(chooseTemplate) * int(128)
                chooseTemplateShow = cv2.addWeighted(chooseTemplate, 0.5, veil, 0.5, 0.)
                cv2.namedWindow('chooseTemplate', cv2.WND_PROP_FULLSCREEN)
                cv2.setWindowProperty('chooseTemplate', cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
                cv2.setMouseCallback('chooseTemplate', chooseTemplateCallback)
                phase = 4
            elif phase == 4:
                show = chooseTemplateShow.copy()
                if bChoosingTemplate:
                    cv2.rectangle(show, point1, point2, (0, 0, 255), 2)
                    show[min(point1[1], point2[1]): max(point1[1], point2[1]),
                    min(point1[0], point2[0]): max(point1[0], point2[0])] = chooseTemplate[
                                                                            min(point1[1], point2[1]): max(point1[1],
                                                                                                           point2[1]),
                                                                            min(point1[0], point2[0]): max(point1[0],
                                                                                                           point2[0])]
                cv2.imshow('chooseTemplate', show)
                if bChooseTemplateDone:
                    bChooseTemplateDone = False
                    print(type(min(point1[1], point2[1])))
                    template = chooseTemplate[min(point1[1], point2[1]): max(point1[1], point2[1]),
                               min(point1[0], point2[0]): max(point1[0], point2[0])]
                    bTemplate = 1
                    cv2.destroyWindow('chooseTemplate')
                    phase = 5
                    bTemplate = 0
            elif phase == 5:
                result = cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)
                pt1, pt2, sco = locate(template, result)
                if not os.path.exists("temp"):
                    os.mkdir("temp")
                cv2.imwrite('temp\\icon_coordinate_target.png', template)
                fp = open('temp\\icon_coordinate_target.png', 'rb')
                base64_date = base64.b64encode(fp.read())
                fp.close()
                # try:
                #     os.remove('temp\\icon_coordinate_target.png')
                # except Exception:
                #     pass
                cv2.namedWindow('Icon Window', cv2.WND_PROP_FULLSCREEN)
                cv2.setWindowProperty('Icon Window', cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
                cv2.rectangle(result, pt1, pt2, (0, 255, 0), 2)
                # matchRate = "Match Rate: " + str(floor(sco*100)) + "%"
                # cv2.putText(result, str(matchRate), (100, 100), cv2.FONT_HERSHEY_COMPLEX, 1., (0, 0, 255))
                phase = 6
            elif phase == 6:
                cv2.imshow('Icon Window', result)
                # 对角线左上角起点A点x轴坐标
                x1 = pt1[0]
                # 对角线左上角起点A点y轴坐标
                y1 = pt1[1]
                # 对角线左上角终点B点x轴坐标
                x2 = pt2[0]
                # 对角线左上角终点B点y轴坐标
                y2 = pt2[1]
                width = x2 - x1
                height = y2 - y1
                if bMatchTemplateDone:
                    x3 = newPoint[0]
                    y3 = newPoint[1]
                    cv2.destroyWindow('Icon Window')
                    time.sleep(2)
                    if x3 < x1 or x3 > x2 or y3 < y1 or y3 > y2:
                        bMatchTemplateDone = False
                        phase = 1
                    else:
                        bMatchTemplateDone = False
                        if reMatch:
                            reMatch = False
                            phase = 1
                        else:
                            phase = 5
            k = cv2.waitKey(5) & 0xFF
            if k in [27, ord('q')]:
                break
            elif k in [13, ord('q')]:
                if base64_date:
                    img_data = str(base64_date, encoding="gbk")
                    # print(img_data)
                    icon_result['img_data'] = img_data
                    pt1_str = "["
                    pt2_str = "["
                    for i in pt1:
                        pt1_str = pt1_str + str(i) + ","
                    for j in pt2:
                        pt2_str = pt2_str + str(j) + ","
                    icon_result['start_coordinates'] = pt1_str.rstrip(",") + "]"  # start_coordinates返回一个[]list
                    icon_result['end_coordinates'] = pt2_str.rstrip(",") + "]"  # end_coordinates返回一个[]list
                    icon_result['width'] = width
                    icon_result['height'] = height
                break
        except Exception as e:
            msg = str(e)
            icon_result['status'] = "error"
            icon_result['msg'] = msg
            break
    cv2.destroyAllWindows()
    return icon_result
