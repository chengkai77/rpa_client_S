import numpy as np
import cv2
import time

from PIL import ImageGrab
from math import floor


def isInsideRect(p, r):
    return p[0] >= r[0] and p[0] <= r[2] + r[0] and p[1] >= r[1] and p[1] <= r[3] + r[1]


mainWindowChooseRect = (2, 2, 382, 42)
bChooseTemplate = False
bChooseTemplateDone = False
bChoosingTemplate = False
bMatchTemplate = False
bMatchTemplateDone = False
bTemplate = -1
phase = 1
chooseTemplate = np.zeros((600, 600, 3), np.uint8)
result = np.zeros((600, 600, 3), np.uint8)
mainWindow = np.zeros((380, 90, 3), np.uint8)
template = np.zeros((60, 60, 3), np.uint8)
point1 = (0, 0)
point2 = (0, 0)


def locate(tp, im):
    grayTp = cv2.cvtColor(tp, cv2.COLOR_BGR2GRAY)
    gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
    gray = cv2.Sobel(gray, cv2.CV_8U, 1, 1)
    (tH, tW) = gray.shape
    found = None
    scales = [0.5, 0.6, 0.7, 0.8, 0.9, 1., 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 2.0]
    for scale in scales:
        resized = cv2.resize(grayTp, (0, 0), fx=scale, fy=scale)
        r = gray.shape[1] / float(resized.shape[1])
        if resized.shape[0] > tH or resized.shape[1] > tW:
            break
        edged = cv2.Sobel(resized, cv2.CV_8U, 1, 1)
        res = cv2.matchTemplate(gray, edged, cv2.TM_CCORR_NORMED)
        (_, maxVal, _, maxLoc) = cv2.minMaxLoc(res)

        clone = np.dstack([edged, edged, edged])
        cv2.rectangle(clone, (maxLoc[0], maxLoc[1]),
                      (maxLoc[0] + tW, maxLoc[1] + tH), (0, 0, 255), 2)
        if found is None or maxVal > found[0]:
            found = (maxVal, maxLoc, scale)
    (_, maxLoc, r) = found
    (startX, startY) = (int(maxLoc[0]), int(maxLoc[1]))
    (endX, endY) = (int((maxLoc[0] + grayTp.shape[1] * r)), int((maxLoc[1] + grayTp.shape[0] * r)))
    print(found[0])
    return (startX, startY), (endX, endY), found[0]


def locate2(tp, im, method=0, blur=False):
    if method == 1:
        grayTp = tp
        gray = im
        (tH, tW, _) = gray.shape
    elif method == 2:
        grayTp = cv2.cvtColor(tp, cv2.COLOR_BGR2GRAY)
        gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
        (tH, tW) = gray.shape
        grayTp = cv2.Sobel(grayTp, cv2.CV_8U, 1, 1)
        gray = cv2.Sobel(gray, cv2.CV_8U, 1, 1)
    else:
        grayTp = cv2.cvtColor(tp, cv2.COLOR_BGR2GRAY)
        gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
        (tH, tW) = gray.shape

    if blur:
        grayTp = cv2.blur(grayTp, (3, 3))
        gray = cv2.blur(gray, (3, 3))

    found = None
    scales = [0.5, 0.6, 0.7, 0.8, 0.9, 1., 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 2.0]
    for scale in scales:
        resized = cv2.resize(grayTp, (0, 0), fx=scale, fy=scale)
        r = gray.shape[1] / float(resized.shape[1])
        if resized.shape[0] > tH or resized.shape[1] > tW:
            break
        res = cv2.matchTemplate(gray, resized, cv2.TM_CCORR_NORMED)
        (_, maxVal, _, maxLoc) = cv2.minMaxLoc(res)

        if found is None or maxVal > found[0]:
            found = (maxVal, maxLoc, scale, resized)
    (_, maxLoc, r, _) = found
    (startX, startY) = (int(maxLoc[0]), int(maxLoc[1]))
    (endX, endY) = (int((maxLoc[0] + grayTp.shape[1] * r)), int((maxLoc[1] + grayTp.shape[0] * r)))

    res = cv2.matchTemplate(gray[startY:endY, startX:endX], found[3], cv2.TM_CCOEFF_NORMED)
    (_, maxVal, _, maxLoc) = cv2.minMaxLoc(res)
    maxVal = found[0] * maxVal

    return (startX, startY), (endX, endY), maxVal


def start(img_data, position, x_axis, y_axis, width, height, waiting, rate):
    position = str(position)
    result = {}
    try:
        ii = ImageGrab.grab()
        rgb = np.array(ii)
        screen_data = cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)
        pt1, pt2, sco = locate(img_data, screen_data)
        matching_rate = float(sco)
        if matching_rate < rate:
            waiting_time = 0
            while waiting_time < waiting:
                waiting_time += 1
                ii = ImageGrab.grab()
                rgb = np.array(ii)
                screen_data = cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)
                pt1, pt2, sco = locate(img_data, screen_data)
                matching_rate = float(sco)
                print(matching_rate)
                if matching_rate >= rate:
                    break
            if waiting_time >= waiting:
                pt1, pt2, sco2 = locate2(img_data, screen_data)
                matching_rate2 = float(sco2)
                print(matching_rate2)
                if matching_rate2 < rate:
                    result["result"] = "error"
                    result[
                        "msg"] = "Unable to find the element, please reduce the match rate or retake the screenshot, current highest matching rate is " + str(
                        floor(max(sco, sco2) * 100)) + "%"
                    print(result)
                    return result
        current_width = int(pt2[0] - pt1[0])
        current_height = int(pt2[1] - pt1[1])
        x_axis = int(current_width / width * x_axis)
        y_axis = int(current_height / height * y_axis)
        if position == "0":
            x = int((pt1[0] + pt2[0]) / 2)
            y = int((pt1[1] + pt2[1]) / 2)
        elif position == "-1":
            x = pt1[0]
            y = pt1[1]
        elif position == "1":
            x = pt2[0]
            y = pt1[1]
        elif position == "-2":
            x = pt1[0]
            y = pt2[1]
        elif position == "2":
            x = pt2[0]
            y = pt2[1]
        result["result"] = "success"
        result["x"] = x + x_axis
        result["y"] = y + y_axis
    except Exception as e:
        print(e)
        result["result"] = "error"
        result["msg"] = str(e)
    return result
