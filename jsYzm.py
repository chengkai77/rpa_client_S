import base64
import sys

import win32gui
from PIL import Image
from pytesseract import *
from fnmatch import fnmatch
from queue import Queue
import matplotlib.pyplot as plt
import cv2
import time
import os
from win32com.client import Dispatch


<<<<<<< HEAD


def clear_border(img,img_name):
  '''去除边框
  '''

  filename = './jushi/' + img_name.split('.')[0] + '-clearBorder.png'
  h, w = img.shape[:2]
  for y in range(0, w):
    for x in range(0, h):
      # if y ==0 or y == w -1 or y == w - 2:
      if y < 4 or y > w -4:
        img[x, y] = 255
      # if x == 0 or x == h - 1 or x == h - 2:
      if x < 4 or x > h - 4:
        img[x, y] = 255

  cv2.imwrite(filename,img)
  return img


def interference_line(img, img_name):
  '''
  干扰线降噪
  '''

  filename =  './jushi/' + img_name.split('.')[0] + '-interferenceline.png'
  h, w = img.shape[:2]
  # ！！！opencv矩阵点是反的
  # img[1,2] 1:图片的高度，2：图片的宽度
  for y in range(1, w - 1):
    for x in range(1, h - 1):
      count = 0
      if img[x, y - 1] > 245:
        count = count + 1
      if img[x, y + 1] > 245:
        count = count + 1
      if img[x - 1, y] > 245:
        count = count + 1
      if img[x + 1, y] > 245:
        count = count + 1
      if count > 2:
        img[x, y] = 255
  cv2.imwrite(filename,img)
  return img

def interference_point(img,img_name, x = 0, y = 0):
=======
def clear_border(img, img_name):
    '''去除边框
    '''

    filename = './jushi/' + img_name.split('.')[0] + '-clearBorder.png'
    h, w = img.shape[:2]
    for y in range(0, w):
        for x in range(0, h):
            # if y ==0 or y == w -1 or y == w - 2:
            if y < 4 or y > w - 4:
                img[x, y] = 255
            # if x == 0 or x == h - 1 or x == h - 2:
            if x < 4 or x > h - 4:
                img[x, y] = 255

    cv2.imwrite(filename, img)
    return img


def interference_line(img, img_name):
    '''
    干扰线降噪
    '''

    filename = './jushi/' + img_name.split('.')[0] + '-interferenceline.png'
    h, w = img.shape[:2]
    # ！！！opencv矩阵点是反的
    # img[1,2] 1:图片的高度，2：图片的宽度
    for y in range(1, w - 1):
        for x in range(1, h - 1):
            count = 0
            if img[x, y - 1] > 245:
                count = count + 1
            if img[x, y + 1] > 245:
                count = count + 1
            if img[x - 1, y] > 245:
                count = count + 1
            if img[x + 1, y] > 245:
                count = count + 1
            if count > 2:
                img[x, y] = 255
    cv2.imwrite(filename, img)
    return img


def interference_point(img, img_name, x=0, y=0):
>>>>>>> 7761156646ef8830f3136e5ef5ab79af81d03fc9
    """点降噪
    9邻域框,以当前点为中心的田字框,黑点个数
    :param x:
    :param y:
    :return:
    """
<<<<<<< HEAD
    filename =  './jushi/' + img_name.split('.')[0] + '-interferencePoint.png'
    # todo 判断图片的长宽度下限
    cur_pixel = img[x,y]# 当前像素点的值
    height,width = img.shape[:2]

    for y in range(0, width - 1):
      for x in range(0, height - 1):
        if y == 0:  # 第一行
            if x == 0:  # 左上顶点,4邻域
                # 中心点旁边3个点
                sum = int(cur_pixel) \
                      + int(img[x, y + 1]) \
                      + int(img[x + 1, y]) \
                      + int(img[x + 1, y + 1])
                if sum <= 2 * 245:
                  img[x, y] = 0
            elif x == height - 1:  # 右上顶点
                sum = int(cur_pixel) \
                      + int(img[x, y + 1]) \
                      + int(img[x - 1, y]) \
                      + int(img[x - 1, y + 1])
                if sum <= 2 * 245:
                  img[x, y] = 0
            else:  # 最上非顶点,6邻域
                sum = int(img[x - 1, y]) \
                      + int(img[x - 1, y + 1]) \
                      + int(cur_pixel) \
                      + int(img[x, y + 1]) \
                      + int(img[x + 1, y]) \
                      + int(img[x + 1, y + 1])
                if sum <= 3 * 245:
                  img[x, y] = 0
        elif y == width - 1:  # 最下面一行
            if x == 0:  # 左下顶点
                # 中心点旁边3个点
                sum = int(cur_pixel) \
                      + int(img[x + 1, y]) \
                      + int(img[x + 1, y - 1]) \
                      + int(img[x, y - 1])
                if sum <= 2 * 245:
                  img[x, y] = 0
            elif x == height - 1:  # 右下顶点
                sum = int(cur_pixel) \
                      + int(img[x, y - 1]) \
                      + int(img[x - 1, y]) \
                      + int(img[x - 1, y - 1])

                if sum <= 2 * 245:
                  img[x, y] = 0
            else:  # 最下非顶点,6邻域
                sum = int(cur_pixel) \
                      + int(img[x - 1, y]) \
                      + int(img[x + 1, y]) \
                      + int(img[x, y - 1]) \
                      + int(img[x - 1, y - 1]) \
                      + int(img[x + 1, y - 1])
                if sum <= 3 * 245:
                  img[x, y] = 0
        else:  # y不在边界
            if x == 0:  # 左边非顶点
                sum = int(img[x, y - 1]) \
                      + int(cur_pixel) \
                      + int(img[x, y + 1]) \
                      + int(img[x + 1, y - 1]) \
                      + int(img[x + 1, y]) \
                      + int(img[x + 1, y + 1])

                if sum <= 3 * 245:
                  img[x, y] = 0
            elif x == height - 1:  # 右边非顶点
                sum = int(img[x, y - 1]) \
                      + int(cur_pixel) \
                      + int(img[x, y + 1]) \
                      + int(img[x - 1, y - 1]) \
                      + int(img[x - 1, y]) \
                      + int(img[x - 1, y + 1])

                if sum <= 3 * 245:
                  img[x, y] = 0
            else:  # 具备9领域条件的
                sum = int(img[x - 1, y - 1]) \
                      + int(img[x - 1, y]) \
                      + int(img[x - 1, y + 1]) \
                      + int(img[x, y - 1]) \
                      + int(cur_pixel) \
                      + int(img[x, y + 1]) \
                      + int(img[x + 1, y - 1]) \
                      + int(img[x + 1, y]) \
                      + int(img[x + 1, y + 1])
                if sum <= 4 * 245:
                  img[x, y] = 0
    cv2.imwrite(filename,img)
    return img

def _get_dynamic_binary_image(filedir, img_name):
  '''
  自适应阀值二值化
  '''

  filename =   './jushi/' + img_name.split('.')[0] + '-binary.png'
  img_name = filedir + '/' + img_name
  print('.....' + img_name)
  im = cv2.imread(img_name)
  im = cv2.cvtColor(im,cv2.COLOR_BGR2GRAY)

  th1 = cv2.adaptiveThreshold(im, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 21, 1)
  cv2.imwrite(filename,th1)
  return th1

def _get_static_binary_image(img, threshold = 140):
  '''
  手动二值化
  '''

  img = Image.open(img)
  img = img.convert('L')
  pixdata = img.load()
  w, h = img.size
  for y in range(h):
    for x in range(w):
      if pixdata[x, y] < threshold:
        pixdata[x, y] = 0
      else:
        pixdata[x, y] = 255

  return img


def cfs(im,x_fd,y_fd):
  '''用队列和集合记录遍历过的像素坐标代替单纯递归以解决cfs访问过深问题
  '''

  # print('**********')

  xaxis=[]
  yaxis=[]
  visited =set()
  q = Queue()
  q.put((x_fd, y_fd))
  visited.add((x_fd, y_fd))
  offsets=[(1, 0), (0, 1), (-1, 0), (0, -1)]#四邻域

  while not q.empty():
      x,y=q.get()

      for xoffset,yoffset in offsets:
          x_neighbor,y_neighbor = x+xoffset,y+yoffset

          if (x_neighbor,y_neighbor) in (visited):
              continue  # 已经访问过了

          visited.add((x_neighbor, y_neighbor))

          try:
              if im[x_neighbor, y_neighbor] == 0:
                  xaxis.append(x_neighbor)
                  yaxis.append(y_neighbor)
                  q.put((x_neighbor,y_neighbor))

          except IndexError:
              pass
  # print(xaxis)
  if (len(xaxis) == 0 | len(yaxis) == 0):
    xmax = x_fd + 1
    xmin = x_fd
    ymax = y_fd + 1
    ymin = y_fd

  else:
    xmax = max(xaxis)
    xmin = min(xaxis)
    ymax = max(yaxis)
    ymin = min(yaxis)
    #ymin,ymax=sort(yaxis)

  return ymax,ymin,xmax,xmin

def detectFgPix(im,xmax):
  '''搜索区块起点
  '''

  h,w = im.shape[:2]
  for y_fd in range(xmax+1,w):
      for x_fd in range(h):
          if im[x_fd,y_fd] == 0:
              return x_fd,y_fd

def CFS(im):
  '''切割字符位置
  '''

  zoneL=[]#各区块长度L列表
  zoneWB=[]#各区块的X轴[起始，终点]列表
  zoneHB=[]#各区块的Y轴[起始，终点]列表

  xmax=0#上一区块结束黑点横坐标,这里是初始化
  for i in range(10):

      try:
          x_fd,y_fd = detectFgPix(im,xmax)
          # print(y_fd,x_fd)
          xmax,xmin,ymax,ymin=cfs(im,x_fd,y_fd)
          L = xmax - xmin
          H = ymax - ymin
          zoneL.append(L)
          zoneWB.append([xmin,xmax])
          zoneHB.append([ymin,ymax])

      except TypeError:
          return zoneL,zoneWB,zoneHB

  return zoneL,zoneWB,zoneHB


def cutting_img(im,im_position,img,xoffset = 1,yoffset = 1):
    filename =  './jushi/' + img.split('.')[0]
=======
    filename = './jushi/' + img_name.split('.')[0] + '-interferencePoint.png'
    # todo 判断图片的长宽度下限
    cur_pixel = img[x, y]  # 当前像素点的值
    height, width = img.shape[:2]

    for y in range(0, width - 1):
        for x in range(0, height - 1):
            if y == 0:  # 第一行
                if x == 0:  # 左上顶点,4邻域
                    # 中心点旁边3个点
                    sum = int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x + 1, y]) \
                          + int(img[x + 1, y + 1])
                    if sum <= 2 * 245:
                        img[x, y] = 0
                elif x == height - 1:  # 右上顶点
                    sum = int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x - 1, y]) \
                          + int(img[x - 1, y + 1])
                    if sum <= 2 * 245:
                        img[x, y] = 0
                else:  # 最上非顶点,6邻域
                    sum = int(img[x - 1, y]) \
                          + int(img[x - 1, y + 1]) \
                          + int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x + 1, y]) \
                          + int(img[x + 1, y + 1])
                    if sum <= 3 * 245:
                        img[x, y] = 0
            elif y == width - 1:  # 最下面一行
                if x == 0:  # 左下顶点
                    # 中心点旁边3个点
                    sum = int(cur_pixel) \
                          + int(img[x + 1, y]) \
                          + int(img[x + 1, y - 1]) \
                          + int(img[x, y - 1])
                    if sum <= 2 * 245:
                        img[x, y] = 0
                elif x == height - 1:  # 右下顶点
                    sum = int(cur_pixel) \
                          + int(img[x, y - 1]) \
                          + int(img[x - 1, y]) \
                          + int(img[x - 1, y - 1])

                    if sum <= 2 * 245:
                        img[x, y] = 0
                else:  # 最下非顶点,6邻域
                    sum = int(cur_pixel) \
                          + int(img[x - 1, y]) \
                          + int(img[x + 1, y]) \
                          + int(img[x, y - 1]) \
                          + int(img[x - 1, y - 1]) \
                          + int(img[x + 1, y - 1])
                    if sum <= 3 * 245:
                        img[x, y] = 0
            else:  # y不在边界
                if x == 0:  # 左边非顶点
                    sum = int(img[x, y - 1]) \
                          + int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x + 1, y - 1]) \
                          + int(img[x + 1, y]) \
                          + int(img[x + 1, y + 1])

                    if sum <= 3 * 245:
                        img[x, y] = 0
                elif x == height - 1:  # 右边非顶点
                    sum = int(img[x, y - 1]) \
                          + int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x - 1, y - 1]) \
                          + int(img[x - 1, y]) \
                          + int(img[x - 1, y + 1])

                    if sum <= 3 * 245:
                        img[x, y] = 0
                else:  # 具备9领域条件的
                    sum = int(img[x - 1, y - 1]) \
                          + int(img[x - 1, y]) \
                          + int(img[x - 1, y + 1]) \
                          + int(img[x, y - 1]) \
                          + int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x + 1, y - 1]) \
                          + int(img[x + 1, y]) \
                          + int(img[x + 1, y + 1])
                    if sum <= 4 * 245:
                        img[x, y] = 0
    cv2.imwrite(filename, img)
    return img


def _get_dynamic_binary_image(filedir, img_name):
    '''
    自适应阀值二值化
    '''

    filename = './jushi/' + img_name.split('.')[0] + '-binary.png'
    img_name = filedir + '/' + img_name
    print('.....' + img_name)
    im = cv2.imread(img_name)
    im = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)

    th1 = cv2.adaptiveThreshold(im, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 21, 1)
    cv2.imwrite(filename, th1)
    return th1


def _get_static_binary_image(img, threshold=140):
    '''
    手动二值化
    '''

    img = Image.open(img)
    img = img.convert('L')
    pixdata = img.load()
    w, h = img.size
    for y in range(h):
        for x in range(w):
            if pixdata[x, y] < threshold:
                pixdata[x, y] = 0
            else:
                pixdata[x, y] = 255

    return img


def cfs(im, x_fd, y_fd):
    '''用队列和集合记录遍历过的像素坐标代替单纯递归以解决cfs访问过深问题
    '''

    # print('**********')

    xaxis = []
    yaxis = []
    visited = set()
    q = Queue()
    q.put((x_fd, y_fd))
    visited.add((x_fd, y_fd))
    offsets = [(1, 0), (0, 1), (-1, 0), (0, -1)]  # 四邻域

    while not q.empty():
        x, y = q.get()

        for xoffset, yoffset in offsets:
            x_neighbor, y_neighbor = x + xoffset, y + yoffset

            if (x_neighbor, y_neighbor) in (visited):
                continue  # 已经访问过了

            visited.add((x_neighbor, y_neighbor))

            try:
                if im[x_neighbor, y_neighbor] == 0:
                    xaxis.append(x_neighbor)
                    yaxis.append(y_neighbor)
                    q.put((x_neighbor, y_neighbor))

            except IndexError:
                pass
    # print(xaxis)
    if (len(xaxis) == 0 | len(yaxis) == 0):
        xmax = x_fd + 1
        xmin = x_fd
        ymax = y_fd + 1
        ymin = y_fd

    else:
        xmax = max(xaxis)
        xmin = min(xaxis)
        ymax = max(yaxis)
        ymin = min(yaxis)
        # ymin,ymax=sort(yaxis)

    return ymax, ymin, xmax, xmin


def detectFgPix(im, xmax):
    '''搜索区块起点
    '''

    h, w = im.shape[:2]
    for y_fd in range(xmax + 1, w):
        for x_fd in range(h):
            if im[x_fd, y_fd] == 0:
                return x_fd, y_fd


def CFS(im):
    '''切割字符位置
    '''

    zoneL = []  # 各区块长度L列表
    zoneWB = []  # 各区块的X轴[起始，终点]列表
    zoneHB = []  # 各区块的Y轴[起始，终点]列表

    xmax = 0  # 上一区块结束黑点横坐标,这里是初始化
    for i in range(10):

        try:
            x_fd, y_fd = detectFgPix(im, xmax)
            # print(y_fd,x_fd)
            xmax, xmin, ymax, ymin = cfs(im, x_fd, y_fd)
            L = xmax - xmin
            H = ymax - ymin
            zoneL.append(L)
            zoneWB.append([xmin, xmax])
            zoneHB.append([ymin, ymax])

        except TypeError:
            return zoneL, zoneWB, zoneHB

    return zoneL, zoneWB, zoneHB


def cutting_img(im, im_position, img, xoffset=1, yoffset=1):
    filename = './jushi/' + img.split('.')[0]
>>>>>>> 7761156646ef8830f3136e5ef5ab79af81d03fc9
    # 识别出的字符个数
    im_number = len(im_position[1])
    # 切割字符
    for i in range(im_number):
        im_start_X = im_position[1][i][0] - xoffset
        im_end_X = im_position[1][i][1] + xoffset
        im_start_Y = im_position[2][i][0] - yoffset
        im_end_Y = im_position[2][i][1] + yoffset
        cropped = im[im_start_Y:im_end_Y, im_start_X:im_end_X]
        cv2.imwrite(filename + '-cutting-' + str(i) + '.png', cropped)


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False

<<<<<<< HEAD
def del_file(path_data):
    for i in os.listdir(path_data) :# os.listdir(path_data)#返回一个列表，里面是当前目录下面的所有东西的相对路径
        file_data = path_data + "\\" + i#当前文件夹的下面的所有东西的绝对路径
        if os.path.isfile(file_data) == True:#os.path.isfile判断是否为文件,如果是文件,就删除.如果是文件夹.递归给del_file.
=======

def del_file(path_data):
    for i in os.listdir(path_data):  # os.listdir(path_data)#返回一个列表，里面是当前目录下面的所有东西的相对路径
        file_data = path_data + "\\" + i  # 当前文件夹的下面的所有东西的绝对路径
        if os.path.isfile(file_data) == True:  # os.path.isfile判断是否为文件,如果是文件,就删除.如果是文件夹.递归给del_file.
>>>>>>> 7761156646ef8830f3136e5ef5ab79af81d03fc9
            os.remove(file_data)
        else:
            del_file(file_data)

<<<<<<< HEAD
def main():
    global ie,doc,pw
    #获取ie解析base64编码
=======

def main():
    global ie, doc, pw
    # 获取ie解析base64编码
>>>>>>> 7761156646ef8830f3136e5ef5ab79af81d03fc9
    hwnd = win32gui.GetForegroundWindow()
    ShellWindowsCLSID = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'
    ShellWindows = Dispatch(ShellWindowsCLSID)
    for shellwindow in ShellWindows:
        if shellwindow.LocationURL != "":
            if "Internet Explorer" in str(shellwindow):
                ie = shellwindow
                doc = ie.Document
                h = shellwindow.HWND
                if h == hwnd:
                    break
    img_str = doc.rpa_yzm
    print(img_str)
<<<<<<< HEAD
    #将base64编码写入图片
    with open("./jushi/yzm.png", "wb") as f:
        # 生成图片
        img_data = base64.b64decode(img_str.replace("data:image/png;base64,",""))
        f.write(img_data)
    #获取图片识别验证码
=======
    # 将base64编码写入图片
    with open("./jushi/yzm.png", "wb") as f:
        # 生成图片
        img_data = base64.b64decode(img_str.replace("data:image/png;base64,", ""))
        f.write(img_data)
    # 获取图片识别验证码
>>>>>>> 7761156646ef8830f3136e5ef5ab79af81d03fc9
    filedir = './jushi'
    for file in os.listdir(filedir):
        if fnmatch(file, '*.png'):
            img_name = file
            # 自适应阈值二值化
            im = _get_dynamic_binary_image(filedir, img_name)
            # 去除边框
            im = clear_border(im, img_name)
            # 对图片进行干扰线降噪
            im = interference_line(im, img_name)
            # 对图片进行点降噪
            im = interference_point(im, img_name)
            # 识别验证码
            try:
                file = './jushi/' + img_name.split('.')[0] + '-interferencePoint.png'
                # 识别验证码
                str_img = image_to_string(Image.open(file), lang='eng')  # 单个字符是10，一行文本是7
            except Exception as err:
                pass
            result = ""
            for i in range(len(str_img)):
                part_str = str_img[i]
                if is_number(part_str):
                    result = result + part_str
            print('识别为：%s' % str_img)
    del_file('./jushi/')
    pw = doc.parentWindow
    js_script = """
        delete document.rpa_yzm;
<<<<<<< HEAD
        document.getElementById("validatecode").value = """ + result +"""
=======
        document.getElementById("validatecode").value = """ + result + """
>>>>>>> 7761156646ef8830f3136e5ef5ab79af81d03fc9
    """
    pw.execScript(js_script)


def cycle(pwd):
    global ie, doc, pw
    js_script = """
        var warning_span = document.getElementById("errorMessage");
        if (warning_span){
            document.warning = warning_span.innerText;
        }else{
            document.warning = "";
        }
    """
    pw.execScript(js_script)
    try:
        yzm = doc.warning
    except Exception:
        print(Exception)
    while yzm == "验证码错误":
        try:
            js1 = 'function getBase64Image(img) {var canvas = document.createElement("canvas");canvas.width = img.width;canvas.height = img.height;var ctx = canvas.getContext("2d");ctx.drawImage(img, 0, 0, img.width, img.height);var dataURL = canvas.toDataURL("image/png");return dataURL;};img = document.getElementById("imgCode2");document.rpa_yzm=getBase64Image(img);'
            pw.execScript(js1)
            time.sleep(1)
            main()
            time.sleep(1)
            js2 = 'document.getElementsByName("userpassword")[0].value = "' + pwd + '";var loginWin = document.getElementById("passwordLogin");var btn = loginWin.getElementsByTagName("button")[0];btn.click();var warning_span = document.getElementById("errorMessage");if (warning_span){document.warning = warning_span.innerText;}else{document.warning = "";}'
            pw.execScript(js2)
            try:
                yzm = doc.warning
            except Exception:
                yzm = ""
        except Exception as e:
<<<<<<< HEAD
            print(e)
=======
            print(e)
>>>>>>> 7761156646ef8830f3136e5ef5ab79af81d03fc9
