import tkinter
import win32con
import win32gui
import win32print


class targeting():
    def __init__(self):
        self.x = 0
        self.y = 0
        self.x1 = 0
        self.y1 = 0
        self.x2 = 0
        self.y2 = 0
        self.root = ""
        self.label = ""
        self.content = ""
        self.method = ""
        self.sequence = 1
        self.exit = "no"

    def callback(self, event):
        self.x = event.x
        self.y = event.y
        if self.method == "moveTo" or self.method == "dragTo":
            self.content = "x: " + str(self.x) + " " + "y: " + str(self.y)
        else:
            if self.sequence == 1:
                self.content = "Position1 x: " + str(self.x) + " " + "y: " + str(self.y)
                self.content = self.content + "\n" + "Position2 x: " + str(self.x2) + " " + "y: " + str(self.y2)
            else:
                self.content = "Position1 x: " + str(self.x1) + " " + "y: " + str(self.y1)
                self.content = self.content + "\n" + "Position2 x: " + str(self.x) + " " + "y: " + str(self.y)
        self.label['text'] = self.content
        self.label.place(x=self.x, y=self.y)

    def on_click(self, event):
        self.x = event.x_root
        self.y = event.y_root
        if self.method == "moveTo" or self.method == "dragTo":
            self.root.destroy()
        else:
            if self.sequence == 1:
                self.sequence = 2
                self.x1 = self.x
                self.y1 = self.y
            else:
                self.x2 = self.x
                self.y2 = self.y
                self.root.destroy()

    def on_right_click(self, event):
        self.exit = "yes"
        self.root.destroy()

    def xFunc(self, event):
        print(event.keycode)
        if event.keycode == 27:
            self.exit = "yes"
            self.root.destroy()

    def main(self, method):
        try:
            self.method = method
            self.root = tkinter.Tk()
            hDC = win32gui.GetDC(0)
            # 横向分辨率
            HORZRES = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
            # 纵向分辨率
            VERTRES = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
            print("resolution_x:", HORZRES, "resolution_y:", VERTRES)
            self.root.attributes('-toolwindow', True, '-alpha', 0.3, '-topmost', True)
            self.root.geometry(str(HORZRES) + "x" + str(VERTRES))
            self.root.overrideredirect(True)
            if method == "moveTo" or method == "dragTo":
                width = 15
                height = 1
            else:
                width = 20
                height = 2
            self.label = tkinter.Label(self.root, width=width, height=height, text=self.content, justify=tkinter.LEFT)
            self.label.place(x=self.x, y=self.y)
            self.root.bind("<Motion>", self.callback)
            self.root.bind("<Key>", self.xFunc)
            self.root.bind("<Button-1>", self.on_click)
            self.root.bind("<Button-3>", self.on_right_click)
            self.root.mainloop()
            self.root.quit()
            del self.root
            result = {}
            result["method"] = method
            result["screenWidth"] = HORZRES
            result["screenHeight"] = VERTRES
            if self.exit == "no":
                result["result"] = "success"
                if method == "moveTo" or method == "dragTo":
                    result["x"] = self.x
                    result["y"] = self.y
                    result["point_coordinates"] = "[" + str(self.x) + "," + str(self.y) + "]"
                else:
                    result["x"] = self.x2 - self.x1
                    result["y"] = self.y2 - self.y1
                    result["point_coordinates"] = "[" + str(self.x2 - self.x1) + "," + str(self.y2 - self.y1) + "]"
            else:
                result["result"] = "exit"
            self.x = 0
            self.y = 0
        except Exception as e:
            print(e)
            result = {}
            result["result"] = "error"
            result["msg"] = str(e)
        print(result)
        return result

# target = targeting()
# target.main("moveTo")
