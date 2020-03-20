import win32clipboard as w
import win32api,win32gui,win32con,time
from PIL import Image
from io import BytesIO
from ctypes import *
def setText(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT,astring)
    w.CloseClipboard()
def getText():
    w.OpenClipboard()
    data = w.getClipboardData()
    w.CloseClipboard()
    return data
def send_message(handle):
    win32gui.PostMessage(handle,win32con.WM_PASTE,0,0)
    win32gui.PostMessage(handle,win32con.WM_KEYDOWN,win32con.VK_RETURN,0)#按下
    win32gui.PostMessage(handle,win32con.WM_KEYUP,win32con.VK_RETURN,0)#松开
    #VK_RETURN 表示虚拟键码16进制代表的的回车键(13)
def pic_ctrl_c(path):
    img = Image.open(path)
    output = BytesIO()
    img.convert("RGB").save(output,"BMP")
    data = output.getvalue()[14:]
    output.close()
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_DIB,data)
    w.CloseClipboard()
    
name = input("请输入备注")
handle = win32gui.FindWindow(None,name)
if handle>0:
    print("找到: %s"%name)
    left,top,right,bottom = win32gui.GetWindowRect(handle)
    print("窗体大小: ",right-left,top-bottom)
    flag = input("发送图片请按1，文字按2：")
    if flag=='1':
        pic_ctrl_c(input("图片路径: ").strip())
    else :
        setText(input("请输入内容: "))
    try:
        i=0;
        while i<1000:
            time.sleep(0.2)
            send_message(handle) 
            i=i+1
    except :
        print("NO")
    finally:
        print("OK")
    '''win32gui.SetForegroundWindow(handle)
    win32gui.MoveWindow(handle,20,20,405,756,True)
    win32gui.SetBkMode(handle,TRANSPARENT)'''
else:
    print("No Find")



