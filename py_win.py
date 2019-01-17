import os
import time
import win32gui
import win32api
import win32con
import socket
import re

# ensp  绝对路径 路径
file_name = "C:\\Program" + ' ' + 'Files\\Huawei\\eNSP\\eNSP_Client.exe'
confi_file = 'C:\Program Files\Huawei\eNSP\cfg\config_server.ini'


class Run_win(object):
    def __init__(self):
        """
        初始化
        获取窗口焦点
        """
        os.startfile(file_name)

        wdname1 = u"eNSP_Client"
        handle = win32gui.FindWindow(0, wdname1)

        if int(handle) <= 0:
            print("没有找到模拟器，退出进程................")
            exit(0)

        win32gui.MoveWindow(handle, 0, 0, 405, 756, True)  # 设定窗口
        time.sleep(1)
        win32gui.SetBkMode(handle, win32con.TRANSPARENT)  # 后台
        time.sleep(0.5)

        win32api.SetCursorPos([526, 274])  # 设置初始化鼠标位置

        self.left, self.top, self.right, self.bottom = win32gui.GetWindowRect(handle)
        # time.sleep(0.5)

    def run_p(self, run_num):
        
        # """拓扑图"""
        # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) # 点击进去拓扑图
        # time.sleep(1)
        # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        print('等待启动拓扑图......')
        time.sleep(50) # 等待启动
        print('开始启动设备......')
        win32api.keybd_event(17, 0, 0, 0)
        win32api.keybd_event(18, 0, 0, 0)
        win32api.keybd_event(run_num, 0, 0, 0)

        win32api.keybd_event(run_num, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)


def vi_ini_config():
    """修改配置文件"""
    ip_s = ''
    print("等待系统分配IP地址......")
    while True:
        ipList = socket.gethostbyname_ex(socket.gethostname())[2]
        for i in ipList:
            time.sleep(1)
            if '10.0' in i:
                ip_s += i
                break

        if '10.0' in ip_s:
            break

    print("等待系统分配IP地址成功,ip为:%s" % ip_s)
    f_r = open(confi_file, 'r')
    ip_str = f_r.readlines()
    f_r.close()

    f_w = open(confi_file, 'w')

    test_num = 0
    for i in ip_str:
        a = re.search(r"(?:[0-9]{1,3}\.){3}[0-9]{1,3}$", i)
        if a and test_num == 0:
            f_w.write(re.sub(r'(?:[0-9]{1,3}\.){3}[0-9]{1,3}$', ip, i))
            test_num += 1

        else:
            f_w.write(i)

    f_w.close()
    print('初始化配置文件成功')


if __name__ == '__main__':
    vi_ini_config()
    time.sleep(5)
    run = Run_win()
    # time.sleep(10)
    run.run_p(65)   # Alt+ + ctrl + a