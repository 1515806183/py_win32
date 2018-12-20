# coding=utf-8
"""
python3.5 + pywin32 224
"""
import os
import time
import win32gui
import win32api
import win32con

# win32api.GetCursorPos()  获取鼠标的位置

# ensp  绝对路径 路径
file_name = "F:\\test_win\\eNSP\\eNSP_Client.exe"
os.startfile(file_name)


class Run_win(object):
    def __init__(self):
        """
        初始化
        获取窗口焦点
        """
        wdname1 = u"eNSP_Client"
        handle = win32gui.FindWindow(0, wdname1)
        print(handle)
        if int(handle) <= 0:
            print("没有找到模拟器，退出进程................")
            exit(0)

        win32gui.MoveWindow(handle, 0, 0, 405, 756, True)  # 设定窗口
        time.sleep(1)
        win32gui.SetBkMode(handle, win32con.TRANSPARENT)  # 后台
        time.sleep(1)

        win32api.SetCursorPos([0, 0])  # 设置初始化鼠标位置

        self.left, self.top, self.right, self.bottom = win32gui.GetWindowRect(handle)
        time.sleep(2)

    def run_p(self, run_num):
        """# 调色板"""
        time.sleep(1)
        win32api.keybd_event(17, 0, 0, 0)
        win32api.keybd_event(18, 0, 0, 0)
        win32api.keybd_event(run_num, 0, 0, 0)

        win32api.keybd_event(run_num, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
        # 关闭调色板
        time.sleep(3)
        if run_num == 80:
            self.run_close_test_all(run_name=u"调色板", run_index=[800, 290])
        if run_num == 68:
            self.run_close_test_all(run_name=u"采集数据报文", run_index=[732, 140])
        if run_num == 69:
            self.run_close_test_all(run_name=u"选项", run_index=[720, 130])

    def close_all(self):
        # 关闭主程序
        win32api.SetCursorPos((self.left + 990, self.bottom - 740))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        time.sleep(0.2)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    def run_close_test_all(self, run_name, run_index):
        # 调色板
        wdname2 = run_name
        handle2 = win32gui.FindWindow(0, wdname2)
        left, top, right, bottom = win32gui.GetWindowRect(handle2)
        win32api.SetCursorPos(run_index)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        time.sleep(0.2)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


if __name__ == '__main__':
    time.sleep(13)  # 程序启动，如果程序是打开的可以取消此操作
    run = Run_win()
    # 调色板快捷键Alt+ + ctrl + p    80
    # 数据抓包快捷键Alt+ + ctrl + d  68
    # 选项快捷键Alt+ + ctrl + e     69
    # 启动设备快捷键Alt+ + ctrl + a  65
    # 停止设备快捷键Alt+ + ctrl + c  67
    run.run_p(80)
    time.sleep(2)
    run.run_p(68)
    time.sleep(2)
    run.run_p(69)
    time.sleep(2)
    run.close_all()
