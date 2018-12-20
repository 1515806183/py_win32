# coding=utf-8
import win32api,win32gui,win32con
import os, time



#
# wdname1 = u"eNSP"
#
# w1hd = win32gui.FindWindow(0, wdname1)
#
# w2hd = win32gui.FindWindowEx(w1hd, None, None, None)
#
# win32gui.SetForegroundWindow(w2hd)


if __name__ == '__main__':
    # file_name = "F:\\test_win\\eNSP\\eNSP_Client.exe"
    #
    # os.startfile(file_name)
    # time.sleep(15)
    wdname1 = u"eNSP_Client"
    handle = win32gui.FindWindow(0, wdname1)
    print(handle)
    if int(handle) <= 0:
        print("没有找到模拟器，退出进程................")
        exit(0)

    win32gui.MoveWindow(handle, 0, 0, 405, 756, True)
    time.sleep(1)
    win32gui.SetBkMode(handle, win32con.TRANSPARENT)
    time.sleep(1)

    win32api.SetCursorPos([0, 0])

    left, top, right, bottom = win32gui.GetWindowRect(handle)
    print(left, top, right, bottom)
    print(right)
    time.sleep(2)

    # 关闭主程序
    # win32api.SetCursorPos((left + 990, bottom - 740))
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    # time.sleep(0.2)
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    # 调色板
    # wdname2 = u"调色板"
    # handle2 = win32gui.FindWindow(0, wdname2)
    # left, top, right, bottom = win32gui.GetWindowRect(handle2)
    # print(left, top, right, bottom)
    # win32api.SetCursorPos([800, 290])
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    # time.sleep(0.2)
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    # 采集数据报文
    # wdname3 = u"采集数据报文"
    # handle2 = win32gui.FindWindow(0, wdname2)
    # left, top, right, bottom = win32gui.GetWindowRect(wdname3)
    # print(left, top, right, bottom)
    # win32api.SetCursorPos([732, 140])
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    # time.sleep(0.2)
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    # 选项
    # wdname4 = u"选项"
    # handle2 = win32gui.FindWindow(0, wdname4)
    # left, top, right, bottom = win32gui.GetWindowRect(handle2)
    # print(left, top, right, bottom)
    # win32api.SetCursorPos([720, 130])
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    # time.sleep(0.2)
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)






