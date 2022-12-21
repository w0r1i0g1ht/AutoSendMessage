# -*- coding:UTF-8 -*-
import datetime
import re
import time
import pyautogui
import psutil
import win32clipboard as clip
import win32con
import win32gui
import win32api
# 不导入pyautogui可能导致鼠标设置定位出错


def TIM_login():
    """
    判断TIM是否登录，没登录则登录
    :return: 返回一个是否登录bool值
    """
    pl = psutil.pids()
    flag = False
    for pid in pl:
        if psutil.Process(pid).name() == 'TIM.exe':
            flag = True
    return flag


def message_get():
    """
    获得message.txt里的内容
    :return: 返回message.txt里的内容
    """
    with open('./message.txt', encoding='UTF-8') as f1:
        message = f1.readlines()
        for msg in message:
            return msg


def people_get():
    """
    获得people.txt里的内容
    :return: 返回people.txt的目标的数组
    """
    with open('./people.txt', encoding='UTF-8') as f2:
        people = f2.readlines()
        return people


def copy(text):
    """
    将消息传入剪贴板
    :param text: 消息的内容
    :return: 无
    """
    clip.OpenClipboard()
    clip.EmptyClipboard()
    clip.SetClipboardData(win32con.CF_UNICODETEXT, text)
    clip.CloseClipboard()


def find_window(title):
    """
    通过正则表达式对窗口标题进行模糊匹配
    :param title: 需要模糊匹配的窗口标题
    :return: 返回匹配到的窗口句柄
    """
    # 正则表达式
    pattern = f".*{title}.*"
    windows = []

    # 回调函数
    def callback(hwnd, windows):
        # 根据窗口句柄获取窗口名字
        name = win32gui.GetWindowText(hwnd)
        if re.match(pattern, name):
            windows.append(hwnd)
        return True

    win32gui.EnumWindows(callback, windows)
    return windows[0]


def TIM_send(title, p, num=1):
    """
    发送消息
    :param title: 需要模糊匹配的窗口标题
    :param p: 需要发送消息的目标
    :param num: 发送次数
    :return: 无
    """
    try:
        # 获取窗口句柄
        handle = find_window(title)
        # 打开窗口
        win32gui.SetForegroundWindow(handle)
        # 窗口最大化
        win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)
        time.sleep(1.5)
        # 鼠标位置设定
        win32api.SetCursorPos((130, 60))
        # 点击一次鼠标
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
        # ctrl+V
        time.sleep(1)
        win32api.keybd_event(0x11, 0, 0, 0)
        win32api.keybd_event(0x56, 0, 0, 0)
        win32api.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        # enter
        time.sleep(1)
        win32api.keybd_event(0x0D, 0, 0, 0)
        win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)
        # 文案
        p = p.strip('\n')
        msg = f"状态: 脚本运行中\n内容: {message_get()}\n对象: {p}\n时间: {datetime.datetime.now()}"
        # 将消息文本复制到剪贴板
        copy(msg)
        for i in range(num):
            # 鼠标位置设定
            win32api.SetCursorPos((1000, 940))
            # 点击一次鼠标
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
            # ctrl+V
            win32api.keybd_event(0x11, 0, 0, 0)
            win32api.keybd_event(0x56, 0, 0, 0)
            win32api.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
            # enter
            time.sleep(1)
            win32api.keybd_event(0x0D, 0, 0, 0)
            win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.5)
            print("[*]目标", p, "发送成功")
    except Exception as e:
        print("[!]窗口句柄获取失败，请检查窗口名字是否正确")


def main():
    if TIM_login():
        print("[*]检测到TIM已登录")
        # title是窗口标题
        title = people_get()[0]
        for p in people_get():
            p = p.strip('\n')
            print("[*]当前发送目标", p)
            copy(p)
            TIM_send(title, p, num=2)
            title = p
        # 最小化窗口
        handle = find_window(p)
        win32gui.SetForegroundWindow(handle)
        win32gui.ShowWindow(handle, win32con.SW_MINIMIZE)
        # 提示完成
        win32api.MessageBox(0, "脚本结束", "提示", win32con.MB_OK)
    else:
        print("[!]未检测到TIM.exe，请重试")

    pass


if __name__ == '__main__':
    main()
