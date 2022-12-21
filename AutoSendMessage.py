#!/usr/bin/python
# -*- coding:UTF-8 -*-
import datetime
import os
import time
import pyautogui
import psutil
import win32clipboard as clip
import win32con
import win32gui
import win32api


# 判断TIM是否登录，没登录则登录
def TIM_login():
    pl = psutil.pids()
    flag = False
    for pid in pl:
        if psutil.Process(pid).name() == 'TIM.exe':
            flag = True
    return flag


def message_get():
    with open('./message.txt',encoding='UTF-8') as f1:
        message = f1.readlines()
        for msg in message:
            return msg


def people_get():
    with open('./people.txt',encoding='UTF-8') as f2:
        people = f2.readlines()
        return people


def copy(text):
    clip.OpenClipboard()
    clip.EmptyClipboard()
    clip.SetClipboardData(win32con.CF_UNICODETEXT,text)
    clip.CloseClipboard()


def TIM_send(title,p,num=1):
    try:
        # 获取窗口句柄
        handle = win32gui.FindWindow(None, title.strip('\n'))  # 通过窗口标题获取窗口句柄
        # print("窗口句柄是：{}".format(handle))
        # 打开窗口
        win32gui.SetForegroundWindow(handle)
        # 窗口最大化
        win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)
        time.sleep(1.5)
        # 鼠标位置设定
        win32api.SetCursorPos((130,60))
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
            win32api.SetCursorPos((1000,940))
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
            TIM_send(title, p, num=1)
            title = p
        # 最小化窗口
        handle = win32gui.FindWindow(None, p.strip('\n'))
        win32gui.SetForegroundWindow(handle)
        win32gui.ShowWindow(handle, win32con.SW_MINIMIZE)
        # 提示完成
        win32api.MessageBox(0, "脚本结束", "提示", win32con.MB_OK)
    else:
        print("[!]未检测到TIM.exe，请重试")
    pass


if __name__ == '__main__':
    main()
