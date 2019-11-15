import time
import win32gui, win32con, win32api
import pyautogui

from PIL import ImageOps, Image, ImageGrab
from numpy import *
import numpy as np
import time
import cv2
import win32gui
from win32api import GetSystemMetrics
MonX = GetSystemMetrics(0)
MonY = GetSystemMetrics(1)

print("App started. Monitor resolution:")
print("Width =", MonX)
print("Height =", MonY)

#WINDOW_SUBSTRING = "Cisco IP Communicator"
#WINDOW_SUBSTRING = "Cisco Agent Desktop"

def getScreen(WINDOW_SUBSTRING):
    win_obj = {}
    window_info = get_window_info(WINDOW_SUBSTRING)
    if 'x' in window_info:
        img = get_screen(
            window_info["x"],
            window_info["y"],
            window_info["x"] + window_info["width"],
            window_info["y"] + window_info["height"]
        )
        win_obj["window_info"] = window_info
        win_obj["img"] = img
        return win_obj
    else:
        return 0

def get_window_info(WINDOW_SUBSTRING):
    window_info = {}
    window_info["winname"]=WINDOW_SUBSTRING
    win32gui.EnumWindows(set_window_coordinates, window_info)
    return window_info

def set_window_coordinates(hwnd, window_info):
    if win32gui.IsWindowVisible(hwnd):
        if window_info["winname"] in win32gui.GetWindowText(hwnd):
            rect = win32gui.GetWindowRect(hwnd)
            x = rect[0]
            y = rect[1]
            w = rect[2] - x
            h = rect[3] - y
            window_info['x'] = x
            window_info['y'] = y
            window_info['width'] = w
            window_info['height'] = h
            window_info['name'] = win32gui.GetWindowText(hwnd)
            win32gui.SetForegroundWindow(hwnd)


def get_screen(x1, y1, x2, y2):
    box = (x1, y1, x2, y2)
    screen = ImageGrab.grab(box)
    img = array(screen.getdata(), dtype=uint8).reshape((screen.size[1], screen.size[0], 3))
    return img


def get_object_coord(template,WINDOW_SUBSTRING):
    win_obj = getScreen(WINDOW_SUBSTRING)
    if win_obj == 0:
        return 0
    method = cv2.TM_SQDIFF_NORMED
    small_image = cv2.imread(template)
    large_image = win_obj["img"]
    result = cv2.matchTemplate(small_image, large_image, method)
    mn, _, mnLoc, _ = cv2.minMaxLoc(result)
    MPx, MPy = mnLoc
    trows, tcols = small_image.shape[:2]
    DistX = win_obj["window_info"]["x"]+round(MPx + (tcols/2))
    DistY = win_obj["window_info"]["y"]+round(MPy + (trows/2))
    return ([DistX, DistY])

def detect_object_screen(template,WINDOW_SUBSTRING):
    win_obj = getScreen(WINDOW_SUBSTRING)
    if win_obj == 0:
        return 0
    method = cv2.TM_SQDIFF_NORMED
    small_image = cv2.imread(template)
    large_image = win_obj["img"]
    result = cv2.matchTemplate(small_image, large_image, cv2.TM_SQDIFF)
    mn, _, mnLoc, _ = cv2.minMaxLoc(result)
    MPx, MPy = mnLoc
    trows, tcols = small_image.shape[:2]
    cv2.rectangle(large_image, (MPx, MPy), (MPx + tcols, MPy + trows), (0, 0, 255), 2)
    return (large_image)

def detect_updating_screen(template,WINDOW_SUBSTRING):
    win_obj = getScreen(WINDOW_SUBSTRING)
    if win_obj == 0:
        return 0
    target_widget_coordinates = {}
    large_image = cv2.cvtColor(win_obj["img"], cv2.COLOR_BGR2GRAY)
    small_image = cv2.imread(template, 0)
    w, h = small_image.shape[::-1]
    result = cv2.matchTemplate(large_image, small_image, cv2.TM_CCOEFF_NORMED)
    threshold = 0.8
    loc = np.where(result >= threshold)
    if count_nonzero(loc) > 0:
        for pt in zip(*loc[::-1]):
            target_widget_coordinates = {"x": pt[0], "y": pt[1]}
            cv2.rectangle(large_image, pt, (pt[0] + w, pt[1] + h), (255, 0, 0), 2)
            DistX = int(win_obj["window_info"]["x"] + round(pt[0] + (w / 2)))
            DistY = int(win_obj["window_info"]["y"] + round(pt[1] + (h / 2)))
            return ([DistX, DistY])
    if not target_widget_coordinates:
        return 0

def getNumbers(WINDOW_SUBSTRING):
    btns = []
    i=0
    while i<10:
        btns.append(get_object_coord('templates/ba'+str(i)+'.png',WINDOW_SUBSTRING))
        #a = detect_object_screen('templates/ba'+str(i)+'.png',win_obj)
        #cv2.imwrite(str(i)+'.png',a)
        i+=1
    return btns

def click(x,y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


