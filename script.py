#  READING SERIAL
import serial.tools.list_ports
import time

#  VOLUME CONTROL
import win32con
import win32api

# BRIGHTNESS CONTROL
import screen_brightness_control as sbc
import asyncio

# MOUSE MOVEMENT
import pyautogui

# SWITCHING WINDOWS
import win32gui
import win32com
import win32com.client

import re

# --------------------------------------------------------------

# CONFIG
selected = 0 
 
isMuted = False

brightRatio = 5
currBright = sbc.get_brightness()[0]

moveDist = 5
moveTime = 0.05
scrollDist = 50

def assignToFunction(code):
    code = code.strip()
    
    # VOLUME CONTROL
    if code == "VOL-":      
        changeVolume(-1)
    elif code == "VOL+":
        changeVolume(1)
    elif code == "EQ":
        muteVolume()
    
    # BRIGHTNESS CONTROL
    elif code == "PREV":    
        asyncio.run((changeBrightness(brightRatio, -1)))
    elif code == "NEXT":
        asyncio.run((changeBrightness(brightRatio, 1)))
    
    # MOUSE MOVEMENT
    elif code == "2":
        moveMouseUp(moveDist)
    elif code == "8":
        moveMouseDown(moveDist)
    elif code == "4":
        moveMouseLeft(moveDist)
    elif code == "6":
        moveMouseRight(moveDist)
    elif code == "100+":
        changeSensitivity(-5)
    elif code == "200+":
        changeSensitivity(5)

    # OTHER MOUSE FUNCTIONS
    elif code == "5":
        mouseClick()
    elif code == "3":
        scroll(scrollDist)
    elif code == "9":
        scroll(scrollDist * -1)

    
    


# READING SERIAL
serialInst = serial.Serial()
port = "COM3"

def readSerial():
    serialInst.baudrate = 4800
    serialInst.port = port
    serialInst.open()

    while True:
        if serialInst.in_waiting:
            packet = serialInst.readline()
            pressed = (packet.decode('utf')).rstrip('\n')
            print(pressed)
            assignToFunction(pressed)


# VOLUME CONTROL

def muteVolume():
    for i in range(0, 50):
        win32api.keybd_event(win32con.VK_VOLUME_DOWN, 0)
        win32api.keybd_event(win32con.VK_VOLUME_DOWN, 0, win32con.KEYEVENTF_KEYUP)

def changeVolume(sign):

    if(sign == 1):
        win32api.keybd_event(win32con.VK_VOLUME_UP, 0)
        win32api.keybd_event(win32con.VK_VOLUME_UP, 0, win32con.KEYEVENTF_KEYUP)
    else:
        win32api.keybd_event(win32con.VK_VOLUME_DOWN, 0)
        win32api.keybd_event(win32con.VK_VOLUME_DOWN, 0, win32con.KEYEVENTF_KEYUP)


# BRIGHTNESS CONTROL

async def changeBrightness(ratio, sign):
    global currBright
    nextBright = currBright + ratio * sign

    nextBright = max(0, nextBright)
    nextBright = min(100, nextBright)
    currBright = nextBright

    sbc.set_brightness(nextBright, display = 10)    

# MOUSE MOVEMENT

def changeSensitivity(diff):
    global moveDist

    moveDist = moveDist + diff
    moveDist = max(5, moveDist)
    moveDist = min(60, moveDist)

def scroll(dist):
    pyautogui.scroll(dist)

def mouseClick():
    pyautogui.click(pyautogui.position())

def moveMouseUp(dist):
    print("Move up")
    pos = pyautogui.position()

    x = pos[0]
    y = pos[1] - dist

    pyautogui.moveTo(x, y, moveTime)

def moveMouseDown(dist):
    print("Move down")
    pos = pyautogui.position()

    x = pos[0]
    y = pos[1] + dist

    pyautogui.moveTo(x, y, moveTime)

def moveMouseLeft(dist):
    print("Move left")
    pos = pyautogui.position()

    x = pos[0] - dist
    y = pos[1] 

    pyautogui.moveTo(x, y, moveTime)

def moveMouseRight(dist):
    print("Move right")
    pos = pyautogui.position()

    x = pos[0] + dist
    y = pos[1] 

    pyautogui.moveTo(x, y, moveTime)


# SWITCHING WINDOWS

class WindowMgr:

    def __init__(self):
        self._hwnd = None

    """Encapsulates some calls to the winapi for window management"""
    def window_enum_handler(self, hwnd, resultList):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
            resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

    def get_app_list(self, handles=[]):
        mlst=[]
        win32gui.EnumWindows(self.window_enum_handler, handles)
        for handle in handles:
            mlst.append(handle)
        return mlst
    



if __name__ == "__main__":
    windowManager = WindowMgr()
    appWindows = windowManager.get_app_list()
    shell = win32com.client.Dispatch("WScript.Shell")
    for i in appWindows:
        try:
            shell.SendKeys('%')
            print(i[1], " sent to the front.")

            win32gui.SetForegroundWindow(i[0])
            win32gui.BringWindowToTop(i[0])
            # win32gui.ShowWindow(i[0], win32con.SW_MAXIMIZE)
            win32gui.ShowWindow(i[0], win32con.SW_MAXIMIZE)
            time.sleep(1)
        except:
            print("%a can't be set as foreground", i[1])
    

    readSerial()