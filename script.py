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

windowSwitcher = None

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

    # SWITCHING WINDOWS
    elif code == "CH":
        windowSwitcher.startSwitchingWindows()
    elif code == "CH+":
        windowSwitcher.goToNextWindow()
    elif code == "CH-":
        windowSwitcher.goToPreviousWindow()



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
    """Encapsulates some calls to the winapi for window management"""
    def __init__(self):
        self.shell = win32com.client.Dispatch("WScript.Shell")

    def setWindowActive(self, hwnd):
        self.shell.SendKeys('%')
        print(win32gui.GetWindowText(hwnd), " sent to the front.")

        win32gui.SetForegroundWindow(hwnd)
        win32gui.BringWindowToTop(hwnd)
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    
    def windowEnumHandler(self, hwnd, resultList):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
            resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

    def getApps(self, handles=[]):
        mlst=[]
        win32gui.EnumWindows(self.windowEnumHandler, handles)
        for handle in handles:
            mlst.append(handle)
        return mlst


class WindowSwitcher:

    def __init__(self):
        self.index = 0
        self.howManyApps = 0
        self.switchingWindowsActivated = False
        
        self.windowManager = None
        self.currentHandler = None
        self.apps = None
    
    def startSwitchingWindows(self):
        self.index = 0
        self.windowManager = WindowMgr()
        self.apps = self.windowManager.getApps()

        self.howManyApps = len(self.apps)
        self.printWindows()

        self.switchingWindowsActivated = True 

    def goToNextWindow(self):
        if self.switchingWindowsActivated is True:
            self.index = (self.index + 1) % self.howManyApps
            
            self.selectWindowByIndex(self.index)
    
    def goToPreviousWindow(self):
        if self.switchingWindowsActivated is True:
            self.index = self.index - 1
            if self.index < 0:
                self.index = self.howManyApps - 1

            self.selectWindowByIndex(self.index)
    
    def selectWindowByIndex(self, index):
        prev = self.currentHandler
        try:
            self.currentHandler = self.apps[index][0]
            
            self.windowManager.setWindowActive(self.currentHandler)
        except:
            self.currnetHandler = prev
    
    def endSwitchingWindows(self):
        self.switchingWindowsActivated = False

    def printWindows(self):
        print(self.howManyApps, " windows active:")
        for app in self.apps:
            print(app[1])


if __name__ == "__main__":
    windowSwitcher = WindowSwitcher()
    readSerial()