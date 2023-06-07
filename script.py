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
    """Encapsulates some calls to the winapi for window management"""

    def __init__ (self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name=None):
        """find a window by its class_name"""
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        """Pass to win32gui.EnumWindows() to check all the opened windows"""
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        """find a window whose title matches the wildcard regex"""
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)
    

def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

def get_app_list(handles=[]):
    mlst=[]
    win32gui.EnumWindows(window_enum_handler, handles)
    for handle in handles:
        mlst.append(handle)
    return mlst

if __name__ == "__main__":
    appwindows = get_app_list()
    for i in appwindows:
        try:
            win32gui.SetForegroundWindow(i[0])
            time.sleep(1)
        except:
            print("%a can't be set as foreground", i[1])
    

    readSerial()