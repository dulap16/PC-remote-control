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

# --------------------------------------------------------------

# CONFIG
commandFinished = True
selected = 0 
 
isMuted = False

brightRatio = 5
currBright = sbc.get_brightness()[0]

moveDist = 5
moveTime = 0.05
scrollDist = 50

switchingManager = None

async def assignToFunction(code):
    commandFinished = False
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
        changeBrightness(brightRatio, -1)
    elif code == "NEXT":
        changeBrightness(brightRatio, 1)
    
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
    
    # PRESS SPACE
    elif code == "0":
        pressSpace()

    # SWITCHING TABS
    elif code == "CH":
        switchingManager.modeSwitched()
    elif code == "CH+":
        switchingManager.nextTab()
    elif code == "CH-":
        switchingManager.previousTab()


    serialInst.flushInput()

    
    

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
            asyncio.run(assignToFunction(pressed))


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

def changeBrightness(ratio, sign):
    global currBright
    nextBright = currBright + ratio * sign

    nextBright = max(0, nextBright)
    nextBright = min(100, nextBright)
    currBright = nextBright

    sbc.set_brightness(nextBright, display = 10)    

# PRESS SPACE

def pressSpace():
    pyautogui.press('space')

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

class SwitchingManager:

    def __init__(self):
        self.switchingOn = False

    def modeSwitched(self):
        if self.switchingOn is False:
            self.startSwitching()
        else:
            self.endSwitching()

    def startSwitching(self):
        self.switchingOn = True

        pyautogui.keyDown('alt')
        self.activateInterface()
    
    def endSwitching(self):
        self.switchingOn = False

        pyautogui.keyUp('alt')

    def activateInterface(self):
        self.nextTab()
        self.previousTab()

    def nextTab(self):
        if self.switchingOn is True:
            pyautogui.press('tab')
    
    def previousTab(self):
        if self.switchingOn is True:
            with pyautogui.hold('ctrlleft'):
                pyautogui.press('tab')


if __name__ == "__main__":
    switchingManager = SwitchingManager()
    readSerial()