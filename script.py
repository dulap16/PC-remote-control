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

# --------------------------------------------------------------

# CONFIG
selected = 0 
 
isMuted = False

brightRatio = 5
currBright = sbc.get_brightness()[0]

moveDist = 5
moveTime = 0.02 * moveDist
scrollDist = 50

async def assignToFunction(code):

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
        moveMouse(0, -moveDist)
    elif code == "8":
        moveMouse(0, moveDist)
    elif code == "4":
        moveMouse(-moveDist, 0)
    elif code == "6":
        moveMouse(moveDist, 0)
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
            pressed = pressed.strip()
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

def moveMouse(x, y):
    pyautogui.move(x, y, moveTime)


if __name__ == "__main__":
    readSerial()