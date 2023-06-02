#  READING SERIAL
import serial.tools.list_ports
import time

#  VOLUME CONTROL
import win32con
import win32api

# BRIGHTNESS CONTROL
import screen_brightness_control as sbc
import asyncio

# --------------------------------------------------------------

# CONFIG
selected = 0 
 
brightRatio = 5
currBright = sbc.get_brightness()[0]

def assignToFunction(code):
    code = code.strip()
    if code == "VOL-":
        changeVolume(-1)
    elif code == "VOL+":
        changeVolume(1)
    elif code == "PREV":
        asyncio.run((changeBrightness(brightRatio, -1)))
    elif code == "NEXT":
        asyncio.run((changeBrightness(brightRatio, 1)))
    elif code == "EQ":
        selected = (selected + 1) % 2


# READING SERIAL
serialInst = serial.Serial()
port = "COM3"

def readSerial():
    serialInst.baudrate = 9600
    serialInst.port = port
    serialInst.open()

    while True:
        if serialInst.in_waiting:
            packet = serialInst.readline()
            pressed = (packet.decode('utf')).rstrip('\n')
            print(pressed)
            assignToFunction(pressed)


# VOLUME CONTROL

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

if __name__ == "__main__":
    readSerial()