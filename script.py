#  READING SERIAL
import serial.tools.list_ports
import time
import math

#  VOLUME CONTROL
import win32con
import win32api
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# BRIGHTNESS CONTROL
import screen_brightness_control as sbc

# --------------------------------------------------------------

# CONFIG
selected = 0 
brightRatio = 5
volMultiplier = 1.7

def assignToFunction(code):
    code = code.strip()
    if code == "VOL-":
        changeVolume(volMultiplier, -1)
    elif code == "VOL+":
        changeVolume(volMultiplier, 1)
    elif code == "PREV":
        changeBrightness(brightRatio, -1)
    elif code == "NEXT":
        changeBrightness(brightRatio, 1)
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

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

def changeVolume(multiplier, sign):
    currVolume = volume.GetMasterVolumeLevel()
    
    win32api.keybd_event(win32con.VK_VOLUME_UP, 0)
    win32api.keybd_event(win32con.VK_VOLUME_UP, 0, win32con.KEYEVENTF_KEYUP)

    ratio = ((currVolume * -1) / 20) * multiplier
    math.floor(ratio)
    ratio = (ratio + 1) * sign

    nextVolume = currVolume + ratio
    
    nextVolume = max(-65.0, nextVolume)
    nextVolume = min(0.0, nextVolume)

    volume.SetMasterVolumeLevel(nextVolume, None)
    print(volume.GetMasterVolumeLevel())



# BRIGHTNESS CONTROL

def changeBrightness(ratio, sign):
    currBright = sbc.get_brightness()[0]
    nextBright = currBright + ratio * sign

    nextBright = max(0, nextBright)
    nextBright = min(100, nextBright)

    sbc.set_brightness(nextBright, display = 0)


if __name__ == "__main__":
    readSerial()