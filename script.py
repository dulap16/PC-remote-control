#  READING SERIAL
import serial.tools.list_ports
import time
import math

#  VOLUME CONTROL
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# --------------------------------------------------------------

# CONFIG
multiplier = 1.7

def assignToFunction(code):
    code = code.strip()
    if code == "VOL-":
        changeVolume(multiplier, -1)
    elif code == "VOL+":
        changeVolume(multiplier, 1)


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

def changeVolume(ratio, sign):
    currVolume = volume.GetMasterVolumeLevel()
    
    ratio = ((currVolume * -1) / 20) * multiplier
    math.floor(ratio)
    ratio = (ratio + 1) * sign

    nextVolume = currVolume + ratio
    
    nextVolume = max(-65.0, nextVolume)
    nextVolume = min(0.0, nextVolume)

    volume.SetMasterVolumeLevel(nextVolume, None)
    print(volume.GetMasterVolumeLevel())

if __name__ == "__main__":
    readSerial()