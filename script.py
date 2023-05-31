#  READING SERIAL
import serial.tools.list_ports
import time

#  VOLUME CONTROL
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# --------------------------------------------------------------

# CONFIG
ratio = 1

def assignToFunction(code):
    code = code.strip()
    if code == "VOL-":
        changeVolume(ratio * -1)
    elif code == "VOL+":
        changeVolume(ratio)


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

def changeVolume(ratio):
    currVolume = volume.GetMasterVolumeLevel()
    nextVolume = currVolume + ratio
    
    nextVolume = max(-65.0, nextVolume)
    nextVolume = min(0.0, nextVolume)

    volume.SetMasterVolumeLevel(nextVolume)

if __name__ == "__main__":
    readSerial()