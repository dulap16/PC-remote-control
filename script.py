#  READING SERIAL
import serial.tools.list_ports
import time

#  VOLUME CONTROL
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# --------------------------------------------------------------

def assignToFunction(code):
    if code == "VOL-":
        print("SET VOLUME DOWN")
    elif code == "VOL+":
        print("SET VOLUME UP")


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


if __name__ == "__main__":
    readSerial()

            