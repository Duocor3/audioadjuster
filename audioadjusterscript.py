from ctypes import cast, POINTER
from tkinter import *

import pyautogui
import pywintypes
import time
import win32api
import win32con
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Get default audio device using PyCAW
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

theaterMode = [False]


def change_refresh_rate(mode):
    # Gets the devmode settings
    devmode = pywintypes.DEVMODEType()

    # Sets the display frequency to the specified refresh rate
    if not mode[0]:
        devmode.DisplayFrequency = 60
        # Changes the button text
        theaterModeButton['text'] = 'Theater Mode: ON'
        theaterModeButton['fg'] = 'green'
    else:
        devmode.DisplayFrequency = 144
        # Changes the button text
        theaterModeButton['text'] = 'Theater Mode: OFF'
        theaterModeButton['fg'] = 'red'

    # Toggles theater mode in dragon center
    dragon_center_shortcut(mode)

    mode[0] = not mode[0]
    devmode.Fields = win32con.DM_DISPLAYFREQUENCY

    # Applies the change
    win32api.ChangeDisplaySettings(devmode, 0)

    # Prints the new refresh rate in the console
    device = win32api.EnumDisplayDevices()
    settings = win32api.EnumDisplaySettings(device.DeviceName, -1)
    for varName in ['Color', 'BitsPerPel', 'DisplayFrequency']: 
        print("%s: %s" % (varName, getattr(settings, varName)))


def dragon_center_shortcut(mode):
    # Activates the keyboard shortcut for theater mode in Dragon Center
    if not mode[0]:
        # Turns on theater mode
        pyautogui.keyDown('alt')
        time.sleep(.2)
        pyautogui.press('2')
        time.sleep(.2)
        pyautogui.keyUp('alt')
    else:
        # Goes back to casual mode
        pyautogui.keyDown('alt')
        time.sleep(.2)
        pyautogui.press('4')
        time.sleep(.2)
        pyautogui.keyUp('alt')


def adjust_volume(adjustment):
    # Get current volume of the right channel
    currentVolumeRight = volume.GetChannelVolumeLevel(1)
    # Set the volume of the left channel to twice of the volume of the right channel
    new_volume = currentVolumeRight + adjustment

    if new_volume > -0.07:
        new_volume = -0.07
    print(new_volume)
    volume.SetChannelVolumeLevel(0, new_volume, None)
    # NOTE: -6.0 dB = half volume !


# sets up the window
window = Tk()
# creates and sets up the adjust audio button
adjustButton = Button(window, text="Adjust Volume Balance", fg="green")
adjustButton.place(x=20, y=10)
adjustButton.bind("<Button-1>", lambda event: adjust_volume(15.5))

# creates and sets up the reset button
resetButton = Button(window, text="Reset Volume Balance", fg="blue")
resetButton.place(x=20, y=60)
resetButton.bind("<Button-1>", lambda event: adjust_volume(0.0))

# creates and sets up the theater mode button
theaterModeButton = Button(window, text="Theater Mode: OFF", fg="red")
theaterModeButton.place(x=20, y=110)
theaterModeButton.bind("<Button-1>", lambda event: change_refresh_rate(theaterMode))

# names and positions the window
window.title("Audio Adjuster")
window.geometry("300x200+800+250")

window.mainloop()
