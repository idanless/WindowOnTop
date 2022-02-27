import win32gui
import win32process
import psutil
import win32con
import winrt.windows.ui.notifications as notifications
import winrt.windows.data.xml.dom as dom
from pynput import keyboard
import os
import time
import win32con
nManager = notifications.ToastNotificationManager
notifier = nManager.create_toast_notifier("Chrome")
#os.system('taskkill /f /im  ontopwin.exe')

def Notification(status,msg):
    tString =f"""
    <toast>
        <visual>
            <binding template='ToastGeneric'>
                <text>{status}</text>
                <text>{msg}</text>
            </binding>
        </visual>
    </toast>
    """

    # convert notification to an XmlDocument
    xDoc = dom.XmlDocument()
    xDoc.load_xml(tString)

    # display notification
    notifier.show(notifications.ToastNotification(xDoc))

def get_active_executable_name():
    try:
        process_id = win32process.GetWindowThreadProcessId(
            win32gui.GetForegroundWindow()
        )

        return ".".join(psutil.Process(process_id[-1]).name().split(".")[:-1])
    except Exception as exception:
        return None
#

def LockWindow():
    nameWindow = get_active_executable_name()
    ClassWindows = win32gui.GetClassName(win32gui.GetForegroundWindow())
    hwnd = win32gui.FindWindow(win32gui.GetClassName(win32gui.GetForegroundWindow()), None)
    Size = win32gui.GetWindowRect(hwnd)
    Notification("*LockOnTop*",f"Window {nameWindow}")
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, Size[2], Size[3], win32con.SWP_NOMOVE + win32con.SWP_NOSIZEAZ)
    #win32con.SWP_NOMOVE + win32con.SWP_NOSIZE
    time.sleep(3)

def UnLockWindow():
    nameWindow = get_active_executable_name()
    ClassWindows = win32gui.GetClassName(win32gui.GetForegroundWindow())
    hwnd = win32gui.FindWindow(win32gui.GetClassName(win32gui.GetForegroundWindow()), None)
    Size = win32gui.GetWindowRect(hwnd)
    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, Size[2], Size[3], win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
    Notification("*UnLock*", f"Window {nameWindow}")
    time.sleep(3)


# The key combinations to look for
COMBINATIONS = [
    {keyboard.Key.shift, keyboard.KeyCode(vk=65)}, # shift + a (see below how to get vks)
    {keyboard.Key.shift, keyboard.KeyCode(vk=90)}  # shift + a (see below how to get vks)
]


def execute(key):
    """ My function to execute when a combination is pressed """
    if key == 65:
        LockWindow()
    if key == 90:
        UnLockWindow()



# The currently pressed keys (initially empty)
pressed_vks = set()

def get_vk(key):
    """Z
    Get the virtual key code from a key.
    These are used so case/shift modifications are ignored.
    """
    return key.vk if hasattr(key, 'vk') else key.value.vk


def is_combination_pressed(combination):
    """ Check if a combination is satisfied using the keys pressed in pressed_vks """
    return all([get_vk(key) in pressed_vks for key in combination])


def on_press(key):
    """ When a key is pressed """
    vk = get_vk(key)
    # Get the key's vk
    pressed_vks.add(vk)  # Add it to the set of currently pressed keys

    for combination in COMBINATIONS:  # Loop though each combination
        if is_combination_pressed(combination):  # And check if all keys are pressedAZAZAZAZ
            execute(vk)  # If tAhey are all pressed, call your function
            break  # Don't allow execute to be called more than once per key press


def on_release(key):
    """ When a key is released """
    vk = get_vk(key)  # Get the key's vk
    try:
        pressed_vks.remove(vk)
    except KeyError:
        pressed_vks.clear()
        # Remove it from the set of currently pressed keysAZ


with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
    listener.join()  # Join the listener thread to the current thread so we don't exit before it stops
