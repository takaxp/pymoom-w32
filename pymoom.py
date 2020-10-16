import keyboard
import win32gui
import math

win32offset=(-7,0,0,0)
shift_amount=(200, 200)
menu_bar=(0, 33)
disable_apps=['emacs', 'File Explorer']

def moom_available_windows(hwnd):
    result = True
    name = win32gui.GetWindowText(hwnd)
    for app in disable_apps:
        if (app in name):
            result = False
    return result;

def moom_move_frame():
    hwnd = win32gui.GetForegroundWindow()
    if not (moom_available_windows(hwnd)):
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win32gui.MoveWindow(hwnd,
                        win32offset[0],
                        win32offset[1],
                        right-left,
                        bottom-top,
                        True)
    print("move to Top-Left")

def moom_move_frame_left():
    hwnd=win32gui.GetForegroundWindow()
    if not (moom_available_windows(hwnd)):
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win32gui.MoveWindow(hwnd,
                        left-shift_amount[0],
                        top+win32offset[1],
                        right-left,
                        bottom-top,
                        True)
    print("move to Left")

def moom_move_frame_right():
    hwnd=win32gui.GetForegroundWindow()
    if not (moom_available_windows(hwnd)):
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win32gui.MoveWindow(hwnd,
                        left+shift_amount[1],
                        top+win32offset[1],
                        right-left,
                        bottom-top,
                        True)
    print("move to Reft")

def moom_move_frame_center():
    hwnd=win32gui.GetForegroundWindow()
    if not (moom_available_windows(hwnd)):
        return
    fleft, ftop, right, bottom = win32gui.GetWindowRect(hwnd)
    fwidth=right-fleft
    fheight=bottom-ftop
    screen=win32gui.GetWindowRect(win32gui.GetDesktopWindow())
    win32gui.MoveWindow(hwnd,
                        math.floor((screen[2]-screen[0])/2.0)-math.floor(fwidth/2.0)+win32offset[0],
                        math.floor((screen[3]-screen[1])/2.0)-math.floor(fheight/2.0)+win32offset[1],
                        fwidth,
                        fheight,
                        True)
    print("move to Center")

def moom_move_frame_to_edge_top():
    hwnd=win32gui.GetForegroundWindow()
    if not (moom_available_windows(hwnd)):
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win32gui.MoveWindow(hwnd,
                        left,
                        win32offset[1],
                        right-left,
                        bottom-top,
                        True)
    print("move to Edge-Top")

def moom_move_frame_to_edge_bottom():
    hwnd=win32gui.GetForegroundWindow()
    if not (moom_available_windows(hwnd)):
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    screen=win32gui.GetWindowRect(win32gui.GetDesktopWindow())
    win32gui.MoveWindow(hwnd,
                        left,
                        screen[3]-(bottom-top)+win32offset[1]-menu_bar[1],
                        right-left,
                        bottom-top,
                        True)
    print("move to Edge-Bottom")

try:
    keyboard.add_hotkey('alt+0', moom_move_frame)
    keyboard.add_hotkey('alt+1', moom_move_frame_left)
    keyboard.add_hotkey('alt+3', moom_move_frame_right)
    keyboard.add_hotkey('alt+2', moom_move_frame_center)
    keyboard.add_hotkey('f1', moom_move_frame_to_edge_top)
    keyboard.add_hotkey('shift+f1', moom_move_frame_to_edge_bottom)

except:
    pass

keyboard.wait('ctrl+tab')
