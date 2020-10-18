import os
import csv
import math
import win32gui
import win32process
import subprocess
import keyboard

pymoom_w32='0.9.1'

###############################################################################
# User variables
win32offset=(-7,0,0,0)
shift_amount=(200, 200)
menu_bar=(0, 33)
exclude_apps=['Emacs'] # 'CabinetWClass'=explore.exe
apps = {'emacs' : [r'C:\Apps\emacs-27.1-x86_64\bin\runemacs.exe', ''],
        'mintty' : [r'C:\cygwin64\bin\mintty.exe',
                    r' -i /Cygwin-Terminal.ico -'],
        'explorer' : [r'C:\Windows\explorer.exe', '']}
###############################################################################

# Utilities
def moom_on_emacs():
    hwnd = win32gui.GetForegroundWindow()
    className = win32gui.GetClassName(hwnd)
    if 'Emacs' in className:
        return True
    return False

def moom_exclude_windows(hwnd):
    className = win32gui.GetClassName(hwnd)
    # print(className)
    for app in exclude_apps:
        if (app in className):
            return True
    return False

def moom_get_hwnds(pid):
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            t, p = win32process.GetWindowThreadProcessId(hwnd)
            if p == pid:
                hwnds.append(hwnd)
        return True
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds

def moom_focus(pid):
    try:
        if pid is not None:
            hwnds = moom_get_hwnds(pid)
            if len(hwnds) > 0:
                win32gui.SetForegroundWindow(hwnds[0])
    except:
        pass

def moom_launch_application(app, dup):
    def moom_process_exist(app):
        # Trust app and apps[app]
        # call = 'TASKLIST', '/FO', 'csv'
        call = 'TASKLIST', '/FO', 'csv', '/FI', 'imagename eq %s*' % app
        proc = subprocess.Popen(call, shell=True, stdout=subprocess.PIPE,
                                universal_newlines=True)
        list = []
        # output = proc.communicate()[0].strip().split('\r\n')
        for p in csv.DictReader(proc.stdout):
            if app in p['Image Name']:
                list.append(p)
        if len(list) > 0:
            return int(list[0]['PID'])
        return None

    if not app in apps:
        print(app+" is not listed")
        return
    if not os.path.exists(apps[app][0]):
        print(apps[app]+" does not exist")
        return

    if not dup:
        p = moom_process_exist(app)
        if p is not None:
            moom_focus(p)
            return

    subprocess.Popen(apps[app][0]+apps[app][1])
    pid = moom_process_exist(app)
    moom_focus(pid)
    print("Starting "+app+".exe ("+str(pid)+")")

# API to move the frame
def moom_move_frame():
    hwnd = win32gui.GetForegroundWindow()
    if moom_exclude_windows(hwnd):
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win32gui.MoveWindow(hwnd,
                        win32offset[0],
                        win32offset[1],
                        right-left,
                        bottom-top,
                        True)
    print("Move to Top-Left")

def moom_move_frame_left():
    hwnd=win32gui.GetForegroundWindow()
    if moom_exclude_windows(hwnd):
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win32gui.MoveWindow(hwnd,
                        left-shift_amount[0],
                        top+win32offset[1],
                        right-left,
                        bottom-top,
                        True)
    print("Move to Left")

def moom_move_frame_right():
    hwnd=win32gui.GetForegroundWindow()
    if moom_exclude_windows(hwnd):
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win32gui.MoveWindow(hwnd,
                        left+shift_amount[1],
                        top+win32offset[1],
                        right-left,
                        bottom-top,
                        True)
    print("Move to Reft")

def moom_move_frame_center():
    hwnd=win32gui.GetForegroundWindow()
    if moom_exclude_windows(hwnd):
        return
    fleft, ftop, right, bottom = win32gui.GetWindowRect(hwnd)
    fwidth=right-fleft
    fheight=bottom-ftop
    screen=win32gui.GetWindowRect(win32gui.GetDesktopWindow())
    win32gui.MoveWindow(hwnd,
                        math.floor((screen[2]-screen[0])/2.0)-
                        math.floor(fwidth/2.0)+win32offset[0],
                        math.floor((screen[3]-screen[1])/2.0)-
                        math.floor(fheight/2.0)+win32offset[1],
                        fwidth,
                        fheight,
                        True)
    print("Move to Center")

def moom_move_frame_to_edge_top():
    hwnd=win32gui.GetForegroundWindow()
    if moom_exclude_windows(hwnd):
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win32gui.MoveWindow(hwnd,
                        left,
                        win32offset[1],
                        right-left,
                        bottom-top,
                        True)
    print("Move to Top-Edge")

def moom_move_frame_to_edge_bottom():
    hwnd=win32gui.GetForegroundWindow()
    if moom_exclude_windows(hwnd):
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    screen=win32gui.GetWindowRect(win32gui.GetDesktopWindow())
    win32gui.MoveWindow(hwnd,
                        left,
                        screen[3]-(bottom-top)+win32offset[1]-menu_bar[1],
                        right-left,
                        bottom-top,
                        True)
    print("Move to Bottom-Edge")

# For windows event
def moom_update_keybinding():
    if moom_on_emacs():
        moom_disable_emacs_keybinding()
    else:
        moom_enable_emacs_keybinding()


# Emacs keybinding
def moom_enable_emacs_keybinding():
    try:
        print("--- Enable emacs keybinding")
        keyboard.remap_hotkey('ctrl+a', 'home')
        keyboard.remap_hotkey('ctrl+e', 'end')
        keyboard.remap_hotkey('ctrl+f', 'right')
        keyboard.remap_hotkey('ctrl+b', 'left')
        keyboard.remap_hotkey('ctrl+d', 'delete')
        #keyboard.remap_hotkey('ctrl+y', 'ctrl+v')
        #keyboard.remap_hotkey('alt+v', 'ctrl+v')
        #keyboard.remap_hotkey('alt+f', 'ctrl+f')
        #keyboard.remap_hotkey('alt+c', 'ctrl+c')
    except:
        pass

def moom_disable_emacs_keybinding():
    try:
        print("--- Disable emacs keybinding")
        keyboard.remove_hotkey('ctrl+a')
        keyboard.remove_hotkey('ctrl+e')
        keyboard.remove_hotkey('ctrl+f')
        keyboard.remove_hotkey('ctrl+b')
        keyboard.remove_hotkey('ctrl+d')
        #keyboard.remove_hotkey('ctrl+y')
        #keyboard.remove_hotkey('alt+v')
        #keyboard.remove_hotkey('alt+f')
        #keyboard.remove_hotkey('alt+c')
    except:
        pass

# Registration
try:
    keyboard.add_hotkey('alt+0',
                        moom_move_frame, suppress=True)
    keyboard.add_hotkey('alt+1',
                        moom_move_frame_left, suppress=True)
    keyboard.add_hotkey('alt+2',
                        moom_move_frame_center, suppress=True)
    keyboard.add_hotkey('alt+3',
                        moom_move_frame_right, suppress=True)
    keyboard.add_hotkey('f1',
                        moom_move_frame_to_edge_top, suppress=True)
    keyboard.add_hotkey('shift+f1',
                        moom_move_frame_to_edge_bottom, suppress=True)
    # Launcher (see also 'apps' above)
    keyboard.add_hotkey('ctrl+alt+e',
                        moom_launch_application, args=['emacs',False],
                        trigger_on_release=True, suppress=True)
    keyboard.add_hotkey('ctrl+alt+i',
                        moom_launch_application, args=['mintty',False],
                        trigger_on_release=True, suppress=True)
    keyboard.add_hotkey('ctrl+alt+enter',
                        moom_launch_application, args=['explorer',True],
                        trigger_on_release=True, suppress=True)
    keyboard.add_hotkey('windows+8',
                        moom_enable_emacs_keybinding, suppress=True)
    keyboard.add_hotkey('windows+0',
                        moom_disable_emacs_keybinding, suppress=True)
except:
    pass

# Starting main loop
print("--- pymoom-w32: "+pymoom_w32)
print("--- Keyboard:   "+keyboard.version)
moom_enable_emacs_keybinding()
print("--- Type C-<tab>-9 to exit")

keyboard.wait('ctrl+tab+9')
