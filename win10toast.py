from __future__ import absolute_import
from __future__ import print_function
from __future__ import unicode_literals
__all__ = ['ToastNotifier']
import logging
import threading
from os import path
from time import sleep
from pkg_resources import Requirement
from pkg_resources import resource_filename
from win32api import GetModuleHandle
from win32api import PostQuitMessage
from win32con import CW_USEDEFAULT
from win32con import IDI_APPLICATION
from win32con import IMAGE_ICON
from win32con import LR_DEFAULTSIZE
from win32con import LR_LOADFROMFILE
from win32con import WM_DESTROY
from win32con import WM_USER
from win32con import WS_OVERLAPPED
from win32con import WS_SYSMENU
from win32gui import CreateWindow
from win32gui import DestroyWindow
from win32gui import LoadIcon
from win32gui import LoadImage
from win32gui import NIF_ICON
from win32gui import NIF_INFO
from win32gui import NIF_MESSAGE
from win32gui import NIF_TIP
from win32gui import NIM_ADD
from win32gui import NIM_DELETE
from win32gui import NIM_MODIFY
from win32gui import RegisterClass
from win32gui import UnregisterClass
from win32gui import Shell_NotifyIcon
from win32gui import UpdateWindow
from win32gui import WNDCLASS

class ToastNotifier(object):
    def __init__(self):
        self._thread = None

    def _show_toast(self, title, msg, icon_path, duration):
        message_map = {WM_DESTROY: self.on_destroy, }
        self.wc = WNDCLASS()
        self.hinst = self.wc.hInstance = GetModuleHandle(None)
        self.wc.lpszClassName = str("PythonTaskbar")
        self.wc.lpfnWndProc = message_map
        try:
            self.classAtom = RegisterClass(self.wc)
        except:
            pass
        style = WS_OVERLAPPED | WS_SYSMENU
        self.hwnd = CreateWindow(self.classAtom, "Taskbar", style, 0, 0, CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, self.hinst, None)
        UpdateWindow(self.hwnd)

        if icon_path is not None:
            icon_path = path.realpath(icon_path)
        else:
            icon_path =  resource_filename("icon.ico")
        icon_flags = LR_LOADFROMFILE | LR_DEFAULTSIZE
        try:
            hicon = LoadImage(self.hinst, icon_path, IMAGE_ICON, 0, 0, icon_flags)
        except Exception as e:
            logging.error("Some trouble with the icon ({}): {}".format(icon_path, e))
            hicon = LoadIcon(0, IDI_APPLICATION)

        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, WM_USER + 20, hicon, "Tooltip")
        Shell_NotifyIcon(NIM_ADD, nid)
        Shell_NotifyIcon(NIM_MODIFY, (self.hwnd, 0, NIF_INFO, WM_USER + 20, hicon, "Balloon Tooltip", msg, 200, title))
        sleep(duration)
        DestroyWindow(self.hwnd)
        UnregisterClass(self.wc.lpszClassName, None)
        return None

    def show_toast(self, title="Notification", msg="Here comes the message", icon_path=None, duration=5, threaded=False):
        if not threaded:
            self._show_toast(title, msg, icon_path, duration)
        else:
            if self.notification_active():
                return False

            self._thread = threading.Thread(target=self._show_toast, args=(title, msg, icon_path, duration))
            self._thread.start()
        return True

    def notification_active(self):
        if self._thread != None and self._thread.is_alive():
            return True
        return False

    def on_destroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0)
        return None