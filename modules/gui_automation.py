# modules/gui_automation.py
import pyautogui
import platform
import time

pyautogui.FAILSAFE = True
pyautogui.PAUSE    = 0.3

IS_MAC = platform.system() == "Darwin"


def press_hotkey(*keys):
    pyautogui.hotkey(*keys)


def type_text(text):
    pyautogui.write(text)


def click_at(x, y):
    pyautogui.click(x, y)


def scroll(amount):
    pyautogui.scroll(amount)


def open_file_dialog():
    press_hotkey("command", "o") if IS_MAC else press_hotkey("ctrl", "o")
    time.sleep(1)


def save_file_dialog():
    press_hotkey("command", "s") if IS_MAC else press_hotkey("ctrl", "s")
    time.sleep(1)


def print_dialog():
    press_hotkey("command", "p") if IS_MAC else press_hotkey("ctrl", "p")
    time.sleep(1)


def pdf_open_and_print(pdf_path):
    open_file_dialog()
    type_text(pdf_path)
    pyautogui.press("enter")
    time.sleep(3)
    print_dialog()
    time.sleep(1)
    pyautogui.press("enter")


def close_active_window():
    press_hotkey("command", "q") if IS_MAC else press_hotkey("alt", "f4")
