import pyautogui
import tkinter as tk
import time
import random
import keyboard
import ctypes
import win32gui
import win32con
import sys
# app = QtWidgets.QApplication(sys.argv)

def pause_program():
    # QApplication.processEvents()
    print("Program paused")
    keyboard.wait("ctrl+alt+p")  # Wait for the keys to be released
    print("Program resumed")

def start_click():
    title = root.title()
    hwnd = ctypes.windll.user32.FindWindowW(None, title)
    if hwnd != 0:
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    else:
        pass
    print(title)
    print(hwnd)
    text = text_entry.get()
    sleep_time = int(text)
    end_time = time.time() + sleep_time * 60
    screen_width, screen_height = pyautogui.size()
    while (flag == True) and time.time() < end_time:
        x = random.randint(800, 1000)
        y = random.randint(400, 600)
        pyautogui.moveTo(x, y)
        pyautogui.click()
        time.sleep(random.uniform(1.5, 2.5))
        if keyboard.is_pressed("ctrl+alt+p"):
            print("pressed")
            
def stop_click():
    flag = not(flag)
        
    
    # label.config(text="You entered: " + text)
flag = True
root = tk.Tk()

    

label = tk.Label(root, text="Enter testing time(minute):")
label.pack()

text_entry = tk.Entry(root)
text_entry.pack()

button = tk.Button(root, text="Start", command=start_click)
button.pack()

button2 = tk.Button(root, text="Stop", command=stop_click)
button2.pack()

# hwnd = ctypes.windll.user32.GetForegroundWindow()
# if flag:
#     hwnd = win32gui.GetForegroundWindow()
#     window_title = win32gui.GetWindowText(hwnd)
#     print(window_title)
root.mainloop()