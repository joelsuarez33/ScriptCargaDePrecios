import pyautogui
import pandas as pd
from pandas import DataFrame
import os
import datetime
import win32com.client
from win32com.client import Dispatch
import time
import subprocess
from dateutil.relativedelta import relativedelta
import ctypes
from contextlib import suppress
import tkinter.messagebox

# --------------------delete functions

def delete_all(path):
    files = [f for f in os.listdir(path)]
    for f in files:
        file = str(path + f)
        if os.path.exists(file) == True:
            os.remove(file)

def delete_pattern(path, pattern):
    files = [f for f in os.listdir(path)]
    files = DataFrame(files, columns=["File"])
    files = files[files['File'].str.contains(pattern)]
    print(files)
    for f in range(0, len(files)):
        z = files['File'].iloc[f]
        file = str(path + z)
        if os.path.exists(file) == True:
            os.remove(file)

def delete_unique(path):
    if os.path.exists(path) == True:
        os.remove(path)

def create_file(path):
    with open(path, "w") as file:
        file.close()

def wait_file(path):
    while not os.path.exists(path):
        time.sleep(1)

CF_TEXT = 1
kernel32 = ctypes.windll.kernel32
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
user32 = ctypes.windll.user32
user32.GetClipboardData.restype = ctypes.c_void_p

def get_clipboard_text():
    user32.OpenClipboard(0)
    try:
        if user32.IsClipboardFormatAvailable(CF_TEXT):
            data = user32.GetClipboardData(CF_TEXT)
            data_locked = kernel32.GlobalLock(data)
            text = ctypes.c_char_p(data_locked)
            value = text.value
            kernel32.GlobalUnlock(data_locked)
            return value
    finally:
        user32.CloseClipboard()

def verify(session=None, control=None):
    while True:
        try:
            session.FindById(control)
            return session.FindById(control)
        except:
            return

# Sample date for execution (change as needed)
start = datetime.datetime(2020, 2, 4, 00, 40, 1)
while datetime.datetime.now() < start:
    time.sleep(1)

# Replace with the path to the application you want to open
path = r"C:\path\to\your\SAP\executable"
subprocess.Popen(path)
time.sleep(3)

# Initialize application shell (SAP GUI simulation)
shell = win32com.client.Dispatch("WScript.Shell")
time.sleep(1)

class cls_SAP_Gui_Scripting:
    def __init__(self, api, conn):
        self.SAPguiAPP = win32com.client.GetObject(api).GetScriptingEngine
        self.Connection = self.SAPguiAPP.OpenConnection(conn, 1)
        self.Session = self.Connection.Children(0)

# Modify the connection name here (do not disclose actual names)
MySAPGUI = cls_SAP_Gui_Scripting("SAPGUI", "INSERT_CONNECTION_NAME_HERE")

# ------------- Executing Process Flow
MySAPGUI.Session.findById("wnd[0]/tbar[0]/okcd").text = "/nTransactionCode"  # Replace with actual transaction code
MySAPGUI.Session.findById("wnd[0]").sendVKey(0)

# Enter data in form fields
MySAPGUI.Session.findById("wnd[0]/usr/inputField1").text = "Value1"
MySAPGUI.Session.findById("wnd[0]/usr/inputField2").text = "Value2"
MySAPGUI.Session.findById("wnd[0]/usr/inputField3").text = "Value3"
MySAPGUI.Session.findById("wnd[0]").sendVKey(0)

# Confirm and proceed with the action
MySAPGUI.Session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Loop through grouped data to fill form
Client_Group = input_table_1.groupby('Client Identifier')

for label, group in Client_Group:
    # Fill form with group-specific data
    MySAPGUI.Session.findById("wnd[0]/usr/inputFieldForClient").text = str(label)
    i = 0
    for label_2, row in group.iterrows():
        MySAPGUI.Session.findById(f"wnd[0]/usr/tableControl/field1[{i}]").text = str(row['Field1'])
        MySAPGUI.Session.findById(f"wnd[0]/usr/tableControl/field2[{i}]").text = str(row['Field2'])
        MySAPGUI.Session.findById(f"wnd[0]/usr/tableControl/field3[{i}]").text = str(row['Field3'])
        MySAPGUI.Session.findById(f"wnd[0]/usr/tableControl/field4[{i}]").text = str(row['Field4'])
        i += 1

    # Save the filled form
    MySAPGUI.Session.findById("wnd[0]/tbar[0]/btnSave").press()

    # Handle any popup messages or warnings
    Flag_Stop = False
    while not Flag_Stop:
        try:
            if MySAPGUI.Session.findById("wnd[1]").text == "Warning Message":
                MySAPGUI.Session.findById("wnd[1]/tbar[0]/btnClose").press()
        except:
            Flag_Stop = True

# -------- Close Application
MySAPGUI.Session.findById("wnd[0]").maximize()
MySAPGUI.Session.findById("wnd[0]").close()
MySAPGUI.Session.findById("wnd[1]/usr/btnConfirm").press()

# Release objects
MySAPGUI = None
Connection = None
shell = None
SAPGuiAPP = None
