# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import cx_Freeze
import os
import re
import json
import random
from tkinter import *
import keyboard
import tkinter as tk
import time
from tkinter import ttk
from tkinter import filedialog
from functools import partial
import random
from tkinter import *
import win32com.client

PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')
base = None

if sys.platform == 'win32':
    base = "Win32GUI"

executables = [cx_Freeze.Executable("PyDictionary.py", base=base, icon='icons\\main_icon.ico')]

cx_Freeze.setup(
    name = "PyDictionary",
    options = {"build_exe":{"packages":["tkinter","json","win32com.client","functools","random","keyboard","re","time"]
              ,"include_files":[(os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll')),(os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'))]}},
    version = "1.55",
    description = "PyDictionary a Dictionary",
    executables = executables
    )
















