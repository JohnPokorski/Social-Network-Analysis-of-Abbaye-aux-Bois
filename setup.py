import sys
import os
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"]}

# GUI applications require a different base on Windows (the default is for a
# console application).

base = "Win32GUI"

os.environ['TCL_LIBRARY'] = "C:\\Users\\John\\AppData\\Local\\Programs\\Python\\Python35\\tcl\\tcl8.6"
os.environ['TK_LIBRARY'] = "C:\\Users\\John\\AppData\\Local\\Programs\\Python\\Python35\\tcl\\tk8.6"

setup(  name = "NetworkAnalyzer",
        version = "0.1",
        description = "Convert XLS into Gephi-Readable networks",
        options = {"build_exe": build_exe_options},
        executables = [Executable("Analysis.py", base=base)])