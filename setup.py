# setup.py
import sys
from cx_Freeze import setup, Executable

# 包含 CapsWriter-Offline 目录
include_files = [(".\\CapsWriter-Offline-Windows-64bit", "CapsWriter-Offline-Windows-64bit")]

build_options = {
    "includes": ["tkinter"],  # 明确包含 tkinter
    "include_files": include_files,
    "excludes": [],
    "optimize": 2
}

base = "Win32GUI"  # 隐藏控制台窗口

setup(
    name="CapsWriter Launcher",
    version="1.0",
    description="启动器 for CapsWriter-Offline",
    options={"build_exe": build_options},
    executables=[Executable("caps_writer_launcher.py", base=base)]
)
