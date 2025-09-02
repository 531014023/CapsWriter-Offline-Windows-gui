# setup.py
import sys,os
from cx_Freeze import setup, Executable

# 包含 CapsWriter-Offline 目录
# 定义要包含的文件和目录，但排除 models 目录
def get_include_files():
    base_dir = ".\\CapsWriter-Offline-Windows-64bit"
    include_list = []
    
    # 遍历基础目录，排除 models 子目录
    for item in os.listdir(base_dir):
        item_path = os.path.join(base_dir, item)
        # 跳过 models 目录
        if item == "models":
            continue
        # 如果是文件，直接添加
        if os.path.isfile(item_path):
            include_list.append((item_path, os.path.join("CapsWriter-Offline-Windows-64bit", item)))
        # 如果是目录（且不是models），递归添加其内容
        elif os.path.isdir(item_path):
            for root, dirs, files in os.walk(item_path):
                for file in files:
                    src_path = os.path.join(root, file)
                    # 计算目标路径
                    rel_path = os.path.relpath(src_path, base_dir)
                    dest_path = os.path.join("CapsWriter-Offline-Windows-64bit", rel_path)
                    include_list.append((src_path, dest_path))
    
    return include_list

include_files = get_include_files()

# 定义要排除的目录或文件
# 常见的排除目录列表
COMMON_EXCLUDE_DIRS = [
    # 开发环境相关
    "__pycache__",
    ".git",
    ".vscode",
    ".idea",
    ".pytest_cache",
    ".mypy_cache",
    
    # Python虚拟环境
    "venv",
    "env",
    "virtualenv",
    
    # 构建产物
    "dist",
    "build",
    "*.egg-info",
    
    # 测试相关
    "test",
    "tests",
    "testing",
    
    # 临时文件
    "temp",
    "tmp",
    "*.log",
    "*.tmp",
    
    # 文档
    "docs",
    "doc",
    
    # 备份文件
    "backup",
    "*.bak"
]

build_options = {
    "includes": ["tkinter"],  # 明确包含 tkinter
    "include_files": include_files,
    "excludes": COMMON_EXCLUDE_DIRS,
    "optimize": 2
}

base = "Win32GUI"  # 隐藏控制台窗口

setup(
    name="CapsWriter Launcher",
    version="1.0",
    description="启动器 for CapsWriter-Offline",
    options={"build_exe": build_options},
    executables=[Executable("caps_writer_launcher_pyqt.py", base=base)]
)
