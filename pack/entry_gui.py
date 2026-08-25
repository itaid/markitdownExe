"""PyInstaller 独立入口：启动 GUI（用于打 EXE，避免打包整个包）"""
from doc2md_tool.gui import run_gui

if __name__ == "__main__":
    run_gui()
