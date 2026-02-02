# iSPA / SASEG Autoexec

图形界面工具，用于快速选择 Z 盘路径并启动 SAS Enterprise Guide。

## 文档

- [使用说明](使用说明.md) — 功能说明与使用方法  
- [打包说明](README_BUILD.md) — 如何打包为 EXE

## 快速开始

1. 安装依赖：`pip install -r requirements.txt`
2. 运行：`python SASEG_GUI.py`
3. 或使用 `build_exe.bat` 打包为 exe 后运行 `dist/SASEG_Autoexec.exe`

## 主要文件

- `SASEG_GUI.py` — 主程序  
- `requirements.txt` — Python 依赖  
- `build_exe.bat` / `build_exe_advanced.bat` — 打包脚本  
- `logo.png` — 程序图标
