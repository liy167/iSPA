@echo off
echo 正在安装依赖包...
pip install -r requirements.txt

echo.
echo 正在打包为exe文件（包含图标和优化选项）...
if not exist logo.png (
    echo 警告: 未找到logo.png文件，将使用默认图标
    python -m PyInstaller --name="SASEG_Autoexec" ^
        --onefile ^
        --windowed ^
        --noconsole ^
        --clean ^
        --noupx ^
        --hidden-import=pywinauto ^
        --hidden-import=pywinauto.application ^
        --hidden-import=pywinauto.keyboard ^
        --hidden-import=comtypes ^
        --hidden-import=comtypes.client ^
        --collect-all=pywinauto ^
        --collect-all=comtypes ^
        SASEG_GUI.py
) else (
    echo 使用logo.png作为exe图标
    python -m PyInstaller --name="SASEG_Autoexec" ^
        --onefile ^
        --windowed ^
        --noconsole ^
        --clean ^
        --noupx ^
        --icon=logo.png ^
        --hidden-import=pywinauto ^
        --hidden-import=pywinauto.application ^
        --hidden-import=pywinauto.keyboard ^
        --hidden-import=comtypes ^
        --hidden-import=comtypes.client ^
        --collect-all=pywinauto ^
        --collect-all=comtypes ^
        SASEG_GUI.py
)

echo.
echo 打包完成！exe文件位于 dist 文件夹中
echo 文件路径: dist\SASEG_Autoexec.exe
pause
