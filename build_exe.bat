@echo off
chcp 65001 >nul
echo ========================================
echo SASEG Autoexec - 打包脚本
echo ========================================
echo.

echo [1/3] 正在检查并安装依赖包...
pip install -r requirements.txt
if errorlevel 1 (
    echo 依赖安装失败，请检查网络连接或pip配置
    pause
    exit /b 1
)

echo.
echo [2/3] 正在清理之前的打包文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist SASEG_Autoexec.spec del /q SASEG_Autoexec.spec

echo.
echo [3/3] 正在打包为exe文件...
if not exist logo.png (
    echo 警告: 未找到logo.png文件，将使用默认图标
    python -m PyInstaller --name="SASEG_Autoexec" ^
        --onefile ^
        --windowed ^
        --noconsole ^
        --clean ^
        --hidden-import=pywinauto ^
        --hidden-import=pywinauto.application ^
        --hidden-import=pywinauto.keyboard ^
        --hidden-import=comtypes ^
        --hidden-import=comtypes.client ^
        --collect-all=pywinauto ^
        SASEG_GUI.py
) else (
    echo 使用logo.png作为exe图标
    python -m PyInstaller --name="SASEG_Autoexec" ^
        --onefile ^
        --windowed ^
        --noconsole ^
        --clean ^
        --icon=logo.png ^
        --hidden-import=pywinauto ^
        --hidden-import=pywinauto.application ^
        --hidden-import=pywinauto.keyboard ^
        --hidden-import=comtypes ^
        --hidden-import=comtypes.client ^
        --collect-all=pywinauto ^
        SASEG_GUI.py
)

if errorlevel 1 (
    echo.
    echo 打包失败！请检查错误信息
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo exe文件位置: dist\SASEG_Autoexec.exe
echo.
echo 您现在可以将 dist\SASEG_Autoexec.exe 分享给团队成员使用
echo.
pause
