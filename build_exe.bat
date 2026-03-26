@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem --------- 资源准备：从 Cursor assets 自动补齐 ---------
rem assets 实际文件名由 Cursor 生成，这里做映射成打包所需的同名资源。
set "ASSET_DIR=%CD%\..\..\.cursor\projects\c-Users-liy167-YuLI-saseg\assets"

if not exist "pdt_manager_icon.png" (
    if exist "%ASSET_DIR%\c__Users_liy167_YuLI_saseg_pdt_manager_icon.png" (
        echo 正在从 assets 拷贝 pdt_manager_icon.png ...
        copy /y "%ASSET_DIR%\c__Users_liy167_YuLI_saseg_pdt_manager_icon.png" "pdt_manager_icon.png" >nul
    )
)

if not exist "logo.png" (
    if exist "%ASSET_DIR%\c__Users_liy167_YuLI_saseg_logo.png" (
        echo 正在从 assets 拷贝 logo.png ...
        copy /y "%ASSET_DIR%\c__Users_liy167_YuLI_saseg_logo.png" "logo.png" >nul
    )
)

rem --------- 将 logo.png 转成 logo.ico 以确保 --icon 生效 ---------
set "ICON_ARG="

if not exist "pdt_manager_icon.png" (
    echo 错误: 缺少 pdt_manager_icon.png（即使尝试从 assets 拷贝也失败）。
    pause
    exit /b 1
)

set "PYEXE=%CD%\.venv\Scripts\python.exe"
if not exist "%PYEXE%" (
    echo 未找到虚拟环境 .venv，正在创建...
    python -m venv .venv
    if errorlevel 1 (
        echo 创建失败：请确认已安装 Python 并已加入 PATH
        pause
        exit /b 1
    )
    set "PYEXE=%CD%\.venv\Scripts\python.exe"
)

if exist "logo.png" (
    "%PYEXE%" -c "from PIL import Image; img=Image.open('logo.png').convert('RGBA'); img.save('logo.ico', format='ICO', sizes=[(16,16),(32,32),(48,48),(64,64)])" >nul 2>&1
    if exist "logo.ico" set "ICON_ARG=--icon=logo.ico"
)

echo ========================================
echo SASEG Autoexec - 打包脚本
echo 使用 Python: %PYEXE%
echo ========================================
echo.

echo [1/3] 正在检查并安装依赖包...
"%PYEXE%" -m pip install --upgrade pip
if errorlevel 1 (
    echo pip 升级失败
    pause
    exit /b 1
)
"%PYEXE%" -m pip install -r requirements.txt
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
if defined ICON_ARG (
    echo 使用 %ICON_ARG% 作为exe图标
) else (
    echo 警告: 未生成 logo.ico，将使用默认图标
)

"%PYEXE%" -m PyInstaller --name="SASEG_Autoexec" ^
    --onefile ^
    --windowed ^
    --noconsole ^
    --clean ^
    --noconfirm ^
    %ICON_ARG% ^
    --add-data "pdt_manager_icon.png;." ^
    --exclude-module=torch ^
    --exclude-module=torchvision ^
    --exclude-module=torchaudio ^
    --exclude-module=tensorflow ^
    --exclude-module=transformers ^
    --exclude-module=sklearn ^
    --hidden-import=pywinauto ^
    --hidden-import=pywinauto.application ^
    --hidden-import=pywinauto.keyboard ^
    --hidden-import=comtypes ^
    --hidden-import=comtypes.client ^
    SASEG_GUI.py

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
