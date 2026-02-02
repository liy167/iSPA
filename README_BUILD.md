# SASEG Autoexec - 打包说明

## 打包为EXE文件

### 方法一：使用批处理脚本（推荐）

1. **双击运行 `build_exe.bat`**
   - 脚本会自动安装依赖并打包
   - 打包完成后，exe文件位于 `dist` 文件夹中

2. **或者使用高级打包脚本 `build_exe_advanced.bat`**
   - 包含更多优化选项
   - 生成的文件可能更小

### 方法二：手动打包

1. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

2. **使用PyInstaller打包**
   ```bash
   pyinstaller --name="SASEG_Autoexec" --onefile --windowed --noconsole SASEG_GUI.py
   ```

3. **打包完成后**
   - exe文件位于 `dist` 文件夹中
   - 文件名为 `SASEG_Autoexec.exe`

## 依赖说明

- **pywinauto**: 用于自动化Windows应用程序（SEGuide.exe）
- **tkinter**: Python内置GUI库（无需安装）
- **pyinstaller**: 用于打包Python程序为exe

## 注意事项

1. **首次打包可能需要较长时间**（下载依赖和编译）
2. **生成的exe文件可能较大**（包含Python解释器和所有依赖）
3. **如果遇到问题**，可以尝试：
   - 使用 `--clean` 选项清理缓存
   - 检查是否有杀毒软件拦截
   - 确保所有依赖都已正确安装

## 分发说明

打包完成后，只需要分发 `dist/SASEG_Autoexec.exe` 文件即可。

**使用要求：**
- Windows操作系统
- 不需要安装Python
- 不需要安装其他依赖
