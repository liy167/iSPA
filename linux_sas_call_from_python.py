import argparse
import os
import re
import subprocess
import sys
import saspy
import pandas as pd

# 是否在 Linux 下运行（无 GUI，直接读 Linux 路径）
IS_LINUX = sys.platform.startswith('linux')

if not IS_LINUX:
    import tkinter as tk
    from tkinter import scrolledtext, font as tkfont, messagebox

# Linux 路径与 Windows 盘符映射：.../DocumentRepository/DDT/ 对应 Z:/
LINUX_PATH_PREFIX = '/u01/app/sas/sas9.4/DocumentRepository/DDT/'
WINDOWS_PATH_PREFIX = 'Z:\\'  # 即 Z:/


# 修改路径为 Windows 格式（用于在 Windows 上读取 Linux 侧生成的日志）
def convert_linux_path_to_windows(linux_path):
    windows_path = linux_path.replace('/u01/app/sas/sas9.4/DocumentRepository/DDT', 'Z:')
    windows_path = windows_path.replace('/', '\\')  # 将斜杠替换为反斜杠
    return windows_path


# Windows 路径转为 Linux 路径（提交给 SAS 的必须是 Linux 路径，否则服务器写不到 Z:）
def convert_windows_path_to_linux(windows_path):
    p = windows_path.replace('\\', '/')
    if p.upper().startswith('Z:/'):
        return LINUX_PATH_PREFIX + p[3:].lstrip('/')
    return p

# 审核：ERROR:: 或 ERROR:；WARNING:: 或 WARNING:
ERROR_PATTERN = re.compile(r'ERROR\s*::|ERROR\s*:', re.IGNORECASE)
WARNING_PATTERN = re.compile(r'WARNING\s*::|WARNING\s*:', re.IGNORECASE)


def _print_log_review_console(lines_with_kind, log_path):
    """在 Linux 下将 ERROR/WARNING 行打印到控制台。"""
    print(f"\n--- 日志审阅: {log_path} ---")
    for line_text, kind in lines_with_kind:
        prefix = "ERROR  " if kind == 'error' else "WARNING"
        print(f"  [{prefix}] {line_text.rstrip()}")
    print("---\n")


def _show_log_review_popup(lines_with_kind, windows_log_path, on_window_close=None):
    """弹出审核窗口，ERROR 行红色，WARNING 行绿色；底部提供是否打开日志文件。关闭窗口时调用 on_window_close 以断开 SAS 会话。（仅 Windows）"""
    root = tk.Tk()
    root.title('日志审核 - ERROR / WARNING')
    line_count = len(lines_with_kind)
    line_height = 22
    btn_area = 70
    content_height = line_count * line_height + btn_area
    win_height = min(max(content_height, 180), 520)
    root.geometry(f'800x{win_height}')
    root.minsize(400, 180)

    def open_log():
        if not os.path.isfile(windows_log_path):
            return
        try:
            subprocess.Popen(['notepad', windows_log_path], shell=True)
        except Exception:
            os.startfile(windows_log_path)
        # 点击“是”只打开日志文件，弹窗不关闭

    def do_close():
        """关闭窗口时先断开 SAS，再销毁窗口。"""
        if on_window_close:
            try:
                on_window_close()
            except Exception:
                pass
        root.destroy()

    def do_not_open():
        do_close()

    root.protocol('WM_DELETE_WINDOW', do_close)

    # 先 pack 底部按钮区，保证始终可见
    btn_frame = tk.Frame(root)
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=8)
    tk.Label(btn_frame, text='是否打开日志文件并审阅？', font=('Microsoft YaHei UI', 10)).pack(side=tk.LEFT, padx=(0, 12))
    tk.Button(btn_frame, text='是', width=8, command=open_log).pack(side=tk.LEFT, padx=4)
    tk.Button(btn_frame, text='否', width=8, command=do_not_open).pack(side=tk.LEFT, padx=4)

    text = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=tkfont.Font(family='Consolas', size=10))
    text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    text.tag_configure('error', foreground='#CC0000')
    text.tag_configure('warning', foreground='#008000')
    for line_text, kind in lines_with_kind:
        text.insert(tk.END, line_text, kind)
    text.config(state=tk.DISABLED)
    root.mainloop()


def _review_log_content(log_content: str, source_label: str, on_window_close=None) -> bool:
    """解析日志内容并审阅 ERROR/WARNING，返回是否存在错误或警告。on_window_close：关闭审阅窗口时调用的回调（用于断开 SAS）。"""
    if not (log_content or "").strip():
        return False
    print(log_content)
    has_error = bool(ERROR_PATTERN.search(log_content))
    has_warning = bool(WARNING_PATTERN.search(log_content))
    if has_error or has_warning:
        lines = log_content.splitlines(keepends=True)
        highlight = []
        for line in lines:
            if ERROR_PATTERN.search(line):
                highlight.append((line, 'error'))
            elif WARNING_PATTERN.search(line):
                highlight.append((line, 'warning'))
        if IS_LINUX:
            _print_log_review_console(highlight, source_label)
        else:
            _show_log_review_popup(highlight, source_label, on_window_close=on_window_close)
    return has_error or has_warning


def check_for_errors_in_log(log_file_path, fallback_log_content=None, on_window_close=None):
    """优先从日志文件读取；若文件不存在且提供了 fallback_log_content（如 saspy 返回的 LOG），则用其审阅。on_window_close：关闭审阅窗口时调用的回调（用于断开 SAS）。"""
    actual_log_path = log_file_path if IS_LINUX else convert_linux_path_to_windows(log_file_path)
    try:
        with open(actual_log_path, 'r', encoding='utf-8', errors='replace') as log_file:
            log_content = log_file.read()
        return _review_log_content(log_content, actual_log_path, on_window_close=on_window_close)
    except FileNotFoundError:
        if fallback_log_content and str(fallback_log_content).strip():
            print(f"日志文件 {actual_log_path} 不可用，改用 SAS 会话返回的 LOG 审阅。\n")
            return _review_log_content(str(fallback_log_content), "(来自 SAS 会话)", on_window_close=on_window_close)
        print(f"日志文件 {actual_log_path} 不存在！")
        return False
    except UnicodeDecodeError as e:
        print(f"读取日志文件时出现编码错误: {e}")
        return False


def run_sas(sas_file_path: str, sas_session=None, check_log=True) -> bool:
    """根据给定的 sas_file_path 在 Linux SAS 上执行并可选择审核日志。
    sas_session: 可选，若传入则复用该会话（用于连续执行多个 SAS 文件）；否则本函数内创建并在结束时关闭。
    check_log: 是否进行日志审阅（ERROR/WARNING）；提交多条 SAS 程序时可设为 False 以跳过。
    返回: 是否有错误或警告（未审阅时返回 False）。
    支持传入 Windows 路径（Z:\\...）或 Linux 路径（/u01/...）；提交给 SAS 时统一转为 Linux 路径，日志才能写到服务器并可通过 Z: 读取。
    """
    macro_file_path = '/u01/app/sas/sas9.4/DocumentRepository/DDT/projects/utility/macros/01_general/autorun.sas'

    # 提交给 SAS 的必须为 Linux 路径，否则日志写不到 Z: 对应目录
    sas_file_path_linux = convert_windows_path_to_linux(sas_file_path)
    sas_file_name_no_ext = os.path.splitext(os.path.basename(sas_file_path_linux))[0]
    base_path = os.path.dirname(sas_file_path_linux)

    if '06_programs' in sas_file_path_linux or '09_validation' in sas_file_path_linux:
        log_output_path = f"{base_path}/07_logs/{sas_file_name_no_ext}.log"
    else:
        log_output_path = f"{base_path}/{sas_file_name_no_ext}.log"

    #print(f"日志保存于 {log_output_path} 。\n")

    sas_code = f"""
proc printto log='{log_output_path}' new;
run;

%let _sasprogramfile = '{sas_file_path_linux}';
%include '{macro_file_path}';
%include '{sas_file_path_linux}';

proc printto; /* 恢复日志输出到默认位置 */
run;
"""

    own_session = sas_session is None
    if own_session:
        sas = saspy.SASsession(cfgname='winiomlinux')
    else:
        sas = sas_session

    session_ended = [False]  # 用列表以便在闭包中修改

    def on_log_window_close():
        """关闭日志审阅窗口时断开 SAS 会话（仅本函数创建的会话）。"""
        if not session_ended[0] and own_session:
            session_ended[0] = True
            try:
                sas.endsas()
            except Exception:
                pass

    try:
        sas_output = sas.submit(sas_code)
        if not check_log:
            print(f"SAS程序 {sas_file_path} 已提交执行。")
            return False
        log_from_sas = sas_output.get('LOG', '') if isinstance(sas_output, dict) else ''
        has_issue = check_for_errors_in_log(
            log_output_path,
            fallback_log_content=log_from_sas,
            on_window_close=on_log_window_close if own_session else None,
        )
        if has_issue:
            print(f"SAS程序 {sas_file_path} 执行时出现错误或警告！")
        else:
            print(f"SAS程序 {sas_file_path} 执行成功。")
            if not IS_LINUX:
                messagebox.showinfo("完成", "恭喜您，程序已运行完成! 无ERROR/WARNING。")
        return has_issue
    finally:
        if own_session and not session_ended[0]:
            sas.endsas()


def main():
    parser = argparse.ArgumentParser(description='从 Python 调用 Linux 上的 SAS 程序（支持多个 SAS 文件）')
    parser.add_argument(
        'sas_file_paths',
        nargs='+',
        help='一个或多个 SAS 程序在 Linux 上的完整路径，例如: /u01/app/sas/sas9.4/DocumentRepository/DDT/projects/.../xxx.sas'
    )
    args = parser.parse_args()
    paths = [os.path.normpath(p) for p in args.sas_file_paths]
    if len(paths) == 1:
        run_sas(paths[0])
        return
    # 多个文件：共用一个 SAS 会话依次执行，不进行日志检查
    sas = saspy.SASsession(cfgname='winiomlinux')
    try:
        for i, sas_file_path in enumerate(paths, 1):
            print(f"\n[{i}/{len(paths)}] 执行: {sas_file_path}")
            run_sas(sas_file_path, sas_session=sas, check_log=False)
        print(f"\n全部 {len(paths)} 个 SAS 程序已提交执行。")
    finally:
        sas.endsas()



if __name__ == '__main__':
# 测试用：无命令行参数时使用以下 6 个 SAS 文件
    _TEST_SAS_PATHS = [
        'Z:/projects/HRS2129/HRS2129_test/csr_01/utility/tools/25_generate_pdt_call.sas'

    ]
    if len(sys.argv) == 1:
        sys.argv[1:] = _TEST_SAS_PATHS
    main()
