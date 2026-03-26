# -*- coding: utf-8 -*-
"""
TFLs 页面 - Initial PGM 按钮逻辑（独立模块）

主界面在 TFLs 页面提供「Initial PGM」按钮，绑定 command=lambda: run_initial_pgm(gui)。
弹窗两步：第一步运行 91_initial_pgm_call.sas；第二步运行 61_ladae_template_call.sas（均通过 PROD 端）。
"""
import os
import re
import threading
import tkinter as tk
from tkinter import messagebox, filedialog


def _get_project_base_path(gui):
    """从 gui 获取当前项目根路径（前四个下拉框拼接）。"""
    base = getattr(gui, "z_drive", "Z:\\")
    for i in range(4):
        if getattr(gui, "selected_paths", None) and i < len(gui.selected_paths) and gui.selected_paths[i]:
            base = os.path.join(base, gui.selected_paths[i])
    return base


def _extract_initial_pgm_calls(sas_file_path):
    """
    从 91_initial_pgm_call.sas 中提取每条 %initial_pgm(...) 调用（包含分号）。
    支持跨行写法，并忽略注释行中的匹配。
    """
    with open(sas_file_path, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()
    # 先去掉块注释，避免注释中的 %initial_pgm 被误识别
    no_block_comment = re.sub(r"/\*.*?\*/", "", content, flags=re.DOTALL)

    # 再去掉整行注释：以 * 或 /* 开头的行
    cleaned_lines = []
    for raw_line in no_block_comment.splitlines():
        stripped = raw_line.lstrip()
        if stripped.startswith("*") or stripped.startswith("/*"):
            continue
        cleaned_lines.append(raw_line)

    # 仅识别“行首（可有空白）是 %initial_pgm”的调用，支持跨行直到分号结束
    calls = []
    start_pat = re.compile(r"^\s*%initial_pgm\b", flags=re.IGNORECASE)
    i = 0
    while i < len(cleaned_lines):
        line = cleaned_lines[i]
        if not start_pat.match(line):
            i += 1
            continue

        buf = [line]
        if ";" in line:
            calls.append("\n".join(buf).strip())
            i += 1
            continue

        i += 1
        while i < len(cleaned_lines):
            buf.append(cleaned_lines[i])
            if ";" in cleaned_lines[i]:
                break
            i += 1
        calls.append("\n".join(buf).strip())
        i += 1

    return calls


def _build_initial_pgm_temp_script(macro_call):
    """为单条 %initial_pgm 调用补齐运行所需的 autorun 初始化块。"""
    return (
        "/* Auto generated: run single %initial_pgm call */\n"
        "data _null_;\n"
        "  if libref('adam') then call execute('%nrstr(%autorun)');\n"
        "run;\n"
        f"{macro_call}\n"
    )


def run_initial_pgm(gui):
    """
    点击 TFLs 页面「Initial PGM」按钮时调用。
    弹出两步弹窗，仿 PDT Gen 风格：第一步运行 91_initial_pgm_call.sas，第二步运行 61_ladae_template_call.sas。
    """
    base_path = _get_project_base_path(gui)
    if not base_path or not os.path.isdir(base_path):
        messagebox.showwarning("Initial PGM", "请先在 TFLs 页面选择有效的项目路径（前四个下拉框）。")
        return

    try:
        from linux_sas_call_from_python import run_sas, review_log_from_content
    except ImportError as e:
        messagebox.showerror("错误", "无法导入 linux_sas_call_from_python（请确保该模块在项目目录下且已安装 saspy）。\n\n%s" % e)
        return

    default_sas_60 = os.path.join(base_path, "utility", "tools", "91_initial_pgm_call.sas")
    default_sas_61 = os.path.join(base_path, "utility", "tools", "61_ladae_template_call.sas")

    dlg = tk.Toplevel(gui.root)
    dlg.title("Initial PGM")
    dlg.geometry("1350x460")
    dlg.resizable(True, False)
    dlg.transient(gui.root)
    dlg.configure(bg="#f0f0f0")

    main = tk.Frame(dlg, padx=20, pady=16, bg="#f0f0f0")
    main.pack(fill=tk.BOTH, expand=True)

    # ---------- 第一步 ----------
    step1_title = tk.Label(
        main,
        text="第一步：运行 PROD 端 91_initial_pgm_call.sas，根据 PDT 在 PROD 端生成初始程序。",
        font=("Microsoft YaHei UI", 10, "bold"),
        fg="#333333",
        bg="#f0f0f0",
        wraplength=820
    )
    step1_title.pack(anchor="w", pady=(0, 10))

    row_sas60 = tk.Frame(main, bg="#f0f0f0")
    row_sas60.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_sas60, text="91_initial_pgm_call.sas：", font=("Microsoft YaHei UI", 9), width=28, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    entry_sas60 = tk.Entry(row_sas60, width=72, font=("Microsoft YaHei UI", 9))
    entry_sas60.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    entry_sas60.insert(0, default_sas_60)

    def browse_sas60():
        path = filedialog.askopenfilename(
            title="选择 91_initial_pgm_call.sas",
            filetypes=[("SAS", "*.sas"), ("All", "*.*")],
            initialdir=os.path.dirname(default_sas_60) or base_path
        )
        if path:
            entry_sas60.delete(0, tk.END)
            entry_sas60.insert(0, path)

    tk.Button(row_sas60, text="浏览...", command=browse_sas60, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    _hint_text_step1 = "初版程序生成中，可前往06_programs/09_validation文件夹下查看细节； 待全部初版程序完成，将跳出日志弹窗，请耐心等待。"

    _hint_text_step2 = _hint_text_step1
    step1_lock = threading.Lock()
    step1_running = [False]

    def run_step1():
        """点击「初版SAS PGMs」：拆分 91 文件中的每个 %initial_pgm 并逐个执行。"""
        with step1_lock:
            if step1_running[0]:
                messagebox.showinfo("提示", "初版SAS PGMs 正在运行中，请等待当前任务完成。")
                return
            step1_running[0] = True
        path = entry_sas60.get().strip()
        if not path or not os.path.isfile(path):
            with step1_lock:
                step1_running[0] = False
            messagebox.showwarning("提示", "请选择有效的 91_initial_pgm_call.sas 文件。")
            return
        # 先展示蓝色提示性文字，再调用 SAS
        hint_step1.config(text=_hint_text_step1)
        dlg.update_idletasks()
        def worker():
            try:
                macro_calls = _extract_initial_pgm_calls(path)
                if not macro_calls:
                    raise ValueError("在 91_initial_pgm_call.sas 中未找到 %initial_pgm 调用。")

                temp_dir = os.path.join(os.path.dirname(path), "_tmp_initial_pgm_calls")
                os.makedirs(temp_dir, exist_ok=True)

                total = len(macro_calls)
                issue_logs = []
                success_count = 0
                for idx, macro_call in enumerate(macro_calls, start=1):
                    temp_sas = os.path.join(temp_dir, f"91_initial_pgm_call_{idx:03d}.sas")
                    temp_content = _build_initial_pgm_temp_script(macro_call)
                    with open(temp_sas, "w", encoding="utf-8", errors="replace") as tf:
                        tf.write(temp_content)

                    has_issue, summary, log_text, source_label = run_sas(
                        temp_sas, check_log=False, return_details=True
                    )
                    if has_issue:
                        issue_logs.append((idx, summary, log_text, source_label))
                    else:
                        success_count += 1
            except Exception as e:
                gui.root.after(0, lambda: messagebox.showerror("错误", "调用 SAS 程序时出错：%s" % e))
                with step1_lock:
                    step1_running[0] = False
                return

            def on_done():
                if issue_logs:
                    # 逐条弹出存在 ERROR/WARNING 的日志审阅窗口
                    for _, _, one_log_text, one_source_label in issue_logs:
                        review_log_from_content(one_log_text, one_source_label)
                else:
                    messagebox.showinfo("完成", "恭喜您，程序已运行完成! 无ERROR/WARNING。")
                gui.update_status(
                    (
                        f"91_initial_pgm_call.sas 已按 %initial_pgm 拆分执行：共 {total} 条，成功 {success_count} 条，异常 {len(issue_logs)} 条（已弹出日志审阅）。"
                        if issue_logs
                        else f"91_initial_pgm_call.sas 已按 %initial_pgm 拆分执行：共 {total} 条，全部成功。"
                    )
                )
                with step1_lock:
                    step1_running[0] = False

            gui.root.after(0, on_done)

        threading.Thread(target=worker, daemon=True).start()

    def open_06_programs():
        folder = os.path.join(base_path, "06_programs")
        if os.path.isdir(folder):
            os.startfile(folder)
            gui.update_status("已打开: 06_programs")
        else:
            messagebox.showwarning("提示", "文件夹不存在：%s" % folder)

    btn_row1 = tk.Frame(main, bg="#f0f0f0")
    btn_row1.pack(anchor="w", pady=(4, 0))
    tk.Button(btn_row1, text="初版SAS PGMs", command=run_step1, width=14, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_row1, text="查看", command=open_06_programs, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)
    # 第一步蓝色提示：放在第二步文字描述前，初始为空，点击「初版SAS PGMs」后显示内容
    hint_step1 = tk.Label(main, text="", font=("Microsoft YaHei UI", 9), fg="#0000CC", bg="#f0f0f0", justify=tk.LEFT)
    hint_step1.pack(anchor="w", pady=(6, 0))

    # ---------- 第二步 ----------
    step2_title = tk.Label(
        main,
        text="第二步：运行 PROD 端 61_ladae_template_call.sas，根据 PDT 在 PROD 端生成 ladae_xx 初始程序。",
        font=("Microsoft YaHei UI", 10, "bold"),
        fg="#333333",
        bg="#f0f0f0"
    )
    step2_title.pack(anchor="w", pady=(14, 10))

    row_sas61 = tk.Frame(main, bg="#f0f0f0")
    row_sas61.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_sas61, text="61_ladae_template_call.sas：", font=("Microsoft YaHei UI", 9), width=28, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    entry_sas61 = tk.Entry(row_sas61, width=72, font=("Microsoft YaHei UI", 9))
    entry_sas61.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    entry_sas61.insert(0, default_sas_61)

    def browse_sas61():
        path = filedialog.askopenfilename(
            title="选择 61_ladae_template_call.sas",
            filetypes=[("SAS", "*.sas"), ("All", "*.*")],
            initialdir=os.path.dirname(default_sas_61) or base_path
        )
        if path:
            entry_sas61.delete(0, tk.END)
            entry_sas61.insert(0, path)

    tk.Button(row_sas61, text="浏览...", command=browse_sas61, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    def run_step2():
        """点击「初版ladae_xx.sas」：先展示蓝色提示，再调用 SAS 程序（与第一步执行顺序一致）。"""
        path = entry_sas61.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("提示", "请选择有效的 61_ladae_template_call.sas 文件。")
            return
        # 先展示蓝色提示性文字，再调用 SAS
        hint_step2.config(text=_hint_text_step2)
        hint_step2.pack(anchor="w", pady=(6, 0))
        dlg.update_idletasks()
        def worker():
            try:
                has_issue, summary, log_text, source_label = run_sas(
                    path, check_log=False, return_details=True
                )
            except Exception as e:
                gui.root.after(0, lambda: messagebox.showerror("错误", "调用 SAS 程序时出错：%s" % e))
                return

            def on_done():
                if has_issue:
                    review_log_from_content(log_text, source_label)
                else:
                    messagebox.showinfo("完成", "恭喜您，程序已运行完成! 无ERROR/WARNING。")
                gui.update_status(
                    f"61_ladae_template_call.sas 已执行完成（{summary}，日志弹窗已显示 ERROR/WARNING 关键行）。"
                    if has_issue
                    else f"已在 PROD 端执行 61_ladae_template_call.sas（{summary}）。"
                )

            gui.root.after(0, on_done)

        threading.Thread(target=worker, daemon=True).start()

    def open_062_safety():
        folder = os.path.join(base_path, "06_programs", "062_safety")
        if os.path.isdir(folder):
            os.startfile(folder)
            gui.update_status("已打开: 06_programs/062_safety")
        else:
            messagebox.showwarning("提示", "文件夹不存在：%s" % folder)

    btn_row2 = tk.Frame(main, bg="#f0f0f0")
    btn_row2.pack(anchor="w", pady=(4, 0))
    tk.Button(btn_row2, text="初版ladae_xx.sas", command=run_step2, width=18, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_row2, text="查看", command=open_062_safety, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)
    hint_step2 = tk.Label(main, text="", font=("Microsoft YaHei UI", 9), fg="#0000CC", bg="#f0f0f0", justify=tk.LEFT)
    # 初始不展示，点击「初版ladae_xx.sas」后再展示

    dlg.focus_set()
