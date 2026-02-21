# -*- coding: utf-8 -*-
"""
TFLs 页面 - Batch Run 按钮逻辑（独立模块）

主界面在 TFLs 页面提供「Batch Run」按钮，绑定 command=lambda: run_batch_run(gui)。
点击后弹出弹窗，仿照 PDT Gen 风格。第一步：解析 92 并生成 (out)_call.sas，再批量运行产生的 sas 程序。
"""
import glob
import os
import re
import tkinter as tk
from tkinter import messagebox, filedialog


def _get_project_base_path(gui):
    """从 gui 获取当前项目根路径（前四个下拉框拼接）。"""
    base = getattr(gui, "z_drive", "Z:\\")
    for i in range(4):
        if getattr(gui, "selected_paths", None) and i < len(gui.selected_paths) and gui.selected_paths[i]:
            base = os.path.join(base, gui.selected_paths[i])
    return base


def _open_xml_with_excel(xml_path, gui):
    """使用 Excel 打开 XML 文件；若 gui 有 _open_with_excel 则调用，否则用系统默认方式打开。"""
    if hasattr(gui, "_open_with_excel"):
        gui._open_with_excel(xml_path)
    else:
        os.startfile(xml_path)


def _show_log_check_xml_list(parent, base_path, gui):
    """展示 base_path/07_logs 下所有 XML 文件，双击用 Excel 打开。"""
    logs_dir = os.path.join(base_path, "07_logs")
    xml_paths = sorted(glob.glob(os.path.join(logs_dir, "*.xml"))) if os.path.isdir(logs_dir) else []
    win = tk.Toplevel(parent)
    win.title("Log Check 生成的 XML 文件")
    win.geometry("325x160")  # 宽度为当前的 1.25 倍（260 * 1.25）
    win.configure(bg="#f0f0f0")
    win.transient(parent)
    tk.Label(
        win, text="双击下方文件使用 Excel 打开。",
        font=("Microsoft YaHei UI", 9), fg="#333333", bg="#f0f0f0"
    ).pack(anchor="w", padx=12, pady=(10, 4))
    list_frame = tk.Frame(win, bg="#f0f0f0")
    list_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 10))
    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox = tk.Listbox(
        list_frame, font=("Consolas", 10), yscrollcommand=scrollbar.set,
        selectmode=tk.SINGLE, activestyle="dotbox", height=12
    )
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)
    for p in xml_paths:
        listbox.insert(tk.END, os.path.basename(p))
    if not xml_paths:
        listbox.insert(tk.END, "  （07_logs 下暂无 XML 文件）")
    def on_double_click(event):
        sel = listbox.curselection()
        if not sel or not xml_paths:
            return
        idx = sel[0]
        if idx < len(xml_paths):
            _open_xml_with_excel(xml_paths[idx], gui)
    listbox.bind("<Double-Button-1>", on_double_click)
    win.update_idletasks()
    w, h = 325, 160
    x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")


# 匹配含 %batch_script_generator 行中的 out= 赋值（type=%str(...) 中可能有括号，故只匹配 out= 值）
_OUT_PARAM_RE = re.compile(r"out\s*=\s*([^,\s\)]+)", re.IGNORECASE)


def _parse_batch_script_generator_outs(sas_92_path):
    """
    读取 92_batch_script_generator_call.sas，遍历含有 %batch_script_generator 的程序行（跳过 /* */ 块注释内），
    返回 [(行内容, out 值), ...]，out 值用于生成文件名 (out)_call.sas。
    """
    if not sas_92_path or not os.path.isfile(sas_92_path):
        return []
    result = []
    in_block_comment = False
    with open(sas_92_path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            # 简单处理块注释：不解析 /* ... */ 内的行
            if "*/" in line:
                in_block_comment = False
                continue
            if "/*" in line:
                in_block_comment = True
                continue
            if in_block_comment:
                continue
            line_stripped = line.strip()
            if "%batch_script_generator" not in line_stripped:
                continue
            m = _OUT_PARAM_RE.search(line)
            if not m:
                continue
            out_val = m.group(1).strip()
            if out_val:
                result.append((line_stripped, out_val))
    return result


def _parse_batch_submit_lines(script_path):
    """
    读取 Batch Run 脚本，提取所有以 %batch_submit 开头的程序行（跳过 /* */ 与 * 注释）。
    返回行内容列表。
    """
    if not script_path or not os.path.isfile(script_path):
        return []
    result = []
    in_block_comment = False
    with open(script_path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            if "*/" in line:
                in_block_comment = False
                continue
            if "/*" in line:
                in_block_comment = True
                continue
            if in_block_comment:
                continue
            line_stripped = line.strip()
            if not line_stripped or line_stripped.startswith("*"):
                continue
            if "%batch_submit" in line_stripped:
                result.append(line_stripped)
    return result


# 匹配 %log_chk 宏调用开始（不区分大小写）
_LOG_CHK_START_RE = re.compile(r"%log_chk\s*\(", re.IGNORECASE)


def _parse_log_chk_calls(script_path):
    """
    读取 Log Check 脚本，识别所有 %log_chk(...) 语句（可跨行），跳过 /* */ 块注释。
    每一个 %log_chk 解析为单独的一条宏调用字符串，返回 [宏调用1, 宏调用2, ...]。
    """
    if not script_path or not os.path.isfile(script_path):
        return []
    with open(script_path, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()
    # 去掉块注释，避免注释内的 %log_chk 被误识别
    content_clean = re.sub(r"/\*.*?\*/", " ", content, flags=re.DOTALL)
    result = []
    for m in _LOG_CHK_START_RE.finditer(content_clean):
        macro_start = m.start()
        paren_start = content_clean.index("(", macro_start)
        depth = 1
        i = paren_start + 1
        while i < len(content_clean) and depth > 0:
            if content_clean[i] == "(":
                depth += 1
            elif content_clean[i] == ")":
                depth -= 1
            i += 1
        if depth != 0:
            continue
        full_call = content_clean[macro_start : i].strip()
        if not full_call.endswith(";"):
            full_call += ";"
        result.append(full_call)
    return result


def _generate_call_sas_files(sas_92_path, tools_dir):
    """
    根据 92 程序中的 %batch_script_generator 行，在 tools_dir 下生成 (out)_call.sas 文件。
    每个文件内容：data _null_/autorun 块 + 该行宏调用。
    返回生成的 .sas 文件路径列表。
    """
    entries = _parse_batch_script_generator_outs(sas_92_path)
    if not entries:
        return []
    os.makedirs(tools_dir, exist_ok=True)
    autorun_block = """data _null_;
  if libref('adam') then call execute('%nrstr(%autorun)');
run;

"""
    generated = []
    for line_content, out_val in entries:
        fname = out_val + "_call.sas"
        fpath = os.path.join(tools_dir, fname)
        content = autorun_block + line_content + "\n"
        with open(fpath, "w", encoding="utf-8", newline="\n") as f:
            f.write(content)
        generated.append(fpath)
    return generated


def run_batch_run(gui):
    """
    点击 TFLs 页面「Batch Run」按钮时调用。
    弹出弹窗，仿照 PDT Gen 风格；第一步：运行 92_batch_script_generator_call.sas。
    """
    base_path = _get_project_base_path(gui)
    if not base_path or not os.path.isdir(base_path):
        messagebox.showwarning("Batch Run", "请先在 TFLs 页面选择有效的项目路径（前四个下拉框）。")
        return

    try:
        from linux_sas_call_from_python import run_sas, convert_windows_path_to_linux
        import saspy
    except ImportError as e:
        messagebox.showerror("错误", "无法导入 linux_sas_call_from_python 或 saspy（请确保该模块在项目目录下且已安装 saspy）。\n\n%s" % e)
        return

    # 默认路径：前四个下拉框 + utility\tools\92_batch_script_generator_call.sas
    default_sas_92 = os.path.join(base_path, "utility", "tools", "92_batch_script_generator_call.sas")

    dlg = tk.Toplevel(gui.root)
    dlg.title("Batch Run")
    dlg.geometry("1400x500")  # 宽度比原 1350 增加 50
    dlg.resizable(True, False)
    dlg.transient(gui.root)
    dlg.grab_set()
    dlg.configure(bg="#f0f0f0")

    main = tk.Frame(dlg, padx=20, pady=16, bg="#f0f0f0")
    main.pack(fill=tk.BOTH, expand=True)

    # ---------- 第一步 ----------
    step1_title = tk.Label(
        main,
        text="第一步：运行 92_batch_script_generator_call.sas，产生数据集和 TFLs 单独 Batch Run 脚本程序。",
        font=("Microsoft YaHei UI", 10, "bold"),
        fg="#333333",
        bg="#f0f0f0"
    )
    step1_title.pack(anchor="w", pady=(0, 10))

    row_sas92 = tk.Frame(main, bg="#f0f0f0")
    row_sas92.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_sas92, text="92_batch_script_generator_call.sas：", font=("Microsoft YaHei UI", 9), width=32, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    entry_sas92 = tk.Entry(row_sas92, width=72, font=("Microsoft YaHei UI", 9))
    entry_sas92.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    entry_sas92.insert(0, default_sas_92)

    def browse_sas92():
        path = filedialog.askopenfilename(
            title="选择 92_batch_script_generator_call.sas",
            filetypes=[("SAS", "*.sas"), ("All", "*.*")],
            initialdir=os.path.dirname(default_sas_92) or base_path
        )
        if path:
            entry_sas92.delete(0, tk.END)
            entry_sas92.insert(0, path)

    def open_tools_edit():
        """打开第一步浏览框中的文件（92_batch_script_generator_call.sas）。"""
        path = entry_sas92.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择或输入 92_batch_script_generator_call.sas 的路径。")
            return
        if os.path.isfile(path):
            os.startfile(path)
            gui.update_status("已打开: %s" % os.path.basename(path))
        else:
            messagebox.showwarning("提示", "文件不存在：%s" % path)

    tk.Button(row_sas92, text="浏览...", command=browse_sas92, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 4))
    tk.Button(row_sas92, text="编辑", command=open_tools_edit, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    _hint_text_step1 = "初版Batch Run脚本生成中，可前往utility\\tools\\文件夹下查看细节。初版Batch Run脚本完成后将跳出日志弹窗，请耐心等待。"

    def run_step1():
        """点击「初版Batch Run脚本」：第一步展示蓝色提示；第二步解析 92 并生成 (out)_call.sas；第三步批量运行生成的 sas。"""
        path = entry_sas92.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("提示", "请选择有效的 92_batch_script_generator_call.sas 文件。")
            return
        # 第一步：展示蓝色提示
        hint_step1.config(text=_hint_text_step1)
        dlg.update_idletasks()
        tools_dir = os.path.join(base_path, "utility", "tools")
        try:
            # 第二步：读取 92 程序，在 utility\tools 下生成 (out)_call.sas
            generated = _generate_call_sas_files(path, tools_dir)
            if not generated:
                messagebox.showwarning("提示", "未在 92 程序中找到包含 %batch_script_generator 的行，或 out= 解析失败。")
                return
            gui.update_status("已生成 %d 个初版 Batch Run 脚本，正在批量运行（Linux 路径，不检查日志）…" % len(generated))
            # 第三步：批量运行。%batch_script_generator 会强制终止 SAS 进程，遇此情况不弹窗，新建会话后继续下一个
            sas = saspy.SASsession(cfgname='winiomlinux')
            try:
                for sas_path in generated:
                    sas_path_linux = convert_windows_path_to_linux(sas_path)
                    try:
                        run_sas(sas_path_linux, sas_session=sas, check_log=False)
                    except Exception as e:
                        err_msg = str(e)
                        if "terminated unexpectedly" in err_msg or "No SAS process attached" in err_msg:
                            # 宏强制终止了 SAS 进程：不进行日志检查、不弹窗，直接运行下一个
                            try:
                                sas.endsas()
                            except Exception:
                                pass
                            sas = saspy.SASsession(cfgname='winiomlinux')
                            gui.update_status("已运行 %s（SAS 进程已由宏终止），继续下一个…" % os.path.basename(sas_path))
                        else:
                            gui.update_status("Batch Run 执行出错：%s" % e)
                            messagebox.showerror("错误", "运行 %s 时出错：%s" % (os.path.basename(sas_path), e))
                            return
            finally:
                try:
                    sas.endsas()
                except Exception:
                    pass
            gui.update_status("初版 Batch Run 脚本已全部执行完成。")
            messagebox.showinfo("完成", "恭喜您，初版Batch Run 脚本已全部执行完成。")
            # 删除第二步产生的 sas 程序文件及对应的日志文件（日志与 sas 同目录，同名 .log）
            for sas_path in generated:
                try:
                    if os.path.isfile(sas_path):
                        os.remove(sas_path)
                except Exception:
                    pass
                log_path = os.path.join(os.path.dirname(sas_path), os.path.splitext(os.path.basename(sas_path))[0] + ".log")
                try:
                    if os.path.isfile(log_path):
                        os.remove(log_path)
                except Exception:
                    pass
        except Exception as e:
            gui.update_status("Batch Run 执行出错。")
            messagebox.showerror("错误", "执行失败：%s" % e)

    btn_row1 = tk.Frame(main, bg="#f0f0f0")
    btn_row1.pack(anchor="w", pady=(4, 0))
    tk.Button(btn_row1, text="初版Batch Run脚本", command=run_step1, width=18, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    hint_step1 = tk.Label(main, text="", font=("Microsoft YaHei UI", 9), fg="#0000CC", bg="#f0f0f0", justify=tk.LEFT)
    hint_step1.pack(anchor="w", pady=(6, 0))

    # ---------- 第二步 ----------
    tools_dir = os.path.join(base_path, "utility", "tools")
    step2_title = tk.Label(
        main,
        text="第二步：请选择Batch Run的脚本",
        font=("Microsoft YaHei UI", 10, "bold"),
        fg="#333333",
        bg="#f0f0f0"
    )
    step2_title.pack(anchor="w", pady=(14, 10))

    row_batch_script = tk.Frame(main, bg="#f0f0f0")
    row_batch_script.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_batch_script, text="Batch Run 脚本：", font=("Microsoft YaHei UI", 9), width=32, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    entry_batch_script = tk.Entry(row_batch_script, width=72, font=("Microsoft YaHei UI", 9))
    entry_batch_script.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

    def browse_batch_script():
        path = filedialog.askopenfilename(
            title="选择 Batch Run 脚本",
            filetypes=[("SAS", "*.sas"), ("All", "*.*")],
            initialdir=tools_dir if os.path.isdir(tools_dir) else base_path
        )
        if path:
            entry_batch_script.delete(0, tk.END)
            entry_batch_script.insert(0, path)

    tk.Button(row_batch_script, text="浏览...", command=browse_batch_script, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    def run_batch_script():
        """点击「运行」：用 run_batch_script_from_python 解析 batch 脚本，base_path 默认为脚本所在目录的上两级，再用 linux_sas_call_from_python 依次执行解析出的 SAS 程序。"""
        batch_script_path = entry_batch_script.get().strip()
        if not batch_script_path or not os.path.isfile(batch_script_path):
            messagebox.showwarning("提示", "请先选择有效的 Batch Run 脚本。")
            return
        # base_path 默认：Batch Run 脚本路径的上两级文件夹（如 .../utility/tools/xx.sas -> .../csr_01）
        base_path = os.path.normpath(os.path.join(os.path.dirname(batch_script_path), "..", ".."))
        try:
            from run_batch_script_from_python import parse_batch_submits, build_sas_paths
            from linux_sas_call_from_python import run_sas
            import saspy
        except ImportError as e:
            messagebox.showerror("错误", "无法导入 run_batch_script_from_python 或 linux_sas_call_from_python。\n\n%s" % e)
            return
        submits = parse_batch_submits(batch_script_path)
        if not submits:
            messagebox.showwarning("提示", "未在批处理脚本中找到任何 %batch_submit(role=..., target=..., pgm=...)。")
            return
        try:
            sas_paths = build_sas_paths(base_path, submits)
        except ValueError as e:
            messagebox.showerror("错误", "解析 batch_submit 失败：%s" % e)
            return
        gui.update_status("已解析 %d 个 SAS 程序，正在批量运行…" % len(sas_paths))
        dlg.update_idletasks()
        sas = saspy.SASsession(cfgname='winiomlinux')
        try:
            for i, sas_path in enumerate(sas_paths, 1):
                gui.update_status("[%d/%d] 执行: %s" % (i, len(sas_paths), os.path.basename(sas_path)))
                dlg.update_idletasks()
                try:
                    run_sas(sas_path, sas_session=sas, check_log=False)
                except Exception as e:
                    err_msg = str(e)
                    if "terminated unexpectedly" in err_msg or "No SAS process attached" in err_msg:
                        try:
                            sas.endsas()
                        except Exception:
                            pass
                        sas = saspy.SASsession(cfgname='winiomlinux')
                        gui.update_status("已运行 %s（SAS 进程已由宏终止），继续下一个…" % os.path.basename(sas_path))
                    else:
                        gui.update_status("Batch Run 执行出错：%s" % e)
                        messagebox.showerror("错误", "运行 %s 时出错：%s" % (os.path.basename(sas_path), e))
                        return
        finally:
            try:
                sas.endsas()
            except Exception:
                pass
        gui.update_status("Batch Run 已全部执行完成（共 %d 个程序）。" % len(sas_paths))
        messagebox.showinfo("完成", "恭喜您，Batch Run 已全部执行完成（共 %d 个程序）。" % len(sas_paths))

    def edit_batch_script():
        """点击「编辑」：打开选中的 Batch Run 脚本。"""
        path = entry_batch_script.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择 Batch Run 脚本。")
            return
        if os.path.isfile(path):
            os.startfile(path)
            gui.update_status("已打开: %s" % os.path.basename(path))
        else:
            messagebox.showwarning("提示", "文件不存在：%s" % path)

    btn_row2 = tk.Frame(main, bg="#f0f0f0")
    btn_row2.pack(anchor="w", pady=(4, 0))
    tk.Button(btn_row2, text="运行", command=run_batch_script, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_row2, text="编辑", command=edit_batch_script, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    # ---------- 第三步 ----------
    default_log_check_sas = os.path.join(base_path, "utility", "tools", "93_log_check_call.sas")
    step3_title = tk.Label(
        main,
        text="第三步：请选择 Log Check 脚本",
        font=("Microsoft YaHei UI", 10, "bold"),
        fg="#333333",
        bg="#f0f0f0"
    )
    step3_title.pack(anchor="w", pady=(14, 10))

    row_log_check = tk.Frame(main, bg="#f0f0f0")
    row_log_check.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_log_check, text="Log Check 脚本：", font=("Microsoft YaHei UI", 9), width=32, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    entry_log_check = tk.Entry(row_log_check, width=72, font=("Microsoft YaHei UI", 9))
    entry_log_check.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    entry_log_check.insert(0, default_log_check_sas)

    def browse_log_check():
        path = filedialog.askopenfilename(
            title="选择 Log Check 脚本",
            filetypes=[("SAS", "*.sas"), ("All", "*.*")],
            initialdir=tools_dir if os.path.isdir(tools_dir) else base_path
        )
        if path:
            entry_log_check.delete(0, tk.END)
            entry_log_check.insert(0, path)

    tk.Button(row_log_check, text="浏览...", command=browse_log_check, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    def run_log_check():
        """点击「运行」：读取 Log Check 脚本，识别每个 %log_chk 为单独 SAS 宏，用 linux_sas_call_from_python 依次运行。"""
        path = entry_log_check.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("提示", "请先选择有效的 Log Check 脚本。")
            return
        log_chk_calls = _parse_log_chk_calls(path)
        if not log_chk_calls:
            messagebox.showwarning("提示", "未在脚本中找到任何 %log_chk(...) 语句。")
            return
        try:
            from linux_sas_call_from_python import run_sas
            import saspy
        except ImportError as e:
            messagebox.showerror("错误", "无法导入 linux_sas_call_from_python。\n\n%s" % e)
            return
        script_dir = os.path.dirname(path)
        autorun_block = """data _null_;
  if libref('adam') then call execute('%nrstr(%autorun)');
run;

"""
        temp_files = []
        try:
            for i, macro_call in enumerate(log_chk_calls, 1):
                fpath = os.path.join(script_dir, "_log_chk_%d_call.sas" % i)
                with open(fpath, "w", encoding="utf-8", newline="\n") as f:
                    f.write(autorun_block + macro_call + "\n")
                temp_files.append(fpath)
        except Exception as e:
            messagebox.showerror("错误", "写入临时 SAS 文件失败：%s" % e)
            return
        gui.update_status("已解析 %d 个 %%log_chk 宏，正在依次运行…" % len(log_chk_calls))
        dlg.update_idletasks()
        sas = saspy.SASsession(cfgname='winiomlinux')
        try:
            for i, sas_path in enumerate(temp_files, 1):
                gui.update_status("[%d/%d] Log Check: %s" % (i, len(temp_files), os.path.basename(sas_path)))
                dlg.update_idletasks()
                try:
                    run_sas(sas_path, sas_session=sas, check_log=False)
                except Exception as e:
                    err_msg = str(e)
                    if "terminated unexpectedly" in err_msg or "No SAS process attached" in err_msg:
                        try:
                            sas.endsas()
                        except Exception:
                            pass
                        sas = saspy.SASsession(cfgname='winiomlinux')
                        gui.update_status("已运行 %s（SAS 进程已由宏终止），继续下一个…" % os.path.basename(sas_path))
                    else:
                        gui.update_status("Log Check 执行出错：%s" % e)
                        messagebox.showerror("错误", "运行 %s 时出错：%s" % (os.path.basename(sas_path), e))
                        return
        finally:
            try:
                sas.endsas()
            except Exception:
                pass
            for fpath in temp_files:
                try:
                    if os.path.isfile(fpath):
                        os.remove(fpath)
                except Exception:
                    pass
                log_path = os.path.join(script_dir, os.path.splitext(os.path.basename(fpath))[0] + ".log")
                try:
                    if os.path.isfile(log_path):
                        os.remove(log_path)
                except Exception:
                    pass
        gui.update_status("Log Check 已全部执行完成（共 %d 个 %%log_chk）。" % len(log_chk_calls))
        messagebox.showinfo("完成", "恭喜您，Log Check 已全部执行完成。" % len(log_chk_calls))
        _show_log_check_xml_list(dlg, base_path, gui)

    def edit_log_check():
        """点击「编辑」：打开选中的 Log Check 脚本。"""
        path = entry_log_check.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择 Log Check 脚本。")
            return
        if os.path.isfile(path):
            os.startfile(path)
            gui.update_status("已打开: %s" % os.path.basename(path))
        else:
            messagebox.showwarning("提示", "文件不存在：%s" % path)

    btn_row3 = tk.Frame(main, bg="#f0f0f0")
    btn_row3.pack(anchor="w", pady=(4, 0))
    tk.Button(btn_row3, text="运行", command=run_log_check, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_row3, text="编辑", command=edit_log_check, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    dlg.focus_set()
    entry_sas92.focus_set()


def _execute_batch_run(gui, project_path):
    """
    在给定项目路径下执行 Batch Run 的具体逻辑（预留，供后续步骤扩展）。
    可在此扩展：遍历 06_programs/09_validation、调用 linux_sas_call_from_python.run_sas 批量执行等。
    """
    try:
        gui.update_status("Batch Run 执行中…")
        programs_dir = os.path.join(project_path, "06_programs")
        validation_dir = os.path.join(project_path, "09_validation")
        if os.path.isdir(programs_dir) or os.path.isdir(validation_dir):
            gui.update_status("Batch Run 已就绪（可在此扩展批量运行 06_programs/09_validation 程序）。")
            messagebox.showinfo("Batch Run", "项目路径有效。\n\n可在 _execute_batch_run 内扩展：批量运行 06_programs 或 09_validation 下的 SAS 程序。")
        else:
            gui.update_status("Batch Run 已执行（项目路径：%s）。" % project_path)
            messagebox.showinfo("Batch Run", "已根据路径执行 Batch Run。\n\n路径：%s" % project_path)
    except Exception as e:
        gui.update_status("Batch Run 执行出错。")
        messagebox.showerror("Batch Run", "执行失败：%s" % e)
