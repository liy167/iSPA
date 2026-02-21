# -*- coding: utf-8 -*-
"""
TFLs 页面 - TFLs Combine 按钮逻辑（独立模块）

主界面在 TFLs 页面提供「TFLs Combine」按钮，绑定 command=lambda: run_tfls_combine(gui)。
点击后弹出弹窗：第一步确认需要 Combined TFLs（PDT Excel 文件），可编辑或确认。
"""
import os
import tkinter as tk
from tkinter import messagebox, filedialog


def _get_project_base_path(gui):
    """从 gui 获取当前项目根路径（前四个下拉框拼接）。"""
    base = getattr(gui, "z_drive", "Z:\\")
    for i in range(4):
        if getattr(gui, "selected_paths", None) and i < len(gui.selected_paths) and gui.selected_paths[i]:
            base = os.path.join(base, gui.selected_paths[i])
    return base


def _get_default_pdt_filename(gui):
    """默认 PDT 文件名：下拉框3的值 + _ + 下拉框4的值 + _PDT.xlsx"""
    paths = getattr(gui, "selected_paths", None) or []
    part3 = (paths[2] or "").strip() if len(paths) > 2 else ""
    part4 = (paths[3] or "").strip() if len(paths) > 3 else ""
    if part3 and part4:
        return part3 + "_" + part4 + "_PDT.xlsx"
    if part3 or part4:
        return (part3 or part4) + "_PDT.xlsx"
    return "PDT.xlsx"


def run_tfls_combine(gui):
    """
    点击 TFLs 页面「TFLs Combine」按钮时调用。
    弹窗第一步：确认需要 Combined TFLs，文件浏览默认为 utility\\documentation 下 Excel，文件名为 下拉框3+下拉框4+_PDT.xlsx。
    """
    base_path = _get_project_base_path(gui)
    if not base_path or not os.path.isdir(base_path):
        messagebox.showwarning("TFLs Combine", "请先在 TFLs 页面选择有效的项目路径（前四个下拉框）。")
        return

    doc_dir = os.path.join(base_path, "utility", "documentation")
    default_pdt_name = _get_default_pdt_filename(gui)
    default_pdt_path = os.path.join(doc_dir, default_pdt_name) if doc_dir else default_pdt_name

    dlg = tk.Toplevel(gui.root)
    dlg.title("TFLs Combine")
    dlg.geometry("1350x269")  # 宽度与 Batch Run 弹窗相同，高度为原 480 的 0.56 倍
    dlg.resizable(True, False)
    dlg.transient(gui.root)
    dlg.configure(bg="#f0f0f0")

    main = tk.Frame(dlg, padx=20, pady=16, bg="#f0f0f0")
    main.pack(fill=tk.BOTH, expand=True)

    # ---------- 第一步 ----------
    tools_dir = os.path.join(base_path, "utility", "tools")
    default_rtf_combine_sas = os.path.join(tools_dir, "31_rtf_combine_call.sas")
    step1_title = tk.Label(
        main,
        text="第一步：运行31_rtf_combine_call.sas，合并TFLs。",
        font=("Microsoft YaHei UI", 10, "bold"),
        fg="#333333",
        bg="#f0f0f0"
    )
    step1_title.pack(anchor="w", pady=(0, 10))

    row_pdt = tk.Frame(main, bg="#f0f0f0")
    row_pdt.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_pdt, text="PDT 文件（Excel）：", font=("Microsoft YaHei UI", 9), width=24, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    entry_pdt = tk.Entry(row_pdt, width=52, font=("Microsoft YaHei UI", 9))
    entry_pdt.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    entry_pdt.insert(0, default_pdt_path)

    def browse_pdt():
        path = filedialog.askopenfilename(
            title="选择 PDT 文件（Excel）",
            filetypes=[("Excel", "*.xlsx"), ("Excel 97-2003", "*.xls"), ("All", "*.*")],
            initialdir=doc_dir if os.path.isdir(doc_dir) else base_path
        )
        if path:
            entry_pdt.delete(0, tk.END)
            entry_pdt.insert(0, path)

    def edit_pdt():
        path = entry_pdt.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择或输入 PDT 文件路径。")
            return
        if not os.path.isfile(path):
            messagebox.showwarning("提示", "文件不存在：%s" % path)
            return
        if hasattr(gui, "_open_with_excel"):
            gui._open_with_excel(path)
        else:
            os.startfile(path)
        gui.update_status("已打开: %s" % os.path.basename(path))

    tk.Button(row_pdt, text="浏览...", command=browse_pdt, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 4))
    tk.Button(row_pdt, text="编辑", command=edit_pdt, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    row_rtf_combine = tk.Frame(main, bg="#f0f0f0")
    row_rtf_combine.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_rtf_combine, text="31_rtf_combine_call.sas：", font=("Microsoft YaHei UI", 9), width=24, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    entry_rtf_combine_sas = tk.Entry(row_rtf_combine, width=52, font=("Microsoft YaHei UI", 9))
    entry_rtf_combine_sas.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    entry_rtf_combine_sas.insert(0, default_rtf_combine_sas)

    def browse_rtf_combine_sas():
        path = filedialog.askopenfilename(
            title="选择 31_rtf_combine_call.sas",
            filetypes=[("SAS", "*.sas"), ("All", "*.*")],
            initialdir=tools_dir if os.path.isdir(tools_dir) else base_path
        )
        if path:
            entry_rtf_combine_sas.delete(0, tk.END)
            entry_rtf_combine_sas.insert(0, path)

    def edit_rtf_combine_sas():
        path = entry_rtf_combine_sas.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择或输入 31_rtf_combine_call.sas 路径。")
            return
        if not os.path.isfile(path):
            messagebox.showwarning("提示", "文件不存在：%s" % path)
            return
        os.startfile(path)
        gui.update_status("已打开: %s" % os.path.basename(path))

    tk.Button(row_rtf_combine, text="浏览...", command=browse_rtf_combine_sas, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 4))
    tk.Button(row_rtf_combine, text="编辑", command=edit_rtf_combine_sas, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    hint_combine = tk.Label(main, text="", font=("Microsoft YaHei UI", 9), fg="#0000CC", bg="#f0f0f0", justify=tk.LEFT)

    def run_combine_tfls():
        """点击「Combine TFLs」：显示蓝色提示，运行第一步浏览框中的 31_rtf_combine_call.sas；SAS 被 terminate 不弹异常窗；完成后打开 03_reports 文件夹。"""
        sas_path = entry_rtf_combine_sas.get().strip()
        if not sas_path or not os.path.isfile(sas_path):
            messagebox.showerror("错误", "未找到程序：%s" % sas_path)
            return
        try:
            from linux_sas_call_from_python import run_sas
        except ImportError as e:
            messagebox.showerror("错误", "无法导入 linux_sas_call_from_python。\n\n%s" % e)
            return
        hint_combine.config(text="TFLs合并中，请耐心等待。")
        hint_combine.pack(anchor="w", pady=(8, 0))
        dlg.update_idletasks()
        gui.update_status("正在运行 31_rtf_combine_call.sas…")
        try:
            run_sas(sas_path, check_log=False)
        except Exception as e:
            err_msg = str(e)
            if "terminated unexpectedly" in err_msg or "No SAS process attached" in err_msg or "terminate" in err_msg.lower():
                pass
            else:
                hint_combine.config(text="")
                gui.update_status("Combine TFLs 执行出错：%s" % e)
                messagebox.showerror("错误", "运行 31_rtf_combine_call.sas 时出错：%s" % e)
                return
        hint_combine.config(text="")
        gui.update_status("Combine TFLs 已执行完成。")
        reports_dir = os.path.join(base_path, "03_reports")
        _show_folder_window(dlg, reports_dir)

    def _show_folder_window(parent, folder_path):
        """弹出新窗口：按钮在文字说明上方，展示文件夹路径。"""
        win = tk.Toplevel(parent)
        win.title("03_reports")
        win.geometry("600x160")  # 宽度为原 480 的 1.25 倍
        win.configure(bg="#f0f0f0")
        win.transient(parent)
        def open_folder():
            if os.path.isdir(folder_path):
                try:
                    os.startfile(folder_path)
                except Exception:
                    pass
        btn_frame = tk.Frame(win, bg="#f0f0f0")
        btn_frame.pack(anchor="w", padx=12, pady=(12, 10))
        tk.Button(btn_frame, text="打开03_reports文件夹", command=open_folder, width=22, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)
        tk.Label(win, text="文件夹路径：", font=("Microsoft YaHei UI", 10), fg="#333333", bg="#f0f0f0").pack(anchor="w", padx=12, pady=(0, 4))
        tk.Label(win, text=folder_path, font=("Consolas", 9), fg="#333333", bg="#f0f0f0", wraplength=560, justify=tk.LEFT).pack(anchor="w", padx=12, pady=(0, 12))
        win.focus_set()

    btn_combine_frame = tk.Frame(main, bg="#f0f0f0")
    btn_combine_frame.pack(anchor="w", pady=(16, 0))
    tk.Button(btn_combine_frame, text="合并TFLs", command=run_combine_tfls, width=14, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    dlg.focus_set()
    entry_pdt.focus_set()
