# -*- coding: utf-8 -*-
"""
TFLs 页面 - Metadata Setup 弹窗：运行 utility/tools/30_generate_tflmeta_call.sas

主界面绑定：command=lambda: show_metadata_setup_dialog(gui)
"""
import os
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


def show_metadata_setup_dialog(gui):
    """
    显示「Metadata Setup」弹窗：默认路径 = 前四级 + utility/tools/30_generate_tflmeta_call.sas，
    浏览/编辑路径，主按钮在 Linux SAS 上执行该程序。
    """
    base_path = _get_project_base_path(gui)
    if not base_path or not os.path.isdir(base_path):
        messagebox.showwarning(
            "Metadata Setup",
            "请先在 TFLs 页面选择有效的项目路径（前四个下拉框）。",
        )
        return

    try:
        from linux_sas_call_from_python import run_sas
    except ImportError as e:
        messagebox.showerror(
            "错误",
            "无法导入 linux_sas_call_from_python（请确保该模块在项目目录下且已安装 saspy）。\n\n%s"
            % e,
        )
        return

    default_sas = os.path.join(base_path, "utility", "tools", "30_generate_tflmeta_call.sas")

    dlg = tk.Toplevel(gui.root)
    dlg.title("Metadata Setup")
    dlg.geometry("1100x220")
    dlg.resizable(True, True)
    dlg.transient(gui.root)
    dlg.configure(bg="#f0f0f0")

    main = tk.Frame(dlg, padx=20, pady=16, bg="#f0f0f0")
    main.pack(fill=tk.BOTH, expand=True)

    step_title = tk.Label(
        main,
        text="第一步：运行 30_generate_tflmeta_call.sas，生成 TFL Metadata。",
        font=("Microsoft YaHei UI", 10, "bold"),
        fg="#333333",
        bg="#f0f0f0",
    )
    step_title.pack(anchor="w", pady=(0, 10))

    row = tk.Frame(main, bg="#f0f0f0")
    row.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(
        row,
        text="30_generate_tflmeta_call.sas：",
        font=("Microsoft YaHei UI", 9),
        width=32,
        anchor="w",
        bg="#f0f0f0",
    ).pack(side=tk.LEFT, padx=(0, 4))
    entry_sas = tk.Entry(row, width=72, font=("Microsoft YaHei UI", 9))
    entry_sas.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    entry_sas.insert(0, default_sas)

    def browse_sas():
        path = filedialog.askopenfilename(
            title="选择 30_generate_tflmeta_call.sas",
            filetypes=[("SAS", "*.sas"), ("All", "*.*")],
            initialdir=os.path.dirname(default_sas) or base_path,
        )
        if path:
            entry_sas.delete(0, tk.END)
            entry_sas.insert(0, path)

    def open_edit():
        path = entry_sas.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择或输入 SAS 程序路径。")
            return
        if os.path.isfile(path):
            os.startfile(path)
            gui.update_status("已打开: %s" % os.path.basename(path))
        else:
            messagebox.showwarning("提示", "文件不存在：%s" % path)

    tk.Button(row, text="浏览...", command=browse_sas, width=8, font=("Microsoft YaHei UI", 9)).pack(
        side=tk.LEFT, padx=(0, 4)
    )
    tk.Button(row, text="编辑", command=open_edit, width=8, font=("Microsoft YaHei UI", 9)).pack(
        side=tk.LEFT
    )

    btn_frame = tk.Frame(main, bg="#f0f0f0")
    btn_frame.pack(anchor="w", pady=(14, 0))

    def run_metadata():
        path = entry_sas.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("提示", "请选择有效的 30_generate_tflmeta_call.sas 文件。")
            return
        if not path.lower().endswith(".sas"):
            messagebox.showwarning("提示", "请选择 .sas 程序文件。")
            return
        def worker():
            try:
                has_issue = run_sas(path, check_log=False)
                gui.root.after(
                    0,
                    lambda: gui.update_status(
                        "30_generate_tflmeta_call.sas 已执行完成。"
                        if has_issue
                        else "已在 Linux 服务器执行 30_generate_tflmeta_call.sas。"
                    ),
                )
            except Exception as e:
                gui.root.after(0, lambda: messagebox.showerror("错误", "调用 SAS 程序时出错：%s" % e))

        threading.Thread(target=worker, daemon=True).start()

    tk.Button(
        btn_frame,
        text="运行 TFL Meta 生成",
        command=run_metadata,
        width=22,
        font=("Microsoft YaHei UI", 9),
    ).pack(side=tk.LEFT)
