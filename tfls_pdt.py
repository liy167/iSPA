# -*- coding: utf-8 -*-
"""
TFLs 页面 - 生成PDT 弹窗逻辑（独立模块，避免 SASEG_GUI 过于庞大）
"""
import os
import tkinter as tk
from tkinter import messagebox, filedialog


def show_pdt_dialog(gui):
    """
    显示「生成PDT」分步弹窗。
    gui: 主窗口实例，需有 .root, .get_current_path(), .selected_paths, .update_status()
    """
    dlg = tk.Toplevel(gui.root)
    dlg.title("生成PDT")
    dlg.geometry("1320x540")  # 1.5× 原宽度 880
    dlg.resizable(True, True)
    dlg.transient(gui.root)
    dlg.grab_set()
    dlg.configure(bg="#f0f0f0")

    main = tk.Frame(dlg, padx=20, pady=16, bg="#f0f0f0")
    main.pack(fill=tk.BOTH, expand=True)

    # ---------- 第一步 ----------
    step1_title = tk.Label(main, text="第一步：填写如下问题", font=("Microsoft YaHei UI", 10, "bold"), fg="#333333", bg="#f0f0f0")
    step1_title.pack(anchor="w", pady=(0, 10))

    step1_content = tk.Frame(main, bg="#f0f0f0")
    step1_content.pack(anchor="w", fill=tk.BOTH, expand=False)

    # 问题1：分析设计类型（多选）
    tk.Label(step1_content, text="问题1  请选择当前分析设计类型（多选框）：", font=("Microsoft YaHei UI", 9), anchor="w", bg="#f0f0f0").pack(anchor="w", pady=(0, 4))
    q1_vars = {}
    q1_frame = tk.Frame(step1_content, bg="#f0f0f0")
    q1_frame.pack(anchor="w", pady=(0, 10))
    for opt in ["SAD", "FE", "MAD", "BE", "MB"]:
        v = tk.BooleanVar(value=False)
        q1_vars[opt] = v
        cb = tk.Checkbutton(q1_frame, text=opt, variable=v, font=("Microsoft YaHei UI", 9), anchor="w", bg="#f0f0f0")
        cb.pack(side=tk.LEFT, padx=(0, 16))

    # 问题2：其他终点（多选）- 第一行 PK浓度+PK参数 共6项，第二行 PD/ADA/QT 共3项
    tk.Label(step1_content, text="问题2  请选择除安全性终点外的其他终点（多选框）：", font=("Microsoft YaHei UI", 9), anchor="w", bg="#f0f0f0").pack(anchor="w", pady=(0, 4))
    q2_row1 = ["PK浓度(血)", "PK浓度(尿)", "PK浓度(粪)", "PK参数(血)", "PK参数(尿)", "PK参数(粪)"]
    q2_row2 = ["PD分析", "ADA分析", "QT分析"]
    q2_vars = {}
    q2_frame = tk.Frame(step1_content, bg="#f0f0f0")
    q2_frame.pack(anchor="w", pady=(0, 10))
    for opt in q2_row1:
        v = tk.BooleanVar(value=False)
        q2_vars[opt] = v
        cb = tk.Checkbutton(q2_frame, text=opt, variable=v, font=("Microsoft YaHei UI", 9), anchor="w", bg="#f0f0f0")
        cb.pack(side=tk.LEFT, padx=(0, 14))
    r2 = tk.Frame(q2_frame, bg="#f0f0f0")
    r2.pack(anchor="w")
    for opt in q2_row2:
        v = tk.BooleanVar(value=False)
        q2_vars[opt] = v
        cb = tk.Checkbutton(r2, text=opt, variable=v, font=("Microsoft YaHei UI", 9), anchor="w", bg="#f0f0f0")
        cb.pack(side=tk.LEFT, padx=(0, 14))

    # 问题3：分析物名称（选中任一 PK浓度 或 PK参数 时显示）
    q3_frame = tk.Frame(step1_content, bg="#f0f0f0")
    tk.Label(q3_frame, text="问题3  若选择PK浓度或PK参数终点，请提供分析物名称（多个分析物用\"|\"分割，如HRS2129|M1|M2）", font=("Microsoft YaHei UI", 9), anchor="w", bg="#f0f0f0").pack(anchor="w", pady=(0, 4))
    q3_entry = tk.Entry(q3_frame, width=80, font=("Microsoft YaHei UI", 9))
    q3_entry.pack(anchor="w", fill=tk.X, pady=(0, 4))

    pk_conc = ["PK浓度(血)", "PK浓度(尿)", "PK浓度(粪)"]
    pk_param = ["PK参数(血)", "PK参数(尿)", "PK参数(粪)"]
    q3_trigger_opts = pk_conc + pk_param

    def toggle_q3(*args):
        if any(q2_vars[o].get() for o in q3_trigger_opts):
            q3_frame.pack(anchor="w", pady=(0, 10))
        else:
            q3_frame.pack_forget()
        dlg.update_idletasks()

    for opt in q3_trigger_opts:
        q2_vars[opt].trace_add("write", toggle_q3)
    q3_frame.pack_forget()  # 初始不显示，勾选任 PK浓度 或 PK参数 后显示

    # ---------- 第二步 ----------
    step2_title = tk.Label(main, text="第二步：基于TOC_template.xlsx，在原项目层面PDT基础上增加对应TFLs，生成最新版PDT。", font=("Microsoft YaHei UI", 10, "bold"), fg="#333333", bg="#f0f0f0", wraplength=1240)
    step2_title.pack(anchor="w", pady=(14, 10))

    default_toc = r"Z:\projects\utility\template\TOC_template.xlsx"
    base_path = gui.get_current_path()
    p3, p4 = (gui.selected_paths[2] or ""), (gui.selected_paths[3] or "")
    default_pdt = os.path.join(base_path, "utility", "documentation", f"{p3}_{p4}_PDT.xlsx" if (p3 or p4) else "项目层面_PDT.xlsx")

    # TOC_template.xlsx
    row_toc = tk.Frame(main, bg="#f0f0f0")
    row_toc.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_toc, text="TOC_template.xlsx：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    toc_entry = tk.Entry(row_toc, width=72, font=("Microsoft YaHei UI", 9))
    toc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    toc_entry.insert(0, default_toc)

    def browse_toc():
        path = filedialog.askopenfilename(title="选择 TOC_template.xlsx", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")], initialdir=os.path.dirname(default_toc))
        if path:
            toc_entry.delete(0, tk.END)
            toc_entry.insert(0, path)

    tk.Button(row_toc, text="浏览...", command=browse_toc, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    # 项目层面PDT.xlsx（加宽以完整展示路径）
    row_pdt = tk.Frame(main, bg="#f0f0f0")
    row_pdt.pack(anchor="w", fill=tk.X, pady=(0, 14))
    tk.Label(row_pdt, text="项目层面PDT.xlsx：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    pdt_entry = tk.Entry(row_pdt, width=72, font=("Microsoft YaHei UI", 9))
    pdt_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    pdt_entry.insert(0, default_pdt)

    def browse_pdt():
        path = filedialog.askopenfilename(title="选择 项目层面PDT.xlsx", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")], initialdir=os.path.dirname(default_pdt) or base_path)
        if path:
            pdt_entry.delete(0, tk.END)
            pdt_entry.insert(0, path)

    tk.Button(row_pdt, text="浏览...", command=browse_pdt, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    # 确定 / 打开并编辑
    btn_frame = tk.Frame(main, bg="#f0f0f0")
    btn_frame.pack(anchor="w", pady=(8, 0))

    def on_ok():
        design_types = [k for k, v in q1_vars.items() if v.get()]
        endpoints = [k for k, v in q2_vars.items() if v.get()]
        analyte_names = q3_entry.get().strip()
        toc_path = toc_entry.get().strip()
        pdt_path = pdt_entry.get().strip()
        setup_path = os.path.join(os.path.dirname(pdt_path), "setup.xlsx")

        if not os.path.isfile(pdt_path):
            messagebox.showerror("错误", f"PDT 文件不存在或无法访问：{pdt_path}")
            return
        if not os.path.isfile(toc_path):
            messagebox.showerror("错误", f"TOC 文件不存在或无法访问：{toc_path}")
            return
        if not os.path.isfile(setup_path):
            messagebox.showerror("错误", f"setup.xlsx 不存在或无法访问：{setup_path}\n（应与 PDT 同目录）")
            return
        if not design_types:
            messagebox.showwarning("提示", "请至少选择一种分析设计类型")
            return

        try:
            from tfls_pdt_update import update_pdt_deliverables
            success, msg = update_pdt_deliverables(
                pdt_path, toc_path, setup_path,
                design_types, endpoints, analyte_names or None
            )
        except Exception as e:
            messagebox.showerror("错误", f"更新失败：{e}")
            return

        if success:
            messagebox.showinfo("成功", msg)
            gui.update_status(msg)
        else:
            messagebox.showerror("错误", msg)

    def on_open_edit():
        pdt_path = pdt_entry.get().strip()
        if pdt_path and os.path.isfile(pdt_path):
            try:
                os.startfile(pdt_path)
                gui.update_status("已打开并编辑: " + os.path.basename(pdt_path))
            except Exception as e:
                messagebox.showerror("错误", "无法打开文件: %s" % e)
        else:
            messagebox.showwarning("提示", "请先选择有效的项目层面PDT.xlsx 路径，或先点击「确定」生成后再打开。")

    tk.Button(btn_frame, text="确定", command=on_ok, width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_frame, text="打开并编辑", command=on_open_edit, width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)
