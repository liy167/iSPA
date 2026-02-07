# -*- coding: utf-8 -*-
"""
TFLs 页面 - 生成PDT 弹窗逻辑（独立模块，避免 SASEG_GUI 过于庞大）
"""
import os
import subprocess
import tkinter as tk
from tkinter import messagebox, filedialog


def show_pdt_dialog(gui):
    """
    显示「生成PDT」分步弹窗。
    gui: 主窗口实例，需有 .root, .get_current_path(), .selected_paths, .update_status()
    """
    dlg = tk.Toplevel(gui.root)
    dlg.title("生成PDT")
    dlg.geometry("1320x810")  # 宽度 1320，高度为原 540 的 1.5 倍，确保三步均可完整展示
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

    base_path = gui.get_current_path()
    p3, p4 = (gui.selected_paths[2] or ""), (gui.selected_paths[3] or "")
    default_pdt = os.path.join(base_path, "utility", "documentation", f"{p3}_{p4}_PDT.xlsx" if (p3 or p4) else "项目层面_PDT.xlsx")
    # 第二步：项目层面 TOC 默认路径 = 前4个下拉框拼接 + \utility\documentation\03_statistics\TOC.xlsx
    default_toc_study = os.path.join(base_path, "utility", "documentation", "03_statistics", "TOC.xlsx")
    default_toc_template = r"Z:\projects\utility\template\TOC_template.xlsx"

    # ---------- 第二步 ----------
    step2_title = tk.Label(main, text="第二步：基于TOC_template.xlsx，生成TOC.xlsx。", font=("Microsoft YaHei UI", 10, "bold"), fg="#333333", bg="#f0f0f0", wraplength=1240)
    step2_title.pack(anchor="w", pady=(14, 10))

    # 第二步：第一个文件 TOC_template.xlsx
    row_toc_template = tk.Frame(main, bg="#f0f0f0")
    row_toc_template.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_toc_template, text="TOC_template.xlsx：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    toc_template_entry = tk.Entry(row_toc_template, width=72, font=("Microsoft YaHei UI", 9))
    toc_template_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    toc_template_entry.insert(0, default_toc_template)

    def browse_toc_template():
        path = filedialog.askopenfilename(title="选择 TOC_template.xlsx", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")], initialdir=os.path.dirname(default_toc_template))
        if path:
            toc_template_entry.delete(0, tk.END)
            toc_template_entry.insert(0, path)

    tk.Button(row_toc_template, text="浏览...", command=browse_toc_template, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    # 第二步：第二个文件 项目层面TOC.xlsx
    row_toc_study = tk.Frame(main, bg="#f0f0f0")
    row_toc_study.pack(anchor="w", fill=tk.X, pady=(0, 14))
    tk.Label(row_toc_study, text="项目层面TOC.xlsx：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    toc_study_entry = tk.Entry(row_toc_study, width=72, font=("Microsoft YaHei UI", 9))
    toc_study_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    toc_study_entry.insert(0, default_toc_study)

    def browse_toc_study():
        path = filedialog.askopenfilename(title="选择 TOC.xlsx", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")], initialdir=os.path.dirname(default_toc_study) or base_path)
        if path:
            toc_study_entry.delete(0, tk.END)
            toc_study_entry.insert(0, path)

    tk.Button(row_toc_study, text="浏览...", command=browse_toc_study, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    # 第二步按钮
    btn_frame = tk.Frame(main, bg="#f0f0f0")
    btn_frame.pack(anchor="w", pady=(8, 0))

    def run_gen_toc_study(template_widget, study_widget):
        """根据前三个问题与 TOC_template，筛选展开后生成 TOC.xlsx（含 TOC sheet：OUTTYPE/OUTREF/OUTTITLE/OUTPOP/OUTNOTE）。"""
        template_path = template_widget.get().strip()
        study_path = study_widget.get().strip()
        if not template_path or not study_path:
            messagebox.showwarning("提示", "请填写 TOC_template 与 TOC 路径。")
            return
        if not os.path.isfile(template_path):
            messagebox.showerror("错误", f"TOC_template 文件不存在或无法访问：{template_path}")
            return

        design_types = [k for k, v in q1_vars.items() if v.get()]
        endpoints = [k for k, v in q2_vars.items() if v.get()]
        analyte_names = q3_entry.get().strip() or None
        if not design_types:
            messagebox.showwarning("提示", "问题1 为必选，请至少选择一种分析设计类型。")
            return
        pk_param_to_conc = {"PK参数(血)": "PK浓度(血)", "PK参数(尿)": "PK浓度(尿)", "PK参数(粪)": "PK浓度(粪)"}
        for param, conc in pk_param_to_conc.items():
            if q2_vars[param].get() and not q2_vars[conc].get():
                messagebox.showwarning("提示", f"选择了「{param}」时，必须同时选择「{conc}」。")
                return

        # setup.xlsx 可选：在 documentation 目录（TOC 的上一级目录）
        setup_path = os.path.join(os.path.dirname(os.path.dirname(study_path)), "setup.xlsx")
        if not os.path.isfile(setup_path):
            setup_path = None

        try:
            from tfls_pdt_gen import gen_toc_study
            success, msg = gen_toc_study(
                template_path, study_path, setup_path,
                design_types, endpoints, analyte_names
            )
        except Exception as e:
            messagebox.showerror("错误", "生成失败：%s" % e)
            return
        if success:
            gui.update_status(msg)
            success_win = tk.Toplevel(dlg)
            success_win.title("成功")
            success_win.transient(dlg)
            success_win.configure(bg="#f0f0f0")
            tk.Label(success_win, text=msg + "\n" + study_path, font=("Microsoft YaHei UI", 9), bg="#f0f0f0", wraplength=420, justify=tk.LEFT).pack(padx=24, pady=(20, 12))
            def close_and_open():
                success_win.destroy()
                if study_path and os.path.isfile(study_path):
                    try:
                        os.startfile(study_path)
                        gui.update_status("已打开审阅: " + os.path.basename(study_path))
                    except Exception as e:
                        messagebox.showerror("错误", "无法打开文件: %s" % e)
            tk.Button(success_win, text="打开审阅", command=close_and_open, width=10, font=("Microsoft YaHei UI", 9)).pack(pady=(0, 20))
        else:
            messagebox.showerror("错误", msg)

    def run_gen(toc_widget, pdt_widget):
        """点击初版PDT：调用 SAS EG 打开并运行 25_generate_pdt_call.sas（路径 = 前四个下拉框 + \\utility\\tools\\25_generate_pdt_call.sas）。"""
        SAS_EG_PATH = r"C:\Program Files\SaS\SASHome\SASEnterpriseGuide\8\SEGuide.exe"
        base_4 = getattr(gui, "z_drive", "Z:\\")
        for i in range(4):
            if getattr(gui, "selected_paths", None) and i < len(gui.selected_paths) and gui.selected_paths[i]:
                base_4 = os.path.join(base_4, gui.selected_paths[i])
        sas_script = os.path.join(base_4, "utility", "tools", "25_generate_pdt_call.sas")
        if not os.path.isfile(sas_script):
            messagebox.showerror("错误", "未找到 SAS 脚本：\n%s" % sas_script)
            return
        if not os.path.isfile(SAS_EG_PATH):
            messagebox.showerror("错误", "未找到 SAS EG：\n%s" % SAS_EG_PATH)
            return
        try:
            subprocess.Popen([SAS_EG_PATH, sas_script], cwd=base_4, shell=False)
            gui.update_status("已调用 SAS EG 打开脚本：25_generate_pdt_call.sas")
            messagebox.showinfo("提示", "已启动 SAS EG 并打开脚本 25_generate_pdt_call.sas，请在 SAS EG 中运行。")
        except Exception as e:
            messagebox.showerror("错误", "调用 SAS EG 时出错：%s" % e)

    def on_open_edit(pdt_widget):
        p = pdt_widget.get().strip()
        if p and os.path.isfile(p):
            try:
                os.startfile(p)
                gui.update_status("已打开: " + os.path.basename(p))
            except Exception as e:
                messagebox.showerror("错误", "无法打开文件: %s" % e)
        else:
            messagebox.showwarning("提示", "请先选择有效的项目层面PDT.xlsx 路径，或先点击「初版PDT」生成后再打开。")

    tk.Button(btn_frame, text="初版TOC", command=lambda: run_gen_toc_study(toc_template_entry, toc_study_entry), width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_frame, text="更新", command=lambda: on_open_edit(toc_study_entry), width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    # ---------- 第三步 ----------
    step3_title = tk.Label(main, text="第三步：基于TOC.xlsx，并在原项目层面PDT基础上添加TFLs行，生成最新版PDT。", font=("Microsoft YaHei UI", 10, "bold"), fg="#333333", bg="#f0f0f0", wraplength=1240)
    step3_title.pack(anchor="w", pady=(20, 10))

    row_toc_s3 = tk.Frame(main, bg="#f0f0f0")
    row_toc_s3.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_toc_s3, text="TOC.xlsx：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    toc_entry_s3 = tk.Entry(row_toc_s3, width=72, font=("Microsoft YaHei UI", 9))
    toc_entry_s3.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    toc_entry_s3.insert(0, default_toc_study)

    def browse_toc_s3():
        path = filedialog.askopenfilename(title="选择 TOC.xlsx", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")], initialdir=os.path.dirname(default_toc_study) or base_path)
        if path:
            toc_entry_s3.delete(0, tk.END)
            toc_entry_s3.insert(0, path)

    tk.Button(row_toc_s3, text="浏览...", command=browse_toc_s3, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    row_pdt_s3 = tk.Frame(main, bg="#f0f0f0")
    row_pdt_s3.pack(anchor="w", fill=tk.X, pady=(0, 14))
    tk.Label(row_pdt_s3, text="项目层面PDT.xlsx：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    pdt_entry_s3 = tk.Entry(row_pdt_s3, width=72, font=("Microsoft YaHei UI", 9))
    pdt_entry_s3.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    pdt_entry_s3.insert(0, default_pdt)

    def browse_pdt_s3():
        path = filedialog.askopenfilename(title="选择 项目层面PDT.xlsx", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")], initialdir=os.path.dirname(default_pdt) or base_path)
        if path:
            pdt_entry_s3.delete(0, tk.END)
            pdt_entry_s3.insert(0, path)

    tk.Button(row_pdt_s3, text="浏览...", command=browse_pdt_s3, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    btn_frame_s3 = tk.Frame(main, bg="#f0f0f0")
    btn_frame_s3.pack(anchor="w", pady=(8, 0))

    tk.Button(btn_frame_s3, text="初版PDT", command=lambda: run_gen(toc_entry_s3, pdt_entry_s3), width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_frame_s3, text="更新", command=lambda: on_open_edit(pdt_entry_s3), width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)
