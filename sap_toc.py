# -*- coding: utf-8 -*-
"""
SAP 页面 - TOC Gen 弹窗逻辑（第一步+第二步：与 PDT Gen 的初版 TOC 一致，无第三步初版PDT）
"""
import os
import tkinter as tk
from tkinter import messagebox, filedialog

from tfls_pdt import gen_toc_study


def show_sap_toc_dialog(gui):
    """
    显示「SAP - TOC Gen」弹窗，仅含第一步（问题1/2/3）与第二步（TOC_template + 项目层面TOC.xlsx，初版TOC）。
    gui: 主窗口实例，需有 .root, .get_current_path(), .selected_paths, .update_status()
    """
    dlg = tk.Toplevel(gui.root)
    dlg.title("TOC Gen")
    dlg.geometry("1320x420")
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

    tk.Label(step1_content, text="问题1  请选择当前分析设计类型（多选框）：", font=("Microsoft YaHei UI", 9), anchor="w", bg="#f0f0f0").pack(anchor="w", pady=(0, 4))
    q1_vars = {}
    q1_frame = tk.Frame(step1_content, bg="#f0f0f0")
    q1_frame.pack(anchor="w", pady=(0, 10))
    for opt in ["SAD", "FE", "MAD", "BE", "MB"]:
        v = tk.BooleanVar(value=False)
        q1_vars[opt] = v
        cb = tk.Checkbutton(q1_frame, text=opt, variable=v, font=("Microsoft YaHei UI", 9), anchor="w", bg="#f0f0f0")
        cb.pack(side=tk.LEFT, padx=(0, 16))

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
    q3_frame.pack_forget()

    base_path = gui.get_current_path()
    default_toc_study = os.path.join(base_path, "utility", "documentation", "03_statistics", "TOC.xlsx")
    default_toc_template = r"Z:\projects\utility\template\TOC_template.xlsx"

    # ---------- 第二步 ----------
    step2_title = tk.Label(main, text="第二步：基于TOC_template.xlsx，生成TOC.xlsx。", font=("Microsoft YaHei UI", 10, "bold"), fg="#333333", bg="#f0f0f0", wraplength=1240)
    step2_title.pack(anchor="w", pady=(14, 10))

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

    btn_frame = tk.Frame(main, bg="#f0f0f0")
    btn_frame.pack(anchor="w", pady=(8, 0))

    def run_gen_toc_study(template_widget, study_widget):
        """根据前三个问题与 TOC_template，筛选展开后生成 TOC.xlsx。"""
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

        setup_path = os.path.join(os.path.dirname(os.path.dirname(study_path)), "setup.xlsx")
        if not os.path.isfile(setup_path):
            setup_path = None

        edcdef_ecrf_path = os.path.join(base_path, "utility", "metadata", "EDCDEF_ecrf.sas7bdat")
        edcdef_code_path = os.path.join(base_path, "utility", "metadata", "EDCDEF_code.sas7bdat")
        if not os.path.isfile(edcdef_code_path):
            edcdef_code_path = os.path.join(base_path, "utility", "metadata", "EDCDEF_code.xlsx")

        try:
            success, msg = gen_toc_study(
                template_path, study_path, setup_path,
                design_types, endpoints, analyte_names,
                edcdef_ecrf_path=edcdef_ecrf_path,
                edcdef_code_path=edcdef_code_path,
            )
        except Exception as e:
            messagebox.showerror("错误", "生成失败：%s" % e)
            return
        if success:
            gui.update_status(msg)
            success_win = tk.Toplevel(dlg)
            success_win.title("成功")
            success_win.geometry("600x175")
            success_win.transient(dlg)
            success_win.resizable(False, False)
            success_win.configure(bg="#eaeaea")
            success_win.attributes("-topmost", True)
            success_win.after(100, lambda: success_win.attributes("-topmost", False))
            success_win.protocol("WM_DELETE_WINDOW", success_win.destroy)

            content = tk.Frame(success_win, bg="#eaeaea", padx=24, pady=20)
            content.pack(fill=tk.BOTH, expand=True)

            tk.Label(content, text=msg, font=("Segoe UI", 10), fg="#000000", bg="#eaeaea", wraplength=420, justify=tk.LEFT).pack(anchor="w", pady=(0, 8))
            tk.Label(content, text=study_path, font=("Segoe UI", 10), fg="#000000", bg="#eaeaea", wraplength=700, justify=tk.LEFT).pack(anchor="w", pady=(0, 8))
            tk.Label(content, text="是否审阅并打开生成文件?", font=("Segoe UI", 10), fg="#000000", bg="#eaeaea", wraplength=420, justify=tk.LEFT).pack(anchor="w", pady=(0, 16))

            def on_yes():
                success_win.destroy()
                if study_path and os.path.isfile(study_path):
                    try:
                        os.startfile(study_path)
                        gui.update_status("已打开审阅: " + os.path.basename(study_path))
                    except Exception as e:
                        messagebox.showerror("错误", "无法打开文件: %s" % e)

            def on_no():
                success_win.destroy()

            btn_frame_win = tk.Frame(content, bg="#eaeaea")
            btn_frame_win.pack(pady=(0, 4))
            tk.Button(btn_frame_win, text="是(Y)", command=on_yes, width=8, font=("Segoe UI", 10),
                      bg="#d8d8d8", fg="#000000", activebackground="#c8c8c8", activeforeground="#000000",
                      relief=tk.RAISED, borderwidth=1, cursor="hand2").pack(side=tk.LEFT, padx=(0, 12))
            tk.Button(btn_frame_win, text="否(N)", command=on_no, width=8, font=("Segoe UI", 10),
                      bg="#d8d8d8", fg="#000000", activebackground="#c8c8c8", activeforeground="#000000",
                      relief=tk.RAISED, borderwidth=1, cursor="hand2").pack(side=tk.LEFT)
            success_win.focus_set()
        else:
            messagebox.showerror("错误", msg)

    def on_open_edit(widget):
        p = widget.get().strip()
        if p and os.path.isfile(p):
            try:
                os.startfile(p)
                gui.update_status("已打开: " + os.path.basename(p))
            except Exception as e:
                messagebox.showerror("错误", "无法打开文件: %s" % e)
        else:
            messagebox.showwarning("提示", "请先选择有效的项目层面TOC.xlsx 路径，或先点击「初版TOC」生成后再打开。")

    tk.Button(btn_frame, text="初版TOC", command=lambda: run_gen_toc_study(toc_template_entry, toc_study_entry), width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_frame, text="编辑", command=lambda: on_open_edit(toc_study_entry), width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    dlg.focus_set()
    toc_template_entry.focus_set()
