# -*- coding: utf-8 -*-
"""
TFLs 页面 - Metadata Setup 弹窗逻辑（独立模块）

主界面在 TFLs 页面提供「Metadata Setup」按钮，绑定 command=lambda: show_metadata_setup_dialog(gui)。
"""
import os
import tkinter as tk
from tkinter import messagebox, filedialog


def show_metadata_setup_dialog(gui):
    """
    显示「Metadata Setup」弹窗。
    第一步：受试者分布 T14_1-1_1.xlsx 初始化设置。
    gui: 主窗口实例，需有 .root, .get_current_path(), .update_status()
    """
    dlg = tk.Toplevel(gui.root)
    dlg.title("Metadata Setup")
    dlg.geometry("1200x280")  # 宽度为原来的 1.5 倍 (800*1.5)
    dlg.resizable(True, True)
    dlg.transient(gui.root)
    dlg.grab_set()
    dlg.configure(bg="#f0f0f0")

    main = tk.Frame(dlg, padx=20, pady=16, bg="#f0f0f0")
    main.pack(fill=tk.BOTH, expand=True)

    # ---------- 第一步：受试者分布 T14_1-1_1.xlsx 初始化设置 ----------
    step1_title = tk.Label(
        main,
        text="第一步：受试者分布 T14_1-1_1.xlsx 初始化设置",
        font=("Microsoft YaHei UI", 10, "bold"),
        fg="#333333",
        bg="#f0f0f0",
    )
    step1_title.pack(anchor="w", pady=(0, 10))

    # 前四个下拉框拼接路径（与 PDT Gen 一致）
    base_path = gui.get_current_path()
    default_t14 = os.path.join(base_path, "utility", "metadata", "T14_1-1_1.xlsx")

    row_t14 = tk.Frame(main, bg="#f0f0f0")
    row_t14.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_t14, text="T14_1-1_1.xlsx：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    t14_entry = tk.Entry(row_t14, width=72, font=("Microsoft YaHei UI", 9))
    t14_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    t14_entry.insert(0, default_t14)

    def browse_t14():
        path = filedialog.askopenfilename(
            title="选择 T14_1-1_1.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("All", "*.*")],
            initialdir=os.path.dirname(default_t14) or base_path,
        )
        if path:
            t14_entry.delete(0, tk.END)
            t14_entry.insert(0, path)

    tk.Button(row_t14, text="浏览...", command=browse_t14, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    btn_frame = tk.Frame(main, bg="#f0f0f0")
    btn_frame.pack(anchor="w", pady=(14, 0))

    def run_init_t14():
        """初版T14_1-1_1：若文件不存在则创建 utility\\metadata 目录并生成空白 T14_1-1_1.xlsx。"""
        path = t14_entry.get().strip()
        if not path:
            messagebox.showwarning("提示", "请填写或选择 T14_1-1_1.xlsx 路径。")
            return
        if os.path.isfile(path):
            messagebox.showinfo("提示", "文件已存在，可直接点击「更新」打开审阅。\n" + path)
            return
        try:
            from openpyxl import Workbook
            d = os.path.dirname(path)
            if d:
                os.makedirs(d, exist_ok=True)
            wb = Workbook()
            wb.active.title = "受试者分布"
            wb.save(path)
            gui.update_status("已初始化 T14_1-1_1.xlsx：%s" % path)
            messagebox.showinfo("成功", "已生成初版 T14_1-1_1.xlsx：\n" + path)
        except Exception as e:
            messagebox.showerror("错误", "初始化失败：%s" % e)

    def on_open_t14():
        p = t14_entry.get().strip()
        if p and os.path.isfile(p):
            try:
                os.startfile(p)
                gui.update_status("已打开: " + os.path.basename(p))
            except Exception as e:
                messagebox.showerror("错误", "无法打开文件: %s" % e)
        else:
            messagebox.showwarning("提示", "请先选择有效的 T14_1-1_1.xlsx 路径，或先点击「初版T14_1-1_1」生成后再打开。")

    tk.Button(btn_frame, text="初版T14_1-1_1", command=run_init_t14, width=14, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_frame, text="更新", command=on_open_t14, width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)
