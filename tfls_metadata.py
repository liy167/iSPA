# -*- coding: utf-8 -*-
"""
TFLs 页面 - Metadata Setup 弹窗逻辑（独立模块）

主界面在 TFLs 页面提供「Metadata Setup」按钮，绑定 command=lambda: show_metadata_setup_dialog(gui)。
第一步：受试者分布 T14_1-1_1.xlsx 初始化设置。
第二步：分析集 XXXX 初始化（从 Word 文档「分析集」章节解析小标题与内容，写入 Excel）。
"""
import os
import re
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog


# ---------- 第二步：分析集 Word 解析与 Excel 生成 ----------

# 常见小黑点/项目符号字符（段落开头可能带这些）
_BULLET_CHARS = ("\u2022", "\u2023", "\u25E6", "\u2043", "\u2219", "\u00B7", "\u25AA", "\u25CF", "-", "*", "·", "•")
_BULLET_PATTERN = re.compile(r"^[\s\u2022\u2023\u25E6\u2043\u2219\u00B7\u25AA\u25CF\-*\u30FB\u2022]+\s*")


def _is_bullet_paragraph(paragraph):
    """判断段落是否为项目符号（小黑点）列表项：Word 编号/列表格式或段首为项目符号。"""
    try:
        pPr = paragraph._element.pPr
        if pPr is not None and pPr.numPr is not None:
            return True
    except Exception:
        pass
    text = (paragraph.text or "").strip()
    if not text:
        return False
    return text[0] in _BULLET_CHARS or _BULLET_PATTERN.match(text) is not None


def _strip_bullet(text):
    """去掉段首的小黑点、编号和空白，得到小段标题。"""
    if not text:
        return text
    t = text.strip()
    # 去掉开头的项目符号及后续空白
    t = _BULLET_PATTERN.sub("", t).strip()
    # 去掉常见编号前缀如 "1. "、"1) "
    t = re.sub(r"^\d+[\.\)\s]+\s*", "", t)
    return t.strip()


def _is_next_chapter_heading(paragraph):
    """
    判断是否为「分析集」之后的下一章节标题，如是则不应再纳入分析集内容。
    满足其一即视为下一章：① 正文形如 "6. 终点指标"（一级编号，后接非数字）；② 样式为标题 1/Heading 1 且不含「分析集」。
    """
    text = (paragraph.text or "").strip()
    if not text:
        return False
    # ① 一级章节编号：如 "6. 终点指标"（"6. " 后不是数字），不含 "分析集"
    if "分析集" in text:
        return False
    if re.match(r"^\d+\.\s+(?!\d)", text):
        return True
    # ② 标题 1 / Heading 1 样式
    try:
        style_name = (paragraph.style and paragraph.style.name) or ""
        s = style_name.strip()
        if re.match(r"^(Heading|标题)\s*1(\s|$|Char)", s, re.I):
            return True
    except Exception:
        pass
    return False


def parse_analysis_set_from_docx(docx_path):
    """
    从 Word 文档中解析「分析集」章节：按小段标题（小黑点后面的）、内容拆分为多行。
    返回 list of (小段标题, 内容)，均为字符串。
    - 按标题内容「分析集」定位章节（不限定章节号，只要是标题内容含「分析集」即可）。
    - 该章节内：以「小黑点」开头的列表项段落视为小段标题（TEXT）；紧随其后的非列表项段落合并为该条的「内容」。
    - 遇到下一章节标题（如 "6. 终点指标" 或 标题 1 样式且不含「分析集」）即停止，不纳入后续章节内容。
    """
    try:
        from docx import Document
    except ImportError:
        raise RuntimeError("请先安装 python-docx：pip install python-docx")

    doc = Document(docx_path)
    paragraphs = doc.paragraphs
    start_idx = None
    for i, p in enumerate(paragraphs):
        t = (p.text or "").strip()
        if "分析集" in t:
            start_idx = i
            break
    if start_idx is None:
        raise ValueError("文档中未找到标题内容为「分析集」的章节。")

    result = []  # [(小段标题, 内容), ...]
    i = start_idx + 1
    while i < len(paragraphs):
        p = paragraphs[i]
        text = (p.text or "").strip()

        # 遇到下一章节标题则停止，分析集章节后面的内容不纳入
        if _is_next_chapter_heading(p):
            break

        if _is_bullet_paragraph(p) and text:
            sub_title = _strip_bullet(text)
            if not sub_title:
                i += 1
                continue
            content_parts = []
            j = i + 1
            while j < len(paragraphs):
                q = paragraphs[j]
                q_text = (q.text or "").strip()
                if _is_next_chapter_heading(q):
                    break
                if _is_bullet_paragraph(q) and q_text:
                    break
                if q_text:
                    content_parts.append(q_text)
                j += 1
            content = "\n".join(content_parts) if content_parts else ""
            result.append((sub_title, content))
            i = j
        else:
            i += 1

    return result


def _strip_parens(s):
    """去掉字符串中括号及其中的内容，支持中英文括号（）。"""
    if not s:
        return s
    return re.sub(r"[（(].*?[）)]", "", s).strip()


def _backup_existing_to_archive(file_path):
    """
    若 file_path 存在，则复制到同目录下的 99_archive 文件夹，文件名加年月日时分秒后缀。
    返回备份后的路径，若原文件不存在则返回 None。
    """
    if not file_path or not os.path.isfile(file_path):
        return None
    dir_name = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)
    name, ext = os.path.splitext(base_name)
    suffix = datetime.now().strftime("%Y%m%d%H%M%S")
    archive_dir = os.path.join(dir_name, "99_archive")
    os.makedirs(archive_dir, exist_ok=True)
    backup_name = "%s_%s%s" % (name, suffix, ext)
    backup_path = os.path.join(archive_dir, backup_name)
    shutil.copy2(file_path, backup_path)
    return backup_path


def write_analysis_set_xlsx(xlsx_path, rows):
    """
    将 (小段标题, 内容) 列表写入 Excel，与完整表格一致：TEXT、ROW、MASK、LINE_BREAK、INDENT、FILTER、FOOTNOTE。
    TEXT：小标题去掉括号及括号内内容；FOOTNOTE：小标题：内容（无小黑点）。
    rows: list of (小段标题, 内容)
    """
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "分析集"
    ws.append(["TEXT", "ROW", "MASK", "LINE_BREAK", "INDENT", "FILTER", "FOOTNOTE"])
    for row_num, (title, content) in enumerate(rows, start=1):
        text_cell = _strip_parens(title)
        footnote = "%s：%s" % (title, content) if content else "%s：" % title
        ws.append([text_cell, row_num, "", "", "", "", footnote])
    d = os.path.dirname(xlsx_path)
    if d:
        os.makedirs(d, exist_ok=True)
    wb.save(xlsx_path)


def show_metadata_setup_dialog(gui):
    """
    显示「Metadata Setup」弹窗。
    第一步：受试者分布 T14_1-1_1.xlsx 初始化设置。
    gui: 主窗口实例，需有 .root, .get_current_path(), .update_status()
    """
    dlg = tk.Toplevel(gui.root)
    dlg.title("Metadata Setup")
    dlg.geometry("1200x420")
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
            messagebox.showinfo("提示", "文件已存在，可直接点击「编辑」打开并手动编辑。\n" + path)
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
            if messagebox.askyesno("成功", "已生成初版 T14_1-1_1.xlsx。\n\n是否审阅并打开生成文件？"):
                try:
                    os.startfile(path)
                    gui.update_status("已打开: " + os.path.basename(path))
                except Exception as e:
                    messagebox.showerror("错误", "无法打开文件：%s" % e)
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
            messagebox.showwarning("提示", "请先选择有效的 T14_1-1_1.xlsx 路径，或先点击「初版T14_1-1_1」生成后再编辑。")

    tk.Button(btn_frame, text="初版T14_1-1_1", command=run_init_t14, width=14, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_frame, text="编辑", command=on_open_t14, width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    # ---------- 第二步：分析集 T14_1-1_2.xlsx 初始化设置 ----------
    step2_title = tk.Label(
        main,
        text="第二步：分析集 T14_1-1_2.xlsx 初始化设置",
        font=("Microsoft YaHei UI", 10, "bold"),
        fg="#333333",
        bg="#f0f0f0",
    )
    step2_title.pack(anchor="w", pady=(24, 10))

    # SAP 文件（.docx）初始路径：前四个下拉框 + utility\documentation\03_statistics\
    default_sap_dir = os.path.join(base_path, "utility", "documentation", "03_statistics")
    default_metadata_dir = os.path.join(base_path, "utility", "metadata")
    default_xlsx_step2 = os.path.join(default_metadata_dir, "T14_1-1_2.xlsx")

    row_sap = tk.Frame(main, bg="#f0f0f0")
    row_sap.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_sap, text="SAP 文件（.docx）：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    sap_entry = tk.Entry(row_sap, width=72, font=("Microsoft YaHei UI", 9))
    sap_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

    def browse_sap():
        path = filedialog.askopenfilename(
            title="选择包含「分析集」章节的 SAP 文档（.docx）",
            filetypes=[("SAP 文档", "*.docx"), ("All", "*.*")],
            initialdir=default_sap_dir if os.path.isdir(default_sap_dir) else base_path,
        )
        if path:
            sap_entry.delete(0, tk.END)
            sap_entry.insert(0, path)

    tk.Button(row_sap, text="浏览...", command=browse_sap, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    row_xlsx = tk.Frame(main, bg="#f0f0f0")
    row_xlsx.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_xlsx, text="T14_1-1_2.xlsx：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    xlsx_entry = tk.Entry(row_xlsx, width=72, font=("Microsoft YaHei UI", 9))
    xlsx_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    xlsx_entry.insert(0, default_xlsx_step2)

    def browse_xlsx():
        path = filedialog.asksaveasfilename(
            title="选择 T14_1-1_2.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("All", "*.*")],
            initialdir=os.path.dirname(default_xlsx_step2) or base_path,
            defaultextension=".xlsx",
        )
        if path:
            xlsx_entry.delete(0, tk.END)
            xlsx_entry.insert(0, path)

    tk.Button(row_xlsx, text="浏览...", command=browse_xlsx, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    btn_frame2 = tk.Frame(main, bg="#f0f0f0")
    btn_frame2.pack(anchor="w", pady=(14, 0))

    def run_init_t14_1_1_2():
        """初版T14_1-1_2：从 SAP 文档解析「分析集」章节并生成 T14_1-1_2.xlsx。"""
        sap_path = sap_entry.get().strip()
        xlsx_path = xlsx_entry.get().strip()
        if not sap_path:
            messagebox.showwarning("提示", "请选择包含「分析集」章节的 SAP 文档（.docx）。")
            return
        if not os.path.isfile(sap_path):
            messagebox.showerror("错误", "SAP 文件不存在：%s" % sap_path)
            return
        if not xlsx_path:
            messagebox.showwarning("提示", "请填写或选择 T14_1-1_2.xlsx 路径。")
            return
        try:
            if os.path.isfile(xlsx_path):
                backup_path = _backup_existing_to_archive(xlsx_path)
                if backup_path:
                    gui.update_status("已备份原文件至：%s" % backup_path)
            rows = parse_analysis_set_from_docx(sap_path)
            if not rows:
                messagebox.showwarning("提示", "未在「分析集」章节下解析到任何小段标题与内容。请确认文档中该章节内的小段标题为「小黑点」列表项（项目符号）。")
                return
            write_analysis_set_xlsx(xlsx_path, rows)
            gui.update_status("已初始化 T14_1-1_2.xlsx：%s" % xlsx_path)
            if messagebox.askyesno("成功", "已生成初版 T14_1-1_2.xlsx（共 %d 条）。\n\n是否审阅并打开生成文件？" % len(rows)):
                try:
                    os.startfile(xlsx_path)
                    gui.update_status("已打开: " + os.path.basename(xlsx_path))
                except Exception as e:
                    messagebox.showerror("错误", "无法打开文件：%s" % e)
        except Exception as e:
            messagebox.showerror("错误", "解析或生成失败：%s" % e)
            gui.update_status("分析集初始化失败：%s" % e)

    def on_open_t14_1_1_2():
        p = xlsx_entry.get().strip()
        if p and os.path.isfile(p):
            try:
                os.startfile(p)
                gui.update_status("已打开: " + os.path.basename(p))
            except Exception as e:
                messagebox.showerror("错误", "无法打开文件: %s" % e)
        else:
            messagebox.showwarning("提示", "请先选择有效的 T14_1-1_2.xlsx 路径，或先点击「初版T14_1-1_2」生成后再编辑。")

    tk.Button(btn_frame2, text="初版T14_1-1_2", command=run_init_t14_1_1_2, width=14, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_frame2, text="编辑", command=on_open_t14_1_1_2, width=10, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)
