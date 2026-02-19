# -*- coding: utf-8 -*-
"""
TFLs 页面 - Metadata Setup 弹窗逻辑（独立模块）

主界面在 TFLs 页面提供「Metadata Setup」按钮，绑定 command=lambda: show_metadata_setup_dialog(gui)。
第一步：受试者分布 T14_1-1_1.xlsx 初始化设置（按 Meta_Data表格制作流程：01/04/05/06 四部分）。
第二步：分析集 XXXX 初始化（从 Word 文档「分析集」章节解析小标题与内容，写入 Excel）。
"""
import os
import re
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog


# ---------- 第一步：受试者分布 T14_1-1_1 数据解析与生成 ----------

# 01部分固定文本与列取值（第1-2行，见 Meta_Data表格制作流程.md）
_T14_01_ROW1 = "筛选受试者"
_T14_01_ROW2 = "筛选失败受试者"
_T14_01_SEC = "01_scr"
_T14_01_DSNIN = "adsl"
_T14_01_TRTSUBN = "trt01pn"
_T14_01_TRTSUBC = "trt01p"
_T14_01_FILTER_ROW1 = "prxmatch('/^(合计|total)\\s*$/i', trt01p)"
_T14_01_FILTER_ROW2 = "prxmatch('/^(合计|total)\\s*$/i', trt01p) and (scfailfl='Y')"
# 04部分固定三行（仅当存在 RANDFL 时）
_T14_04_ROWS = ("随机化", "随机接受研究治疗", "随机未接受研究治疗")
# 05部分
_T14_05_HEADER = "完成研究治疗汇总"
_T14_05_TITLE = "终止研究治疗"
# 06部分
_T14_06_COMPLETE = "完成随访"
_T14_06_WITHDRAW = "退出随访"
_T14_06_EXTRA = "随机未接受研究治疗"

# EDCDEF_code 中 治疗结束原因 的 CODE_NAME_CHN 匹配
_EDC_DCTREAS_NAMES = ("治疗结束主要原因", "治疗结束原因")
# EDCDEF_code 中 随访结束原因 的 CODE_NAME_CHN 匹配
_EDC_FOLLOWUP_NAMES = ("随访结束原因", "随访结束主要原因", "研究结束原因", "原因结束主要原因")


def _find_excel_column(df, candidates):
    """在 DataFrame 列名中查找匹配项（忽略大小写、首尾空格）。返回列名或 None。"""
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        k = cand.strip().lower()
        for col_key, col_orig in cols.items():
            if k in col_key or col_key in k:
                return col_orig
    return None


def parse_adam_spec_for_randfl_enrlfl(adam_excel_path):
    """
    从 ADaM 数据集说明 Excel 的 variables sheet 中判断 ADSL 是否存在 RANDFL/ENRLFL。
    检查路径：variables sheet → ADSL 数据集 → RANDFL 且 Study Specific = Y，或 ENRLFL。
    返回: "randfl" | "enrlfl" | None
    """
    try:
        import pandas as pd
    except ImportError:
        raise RuntimeError(
            "请先安装 pandas：pip install pandas\n"
            "若提示权限错误，请以管理员身份打开命令行再执行，或在项目目录使用：python -m venv venv 后激活 venv 再 pip install pandas"
        )

    xl = pd.ExcelFile(adam_excel_path)
    # 查找 variables 或 Variables 等 sheet
    sheet_name = None
    for s in xl.sheet_names:
        if "variable" in s.lower():
            sheet_name = s
            break
    if sheet_name is None:
        raise ValueError("ADaM 说明文件中未找到 variables 相关 sheet。")

    df = pd.read_excel(adam_excel_path, sheet_name=sheet_name, header=0)
    if df.empty:
        raise ValueError("variables sheet 为空。")

    col_dataset = _find_excel_column(df, ("Dataset", "Data Set", "数据集", "Dataset Name"))
    col_var = _find_excel_column(df, ("Variable", "变量", "Variable Name"))
    col_study_spec = _find_excel_column(df, ("Study Specific", "StudySpecific", "Study Specific Flag"))

    if col_dataset is None or col_var is None:
        raise ValueError("variables sheet 中未找到 Dataset 或 Variable 列。")

    # 筛选 ADSL
    ds_col = df[col_dataset].astype(str).str.strip()
    adsl_mask = ds_col.str.upper().str.contains("ADSL", na=False)
    adsl_df = df.loc[adsl_mask]

    if adsl_df.empty:
        return None

    var_col = adsl_df[col_var].astype(str).str.strip()

    # 优先检查 RANDFL 且 Study Specific = Y
    randfl_mask = var_col.str.upper() == "RANDFL"
    if randfl_mask.any():
        if col_study_spec is not None:
            ss = adsl_df.loc[randfl_mask, col_study_spec].astype(str).str.strip().str.upper()
            if (ss == "Y").any():
                return "randfl"
        else:
            return "randfl"

    # 否则检查 ENRLFL
    enrlfl_mask = var_col.str.upper() == "ENRLFL"
    if enrlfl_mask.any():
        return "enrlfl"

    return None


def read_edcdef_code(edc_path):
    """
    读取 EDCDEF_code 数据集（SAS 或 Excel 导出），按 CODE_NAME_CHN 提取 CODE_ORDER、CODE_LABEL。
    返回: dict[str, list[(order, label)]]
    """
    try:
        import pandas as pd
    except ImportError:
        raise RuntimeError(
            "请先安装 pandas：pip install pandas\n"
            "若提示权限错误，请以管理员身份打开命令行再执行，或在项目目录使用：python -m venv venv 后激活 venv 再 pip install pandas"
        )

    if not edc_path or not os.path.isfile(edc_path):
        return {}

    ext = os.path.splitext(edc_path)[1].lower()
    if ext == ".sas7bdat":
        try:
            import pyreadstat
            df, _ = pyreadstat.read_sas7bdat(edc_path)
        except ImportError:
            raise RuntimeError("读取 SAS 数据集需要 pyreadstat：pip install pyreadstat")
    elif ext in (".xlsx", ".xls"):
        df = pd.read_excel(edc_path, header=0)
    else:
        return {}

    if df is None or df.empty:
        return {}

    col_name = _find_excel_column(df, ("CODE_NAME_CHN", "Code_Name_Chn", "code_name_chn"))
    col_order = _find_excel_column(df, ("CODE_ORDER", "Code_Order", "code_order"))
    col_label = _find_excel_column(df, ("CODE_LABEL", "Code_Label", "code_label"))

    if col_name is None or col_label is None:
        return {}

    result = {}
    for _, row in df.iterrows():
        name = str(row.get(col_name, "") or "").strip()
        if not name:
            continue
        order_val = row.get(col_order) if col_order else 0
        try:
            order_val = float(order_val) if order_val is not None and str(order_val).strip() else 0
        except (ValueError, TypeError):
            order_val = 0
        label = str(row.get(col_label, "") or "").strip()
        if name not in result:
            result[name] = []
        result[name].append((order_val, label))

    for k in result:
        result[k].sort(key=lambda x: x[0])

    return result


def _get_dctreas_reasons(edc_data):
    """从 EDCDEF 中提取治疗结束原因列表（按 CODE_ORDER 排序）。"""
    for k, items in edc_data.items():
        for name in _EDC_DCTREAS_NAMES:
            if name in k or k in name:
                return [lb for _, lb in items]
    return []


def _get_followup_reasons(edc_data):
    """从 EDCDEF 中提取随访结束原因列表（按 CODE_ORDER 排序）。"""
    for k, items in edc_data.items():
        for name in _EDC_FOLLOWUP_NAMES:
            if name in k or k in name:
                return [lb for _, lb in items]
    return []


def build_t14_1_1_1_rows(randfl_or_enrlfl, dct_reasons, followup_reasons):
    """
    按 Meta_Data 流程构建 T14_1-1_1 受试者分布的所有行。
    randfl_or_enrlfl: "randfl" | "enrlfl" | None
    dct_reasons: 治疗结束原因列表
    followup_reasons: 随访结束原因列表
    返回: list of dict with keys: TEXT, ROW, MASK, LINE_BREAK, INDENT, SEC, TRT_I, DSNIN, TRTSUBN, TRTSUBC, FILTER, FOOTNOTE
    """
    def _empty_meta():
        return {"SEC": "", "TRT_I": "", "DSNIN": "", "TRTSUBN": "", "TRTSUBC": ""}

    rows = []
    row_num = 0

    # 01部分（第1-2行按 Meta_Data 流程：SEC/DSNIN/TRTSUBN/TRTSUBC/FILTER）
    row_num += 1
    rows.append({
        "TEXT": _T14_01_ROW1, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "",
        "SEC": _T14_01_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC,
        "FILTER": _T14_01_FILTER_ROW1, "FOOTNOTE": "",
    })
    row_num += 1
    rows.append({
        "TEXT": _T14_01_ROW2, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "",
        "SEC": _T14_01_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC,
        "FILTER": _T14_01_FILTER_ROW2, "FOOTNOTE": "",
    })
    if randfl_or_enrlfl == "randfl":
        text_3 = "筛选成功为随机受试者"
    elif randfl_or_enrlfl == "enrlfl":
        text_3 = "筛选成功为入组受试者"
    else:
        text_3 = "筛选成功为随机受试者"  # 默认
    row_num += 1
    em = _empty_meta()
    rows.append({"TEXT": text_3, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "", **em, "FILTER": "", "FOOTNOTE": ""})

    # 04部分（仅当存在 RANDFL 时）
    if randfl_or_enrlfl == "randfl":
        for t in _T14_04_ROWS:
            row_num += 1
            em = _empty_meta()
            rows.append({"TEXT": t, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "", **em, "FILTER": "", "FOOTNOTE": ""})

    # 05部分
    row_num += 1
    em = _empty_meta()
    rows.append({"TEXT": _T14_05_HEADER, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "", **em, "FILTER": "", "FOOTNOTE": ""})
    row_num += 1
    rows.append({"TEXT": _T14_05_TITLE, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "", **_empty_meta(), "FILTER": "", "FOOTNOTE": ""})
    for idx, reason in enumerate(dct_reasons):
        row_num += 1
        rows.append({"TEXT": reason, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "", **_empty_meta(), "FILTER": "DCTREAS=%d" % idx, "FOOTNOTE": ""})

    # 06部分
    row_num += 1
    rows.append({"TEXT": _T14_06_COMPLETE, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "", **_empty_meta(), "FILTER": "", "FOOTNOTE": ""})
    row_num += 1
    rows.append({"TEXT": _T14_06_WITHDRAW, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "", **_empty_meta(), "FILTER": "", "FOOTNOTE": ""})
    for idx, reason in enumerate(followup_reasons):
        row_num += 1
        reason_filter = "退出原因=%d" % idx
        rows.append({"TEXT": reason, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "", **_empty_meta(), "FILTER": reason_filter, "FOOTNOTE": ""})
        row_num += 1
        extra_filter = (reason_filter + " and " if reason_filter else "") + "RANDFL='Y' and TRTSDT NE ."
        rows.append({"TEXT": _T14_06_EXTRA, "ROW": row_num, "MASK": "", "LINE_BREAK": "", "INDENT": "", **_empty_meta(), "FILTER": extra_filter, "FOOTNOTE": ""})

    return rows


def write_t14_1_1_1_xlsx(xlsx_path, rows):
    """将受试者分布行写入 Excel，列：TEXT, ROW, MASK, LINE_BREAK, INDENT, SEC, TRT_I, DSNIN, TRTSUBN, TRTSUBC, FILTER, FOOTNOTE。"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "受试者分布"
    ws.append(["TEXT", "ROW", "MASK", "LINE_BREAK", "INDENT", "SEC", "TRT_I", "DSNIN", "TRTSUBN", "TRTSUBC", "FILTER", "FOOTNOTE"])
    for r in rows:
        ws.append([
            r.get("TEXT", ""), r.get("ROW", 0), r.get("MASK", ""), r.get("LINE_BREAK", ""), r.get("INDENT", ""),
            r.get("SEC", ""), r.get("TRT_I", ""), r.get("DSNIN", ""), r.get("TRTSUBN", ""), r.get("TRTSUBC", ""),
            r.get("FILTER", ""), r.get("FOOTNOTE", ""),
        ])
    d = os.path.dirname(xlsx_path)
    if d:
        os.makedirs(d, exist_ok=True)
    wb.save(xlsx_path)


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
    dlg.geometry("1200x520")
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
    default_adam_dir = os.path.join(base_path, "utility", "documents")
    if not os.path.isdir(default_adam_dir):
        default_adam_dir = os.path.join(base_path, "utility", "documentation")
    default_edc_dir = os.path.join(base_path, "utility", "metadata")
    if not os.path.isdir(default_edc_dir):
        default_edc_dir = os.path.join(base_path, "metadata")

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

    row_adam = tk.Frame(main, bg="#f0f0f0")
    row_adam.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_adam, text="ADaM 数据集说明（Excel）：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    adam_entry = tk.Entry(row_adam, width=72, font=("Microsoft YaHei UI", 9))
    adam_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

    def browse_adam():
        path = filedialog.askopenfilename(
            title="选择 ADaM 数据集说明文件（Excel）",
            filetypes=[("Excel", "*.xlsx"), ("Excel 97", "*.xls"), ("All", "*.*")],
            initialdir=default_adam_dir if os.path.isdir(default_adam_dir) else base_path,
        )
        if path:
            adam_entry.delete(0, tk.END)
            adam_entry.insert(0, path)

    tk.Button(row_adam, text="浏览...", command=browse_adam, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    row_edc = tk.Frame(main, bg="#f0f0f0")
    row_edc.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_edc, text="EDCDEF_code（SAS/Excel）：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    edc_entry = tk.Entry(row_edc, width=72, font=("Microsoft YaHei UI", 9))
    edc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    default_edc = os.path.join(default_edc_dir, "EDCDEF_code.sas7bdat")
    if not os.path.isfile(default_edc):
        default_edc = os.path.join(default_edc_dir, "EDCDEF_code.xlsx")
    edc_entry.insert(0, default_edc if os.path.isfile(default_edc) else "")

    def browse_edc():
        path = filedialog.askopenfilename(
            title="选择 EDCDEF_code（.sas7bdat 或 .xlsx）",
            filetypes=[("SAS 数据集", "*.sas7bdat"), ("Excel", "*.xlsx"), ("All", "*.*")],
            initialdir=default_edc_dir if os.path.isdir(default_edc_dir) else base_path,
        )
        if path:
            edc_entry.delete(0, tk.END)
            edc_entry.insert(0, path)

    tk.Button(row_edc, text="浏览...", command=browse_edc, width=8, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

    btn_frame = tk.Frame(main, bg="#f0f0f0")
    btn_frame.pack(anchor="w", pady=(14, 0))

    def run_init_t14():
        """初版T14_1-1_1：按 Meta_Data 流程生成 01/04/05/06 四部分。若文件已存在则备份后覆盖。"""
        path = t14_entry.get().strip()
        if not path:
            messagebox.showwarning("提示", "请填写或选择 T14_1-1_1.xlsx 路径。")
            return
        adam_path = adam_entry.get().strip()
        edc_path = edc_entry.get().strip()

        randfl_or_enrlfl = None
        if adam_path and os.path.isfile(adam_path):
            try:
                randfl_or_enrlfl = parse_adam_spec_for_randfl_enrlfl(adam_path)
                gui.update_status("ADaM 解析：%s" % (randfl_or_enrlfl or "未检测到 RANDFL/ENRLFL"))
            except Exception as e:
                messagebox.showwarning("ADaM 解析", "无法解析 ADaM 说明文件，将使用默认（随机受试者）：%s" % e)
                randfl_or_enrlfl = "randfl"
        else:
            randfl_or_enrlfl = "randfl"

        edc_data = {}
        if edc_path and os.path.isfile(edc_path):
            try:
                edc_data = read_edcdef_code(edc_path)
                gui.update_status("EDCDEF 已读取")
            except Exception as e:
                messagebox.showwarning("EDCDEF 读取", "无法读取 EDCDEF_code，05/06 部分将为空：%s" % e)

        dct_reasons = _get_dctreas_reasons(edc_data)
        followup_reasons = _get_followup_reasons(edc_data)

        try:
            if os.path.isfile(path):
                backup_path = _backup_existing_to_archive(path)
                if backup_path:
                    gui.update_status("已备份原文件至：%s" % backup_path)
            d = os.path.dirname(path)
            if d:
                os.makedirs(d, exist_ok=True)
            rows = build_t14_1_1_1_rows(randfl_or_enrlfl, dct_reasons, followup_reasons)
            write_t14_1_1_1_xlsx(path, rows)
            gui.update_status("已初始化 T14_1-1_1.xlsx（共 %d 行）：%s" % (len(rows), path))
            if messagebox.askyesno("成功", "已生成初版 T14_1-1_1.xlsx（01/04/05/06 四部分，共 %d 行）。\n\n是否审阅并打开生成文件？" % len(rows)):
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
