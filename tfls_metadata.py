# -*- coding: utf-8 -*-
"""
TFLs 页面 - Metadata Setup 弹窗逻辑（独立模块）

主界面在 TFLs 页面提供「Metadata Setup」按钮，绑定 command=lambda: show_metadata_setup_dialog(gui)。
第一步：受试者分布 T14_1-1_1.xlsx 初始化设置（按 Meta_Data表格制作流程：01/04/05/06 四部分）。
第二步：分析集 XXXX 初始化（从 Word 文档「分析集」章节解析小标题与内容，写入 Excel）。
"""
import logging
import os
import re
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog

logger = logging.getLogger(__name__)


def _setup_log_file():
    """将本模块日志同时输出到本地文件（logs/tfls_metadata.log），仅添加一次。"""
    if any(isinstance(h, logging.FileHandler) for h in logger.handlers):
        return
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    try:
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, "tfls_metadata.log")
        fh = logging.FileHandler(log_path, mode="a", encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
        logger.addHandler(fh)
        logger.setLevel(logging.DEBUG)
    except Exception as e:
        logger.debug("无法创建日志文件 %s: %s", log_dir, e)


_setup_log_file()


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
# 第3行之前一行（第10行）：根据 RANDFL/ENRLFL 二选一
_T14_01_ROW_BEFORE_ROW3_RANDFL = "筛选成功未随机受试者"
_T14_01_FILTER_ROW_BEFORE_ROW3_RANDFL = "prxmatch('/^(合计|total)\\s*$/i', trt01p) and (scfailfl='N') and randfl='N'"
_T14_01_ROW_BEFORE_ROW3_ENRLFL = "筛选成功未入组受试者"
_T14_01_FILTER_ROW_BEFORE_ROW3_ENRLFL = "prxmatch('/^(合计|total)\\s*$/i', trt01p) and (scfailfl='N') and enrlfl='N'"
# 04部分：第10行后/第14行后各 3 行（RANDFL 与 ENRLFL 可同时存在）
_T14_04_SEC = "04_rnd"
_T14_04_ROWS_RANDFL = ("随机受试者", "随机未接受研究治疗", "随机且接受研究治疗")
_T14_04_FILTERS_RANDFL = ("randfl='Y' and scfailfl='N'", "randfl='Y' and scfailfl='N' and saffl='N'", "randfl='Y' and scfailfl='N' and saffl='Y'")
_T14_04_ROWS_ENRLFL = ("入组受试者", "入组未接受研究治疗", "入组且接受研究治疗")
_T14_04_FILTERS_ENRLFL = ("enrlfl='Y' and scfailfl='N'", "enrlfl='Y' and scfailfl='N' and saffl='N'", "enrlfl='Y' and scfailfl='N' and saffl='Y'")
# 05部分（完成/终止研究治疗，前 3 行 TEXT 由变量标签动态生成）
# 规则：从变量标签中保留「治疗」两字作为 base（去掉其他字样），第1行=「完成研究」+base，第2行=「终止研究」+base，第3行=「终止研究」+base+「原因」
_T14_05_SEC = "05_trt"
_T14_05_DATASET = "ADSL"          # 宏：Variables sheet 中用于查找治疗结束状态的数据集名，可配置
_T14_05_VAR_EOTSTT = "EOTSTT"     # 宏：治疗结束状态变量名前缀，可配置；匹配以该前缀开头的变量（如 EOTSTT、EOTSTT1、EOTSTT2）
_T14_05_BASE_KEEP = "治疗"        # 从变量标签中保留此二字作为 base，去掉其余字样（不固定后缀）
_T14_05_PREFIX_COMPLETE = "完成研究"
_T14_05_PREFIX_TERMINATE = "终止研究"
_T14_05_DEFAULT_LABEL = "治疗结束状态"  # 未从 ADaM 解析到时的默认标签
_T14_05_EXCLUDE_REASONS = ("已完成",)    # 05 部分原因行中排除这些（不输出对应行），可配置
_T14_06_EXCLUDE_REASONS = ("已完成",)    # 06 部分随访原因中排除这些（不输出对应行），可配置
# 06部分（前3行强制赋值，后面每个随访原因一行+隔行增加「随机未接受研究治疗」）
_T14_06_SEC = "06_fup"
_T14_06_ROW1 = "完成研究"
_T14_06_FILTER_ROW1 = "saffl='Y' and EOSSTT='完成研究'"
_T14_06_ROW2 = "退出研究"
_T14_06_FILTER_ROW2 = "saffl='Y' and EOSSTT='退出研究'"
_T14_06_ROW3 = "退出研究原因"
_T14_06_EXTRA = "随机未接受研究治疗"
_T14_06_EXTRA_SUFFIX = " and randfl='Y' and saffl='N'"  # 隔行「随机未接受研究治疗」在原因 FILTER 基础上追加

# EDCDEF_code 中 治疗结束原因 的 CODE_NAME_CHN 匹配
_EDC_DCTREAS_NAMES = ("治疗结束主要原因", "治疗结束原因")
# EDCDEF_code 中 随访结束原因 的 CODE_NAME_CHN 匹配
_EDC_FOLLOWUP_NAMES = ("随访结束原因", "随访结束主要原因", "研究结束原因", "原因结束主要原因")
# EDCDEF_code 中 筛选失败原因 的 CODE_NAME_CHN 匹配（TEXT=CODE_LABEL，顺序=CODE_ORDER）
# 支持「筛选结束原因」「筛选失败原因」「筛选原因」等常见命名（匹配时列名优先，其次为行值）
_EDC_SCREEN_FAIL_NAMES = ("筛选结束原因", "筛选失败原因", "筛选原因")

# 所有“原因”类名称，用于 read_edcdef_code 中按列名匹配（列名=表头，非行内 CODE_NAME_CHN 值）
_EDC_ALL_REASON_NAMES = set(_EDC_DCTREAS_NAMES) | set(_EDC_FOLLOWUP_NAMES) | set(_EDC_SCREEN_FAIL_NAMES)


def _find_excel_column(df, candidates):
    """在 DataFrame 列名中查找匹配项（忽略大小写、首尾空格）。返回列名或 None。
    优先精确匹配，避免短列名误匹配（如查找 CODE_NAME_CHN 时不能匹配成 CODE_NAME）。"""
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        k = cand.strip().lower()
        for col_key, col_orig in cols.items():
            if col_key == k:
                return col_orig
        for col_key, col_orig in cols.items():
            # 子串匹配时要求列名不少于候选长度，避免 CODE_NAME 匹配到 CODE_NAME_CHN 的查找
            if k in col_key and len(col_key) >= len(k):
                return col_orig
            if col_key in k and len(col_key) >= len(k):
                return col_orig
    return None


def parse_adam_spec_for_randfl_enrlfl(adam_excel_path):
    """
    从 ADaM 数据集说明 Excel 的 variables sheet 中判断 ADSL 是否存在 RANDFL/ENRLFL。
    检查路径：variables sheet → ADSL 数据集 → RANDFL 且 Study Specific = Y，或 ENRLFL。
    若既有 RANDFL 也有 ENRLFL 则都返回，顺序为 RANDFL 先、ENRLFL 后。
    返回: tuple of "randfl" 和/或 "enrlfl"，如 ("randfl",)、("enrlfl",)、("randfl", "enrlfl")；均不存在时返回 ("randfl",) 作为默认。
    详细日志写入 logs/tfls_metadata.log，便于排查为何出现 ENRLFL 块。
    """
    _setup_log_file()
    logger.info("[RANDFL/ENRLFL] 开始解析 ADaM 说明文件：%s", adam_excel_path)
    try:
        import pandas as pd
    except ImportError:
        raise RuntimeError(
            "请先安装 pandas：pip install pandas\n"
            "若提示权限错误，请以管理员身份打开命令行再执行，或在项目目录使用：python -m venv venv 后激活 venv 再 pip install pandas"
        )

    xl = pd.ExcelFile(adam_excel_path)
    sheet_name = None
    for s in xl.sheet_names:
        if "variable" in s.lower():
            sheet_name = s
            break
    if sheet_name is None:
        logger.warning("[RANDFL/ENRLFL] 未找到 variables 相关 sheet，sheet 列表：%s", xl.sheet_names)
        raise ValueError("ADaM 说明文件中未找到 variables 相关 sheet。")

    logger.info("[RANDFL/ENRLFL] 使用 sheet：%s", sheet_name)
    df = pd.read_excel(adam_excel_path, sheet_name=sheet_name, header=0)
    if df.empty:
        raise ValueError("variables sheet 为空。")

    col_dataset = _find_excel_column(df, ("Dataset", "Data Set", "数据集", "Dataset Name"))
    col_var = _find_excel_column(df, ("Variable", "变量", "Variable Name"))
    col_study_spec = _find_excel_column(df, ("Study Specific", "StudySpecific", "Study Specific Flag"))
    logger.info("[RANDFL/ENRLFL] 列名映射：Dataset=%s, Variable=%s, Study Specific=%s",
                col_dataset, col_var, col_study_spec)

    if col_dataset is None or col_var is None:
        raise ValueError("variables sheet 中未找到 Dataset 或 Variable 列。")

    ds_col = df[col_dataset].astype(str).str.strip()
    # 仅匹配 Dataset 列等于 "ADSL" 的行，避免含 "ADSL" 子串的其它数据集（如 ADSL_SUPP）导致误判 ENRLFL
    adsl_mask = ds_col.str.upper() == "ADSL"
    adsl_df = df.loc[adsl_mask]

    adsl_dataset_values = ds_col.loc[adsl_mask].unique().tolist()
    logger.info("[RANDFL/ENRLFL] 筛选条件：Dataset 列等于 'ADSL'（精确）；匹配到的行数=%d，Dataset 取值：%s",
                len(adsl_df), adsl_dataset_values)

    if adsl_df.empty:
        logger.warning("[RANDFL/ENRLFL] 无 ADSL 行，返回默认 ('randfl',)")
        return ("randfl",)  # 默认

    var_col = adsl_df[col_var].astype(str).str.strip()
    # 记录 ADSL 下 Variable 列的全部取值（用于核对是否含 RANDFL/ENRLFL）
    unique_vars = sorted(var_col.unique().tolist())
    logger.info("[RANDFL/ENRLFL] ADSL 行数=%d，Variable 列唯一值（共 %d 个）：%s",
                len(adsl_df), len(unique_vars), unique_vars)

    out = []

    # 检查 RANDFL 且 Study Specific = Y
    randfl_mask = var_col.str.upper() == "RANDFL"
    randfl_found = randfl_mask.any()
    if randfl_found:
        if col_study_spec is not None:
            ss = adsl_df.loc[randfl_mask, col_study_spec].astype(str).str.strip().str.upper()
            study_spec_y = (ss == "Y").any()
            logger.info("[RANDFL/ENRLFL] 检测到 RANDFL 行，Study Specific 列含 Y：%s", study_spec_y)
            if study_spec_y:
                out.append("randfl")
        else:
            logger.info("[RANDFL/ENRLFL] 检测到 RANDFL 行，无 Study Specific 列，按存在即采纳")
            out.append("randfl")
    else:
        logger.info("[RANDFL/ENRLFL] 未在 ADSL 的 Variable 列中找到 RANDFL")

    # 检查 ENRLFL 且 Study Specific = Y（与 RANDFL 一致，均需 Study Specific 列=Y）
    enrlfl_mask = var_col.str.upper() == "ENRLFL"
    enrlfl_found = enrlfl_mask.any()
    if enrlfl_found:
        if col_study_spec is not None:
            ss_enrl = adsl_df.loc[enrlfl_mask, col_study_spec].astype(str).str.strip().str.upper()
            enrlfl_study_spec_y = (ss_enrl == "Y").any()
            logger.info("[RANDFL/ENRLFL] 检测到 ENRLFL 行，Study Specific 列含 Y：%s（若为 True 则会输出第14-17行）", enrlfl_study_spec_y)
            if enrlfl_study_spec_y:
                out.append("enrlfl")
        else:
            logger.info("[RANDFL/ENRLFL] 检测到 ENRLFL 行，无 Study Specific 列，按存在即采纳")
            out.append("enrlfl")
    else:
        logger.info("[RANDFL/ENRLFL] 未在 ADSL 的 Variable 列中找到 ENRLFL，将不输出入组块（第14-17行）")

    result = tuple(out) if out else ("randfl",)
    blocks = []
    if "randfl" in result:
        blocks.append("第10-13行(RANDFL)")
    if "enrlfl" in result:
        blocks.append("第14-17行(ENRLFL)")
    logger.info("[RANDFL/ENRLFL] 解析结果：%s → 将输出块：%s", result, "；".join(blocks))
    return result


def _t14_05_texts_from_label(label):
    """
    从变量标签生成 05 部分前 3 行 TEXT。
    规则：从标签中保留「治疗」两字作为 base（去掉其他字样，不固定后缀），第1行=「完成研究」+base，第2行=「终止研究」+base，第3行=「终止研究」+base+「原因」。
    返回: (row1_text, row2_text, row3_text)
    """
    label = (label or "").strip() or _T14_05_DEFAULT_LABEL
    keep = _T14_05_BASE_KEEP
    base = keep if keep and keep in label else (label or keep)
    if not base:
        base = "治疗"
    row1 = _T14_05_PREFIX_COMPLETE + base
    row2 = _T14_05_PREFIX_TERMINATE + base
    row3 = _T14_05_PREFIX_TERMINATE + base + "原因"
    return (row1, row2, row3)


def parse_adam_spec_for_eotstt_label(adam_excel_path):
    """
    从 ADaM 数据集说明 Excel 的 variables sheet 中，在 _T14_05_DATASET 下查找变量（治疗结束状态）：
    优先精确匹配 _T14_05_VAR_EOTSTT，否则匹配以该前缀开头的变量（如 EOTSTT1、EOTSTT2）。
    返回 (Variable Label/标签列取值, 实际匹配到的变量名)，用于 05 部分 TEXT 与 FILTER；未找到则返回 (默认标签, 宏变量名)。
    """
    default_label = _T14_05_DEFAULT_LABEL
    default_var = _T14_05_VAR_EOTSTT.strip()
    try:
        import pandas as pd
    except ImportError:
        return (default_label, default_var or "EOTSTT")

    if not adam_excel_path or not os.path.isfile(adam_excel_path):
        return (default_label, default_var or "EOTSTT")

    try:
        xl = pd.ExcelFile(adam_excel_path)
        sheet_name = None
        for s in xl.sheet_names:
            if "variable" in s.lower():
                sheet_name = s
                break
        if sheet_name is None:
            return (default_label, default_var or "EOTSTT")
        df = pd.read_excel(adam_excel_path, sheet_name=sheet_name, header=0)
        if df.empty:
            return (default_label, default_var or "EOTSTT")

        col_dataset = _find_excel_column(df, ("Dataset", "Data Set", "数据集", "Dataset Name"))
        col_var = _find_excel_column(df, ("Variable", "变量", "Variable Name"))
        col_label = _find_excel_column(df, ("Variable Label", "Label", "变量标签", "标签", "VariableLabel"))
        if col_dataset is None or col_var is None:
            return (default_label, default_var or "EOTSTT")

        ds_col = df[col_dataset].astype(str).str.strip()
        adsl_mask = ds_col.str.upper() == _T14_05_DATASET.upper()
        adsl_df = df.loc[adsl_mask]
        if adsl_df.empty:
            return (default_label, default_var or "EOTSTT")

        var_col = adsl_df[col_var].astype(str).str.strip()
        var_upper = var_col.str.upper()
        prefix = (default_var or "EOTSTT").upper()
        # 优先精确匹配（如 EOTSTT），否则匹配以该前缀开头的变量（如 EOTSTT1、EOTSTT2）
        eotstt_mask = (var_upper == prefix) if prefix else (var_col == "")
        if not eotstt_mask.any() and prefix:
            eotstt_mask = var_upper.str.startswith(prefix)
        if not eotstt_mask.any():
            return (default_label, default_var or "EOTSTT")
        if col_label is None:
            return (default_label, default_var or "EOTSTT")
        matched_row = adsl_df.loc[eotstt_mask].iloc[0]
        label_val = str(matched_row[col_label] or "").strip()
        actual_var = str(matched_row[col_var] or "").strip() or default_var
        return (label_val if label_val else default_label, actual_var)
    except Exception:
        return (default_label, default_var or "EOTSTT")


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

    # 与 EDC 表结构一致：CODE_GRP, CODE_NAME_CHN, CODE_LABEL, CODE_ORDER 等；顺序列支持 CODE_ORDER 或 CODE_ORDER_R
    col_name = _find_excel_column(df, ("CODE_NAME_CHN", "Code_Name_Chn", "code_name_chn"))
    col_order = _find_excel_column(df, ("CODE_ORDER", "Code_Order", "code_order", "CODE_ORDER_R", "Code_Order_R"))
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

    # 按列名（表头）再匹配一轮：EDCDEF 中可能用「筛选失败原因」等作为列名，而非 CODE_NAME_CHN 的行值
    for col in df.columns:
        col_strip = str(col).strip()
        for name in _EDC_ALL_REASON_NAMES:
            if col_strip != name and not (name in col_strip and len(col_strip) >= len(name)):
                continue
            # 用该列的值作为 CODE_LABEL，同一行的 CODE_ORDER 作为顺序
            pairs = []
            for idx, row in df.iterrows():
                try:
                    o = row.get(col_order) if col_order is not None else idx
                    o = float(o) if o is not None and str(o).strip() else idx
                except (ValueError, TypeError):
                    o = idx
                lb = str(row.get(col, "") or "").strip()
                if lb:
                    pairs.append((o, lb))
            pairs.sort(key=lambda x: x[0])
            if pairs:
                result[col_strip] = pairs
            break

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


def _get_screen_fail_reasons(edc_data):
    """从 EDCDEF 中提取筛选失败原因列表。优先按列名（表头）匹配「筛选结束原因/筛选失败原因」；否则按行内 CODE_NAME_CHN 值匹配；按 CODE_ORDER 排序，TEXT 取 CODE_LABEL。"""
    if not edc_data:
        logger.warning("[筛选失败原因] edc_data 为空，无法匹配列名或 CODE_NAME_CHN，返回空列表。")
        return []
    edc_keys = list(edc_data.keys())
    logger.info("[筛选失败原因] EDCDEF 候选键（列名或 CODE_NAME_CHN 行值，共 %d 个）：%s", len(edc_keys), edc_keys)
    logger.info("[筛选失败原因] 待匹配名称 _EDC_SCREEN_FAIL_NAMES：%s", _EDC_SCREEN_FAIL_NAMES)
    for k, items in edc_data.items():
        k_strip = (k or "").strip()
        for name in _EDC_SCREEN_FAIL_NAMES:
            # 精确匹配或数据键包含完整名称，避免短子串误匹配（如 k='因'）
            if k_strip == name or (name in k_strip and len(k_strip) >= len(name)):
                labels = [lb for _, lb in items]
                logger.info("[筛选失败原因] 匹配到 key=%r，共 %d 条：%s", k_strip, len(labels), labels)
                return labels
        logger.debug("[筛选失败原因] key=%r 与任一 %s 未匹配。", k_strip, _EDC_SCREEN_FAIL_NAMES)
    logger.warning("[筛选失败原因] 未找到匹配的列名或 CODE_NAME_CHN（期望含：%s），返回空列表。", _EDC_SCREEN_FAIL_NAMES)
    return []


def build_t14_1_1_1_rows(randfl_enrlfl_flags, dct_reasons, followup_reasons, screen_fail_reasons=None, treatment_end_label=None, treatment_end_var_name=None):
    """
    按 Meta_Data 流程构建 T14_1-1_1 受试者分布的所有行。
    randfl_enrlfl_flags: tuple of "randfl" 和/或 "enrlfl"。每个 flag 输出「第10行+3行」块。
    dct_reasons: 治疗结束原因列表
    followup_reasons: 随访结束原因列表
    screen_fail_reasons: 筛选失败原因列表
    treatment_end_label: 「治疗结束状态」类变量标签，用于动态生成 05 部分前 3 行 TEXT；None 时用默认。
    treatment_end_var_name: 实际变量名（如 EOTSTT、EOTSTT1），用于 05 部分 FILTER；None 时用 _T14_05_VAR_EOTSTT。
    返回: list of dict with keys: TEXT, MASK, LINE_BREAK, INDENT, SEC, TRT_I, DSNIN, TRTSUBN, TRTSUBC, FILTER（第一步不含 ROW、FOOTNOTE）
    """
    def _empty_meta():
        return {"SEC": "", "TRT_I": "", "DSNIN": "", "TRTSUBN": "", "TRTSUBC": ""}

    rows = []
    row_num = 0

    # 01部分（第1-2行按 Meta_Data 流程：SEC/DSNIN/TRTSUBN/TRTSUBC/FILTER）
    row_num += 1
    rows.append({
        "TEXT": _T14_01_ROW1, "MASK": "", "LINE_BREAK": "", "INDENT": "",
        "SEC": _T14_01_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC,
        "FILTER": _T14_01_FILTER_ROW1,
    })
    row_num += 1
    rows.append({
        "TEXT": _T14_01_ROW2, "MASK": "", "LINE_BREAK": "", "INDENT": "",
        "SEC": _T14_01_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC,
        "FILTER": _T14_01_FILTER_ROW2,
    })

    # 筛选失败原因（第2行之后、第3行之前）：标题行 + 各原因子行，数据来源 EDCDEF CODE_NAME_CHN=筛选结束原因
    if screen_fail_reasons is None:
        screen_fail_reasons = []
    row_num += 1
    rows.append({
        "TEXT": "筛选失败原因", "MASK": "", "LINE_BREAK": "", "INDENT": "",
        "SEC": _T14_01_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC,
        "FILTER": "0",
    })
    for reason in screen_fail_reasons:
        row_num += 1
        # SAS 字符串内单引号需双写
        reason_esc = (reason or "").replace("'", "''")
        filter_val = "%s and SCFAILRE='%s'" % (_T14_01_FILTER_ROW2, reason_esc)
        rows.append({
            "TEXT": reason or "", "MASK": "", "LINE_BREAK": "", "INDENT": "1",
            "SEC": _T14_01_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC,
            "FILTER": filter_val,
        })

    # 第10行+3行 / 第14行+3行：若既有 RANDFL 也有 ENRLFL 则都处理。每个 flag 一块：1 行「筛选成功未随机/未入组」+ 3 行 04（随机受试者/入组受试者等）
    flags = tuple(randfl_enrlfl_flags) if randfl_enrlfl_flags else ("randfl",)
    for flag in flags:
        if flag == "enrlfl":
            text_before = _T14_01_ROW_BEFORE_ROW3_ENRLFL
            filter_before = _T14_01_FILTER_ROW_BEFORE_ROW3_ENRLFL
            four_rows_text = _T14_04_ROWS_ENRLFL
            four_filters = _T14_04_FILTERS_ENRLFL
        else:
            text_before = _T14_01_ROW_BEFORE_ROW3_RANDFL
            filter_before = _T14_01_FILTER_ROW_BEFORE_ROW3_RANDFL
            four_rows_text = _T14_04_ROWS_RANDFL
            four_filters = _T14_04_FILTERS_RANDFL
        row_num += 1
        rows.append({
            "TEXT": text_before, "MASK": "", "LINE_BREAK": "", "INDENT": "",
            "SEC": _T14_01_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC,
            "FILTER": filter_before,
        })
        for i, (t, f) in enumerate(zip(four_rows_text, four_filters)):
            row_num += 1
            # 04 块第一行：LINE_BREAK="1"、INDENT 为空；第2、3行：LINE_BREAK 为空、INDENT="1"
            line_break = "1" if i == 0 else ""
            indent = "" if i == 0 else "1"
            rows.append({
                "TEXT": t, "MASK": "", "LINE_BREAK": line_break, "INDENT": indent,
                "SEC": _T14_04_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC,
                "FILTER": f,
            })

    # 05部分：前 3 行 TEXT 由 treatment_end_label 动态生成（去掉「结束状态」得 base，再拼「完成研究」/「终止研究」）
    row1_text, row2_text, row3_text = _t14_05_texts_from_label(treatment_end_label)
    row1_esc = (row1_text or "").replace("'", "''")
    row2_esc = (row2_text or "").replace("'", "''")
    var_eotstt = (treatment_end_var_name or _T14_05_VAR_EOTSTT or "EOTSTT").strip()
    filter_row1 = "saffl='Y' and %s='%s'" % (var_eotstt, row1_esc)
    filter_row2 = "saffl='Y' and %s='%s'" % (var_eotstt, row2_esc)
    em_05 = {"SEC": _T14_05_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC}
    row_num += 1
    rows.append({
        "TEXT": row1_text, "MASK": "", "LINE_BREAK": "1", "INDENT": "",
        **em_05, "FILTER": filter_row1,
    })
    row_num += 1
    rows.append({
        "TEXT": row2_text, "MASK": "", "LINE_BREAK": "", "INDENT": "",
        **em_05, "FILTER": filter_row2,
    })
    row_num += 1
    rows.append({
        "TEXT": row3_text, "MASK": "", "LINE_BREAK": "", "INDENT": "",
        **em_05, "FILTER": "0",
    })
    for reason in dct_reasons:
        if (reason or "").strip() in _T14_05_EXCLUDE_REASONS:
            continue
        row_num += 1
        reason_esc = (reason or "").replace("'", "''")
        filter_dct = "%s and dctreas='%s'" % (filter_row2, reason_esc)
        rows.append({
            "TEXT": reason or "", "MASK": "", "LINE_BREAK": "", "INDENT": "1",
            **em_05, "FILTER": filter_dct,
        })

    # 06部分：前3行强制赋值（完成研究、退出研究、退出研究原因），后面类似05后半部分且隔行增加「随机未接受研究治疗」
    em_06 = {"SEC": _T14_06_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC}
    row_num += 1
    rows.append({
        "TEXT": _T14_06_ROW1, "MASK": "", "LINE_BREAK": "1", "INDENT": "",
        **em_06, "FILTER": _T14_06_FILTER_ROW1,
    })
    row_num += 1
    rows.append({
        "TEXT": _T14_06_ROW2, "MASK": "", "LINE_BREAK": "", "INDENT": "",
        **em_06, "FILTER": _T14_06_FILTER_ROW2,
    })
    row_num += 1
    rows.append({
        "TEXT": _T14_06_ROW3, "MASK": "", "LINE_BREAK": "", "INDENT": "",
        **em_06, "FILTER": "0",
    })
    for reason in followup_reasons:
        if (reason or "").strip() in _T14_06_EXCLUDE_REASONS:
            continue
        row_num += 1
        reason_esc = (reason or "").replace("'", "''")
        reason_filter = "%s and dcsreas='%s'" % (_T14_06_FILTER_ROW2, reason_esc)
        rows.append({
            "TEXT": reason or "", "MASK": "", "LINE_BREAK": "", "INDENT": "1",
            **em_06, "FILTER": reason_filter,
        })
        row_num += 1
        extra_filter = reason_filter + _T14_06_EXTRA_SUFFIX
        rows.append({
            "TEXT": _T14_06_EXTRA, "MASK": "", "LINE_BREAK": "", "INDENT": "2",
            **em_06, "FILTER": extra_filter,
        })

    return rows


# 第一步 T14_1-1_1.xlsx 表头列名（不含 ROW、FOOTNOTE），与 build_t14_1_1_1_rows 返回的 dict 部分键对应
_T14_1_1_1_COLUMNS = (
    "TEXT", "MASK", "LINE_BREAK", "INDENT", "SEC", "TRT_I", "DSNIN", "TRTSUBN", "TRTSUBC", "FILTER",
)


def write_t14_1_1_1_xlsx(xlsx_path, rows):
    """将受试者分布行写入 Excel，不包含 ROW、FOOTNOTE 列：TEXT, MASK, LINE_BREAK, INDENT, SEC, TRT_I, DSNIN, TRTSUBN, TRTSUBC, FILTER。"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "受试者分布"
    ws.append(list(_T14_1_1_1_COLUMNS))
    defaults = {"TEXT": "", "MASK": "", "LINE_BREAK": "", "INDENT": "", "SEC": "", "TRT_I": "", "DSNIN": "", "TRTSUBN": "", "TRTSUBC": "", "FILTER": ""}
    for r in rows:
        row = [r.get(k, defaults.get(k, "")) for k in _T14_1_1_1_COLUMNS]
        ws.append(row)
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

    # ADaM 数据集说明：在该文件夹下自动查找文件名同时含 ADAM、PDS 的 .xlsx，没有则留空
    adam_doc_dir = os.path.join(base_path, "utility", "documentation")
    if not os.path.isdir(adam_doc_dir):
        adam_doc_dir = os.path.join(base_path, "utility", "documents")

    def _find_adam_pds_xlsx(d):
        if not d or not os.path.isdir(d):
            return None
        cands = []
        for name in os.listdir(d):
            if not name.lower().endswith(".xlsx"):
                continue
            n = name.lower()
            if "adam" in n and "pds" in n:
                p = os.path.join(d, name)
                if os.path.isfile(p):
                    cands.append((os.path.getmtime(p), p))
        if not cands:
            return None
        cands.sort(key=lambda x: -x[0])
        return cands[0][1]

    default_adam = _find_adam_pds_xlsx(adam_doc_dir)
    if not default_adam:
        default_adam = _find_adam_pds_xlsx(os.path.join(base_path, "utility", "documents"))
    if not default_adam:
        default_adam = ""  # 未找到则留空

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
    if default_adam:
        adam_entry.insert(0, default_adam)

    def browse_adam():
        path = filedialog.askopenfilename(
            title="选择 ADaM 数据集说明文件（Excel）",
            filetypes=[("Excel", "*.xlsx"), ("Excel 97", "*.xls"), ("All", "*.*")],
            initialdir=adam_doc_dir if os.path.isdir(adam_doc_dir) else base_path,
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

        randfl_enrlfl_flags = ("randfl",)
        if adam_path and os.path.isfile(adam_path):
            try:
                randfl_enrlfl_flags = parse_adam_spec_for_randfl_enrlfl(adam_path)
                gui.update_status("ADaM 解析：%s" % (", ".join(randfl_enrlfl_flags) if randfl_enrlfl_flags else "未检测到 RANDFL/ENRLFL"))
            except Exception as e:
                messagebox.showwarning("ADaM 解析", "无法解析 ADaM 说明文件，将使用默认（随机受试者）：%s" % e)
                randfl_enrlfl_flags = ("randfl",)
        else:
            randfl_enrlfl_flags = ("randfl",)

        edc_data = {}
        if edc_path and os.path.isfile(edc_path):
            try:
                edc_data = read_edcdef_code(edc_path)
                gui.update_status("EDCDEF 已读取")
            except Exception as e:
                messagebox.showwarning("EDCDEF 读取", "无法读取 EDCDEF_code，05/06 部分将为空：%s" % e)

        dct_reasons = _get_dctreas_reasons(edc_data)
        followup_reasons = _get_followup_reasons(edc_data)
        screen_fail_reasons = _get_screen_fail_reasons(edc_data)

        treatment_end_label = None
        treatment_end_var_name = None
        if adam_path and os.path.isfile(adam_path):
            treatment_end_label, treatment_end_var_name = parse_adam_spec_for_eotstt_label(adam_path)

        try:
            if os.path.isfile(path):
                backup_path = _backup_existing_to_archive(path)
                if backup_path:
                    gui.update_status("已备份原文件至：%s" % backup_path)
            d = os.path.dirname(path)
            if d:
                os.makedirs(d, exist_ok=True)
            rows = build_t14_1_1_1_rows(randfl_enrlfl_flags, dct_reasons, followup_reasons, screen_fail_reasons, treatment_end_label, treatment_end_var_name)
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

    # SAP 文件（.docx）默认：在默认目录下查找文件名中包含 SAP 的 .docx（不区分大小写）
    default_sap_dir = os.path.join(base_path, "utility", "documentation", "03_statistics")
    default_metadata_dir = os.path.join(base_path, "utility", "metadata")
    default_xlsx_step2 = os.path.join(default_metadata_dir, "T14_1-1_2.xlsx")

    def _find_sap_docx_in_dir(d):
        if not d or not os.path.isdir(d):
            return None
        cands = []
        for name in os.listdir(d):
            if name.lower().endswith(".docx") and "sap" in name.lower():
                p = os.path.join(d, name)
                if os.path.isfile(p):
                    cands.append((os.path.getmtime(p), p))
        if not cands:
            return None
        cands.sort(key=lambda x: -x[0])  # 按修改时间取最新
        return cands[0][1]

    default_sap = _find_sap_docx_in_dir(default_sap_dir)
    if not default_sap:
        default_sap = _find_sap_docx_in_dir(os.path.join(base_path, "utility", "documentation"))
    if not default_sap:
        default_sap = os.path.join(default_sap_dir, "SAP.docx")  # 无匹配时仍显示默认路径供浏览

    row_sap = tk.Frame(main, bg="#f0f0f0")
    row_sap.pack(anchor="w", fill=tk.X, pady=(0, 6))
    tk.Label(row_sap, text="SAP 文件（.docx）：", font=("Microsoft YaHei UI", 9), width=22, anchor="w", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 4))
    sap_entry = tk.Entry(row_sap, width=72, font=("Microsoft YaHei UI", 9))
    sap_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    sap_entry.insert(0, default_sap)

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
