# -*- coding: utf-8 -*-
"""
生成受试者分布metadata t14_1-1_1.xlsx。

按 Meta_Data 表格制作流程生成 01/04/05/06 四部分。
输入方式：输入宏参数，用户调用时在下方设置路径，或通过命令行传入。
05_trt 簇数始终由 PATH_EDCDEF_CASEBOOK 解析。可直接运行本脚本或作为模块调用 run_tadsl_disp_init(path_t141_11, path_adam_spec, path_edc_code, path_edc_casebook)。
"""

print(">>>>>>>>>>>>>生成受试者分布metadata t14_1-1_1.xlsx>>>>>>>>>>>>>")

import os
import sys
import warnings

# 抑制 openpyxl 读取含页眉/页脚的 xlsx 时的解析警告（如 ADaM PDS）
warnings.filterwarnings("ignore", message="Cannot parse header or footer so it will be ignored")

# ---------- 输入宏参数：用户调用时设置以下路径（或通过命令行参数传入） ----------
PATH_T14_1_1_1 = ""   # t14_1-1_1.xlsx 输出路径
PATH_ADAM_SPEC = ""   # ADaM PDS
PATH_EDCDEF_CODE = "" # EDCDEF_code（.sas7bdat 或 .xlsx）
PATH_EDCDEF_CASEBOOK = ""  # edcdef_casebook.sas7bdat（用户必填）。05_trt 簇数始终由此解析：EDC_DATA 以 DSEOT 开头的观测数，EDC_FORM 取「结束页」前作为 base


# ---------- T14_1-1_1 受试者分布数据解析与生成 ----------

# 01部分固定文本与列取值（第1-2行）
_T14_01_ROW1 = "筛选受试者"
_T14_01_ROW2 = "筛选失败受试者"
_T14_01_SEC = "01_scr"
_T14_01_DSNIN = "adsl"
_T14_01_TRTSUBN = "trt01pn"
_T14_01_TRTSUBC = "trt01p"
_T14_01_FILTER_ROW1 = "prxmatch('/^(合计|total)\\s*$/i', trt01p)"
_T14_01_FILTER_ROW2 = "prxmatch('/^(合计|total)\\s*$/i', trt01p) and (scfailfl='Y')"
# 第3行之前一行（第10行）：根据 RANDFL/ENRLFL 二选一 筛选成功未随机受试者 vs 筛选成功未入组受试者
_T14_01_ROW_BEFORE_ROW3_RANDFL = "筛选成功未随机受试者"
_T14_01_FILTER_ROW_BEFORE_ROW3_RANDFL = "prxmatch('/^(合计|total)\\s*$/i', trt01p) and (scfailfl='N') and randfl='N'"
_T14_01_ROW_BEFORE_ROW3_ENRLFL = "筛选成功未入组受试者"
_T14_01_FILTER_ROW_BEFORE_ROW3_ENRLFL = "prxmatch('/^(合计|total)\\s*$/i', trt01p) and (scfailfl='N') and enrlfl='N'"
# 04部分：随机或入组受试者分布（RANDFL 与 ENRLFL 可同时存在）
_T14_04_SEC = "04_rnd"
_T14_04_ROWS_RANDFL = ("随机受试者", "随机未接受研究治疗", "随机且接受研究治疗")
_T14_04_FILTERS_RANDFL = ("randfl='Y' and scfailfl='N'", "randfl='Y' and scfailfl='N' and saffl='N'", "randfl='Y' and scfailfl='N' and saffl='Y'")
_T14_04_ROWS_ENRLFL = ("入组受试者", "入组未接受研究治疗", "入组且接受研究治疗")
_T14_04_FILTERS_ENRLFL = ("enrlfl='Y' and scfailfl='N'", "enrlfl='Y' and scfailfl='N' and saffl='N'", "enrlfl='Y' and scfailfl='N' and saffl='Y'")
# 05部分：TEXT = 「完成」/「终止」+ EDC_FORM（「结束页」前）；FILTER 中 EOTSTT/EOTSTTn 的取值固定为「完成治疗」或「终止治疗」
_T14_05_SEC = "05_trt"
_T14_05_DATASET = "ADSL"
_T14_05_VAR_EOTSTT = "EOTSTT"
_T14_05_BASE_KEEP = "治疗"
_T14_05_PREFIX_COMPLETE = "完成"
_T14_05_PREFIX_TERMINATE = "终止"
_T14_05_FILTER_VALUE_COMPLETE = "完成治疗"   # 完成治疗行 FILTER 中变量取值
_T14_05_FILTER_VALUE_TERMINATE = "终止治疗"  # 终止治疗行及终止治疗原因行 FILTER 中变量取值
# 05_trt 仅一簇时：完成/终止行固定为「完成研究治疗」「终止研究治疗」，FILTER 固定用 EOTSTT
_T14_05_SINGLE_CLUSTER_ROW1_TEXT = "完成研究治疗"
_T14_05_SINGLE_CLUSTER_ROW1_FILTER = "saffl='Y' and EOTSTT='完成治疗'"
_T14_05_SINGLE_CLUSTER_ROW2_TEXT = "终止研究治疗"
_T14_05_SINGLE_CLUSTER_ROW2_FILTER = "saffl='Y' and EOTSTT='终止治疗'"
_T14_05_DEFAULT_LABEL = "治疗结束状态"
_T14_05_EXCLUDE_REASONS = ("已完成",)

_T14_06_EXCLUDE_REASONS = ("已完成",)
# 06部分
_T14_06_SEC = "06_fup"
_T14_06_ROW1 = "完成研究"
_T14_06_FILTER_ROW1 = "saffl='Y' and EOSSTT='完成研究'"
_T14_06_ROW2 = "退出研究"
_T14_06_FILTER_ROW2 = "saffl='Y' and EOSSTT='退出研究'"
_T14_06_ROW3 = "退出研究原因"
_T14_06_EXTRA = "随机未接受研究治疗"
_T14_06_EXTRA_SUFFIX = " and randfl='Y' and saffl='N'"

# 筛选失败原因/治疗结束原因/退出研究原因：仅当 CODE_GRP=DSENROLL/DSEOT/DSEOS 且 CODE_NAME='DSDECOD'，且 CODE_LABEL 不为「已完成」「COMPLETED」时纳入
_EDC_CODE_GRP_SCREEN = "DSENROLL"
_EDC_CODE_GRP_DCTREAS = "DSEOT"
_EDC_CODE_GRP_FOLLOWUP = "DSEOS"
_EDC_CODE_NAME_DSDECOD = "DSDECOD"
_EDC_EXCLUDE_LABELS = ("已完成", "COMPLETED")  # CODE_LABEL 等于这些时排除
# Casebook：EDC_FORM（表单名称/标签列）仅保留「结束页」之前的文本作为 05_trt 簇前三行的 base
_CASEBOOK_EDC_FORM_TRUNCATE_AT = "结束页"


def _find_excel_column(df, candidates):
    """在 DataFrame 列名中查找匹配项（忽略大小写、首尾空格）。返回列名或 None。"""
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        k = cand.strip().lower()
        for col_key, col_orig in cols.items():
            if col_key == k:
                return col_orig
        for col_key, col_orig in cols.items():
            if k in col_key and len(col_key) >= len(k):
                return col_orig
            if col_key in k and len(col_key) >= len(k):
                return col_orig
    return None


def parse_adam_spec_for_randfl_enrlfl(adam_excel_path):
    """
    从 ADaM PDS 的 variables sheet 中判断 ADSL 是否存在 RANDFL/ENRLFL （StudySpecific Flag为Y为前提筛选条件）。
    返回: tuple of "randfl" 和/或 "enrlfl"；均不存在时返回 ("randfl",) 作为默认。
    """

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
        print("RANDFL/ENRLFL exists in ADSL？未找到 variables 相关 sheet，sheet 列表：", xl.sheet_names)
        raise ValueError("ADaM 说明文件中未找到 variables 相关 sheet。")

    df = pd.read_excel(adam_excel_path, sheet_name=sheet_name, header=0)
    if df.empty:
        raise ValueError("variables sheet 为空。")

    col_dataset = _find_excel_column(df, ("Dataset", "Data Set", "数据集", "Dataset Name"))
    col_var = _find_excel_column(df, ("Variable", "变量", "Variable Name"))
    col_study_specific = _find_excel_column(df, ("Study Specific", "StudySpecific", "Study Specific Flag"))

    if col_dataset is None or col_var is None:
        raise ValueError("variables sheet 中未找到 Dataset 或 Variable 列。")

    ds_col = df[col_dataset].astype(str).str.strip()
    adsl_mask = ds_col.str.upper() == "ADSL"
    adsl_df = df.loc[adsl_mask]

    if adsl_df.empty:
        print("[RANDFL/ENRLFL] 无 ADSL 行，返回默认 ('randfl',)")
        return ("randfl",)

    var_col = adsl_df[col_var].astype(str).str.strip()
    out = []

    randfl_mask = var_col.str.upper() == "RANDFL"
    if randfl_mask.any():
        if col_study_specific is not None:
            ss = adsl_df.loc[randfl_mask, col_study_specific].astype(str).str.strip().str.upper()
            if (ss == "Y").any():
                out.append("randfl")
        else:
            out.append("randfl")

    enrlfl_mask = var_col.str.upper() == "ENRLFL"
    if enrlfl_mask.any():
        if col_study_specific is not None:
            ss_enrl = adsl_df.loc[enrlfl_mask, col_study_specific].astype(str).str.strip().str.upper()
            if (ss_enrl == "Y").any():
                out.append("enrlfl")
        else:
            out.append("enrlfl")

    result = tuple(out) if out else ("randfl",)
    print("RANDFL/ENRLFL exists in ADSL？ 解析结果：", result)
    return result


def _t14_05_texts_from_label(label):
    """从变量标签生成 05 部分前 3 行 TEXT。返回 (row1_text, row2_text, row3_text)。"""
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
    """从 ADaM PDS的 variables sheet 中查找治疗结束状态变量标签与变量名。返回 (label, var_name)。"""
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
    读取 EDC Metadata EDCDEF_code，仅当 CODE_GRP in (DSENROLL,DSEOT,DSEOS) 且 CODE_NAME='DSDECOD'，
    且 CODE_LABEL 不为「已完成」「COMPLETED」时纳入。返回 dict[str, list[(order, label)]]，key 为 CODE_GRP。
    """
    try:
        import pandas as pd
    except ImportError:
        raise RuntimeError("请先安装 pandas：pip install pandas")

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

    col_grp = _find_excel_column(df, ("CODE_GRP", "Code_Grp", "code_grp"))
    col_name = _find_excel_column(df, ("CODE_NAME", "Code_Name", "code_name"))
    col_order = _find_excel_column(df, ("CODE_ORDER", "Code_Order", "code_order", "CODE_ORDER_R", "Code_Order_R"))
    col_label = _find_excel_column(df, ("CODE_LABEL", "Code_Label", "code_label"))

    if col_grp is None or col_label is None:
        return {}
    if col_name is None:
        return {}

    code_grps_want = {_EDC_CODE_GRP_SCREEN, _EDC_CODE_GRP_DCTREAS, _EDC_CODE_GRP_FOLLOWUP}
    code_name_want = _EDC_CODE_NAME_DSDECOD.upper()

    result = {_EDC_CODE_GRP_SCREEN: [], _EDC_CODE_GRP_DCTREAS: [], _EDC_CODE_GRP_FOLLOWUP: []}
    for _, row in df.iterrows():
        grp = str(row.get(col_grp, "") or "").strip().upper()
        name = str(row.get(col_name, "") or "").strip().upper()
        if grp not in code_grps_want or name != code_name_want:
            continue
        label = str(row.get(col_label, "") or "").strip()
        if not label:
            continue
        if label == "已完成" or label.upper() == "COMPLETED":
            continue
        order_val = row.get(col_order) if col_order else 0
        try:
            order_val = float(order_val) if order_val is not None and str(order_val).strip() else 0
        except (ValueError, TypeError):
            order_val = 0
        result[grp].append((order_val, label))

    for k in result:
        result[k].sort(key=lambda x: x[0])

    return result

def _get_screen_fail_reasons(edc_data):
    """从 EDCDEF 中提取筛选失败原因列表（CODE_GRP=DSENROLL，CODE_NAME=DSDECOD，已排除已完成/COMPLETED）。"""
    items = edc_data.get(_EDC_CODE_GRP_SCREEN, [])
    return [lb for _, lb in items]

def _get_dctreas_reasons(edc_data):
    """从 EDCDEF 中提取治疗结束原因列表（CODE_GRP=DSEOT，CODE_NAME=DSDECOD，已排除已完成/COMPLETED）。"""
    items = edc_data.get(_EDC_CODE_GRP_DCTREAS, [])
    return [lb for _, lb in items]


def _get_followup_reasons(edc_data):
    """从 EDCDEF 中提取退出研究原因列表（CODE_GRP=DSEOS，CODE_NAME=DSDECOD，已排除已完成/COMPLETED）。"""
    items = edc_data.get(_EDC_CODE_GRP_FOLLOWUP, [])
    return [lb for _, lb in items]


def read_edcdef_casebook(casebook_path):
    """
    读取 edcdef_casebook.sas7bdat，筛选 EDC_DATA 以 DSEOT 开头的观测，得到每簇 05_trt 的 base 与变量名。
    返回: list of (base, var_name)。base 为 EDC_FORM 列「结束页」之前的文本；var_name 为 EDC_DATA 去掉 DS 前缀（DSEOT→EOTSTT, DSEOT1→EOTSTT1）。
    """
    if not casebook_path or not os.path.isfile(casebook_path):
        return []
    ext = os.path.splitext(casebook_path)[1].lower()
    if ext != ".sas7bdat":
        print("Casebook 仅支持 .sas7bdat，当前：", casebook_path)
        return []
    try:
        import pandas as pd
        import pyreadstat
    except ImportError as e:
        print("读取 Casebook 需要 pyreadstat：", e)
        return []
    try:
        df, _ = pyreadstat.read_sas7bdat(casebook_path)
    except Exception as e:
        print("无法读取 Casebook %s：%s" % (casebook_path, e))
        return []
    if df is None or df.empty:
        return []
    col_data = _find_excel_column(df, ("EDC_DATA", "Edc_Data", "edc_data"))
    col_form = _find_excel_column(df, ("EDC_FORM", "Edc_Form", "edc_form"))
    if col_data is None or col_form is None:
        print("Casebook 中未找到 EDC_DATA 或 EDC_FORM 列")
        return []
    out = []
    truncate_at = _CASEBOOK_EDC_FORM_TRUNCATE_AT
    for _, row in df.iterrows():
        edc_data_val = str(row.get(col_data, "") or "").strip().upper()
        if not edc_data_val.startswith("DSEOT"):
            continue
        # 变量名：去掉 DS 前缀，EOT/EOT1/EOT2 转为 EOTSTT/EOTSTT1/EOTSTT2
        rest = edc_data_val[2:] if len(edc_data_val) >= 2 else edc_data_val
        if rest.startswith("EOT"):
            var_name = "EOTSTT" + rest[3:]  # EOT->EOTSTT, EOT1->EOTSTT1, EOT2->EOTSTT2
        else:
            var_name = rest
        edc_form_val = str(row.get(col_form, "") or "").strip()
        base = edc_form_val.split(truncate_at)[0].strip() if truncate_at in edc_form_val else edc_form_val
        if not base:
            base = _T14_05_BASE_KEEP
        out.append((base, var_name))
    return out


def get_adsl_study_specific_variables(adam_excel_path):
    """返回 ADaM PDS 中 ADSL 数据集、Study_specific='Y' 的 Variable 列集合（大写）。用于校验 05_trt 变量名。"""
    if not adam_excel_path or not os.path.isfile(adam_excel_path):
        return set()
    try:
        import pandas as pd
    except ImportError:
        return set()
    try:
        xl = pd.ExcelFile(adam_excel_path)
        sheet_name = None
        for s in xl.sheet_names:
            if "variable" in s.lower():
                sheet_name = s
                break
        if sheet_name is None:
            return set()
        df = pd.read_excel(adam_excel_path, sheet_name=sheet_name, header=0)
        if df.empty:
            return set()
        col_dataset = _find_excel_column(df, ("Dataset", "Data Set", "数据集", "Dataset Name"))
        col_var = _find_excel_column(df, ("Variable", "变量", "Variable Name"))
        col_ss = _find_excel_column(df, ("Study Specific", "StudySpecific", "Study Specific Flag"))
        if col_dataset is None or col_var is None:
            return set()
        ds_col = df[col_dataset].astype(str).str.strip().str.upper()
        adsl_mask = ds_col == _T14_05_DATASET.upper()
        adsl_df = df.loc[adsl_mask]
        if adsl_df.empty:
            return set()
        var_col = adsl_df[col_var].astype(str).str.strip().str.upper()
        if col_ss is not None:
            ss_col = adsl_df[col_ss].astype(str).str.strip().str.upper()
            mask_y = ss_col == "Y"
            vars_y = set(var_col.loc[mask_y].tolist())
        else:
            vars_y = set(var_col.tolist())
        return set(v.strip().upper() for v in vars_y if v)
    except Exception:
        return set()


def build_t14_1_1_1_rows(randfl_enrlfl_flags, dct_reasons, followup_reasons, screen_fail_reasons=None, otr_clusters=None):
    """
    按 Meta_Data 流程构建 t14_1-1_1 受试者分布的所有行。返回 list of dict。
    05_trt 簇数始终由 otr_clusters 决定（来自 PATH_EDCDEF_CASEBOOK）：list of (base, var_name)，每项生成一簇 SEC=05_trt；空列表则无 05_trt 行。
    """
    rows = []
    row_num = 0

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
        reason_esc = (reason or "").replace("'", "''")
        filter_val = "%s and SCFAILRE='%s'" % (_T14_01_FILTER_ROW2, reason_esc)
        rows.append({
            "TEXT": reason or "", "MASK": "", "LINE_BREAK": "", "INDENT": "1",
            "SEC": _T14_01_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC,
            "FILTER": filter_val,
        })

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
            line_break = "1" if i == 0 else ""
            indent = "" if i == 0 else "1"
            rows.append({
                "TEXT": t, "MASK": "", "LINE_BREAK": line_break, "INDENT": indent,
                "SEC": _T14_04_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC,
                "FILTER": f,
            })

    em_05 = {"SEC": _T14_05_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC}
    otr_list = otr_clusters or []
    for idx, (base, var_eotstt) in enumerate(otr_list):
        var_eotstt = (var_eotstt or "").strip() or _T14_05_VAR_EOTSTT
        # DCTREAS 与 EOTSTT 使用相同序号后缀：EOTSTT→dctreas，EOTSTT1→dctreas1，EOTSTT2→dctreas2
        var_upper = var_eotstt.upper()
        suffix = var_upper[6:] if var_upper.startswith("EOTSTT") else ""
        dctreas_var = "dctreas" + suffix
        single_cluster = len(otr_list) == 1
        if single_cluster:
            row1_text = _T14_05_SINGLE_CLUSTER_ROW1_TEXT
            row2_text = _T14_05_SINGLE_CLUSTER_ROW2_TEXT
            filter_row1 = _T14_05_SINGLE_CLUSTER_ROW1_FILTER
            filter_row2 = _T14_05_SINGLE_CLUSTER_ROW2_FILTER
        else:
            row1_text = _T14_05_PREFIX_COMPLETE + (base or _T14_05_BASE_KEEP)
            row2_text = _T14_05_PREFIX_TERMINATE + (base or _T14_05_BASE_KEEP)
            filter_row1 = "saffl='Y' and %s='%s'" % (var_eotstt, _T14_05_FILTER_VALUE_COMPLETE)
            filter_row2 = "saffl='Y' and %s='%s'" % (var_eotstt, _T14_05_FILTER_VALUE_TERMINATE)
        row3_text = _T14_05_PREFIX_TERMINATE + (base or _T14_05_BASE_KEEP) + "原因"
        row_num += 1
        rows.append({"TEXT": row1_text, "MASK": "", "LINE_BREAK": "1", "INDENT": "", **em_05, "FILTER": filter_row1})
        row_num += 1
        rows.append({"TEXT": row2_text, "MASK": "", "LINE_BREAK": "", "INDENT": "", **em_05, "FILTER": filter_row2})
        row_num += 1
        rows.append({"TEXT": row3_text, "MASK": "", "LINE_BREAK": "", "INDENT": "", **em_05, "FILTER": "0"})
        for reason in dct_reasons:
            if (reason or "").strip() in _T14_05_EXCLUDE_REASONS:
                continue
            row_num += 1
            reason_esc = (reason or "").replace("'", "''")
            filter_dct = "%s and %s='%s'" % (filter_row2, dctreas_var, reason_esc)
            rows.append({"TEXT": reason or "", "MASK": "", "LINE_BREAK": "", "INDENT": "1", **em_05, "FILTER": filter_dct})

    em_06 = {"SEC": _T14_06_SEC, "TRT_I": "", "DSNIN": _T14_01_DSNIN, "TRTSUBN": _T14_01_TRTSUBN, "TRTSUBC": _T14_01_TRTSUBC}
    row_num += 1
    rows.append({"TEXT": _T14_06_ROW1, "MASK": "", "LINE_BREAK": "1", "INDENT": "", **em_06, "FILTER": _T14_06_FILTER_ROW1})
    row_num += 1
    rows.append({"TEXT": _T14_06_ROW2, "MASK": "", "LINE_BREAK": "", "INDENT": "", **em_06, "FILTER": _T14_06_FILTER_ROW2})
    row_num += 1
    rows.append({"TEXT": _T14_06_ROW3, "MASK": "", "LINE_BREAK": "", "INDENT": "", **em_06, "FILTER": "0"})
    for reason in followup_reasons:
        if (reason or "").strip() in _T14_06_EXCLUDE_REASONS:
            continue
        row_num += 1
        reason_esc = (reason or "").replace("'", "''")
        reason_filter = "%s and dcsreas='%s'" % (_T14_06_FILTER_ROW2, reason_esc)
        rows.append({"TEXT": reason or "", "MASK": "", "LINE_BREAK": "", "INDENT": "1", **em_06, "FILTER": reason_filter})
        row_num += 1
        extra_filter = reason_filter + _T14_06_EXTRA_SUFFIX
        rows.append({"TEXT": _T14_06_EXTRA, "MASK": "", "LINE_BREAK": "", "INDENT": "2", **em_06, "FILTER": extra_filter})

    return rows


_T14_1_1_1_COLUMNS = (
    "TEXT", "MASK", "LINE_BREAK", "INDENT", "SEC", "TRT_I", "DSNIN", "TRTSUBN", "TRTSUBC", "FILTER",
)


def write_tadsl_disp_xlsx(xlsx_path, rows):
    """将受试者分布行写入 Excel。"""
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


def run_tadsl_disp_init(path_t141_11, path_adam_spec, path_edc_code, path_edc_casebook):
    """
    初版 T14_1-1_1：按 Meta_Data 流程生成 01/04/05/06 四部分。若 t14_1-1_1.xlsx 已存在则仅告警不覆盖；不存在则生成。

    path_t141_11: t14_1-1_1.xlsx 输出路径
    path_adam_spec: ADaM PDS 路径（可为空，则使用默认 RANDFL）
    path_edc_code: EDCDEF_code 路径（.sas7bdat 或 .xlsx，可为空，则 05/06 原因为空）
    path_edc_casebook: edcdef_casebook.sas7bdat 路径（用户提供）。据此解析 EDC_DATA 以 DSEOT 开头的观测数，生成对应簇数的 SEC=05_trt；校验变量名在 ADSL Study_specific='Y' 中（不通过仅 warning 仍生成）
    """
    path_t141_11 = (path_t141_11 or "").strip()
    if not path_t141_11:
        raise ValueError("请设置 t14_1-1_1.xlsx 输出路径（宏参数 PATH_T14_1_1_1 或命令行第 1 个参数）。")

    randfl_enrlfl_flags = ("randfl",)
    if path_adam_spec and os.path.isfile(path_adam_spec):
        try:
            randfl_enrlfl_flags = parse_adam_spec_for_randfl_enrlfl(path_adam_spec)
        except Exception as e:
            print("无法解析 ADaM 说明文件，将使用默认（随机受试者）：", e)

    edc_data = {}
    if path_edc_code and os.path.isfile(path_edc_code):
        try:
            edc_data = read_edcdef_code(path_edc_code)
            print("edcdef_cose.sas7bdat 已读取DS相关表单，获取失败原因列表")
        except Exception as e:
            print("无法读取edcdef_cose.sas7bdat，05/06 部分原因将为空：", e)

    dct_reasons = _get_dctreas_reasons(edc_data)
    followup_reasons = _get_followup_reasons(edc_data)
    screen_fail_reasons = _get_screen_fail_reasons(edc_data)

    # 05_trt 簇数始终由 casebook 解析；用户始终提供 PATH_EDCDEF_CASEBOOK
    path_edc_casebook = (path_edc_casebook or "").strip()
    otr_clusters = read_edcdef_casebook(path_edc_casebook) if path_edc_casebook else []
    if otr_clusters:
        print("edcdef_casebook.sas7bdat中解析到有%d 簇 SEC=05_trt：%s" % (len(otr_clusters), [v for _, v in otr_clusters]))
        if len(otr_clusters) == 1:
            # 05_trt 簇数为 1 时，EDC_FORM 直接赋值为「治疗」作为 base
            otr_clusters = [(_T14_05_BASE_KEEP, var_name) for _, var_name in otr_clusters]
        adsl_study_vars = get_adsl_study_specific_variables(path_adam_spec) if path_adam_spec and os.path.isfile(path_adam_spec) else set()
        for base, var_name in otr_clusters:
            if var_name and var_name.upper() not in adsl_study_vars:
                print("05_trt 变量 %s 不在 ADSL Study_specific='Y' 中，仍生成该簇" % var_name)
    else:
        print("Casebook 未提供或未解析到 DSEOT 观测，05_trt 簇数为 0")

    '''
    if os.path.isfile(path_t141_11):
        print("WARNING:: t14_1-1_1.xlsx 已存在,不产生新文件。" )
        return path_t141_11
    '''

    d = os.path.dirname(path_t141_11)
    if d:
        os.makedirs(d, exist_ok=True)
    rows = build_t14_1_1_1_rows(randfl_enrlfl_flags, dct_reasons, followup_reasons, screen_fail_reasons, otr_clusters)
    write_tadsl_disp_xlsx(path_t141_11, rows)
    print("t14_1-1_1.xlsx已生成。共 %d 行：%s" % (len(rows), path_t141_11))
    return path_t141_11


if __name__ == "__main__":  
    
    #path_t141_11 = sys.argv[1]
    #path_adam_spec = sys.argv[2]
    #path_edc_code = sys.argv[3]
    #path_edc_casebook = sys.argv[4]
    path_t141_11 = r'.\SHR1905_202\t14_1-1_1.xlsx'
    path_adam_spec = r'.\SHR1905_202\ADAM_PDS.xlsx'
    path_edc_code = r'.\SHR1905_202\EDCDEF_code.sas7bdat'
    path_edc_casebook = r'.\SHR1905_202\edcdef_casebook.sas7bdat'
    
    run_tadsl_disp_init(path_t141_11, path_adam_spec, path_edc_code, path_edc_casebook)
    
