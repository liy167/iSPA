# TFLs PDT 功能文档

本文档记录 SASEG 中「生成PDT」功能的实现约定与沟通结论，用于存储记忆与后续维护参考。

---

## 1. 功能概述

在 TFLs 页面点击「生成PDT」按钮，弹出分步对话框。用户选择设计类型与终点后，程序基于 TOC_template.xlsx 与 setup.xlsx，更新项目层面 PDT 的 Deliverables sheet 中 Category=Output 的行。

---

## 2. 技术路线与流程

```
输入：TOC_template.xlsx (PH1) + 项目PDT + setup.xlsx + 用户选择
    ↓
1. 备份 PDT → 99_archive/{文件名}_{时间戳}.xlsx
    ↓
2. 从 setup.xlsx 读取 LNG（B列=LNG 时取 C列）
    ↓
3. 从 TOC PH1 sheet 读取行
    ↓
4. 按用户选择筛选行、按设计类型展开
    ↓
5. 删除 PDT Deliverables 中 Category=Output 的原有行
    ↓
6. 追加新行，并应用数据验证
```

---

## 3. 文件与路径约定

| 文件 | 路径/位置 |
|------|-----------|
| TOC_template.xlsx | Z:\projects\utility\template\TOC_template.xlsx（默认） |
| 项目PDT | {当前路径}\utility\documentation\{p3}_{p4}_PDT.xlsx |
| setup.xlsx | 与 PDT 同目录 |
| 备份目录 | 与 PDT 同目录下的 99_archive |

---

## 4. TOC 读取约定

- **仅使用 PH1 sheet**，无 fallback 到其他 sheet
- 列名：Template#, Output Type, Title_CN, Title_EN, Population, Footnotes_CN, Footnotes_EN, Category_CN, SAD, FE, MAD, BE, MB
- 支持别名：Footnote_CN → Footnotes_CN, Footnote_EN → Footnotes_EN

---

## 5. LNG 与 setup.xlsx

- **LNG 读取**：setup.xlsx 的 Macro Variables sheet，B 列值="LNG" 时，取 C 列值
- **用途**：决定 Title/Footnotes 使用中文还是英文列
  - LNG 为 CHN/CN/Chinese/中文/ZH 等 → 用 Title_CN、Footnotes_CN
  - 否则 → 用 Title_EN、Footnotes_EN
- **括号**：Title 后追加设计类型时，LNG 为中文用「（）」，英文用「()」

---

## 6. PDT Deliverables 列映射（PDT_COL_ALIASES）

| PDT 实际列名 | 逻辑列名 |
|--------------|----------|
| OUTCAT | Category |
| OUTTYPE | Output Type |
| OUTREF | Output Reference |
| OUTTITLE | Title |
| OUTPOP | Population |
| OUTFNOTE | Footnotes |
| PGMLEVEL | Validation Level |
| USERDEV | Developers |
| USERQC | Validators |
| OUTSTS | Output Status |
| STASCHK | Validated by Programmer/Statistician |

---

## 7. Output Reference 后缀与 Title

- **仅选一个设计类型时**：Output Reference 与 Title 均不加后缀（例：14.1-1.1、受试者分布）
- **选多个设计类型时**：
  - **后缀**：SAD→a, FE→b, MAD→c, BE→d, MB→e  
    例：14.1-1.1a, 14.1-1.1b
  - **Title**：在原有 Title 后追加设计类型括号  
    中文：受试者分布（SAD）；英文：Subject Disposition (SAD)

---

## 8. 新行初始值与数据验证

| 列 | 初始值 | 数据验证来源 |
|----|--------|--------------|
| Validation Level | Non-critical | List Values!$E$2:$E$4 |
| Developers | Gang Cheng | - |
| Validators | Yi Yang | - |
| Output Status | （无初始值） | List Values!$F$2:$F$3 |
| Validated by Programmer/Statistician | Not Started | List Values!$G$2:$G$4 |

---

## 9. 问题3 分析物占位符替换

- **占位符**：Title_CN 中的 `<Analyte分析物>`，Title_EN 中的 `<Analyte>`
- **问题3 有值时**：用 `|` 分割多个分析物（如 `HRS2129|M1`），每个分析物输出一行
  - Title：将占位符替换为对应分析物
  - Output Reference：追加 `_{分析物}`（如 `14.1-1.1a_HRS2129`、`14.1-1.1a_M1`）
- **问题3 无值时**：占位符替换为空字符串

---

## 10. 行筛选规则

- **基准行**：排除 Category_CN ∈ {QT分析, C-QT分析, PK浓度, PK参数, PD分析, ADA分析}（除非用户勾选对应终点）
- **PK浓度/PK参数 血/尿/粪 过滤**（Title_CN/Title_EN 不含未选子类型字眼）：
  - 选(血)：仅保留无「尿」「粪」及 Urine、Feces、Stool 等的观测
  - 选(尿)：仅保留无「血」「粪」及 Blood、Feces、Stool 等的观测
  - 选(粪)：仅保留无「血」「尿」及 Blood、Urine 等的观测
  - 多选时：仅排除未选子类型的字眼
- **终点额外添加**：勾选 PK浓度(血) 等 → 加入 Category_CN=PK浓度；勾选 PK参数(血) 等 → 加入 Category_CN=PK参数；PD/ADA/QT 同理
- **设计类型展开**：TOC 中 SAD/FE/MAD/BE/MB 列非空且被选中 → 每类型生成一行

---

## 11. 弹窗与界面约定

- **成功时**：保留主弹窗不关闭，仅关闭成功提示框，便于继续操作
- **测试阶段默认路径**：前 4 个下拉框默认 projects、HRS2129、HRS2129_test、csr_01

---

## 12. 公式自动刷新

保存 PDT 时设置：
- `fullCalcOnLoad = True`：打开时执行全量重算
- `calcMode = "auto"`：自动计算模式
- `calcId` 递增：使 Excel 认为计算链已变更，触发重算

避免因 openpyxl 修改后公式显示 #VALUE! 需手动双击才能刷新的问题。

---

## 13. 相关代码文件

| 文件 | 说明 |
|------|------|
| tfls_pdt.py | 弹窗 UI 与 on_ok 调用逻辑 |
| tfls_pdt_update.py | PDT 更新核心：备份、读 TOC/setup、筛选、写 Deliverables、数据验证 |

---

## 14. 更新记录

- 2025-02：初始整理，基于 TFLs_PDT 沟通结论
