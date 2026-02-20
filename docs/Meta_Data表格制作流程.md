# Meta Data 表格制作流程

## 一、基础信息

| 项目 | 说明 |
|:---|:---|
| 文件位置 | 指定路径下的 metadata 文件 |
| 文件名 | T14.1-1.2.xlsx |
| 核心要求 | 必须包含固定的几列（列名已固定） |
| 表格结构 | 分为4个部分：01部分、04部分、05部分、06部分 |

## 二、数据查找路径

### ADaM数据集说明文件位置
下拉框1 → 下拉框2 → 下拉框3 → 下拉框4 → utility → documents → adam数据集说明文件(excel 格式)

### EDC定义数据集位置
下拉框4 → metadata → EDCDEF_code (SAS数据集)

### EDCDEF_code 表结构（与 EDC 一致）

程序读取 EDCDEF_code 时依赖以下列（列名不区分大小写，支持常见变体）：

| 列名 | 用途 | 说明 |
|:---|:---|:---|
| **CODE_NAME_CHN** | 必选 | 代码中文名称，用于区分不同原因类型（如「治疗结束主要原因」「筛选结束原因」） |
| **CODE_LABEL** | 必选 | 具体原因文本，写入 T14 的 TEXT 列 |
| **CODE_ORDER** 或 **CODE_ORDER_R** | 可选 | 排序顺序，升序；缺省时按行顺序 |
| CODE_GRP / CODE_GRP_CHN / CODE_NAME / CODE_RAW_VALI / CODE_NAME_RAV | 不参与提取 | 表中可存在，程序不读取 |

同一类原因由多行组成：**CODE_NAME_CHN** 相同，每行一个 **CODE_LABEL**，**CODE_ORDER** 决定先后。例如「治疗结束主要原因」多行对应：已完成、不良事件、失访、死亡、受试者要求终止…；**筛选失败原因**需在表中存在 **CODE_NAME_CHN** =「筛选结束原因」或「筛选失败原因」或「筛选原因」的若干行，其 **CODE_LABEL** 将用于 01 部分第 2 行后的「筛选失败原因」子行。

## 三、各部分内容制作步骤

### 第一部分：01部分（筛选/入组信息）

| 行号 | 内容 | 赋值方式 | 特殊说明 |
|:---|:---|:---|:---|
| 第1-2行 | 固定文本 | 直接硬写 | 见下表列结构及取值 |
| 第2行后 | 筛选失败原因 | 来自 EDCDEF | 标题行 + 各原因子行，TEXT/顺序取自 CODE_NAME_CHN=「筛选结束原因」的 CODE_LABEL/CODE_ORDER |
| 第3行 | 动态判断 | 条件赋值 | 根据ADSL数据集说明文件中，是否存在RANDFL变量决定 |

#### 固定列结构（必含列名）

**列名（从左到右）：** `TEXT` | `MASK` | `LINE_BREAK` | `INDENT` | `SEC` | `TRT_I` | `DSNIN` | `TRTSUBN` | `TRTSUBC` | `FILTER`

#### 第1-2行各列取值

| 行 | TEXT | MASK | LINE_BREAK | INDENT | SEC | TRT_I | DSNIN | TRTSUBN | TRTSUBC | FILTER |
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
| 第1行 | 筛选受试者 | | | | 01_scr | | adsl | trt01pn | trt01p | `prxmatch('/^(合计\|total)\s*$/i', trt01p)` |
| 第2行 | 筛选失败受试者 | | | | 01_scr | | adsl | trt01pn | trt01p | `prxmatch('/^(合计\|total)\s*$/i', trt01p) and (scfailfl='Y')` |

- **第1行**：筛选受试者（Screened Subjects），FILTER 为合计/Total 匹配。
- **第2行**：筛选失败受试者（Screening Failed Subjects），FILTER 在上一行基础上增加 `scfailfl='Y'` 条件。

#### 筛选失败原因（第2行之后、第3行之前）

本块紧接第2行，在「筛选成功为随机/入组受试者」（第3行）之前插入：先一行**标题行**「筛选失败原因」，再按顺序若干行**具体原因**。

**数据来源：** EDCDEF_code 数据集（与 EDC 中表结构一致，见上文「EDCDEF_code 表结构」）。  
- 筛选条件：表中需有若干行 **CODE_NAME_CHN** 为 **「筛选结束原因」**、**「筛选失败原因」** 或 **「筛选原因」**（任一会参与本块；若 EDC 中暂无，可在 EDCDEF_code 中增加该类别及多行 CODE_LABEL）。  
- **TEXT 列**：取该条件下的 `CODE_LABEL` 值（如：不良事件、不符合入选标准/符合排除标准、失访、受试者要求退出、其他）。  
- **排列顺序**：按 `CODE_ORDER`（或 `CODE_ORDER_R`）升序。

| 行类型 | TEXT | INDENT | SEC | DSNIN | TRTSUBN | TRTSUBC | FILTER |
|:---|:---|:---|:---|:---|:---|:---|:---|
| 标题行 | 筛选失败原因 | 空 | 01_scr | adsl | trt01pn | trt01p | 0 |
| 各原因行 | CODE_LABEL 值 | 1 | 01_scr | adsl | trt01pn | trt01p | `prxmatch(...) and (scfailfl='Y') and SCFAILRE='该行原因文本'` |

- 标题行 FILTER 为 `0`，表示仅作分组标题。
- 各原因行 FILTER 在「筛选失败受试者」条件基础上增加 `SCFAILRE='[该行 TEXT]'`，与 ADSL 中筛选失败原因变量一致。

#### 第10行+3行 / 第14行+3行（按 RANDFL/ENRLFL）

**检查路径：** variables sheet → ADSL 数据集（Dataset 列等于 "ADSL"）→ 查找 **Variable 列 = "RANDFL" 且 Study Specific 列 = "Y"**，以及 **Variable 列 = "ENRLFL" 且 Study Specific 列 = "Y"**。**RANDFL 与 ENRLFL 均需满足 Study Specific = Y 才输出对应块**。若两者都满足，则先输出 RANDFL 块（第10行+3行），再输出 ENRLFL 块（第14行+3行）。

**RANDFL 块（第10行+3行，仅当 Variable=RANDFL 且 Study Specific=Y 时）：**

| 行 | TEXT | INDENT | SEC | FILTER |
|:---|:---|:---|:---|:---|
| 第10行 | 筛选成功未随机受试者 | 空 | 01_scr | `prxmatch('/^(合计\|total)\s*$/i', trt01p) and (scfailfl='N') and randfl='N'` |
| 第11行 | 随机受试者 | 1 | 04_rnd | `randfl='Y' and scfailfl='N'` |
| 第12行 | 随机未接受研究治疗 | 1 | 04_rnd | `randfl='Y' and scfailfl='N' and saffl='N'` |
| 第13行 | 随机且接受研究治疗 | 1 | 04_rnd | `randfl='Y' and scfailfl='N' and saffl='Y'` |

**ENRLFL 块（第14行+3行，仅当 Variable=ENRLFL 且 Study Specific=Y 时）：**

| 行 | TEXT | INDENT | SEC | FILTER |
|:---|:---|:---|:---|:---|
| 第14行 | 筛选成功未入组受试者 | 空 | 01_scr | `prxmatch('/^(合计\|total)\s*$/i', trt01p) and (scfailfl='N') and enrlfl='N'` |
| 第15行 | 入组受试者 | 1 | 04_rnd | `enrlfl='Y' and scfailfl='N'` |
| 第16行 | 入组未接受研究治疗 | 1 | 04_rnd | `enrlfl='Y' and scfailfl='N' and saffl='N'` |
| 第17行 | 入组且接受研究治疗 | 1 | 04_rnd | `enrlfl='Y' and scfailfl='N' and saffl='Y'` |

以上 DSNIN/TRTSUBN/TRTSUBC 同 01 部分（adsl, trt01pn, trt01p）。

### 第二部分：04部分（随机/入组信息）

每个 RANDFL 或 ENRLFL 块包含 3 行（随机受试者/入组受试者、未接受研究治疗、且接受研究治疗），SEC=04_rnd，FILTER 见上表。仅当 Variables sheet 中 ADSL 下对应变量存在且 **Study Specific 列 = Y** 时才输出该块；若 RANDFL 与 ENRLFL 均满足，则先输出 RANDFL 的 3 行，再输出 ENRLFL 的 3 行。

### 第三部分：05部分（完成/终止研究治疗）

#### 前 3 行 TEXT 动态生成规则

前 3 行 TEXT **不固定**，由 ADaM Variables sheet 中 **指定数据集下指定变量（治疗结束状态）的 Variable Label** 动态生成。

**可配置宏（代码 `tfls_metadata.py` 05 部分）：** `_T14_05_DATASET`（默认 "ADSL"）、`_T14_05_VAR_EOTSTT`（默认 "EOTSTT"），修改后即可更换数据集名与变量名。

1. **数据来源：** Variables sheet → Dataset=*_T14_05_DATASET* → Variable=*_T14_05_VAR_EOTSTT* → 取「Variable Label」/「标签」列的值（如「治疗结束状态」）。未找到时使用默认「治疗结束状态」。
2. **生成规则：** 从标签中**保留「治疗」两字**作为 base（去掉其余字样，不固定后缀），则：
   - **第 1 行 TEXT** = 「完成研究」+ base，FILTER = `saffl='Y' and [变量名]='[第1行TEXT]'`
   - **第 2 行 TEXT** = 「终止研究」+ base，FILTER = `saffl='Y' and [变量名]='[第2行TEXT]'`
   - **第 3 行 TEXT** = 「终止研究」+ base + 「原因」，FILTER = 0（标题行）

| 行 | LINE_BREAK | INDENT | SEC | DSNIN | TRTSUBN | TRTSUBC |
|:---|:---|:---|:---|:---|:---|:---|
| 第1行 | 1 | 空 | 05_trt | adsl | trt01pn | trt01p |
| 第2行 | 空 | 空 | 05_trt | adsl | trt01pn | trt01p |
| 第3行 | 空 | 空 | 05_trt | adsl | trt01pn | trt01p |

#### 终止原因详细行（第 4 行起，来自 EDCDEF）

- **数据来源：** EDCDEF_code，CODE_NAME_CHN =「治疗结束主要原因」或「治疗结束原因」，按 CODE_ORDER 顺序取 CODE_LABEL。
- **每行：** TEXT = CODE_LABEL 值，INDENT = 1，LINE_BREAK = 空，SEC/DSNIN/TRTSUBN/TRTSUBC 同前 3 行。
- **FILTER：** `saffl='Y' and EOTSTT='[第2行TEXT]' and DCTREAS='该行原因文本'`（与 ADSL 变量 DCTREAS 一致，字符串内单引号需双写）。

**对应变量名：** EOTSTT（治疗结束状态）、DCTREAS（研究治疗终止的具体原因）

### 第四部分：06部分（随访结束/完成研究）

#### 核心概念映射
- 完成随访 = 完成研究
- 退出随访 = 退出研究

#### 结构说明
| 行类型 | 内容 | 特殊要求 |
|:---|:---|:---|
| 基础行 | 完成随访/退出随访 | 同05部分逻辑 |
| 原因详述行 | 退出原因 = 0 | 来自EDC数据集 |
| 特殊行 | 随机未接受研究治疗 | 每个原因后额外增加一行 |

#### 06部分结构
- 原因1
  - 随机未接受研究治疗（对应原因1）
- 原因2
  - 随机未接受研究治疗（对应原因2）
- 原因3...
  - 随机未接受研究治疗（对应原因3）

#### 制作步骤
1. 打开 EDCDEF_code SAS数据集
2. 查找变量：CODE_NAME_CHN = 随访结束原因 或 随访结束主要原因 或 研究结束原因 或 原因结束主要原因
3. 按CODE_ORDER值的顺序提取原因（原因即CODE_LABEL列的值）
4. 每个原因后增加一行"随机未接受研究治疗"
5. Filter条件：在上一条原因的filter条件基础上，额外增加筛选条件 `RANDFL='Y' and TRTSDT NE .`

## 四、关键变量对照表

| 显示名称 | 数据集变量名 | 所在数据集 |
|:---|:---|:---|
| 随机受试者标记 | RANDFL | ADSL |
| 研究治疗终止原因，即治疗结束原因 | CODE_LABEL | EDCDEF-code 数据集，且CODE_NAME_CHN = 治疗结束原因 或 治疗结束主要原因 |
| 随访结束原因 | CODE_LABEL | EDCDEF-code 数据集，且CODE_NAME_CHN = 随访结束原因 或 随访结束主要原因 或 研究结束原因 或 原因结束主要原因 |

## 五、制作检查清单

- [ ] 确认ADSL数据集中是否存在RANDFL变量（Study Specific = Y）
- [ ] 01部分第3行根据判断结果正确赋值
- [ ] 04部分三行直接硬写（仅当有随机变量时）
- [ ] 05部分原因按EDC数据集order顺序排列
- [ ] 06部分每个原因后增加"随机未接受研究治疗"行
- [ ] 06部分filter条件正确叠加
- [ ] 所有列名使用固定命名规范