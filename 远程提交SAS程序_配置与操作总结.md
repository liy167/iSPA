# 远程提交 SAS 程序：配置与操作总结

## 一、配置流程（一次性）

### 1. 安装 SASPy

```bash
pip install saspy
```

### 2. 准备配置文件

- **推荐**：复制 saspy 自带的 `sascfg.py` 为 **`sascfg_personal.py`**，只改 personal 文件，避免升级被覆盖。
- **放置位置**（任选其一）：
  - saspy 安装目录
  - 当前工作目录
  - **`~/.config/saspy/`**（Windows：`%USERPROFILE%\.config\saspy\`）
  - 或任意在 Python 搜索路径中的目录；若不在，则在代码里用 **`cfgfile=`** 指定完整路径，例如：  
    `C:\Users\liy167\.config\sascfg_personal.py`

### 3. 在配置文件里写“连接定义”

- 在配置文件中：
  - 在 **`SAS_config_names`** 里写上配置名（如 `'linux_sas'`）。
  - 定义一个**同名变量**，赋值为一个**字典**，里面是该连接方式的参数。

**示例（Windows 本机 → Linux 服务器 SAS，SSH 方式）：**

```python
SAS_config_names = ['linux_sas']

linux_sas = {
    'saspath'  : '/u01/app/sas/sas9.4/SASHome/SASFoundation/9.4/bin/sas_u8',  # 服务器上 SAS 可执行路径
    'ssh'      : 'C:\\Windows\\System32\\OpenSSH\\ssh.exe',                   # 本机 ssh 路径
    'host'     : 'SAS服务器的主机名或IP',
    'options'  : ['-fullstimer'],
    # 若 Linux 用户名与 Windows 不同：
    # 'luser'    : 'linux_username',
    # 若用密钥认证：
    # 'identity': r'C:\Users\liy167\.ssh\id_rsa',
    # 若 SSH 端口非 22：
    # 'port'     : 22,
}
```

- **`cfgname='linux_sas'`** 的含义：使用配置文件中名为 **`linux_sas`** 的这组设置建会话。

### 4. 保证能 SSH 到服务器

- 本机执行：`ssh 用户名@SAS服务器主机名`（或使用上面配置的 `host`、`luser`、`identity`）能免密登录，或已配置好 sshpass 等，以便 saspy 通过 SSH 在服务器上启动 SAS。

---

## 二、操作流程（每次提交 SAS 程序）

### 1. 路径约定（若从 Windows 选路径）

- **本机（Windows）**：例如 `Z:\projects\...\utility\tools\25_generate_pdt_call.sas`
- **服务器（Linux）**：例如 `/u01/app/sas/sas9.4/DocumentRepository/DDT/projects/.../utility/tools/25_generate_pdt_call.sas`
- **转换规则**：  
  - `Z:` → `/u01/app/sas/sas9.4/DocumentRepository/DDT`（按你们实际挂载点调整）  
  - `\` → `/`

### 2. 建会话并提交代码

```python
import saspy

# 建会话（使用上面配置的 linux_sas）
sas = saspy.SASsession(cfgname='linux_sas')
# 若配置文件不在默认路径：
# sas = saspy.SASsession(cfgfile=r'C:\Users\liy167\.config\sascfg_personal.py', cfgname='linux_sas')

# 要执行的 .sas 在服务器上的路径（若从 Windows 路径转换，先按上面规则得到 linux_path）
sas_script_linux = '/u01/app/sas/sas9.4/DocumentRepository/DDT/utility/tools/25_generate_pdt_call.sas'

# 提交：在服务器上执行 %include
sas.set_batch(True)  # 可选，批处理模式
result = sas.submit('%include "' + sas_script_linux + '"; ')

# 可选：查看返回的 LOG
print(result['LOG'])

# 关闭会话
sas.endsas()
```

### 3. 与「初版PDT」的衔接

- **配置**：在 `sascfg_personal.py` 里配好 **`linux_sas`**（或你用的配置名），并保证 `cfgfile`/`cfgname` 指向它。
- **操作**：点击「初版PDT」时，用前四个下拉框拼出 `25_generate_pdt_call.sas` 的 Windows 路径 → 转成 Linux 路径 → 用 saspy **`SASsession(cfgname='linux_sas')`** + **`submit('%include "' + linux_path + '"; ')`** → 可选检查 LOG → **`endsas()`**。

---

## 三、流程对照表

| 阶段       | 配置流程（一次性）                         | 操作流程（每次）                                       |
|------------|--------------------------------------------|--------------------------------------------------------|
| 环境       | 安装 saspy；准备 sascfg_personal.py       | 无                                                    |
| 连接定义   | 在配置里写 SAS_config_names + linux_sas 等 | 无                                                    |
| 连接方式   | SSH：填 saspath / ssh / host 等           | 无                                                    |
| 网络/登录  | 保证本机可 SSH 到 SAS 服务器              | 无                                                    |
| 路径       | 无                                        | Windows 路径 → 按规则转成 Linux 路径（若需要）        |
| 提交       | 无                                        | SASsession(cfgname='linux_sas') → submit('%include "…";') |
| 收尾       | 无                                        | 查看 result['LOG']，endsas()                         |

---

## 四、参考代码（完整流程）

以下为从建会话到提交、关会话的完整示例（配置名使用 `linux_sas`，脚本为 `25_generate_pdt_call.sas`）：

```python
import saspy

sas = saspy.SASsession(cfgname='linux_sas')
# 若配置文件路径固定，可写：
# sas = saspy.SASsession(cfgfile=r'C:\Users\liy167\.config\sascfg_personal.py', cfgname='linux_sas')

autoexec = '/u01/app/sas/sas9.4/DocumentRepository/DDT/utility/tools/25_generate_pdt_call.sas'
ps = sas.submit('%include "' + autoexec + '"; ')

# 可选：查看提交返回的 LOG
print(ps['LOG'])

sas.endsas()
```

说明：本文档不包含日志重定向（proc printto）与本地日志文件错误检查，仅通过 `submit()` 返回的 `LOG` 查看结果。

---

## 五、Windows 路径转 Linux 路径（Python 示例）

在「初版PDT」等场景中，界面给出的是 Windows 路径（如 `Z:\projects\...\utility\tools\25_generate_pdt_call.sas`），提交到服务器前需转为 Linux 路径：

```python
def convert_windows_path_to_linux(win_path):
    """将 Windows 路径（Z:\...）转为服务器 Linux 路径。"""
    if not win_path:
        return win_path
    # Z: 对应服务器挂载根路径（按实际环境修改）
    linux_base = '/u01/app/sas/sas9.4/DocumentRepository/DDT'
    s = win_path.strip()
    if s.upper().startswith('Z:\\') or s.upper().startswith('Z:/'):
        s = linux_base + s[2:]  # 去掉 Z:
    s = s.replace('\\', '/')
    return s

# 示例
win_path = r'Z:\projects\utility\tools\25_generate_pdt_call.sas'
linux_path = convert_windows_path_to_linux(win_path)
# 得到：/u01/app/sas/sas9.4/DocumentRepository/DDT/projects/utility/tools/25_generate_pdt_call.sas
```

---

## 六、cfgname 配置名说明

- **cfgname** 的值（如 `'linux_sas'`）必须与配置文件里**某个变量的名字**一致。
- 该变量是一个**字典**，内容为连接方式（如 SSH）所需键值：`saspath`、`ssh`、`host` 等。
- 该名字还必须出现在配置文件顶部的 **`SAS_config_names`** 列表中，否则无法被选用。
- 若配置文件中只有一个连接定义且已写入 `SAS_config_names`，建会话时可省略 `cfgname`；若有多个，则必须指定 `cfgname='配置名'`。

---

## 七、注意事项

- **saspy 与 SAS 版本**：需 SAS 9.4 或更高（或 Viya），且 saspy 与服务器 SAS 兼容。
- **SSH**：从 Windows 连 Linux 时需本机有 SSH 客户端（如 OpenSSH），并配置免密或 sshpass；`ssh`、`host`、`saspath` 必须正确。
- **路径**：`%include` 中的路径必须是**服务器上可访问的路径**，不能写本机 Windows 路径。
- **资源**：用完后调用 `sas.endsas()` 关闭会话，避免占用服务器资源。
