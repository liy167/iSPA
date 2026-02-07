import os
import saspy
import re

# 指定自定义的 sascfg.py 路径
cfg_path = r'C:\Users\yangy102\.config\sascfg.py'

# 创建SAS会话并使用自定义配置文件
sas = saspy.SASsession(cfgfile=cfg_path)

# 测试连接
print(sas)

# 设置批处理模式
sas.set_batch(True)

# 定义SAS宏程序的路径和SAS主程序的路径
macro_file_path = '/u01/app/sas/sas9.4/DocumentRepository/DDT/projects/utility/macros/01_general/autorun.sas'
sas_file_path = '/u01/app/sas/sas9.4/DocumentRepository/DDT/projects/HRS9432/HRS9432_201/csr_01/06_programs/061_data/test.sas'

# 提取文件名（不包括路径和扩展名），只取纯文件名
sas_file_name_no_ext = os.path.splitext(os.path.basename(sas_file_path))[0]  # 'test'

# 检查 sas_file_path 中是否包含 '/06_programs' 或 '/09_validation'
if '/06_programs' in sas_file_path:
    base_path = os.path.dirname(sas_file_path).split('/06_programs')[0]  # 保留到 '/csr_01'
elif '/09_validation' in sas_file_path:
    base_path = os.path.dirname(sas_file_path).split('/09_validation')[0]  # 保留到 '/csr_01'
else:
    raise ValueError("sas_file_path 中既不包含 '/06_programs' 也不包含 '/09_validation'")

# 构造完整的日志文件路径，使用拼接
log_output_path = f"{base_path}/07_logs/{sas_file_name_no_ext}.log"

print(f"日志保存于 {log_output_path} 。\n")

# 构造包含 %INCLUDE 语句和日志重定向的SAS代码
sas_code = f"""
proc printto log='{log_output_path}' new;
run;

%let _sasprogramfile = '{sas_file_path}';
%include '{macro_file_path}';
%include '{sas_file_path}';

proc printto; /* 恢复日志输出到默认位置 */
run;
"""

# 提交SAS代码
sas_output = sas.submit(sas_code)

# 修改路径为 Windows 格式
def convert_linux_path_to_windows(linux_path):
    windows_path = linux_path.replace('/u01/app/sas/sas9.4/DocumentRepository/DDT', 'Z:')
    windows_path = windows_path.replace('/', '\\')  # 将斜杠替换为反斜杠
    return windows_path

# 检查日志文件中的错误信息
def check_for_errors_in_log(log_file_path):
    # 将路径转换为 Windows 格式
    windows_log_file_path = convert_linux_path_to_windows(log_file_path)
    
    try:
        with open(windows_log_file_path, 'r', encoding='utf-8') as log_file:  # 使用 utf-8 编码读取文件
            log_content = log_file.read()
            print(log_content)  # 打印日志文件内容进行检查

            # 改进正则表达式，考虑到可能有空格或其它字符在 ERROR 后面
            error_match = re.search(r'ERROR\s*\d*-*\d*:?', log_content, re.IGNORECASE)
            
            if error_match:
                print(f"找到错误: {error_match.group()}")  # 打印匹配到的错误
                return True
            else:
                return False
    except FileNotFoundError:
        print(f"日志文件 {windows_log_file_path} 不存在！")
        return False
    except UnicodeDecodeError as e:
        print(f"读取日志文件时出现编码错误: {e}")
        return False
    
# 检查日志文件是否包含ERROR
if check_for_errors_in_log(log_output_path):
    print(f"SAS程序 {sas_file_path} 执行时出现错误！")
else:
    print(f"SAS程序 {sas_file_path} 执行成功。")

# 输出SAS日志路径
print(f"SAS日志文件路径：{log_output_path}")
