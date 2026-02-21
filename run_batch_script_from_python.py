"""
解析批处理 .sas 脚本中的 %batch_submit()，根据 role/target/pgm 推导出 SAS 程序路径，
再使用 linux_sas_call_from_python.run_sas 依次执行这些程序。
"""
import argparse
import os
import re
import sys

# role -> 顶层目录名
ROLE_DIR = {
    'developer': '06_programs',
    'validator': '09_validation',
}

# (role, target) -> 子目录名
TARGET_DIR = {
    ('developer', 'data'): '061_data',
    ('developer', 'safety'): '062_safety',
    ('developer', 'efficacy'): '063_efficacy',
    ('developer', 'pkpd'): '064_pkpd',
    ('developer', 'stats'): '065_stats',
    ('validator', 'data'): '091_data',
    ('validator', 'safety'): '092_safety',
    ('validator', 'efficacy'): '093_efficacy',
    ('validator', 'pkpd'): '094_pkpd',
    ('validator', 'stats'): '095_stats',
}

# 单行内 %batch_submit( role=xxx, target=yyy, pgm=zzz , ... )
BATCH_SUBMIT_RE = re.compile(
    r'%batch_submit\s*\(\s*role\s*=\s*(\w+)\s*,\s*target\s*=\s*(\w+)\s*,\s*pgm\s*=\s*(\w+)',
    re.IGNORECASE
)


def parse_batch_submits(batch_script_path: str) -> list[tuple[str, str, str]]:
    """
    从批处理脚本中解析出所有 %batch_submit(role=..., target=..., pgm=...)，
    返回 [(role, target, pgm), ...]，保持顺序、不去重。
    """
    with open(batch_script_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    found = []
    for m in BATCH_SUBMIT_RE.finditer(content):
        role, target, pgm = m.group(1).lower(), m.group(2).lower(), m.group(3)
        found.append((role, target, pgm))
    return found


def build_sas_paths(base_path: str, submits: list[tuple[str, str, str]]) -> list[str]:
    """
    根据 (role, target, pgm) 列表和 base_path 拼出完整 .sas 路径列表。
    未知 role/target 会抛出 ValueError。
    """
    base_path = base_path.strip().rstrip('/\\').replace('\\', '/')
    paths = []
    for role, target, pgm in submits:
        if role not in ROLE_DIR:
            raise ValueError(f"未知 role={role}，允许: {list(ROLE_DIR)}")
        key = (role, target)
        if key not in TARGET_DIR:
            raise ValueError(f"未知 target={target}（role={role}），允许: data|safety|efficacy|pkpd|stats")
        role_dir = ROLE_DIR[role]
        target_dir = TARGET_DIR[key]
        # 统一用 / 拼接，run_sas 内会按需转换
        path = f"{base_path}/{role_dir}/{target_dir}/{pgm}.sas"
        paths.append(path)
    return paths


def main():
    parser = argparse.ArgumentParser(
        description='解析批处理脚本中的 %batch_submit，推导 SAS 程序路径并调用 run_sas 执行。'
    )
    parser.add_argument(
        'batch_script',
        help='批处理 .sas 文件路径，如 05_batch_script_tfl_dev.sas 或 Z:\\...\\05_batch_script_tfl_dev.sas',
    )
    parser.add_argument(
        '--base-path',
        required=True,
        help='包含 06_programs/09_validation 的项目目录，如 Z:\\projects\\HRS2129\\HRS2129_test\\csr_01',
    )
    args = parser.parse_args()

    batch_script_path = os.path.normpath(args.batch_script)
    if not os.path.isfile(batch_script_path):
        print(f"错误：批处理文件不存在: {batch_script_path}")
        sys.exit(1)

    submits = parse_batch_submits(batch_script_path)
    if not submits:
        print("未找到任何 %batch_submit(...)，退出。")
        sys.exit(1)

    try:
        paths = build_sas_paths(args.base_path, submits)
    except ValueError as e:
        print(f"错误：{e}")
        sys.exit(1)

    print("解析到的 SAS 程序及路径：")
    for i, (role, target, pgm) in enumerate(submits, 1):
        print(f"  {i}. role={role}, target={target}, pgm={pgm} -> {paths[i-1]}")
    print()

    # 复用 linux_sas_call_from_python 的 run_sas（同目录导入）
    from linux_sas_call_from_python import run_sas
    import saspy

    if len(paths) == 1:
        run_sas(paths[0])
        return

    sas = saspy.SASsession(cfgname='winiomlinux')
    try:
        for i, sas_file_path in enumerate(paths, 1):
            print(f"\n[{i}/{len(paths)}] 执行: {sas_file_path}")
            run_sas(sas_file_path, sas_session=sas, check_log=False)
        print(f"\n全部 {len(paths)} 个 SAS 程序已提交执行。")
    finally:
        sas.endsas()


if __name__ == '__main__':
    main()
