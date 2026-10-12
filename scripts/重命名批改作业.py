"""将批改 PDF 重命名为 学号_姓名_第N周作业批改.pdf。"""
import argparse
import json
from pathlib import Path
import re

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def make_plan(folder, week):
    folder = Path(folder).resolve()
    if not folder.is_dir():
        raise ValueError(f'文件夹不存在：{folder}')
    files = sorted(p for p in folder.rglob('*')
                   if p.is_file() and p.suffix.lower() == '.pdf'
                   and not p.name.startswith('._') and '__MACOSX' not in p.parts)
    reserved = {str(p).casefold() for p in folder.rglob('*')}
    plan, skipped = [], []
    for path in files:
        if path.is_symlink() or not path.resolve().is_relative_to(folder):
            raise ValueError(f'文件指向目录外或为符号链接：{path.name}')
        match = re.match(r'^(\d{10})_([^_]+)_', path.name)
        if not match or not match[2].strip():
            skipped.append(path.name)
            continue
        sid = match[1]
        name = re.sub(r'\s*[（(]\d+[）)]\s*$', '', match[2]).strip()
        if not name:
            skipped.append(path.name)
            continue
        stem = f'{sid}_{name}_第{week}周作业批改'
        if re.fullmatch(re.escape(stem) + r'(?:_[2-9]|_[1-9]\d+)?\.pdf', path.name):
            continue
        target = path.with_name(stem + '.pdf')
        suffix = 2
        while str(target).casefold() in reserved:
            target = path.with_name(f'{stem}_{suffix}.pdf')
            suffix += 1
        reserved.add(str(target).casefold())
        plan.append((path, target))
    return plan, skipped


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--week', type=int, required=True, help='作业周次，例如 4')
    parser.add_argument('--output', type=Path, default=PROJECT_ROOT / '作业收集_2026秋', help='作业数据目录')
    parser.add_argument('--folder', type=Path, help='批改 PDF 文件夹，默认数据目录下的第N周作业批改')
    parser.add_argument('--dry-run', action='store_true', help='仅预览，不修改文件')
    args = parser.parse_args()
    if args.week <= 0:
        parser.error('--week 必须大于 0')
    folder = (args.folder or args.output / f'第{args.week}周作业批改').resolve()
    try:
        plan, skipped = make_plan(folder, args.week)
        for old, new in plan:
            print(f'{old.name} -> {new.name}')
        for name in skipped:
            print(f'[跳过] 无法识别学号和姓名：{name}')
        if args.dry_run:
            print(f'预览：拟重命名 {len(plan)} 份 PDF，跳过 {len(skipped)} 份。')
            return
        # 先保存旧名映射，供恢复文件名及核对原邮件标识使用。
        log_path = folder / '重命名记录.json'
        history = json.loads(log_path.read_text(encoding='utf-8')) if log_path.exists() else []
        if not isinstance(history, list):
            raise ValueError('重命名记录格式不正确')
        for old, new in plan:
            if new.exists():
                raise FileExistsError(f'目标文件已存在：{new.name}')
            entry = {'old': str(old.relative_to(folder)), 'new': str(new.relative_to(folder)), 'week': args.week}
            if entry not in history:
                history.append(entry)
                temporary = log_path.with_suffix('.tmp')
                temporary.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding='utf-8')
                temporary.replace(log_path)
            old.rename(new)
        print(f'完成：重命名 {len(plan)} 份 PDF，跳过 {len(skipped)} 份。')
    except (OSError, ValueError) as exc:
        parser.exit(1, f'运行失败：{exc}\n')


if __name__ == '__main__':
    main()
