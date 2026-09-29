"""按学号将批改 PDF 发回真实发件邮箱；默认预览，--send 才发送。"""
import argparse
from collections import defaultdict
from datetime import datetime
from email.message import EmailMessage
from email import policy
from email.utils import getaddresses, make_msgid, formatdate
import hashlib
import json
import os
from pathlib import Path
import re
import smtplib
import ssl
import time

from openpyxl import load_workbook

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def save_json(path, data):
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    temporary.replace(path)


def load_config():
    path = PROJECT_ROOT / '邮箱配置.local.ps1'
    if path.exists():
        for line in path.read_text(encoding='utf-8-sig').splitlines():
            m = re.fullmatch(r"\s*\$env:(HOMEWORK_[A-Z_]+)\s*=\s*'((?:[^']|'')*)'\s*", line)
            if m:
                os.environ.setdefault(m[1], m[2].replace("''", "'"))
    sender = os.getenv('HOMEWORK_EMAIL', '')
    code = os.getenv('HOMEWORK_SMTP_AUTH_CODE') or os.getenv('HOMEWORK_AUTH_CODE', '')
    if not sender or not code:
        raise ValueError('请在本地邮箱配置中填写邮箱和授权码')
    return sender, code


def addresses(value):
    return {address.lower() for _, address in getaddresses([str(value or '').replace('；', ',').replace(';', ',')])
            if re.fullmatch(r'[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+', address)}


def make_plan(root, folder, week):
    wb = load_workbook(root / '学生邮箱.xlsx', read_only=True, data_only=True)
    try:
        rows = iter(wb['学生邮箱'].values)
        headers = list(next(rows))
        students = {}
        for row in rows:
            sid = str(row[headers.index('学号')])
            students[sid] = (str(row[headers.index('姓名')] or ''), addresses(row[headers.index('邮箱（多个用分号分隔）')]), str(row[headers.index('备注')] or ''))
    finally:
        wb.close()
    state = json.loads((root / '收取记录.json').read_text(encoding='utf-8'))
    origins = {hashlib.sha256(k.encode()).hexdigest()[:20]: r for k, r in state.items()}
    groups = defaultdict(list)
    problems = []
    for path in sorted(folder.rglob('*')):
        if not path.is_file() or path.suffix.lower() != '.pdf' or path.name.startswith('._') or '__MACOSX' in path.parts:
            continue
        m = re.match(r'^(\d{10})_', path.name)
        sid = m[1] if m else ''
        if sid not in students:
            problems.append({'file': str(path), 'reason': '文件名未以名单学号开头'})
            continue
        name, emails, notes = students[sid]
        if not path.name.startswith(f'{sid}_{name}_'):
            problems.append({'file': str(path), 'reason': '文件姓名与邮箱名单不一致'})
            continue
        if '需确认' in notes:
            problems.append({'file': str(path), 'reason': '邮箱表标记需确认：' + notes})
            continue
        tokens = re.findall(r'_([a-f0-9]{20})_', path.name)
        origin = origins.get(tokens[0]) if len(tokens) == 1 else None
        recipient = next(iter(emails)) if len(emails) == 1 else None
        basis = '邮箱表唯一地址'
        if origin:
            original_emails = addresses(origin.get('sender'))
            if origin.get('sid') != sid or origin.get('number') != week or len(original_emails) != 1:
                problems.append({'file': str(path), 'reason': '原邮件标识与学生或周次不一致'})
                continue
            original_email = next(iter(original_emails))
            if len(emails) > 1 and original_email not in emails:
                problems.append({'file': str(path), 'reason': '原邮件发件地址与邮箱表不一致'})
                continue
            if len(emails) > 1:
                recipient = original_email
                basis = '按附件原邮件标识匹配发件地址'
        if not recipient:
            problems.append({'file': str(path), 'reason': '邮箱缺失或多邮箱无法确定'})
            continue
        with path.open('rb') as stream:
            if b'%PDF-' not in stream.read(1024):
                problems.append({'file': str(path), 'reason': '文件不是可识别的PDF'})
                continue
        groups[(sid, name, recipient, basis)].append(path)
    # 同一学生同一收件地址合并附件，不受匹配方式不同影响。
    merged = defaultdict(list)
    for (sid, name, recipient, basis), files in groups.items():
        merged[(sid, name, recipient)].extend(files)
    plan = []
    for (sid, name, recipient), files in sorted(merged.items()):
        items = [{'path': str(p.resolve()), 'name': p.name, 'sha256': hashlib.sha256(p.read_bytes()).hexdigest(), 'bytes': p.stat().st_size} for p in files]
        plan.append({'sid': sid, 'name': name, 'recipient': recipient, 'week': week, 'files': items})
    return plan, problems


def message_for(item, sender):
    msg = EmailMessage()
    msg['From'] = sender
    msg['To'] = item['recipient']
    msg['Subject'] = f"第{item['week']}周作业批改"
    msg['Date'] = formatdate(localtime=True)
    msg['Message-ID'] = make_msgid()
    msg.set_content(f"{item['name']}同学你好：\n\n附件是你的第{item['week']}周作业批改，请查收。\n如有疑问，请直接回复这封邮件。\n\n祝好，\n孙谌劼")
    for entry in item['files']:
        content = Path(entry['path']).read_bytes()
        if hashlib.sha256(content).hexdigest() != entry['sha256']:
            raise ValueError('附件在预览后被修改，请重新运行')
        msg.add_attachment(content, maintype='application', subtype='pdf', filename=entry['name'])
    return msg


def delivery_key(item, sender):
    identity = {'sender': sender, 'sid': item['sid'], 'recipient': item['recipient'], 'week': item['week'],
                'files': sorted((entry['name'], entry['sha256']) for entry in item['files'])}
    return hashlib.sha256(json.dumps(identity, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


def send_plan(plan, root, week, interval):
    sender, code = load_config()
    log_path = root / f'第{week}周返还记录.json'
    log = json.loads(log_path.read_text(encoding='utf-8')) if log_path.exists() else {}
    sent = skipped = failed = 0
    for item in plan:
        key = delivery_key(item, sender)
        old = log.get(key, {})
        if old.get('status') in ('已发送', '发送中', '结果不确定'):
            print(f"[跳过] {item['sid']}：{old['status']}，避免重复发送", flush=True)
            skipped += 1
            continue
        server = None
        record = {**item, 'status': '准备', 'time': datetime.now().astimezone().isoformat()}
        try:
            msg = message_for(item, sender)
            raw = msg.as_bytes(policy=policy.SMTP)
            server = smtplib.SMTP('smtp.qq.com', 587, timeout=90)
            server.ehlo()
            server.starttls(context=ssl.create_default_context())
            server.ehlo()
            limit = int(server.esmtp_features.get('size', '0') or 0)
            if limit and len(raw) > limit:
                raise ValueError(f'邮件大小 {len(raw)} 超过服务器限制 {limit} 字节')
            server.login(sender, code)
            record['status'] = '发送中'
            record['message_id'] = str(msg['Message-ID'])
            log[key] = record
            save_json(log_path, log)
            server.sendmail(sender, [item['recipient']], raw)
            record['status'] = '已发送'
            sent += 1
            print(f"[已发送] {item['sid']} {item['name']}，{len(item['files'])} 个PDF", flush=True)
        except (smtplib.SMTPException, OSError, ValueError) as exc:
            # 服务器明确拒绝可重试；发送中断则先查邮箱，避免重复投递。
            uncertain = record['status'] == '发送中' and not isinstance(exc, (smtplib.SMTPResponseException, smtplib.SMTPRecipientsRefused, smtplib.SMTPSenderRefused))
            record['status'] = '结果不确定' if uncertain else '失败'
            record['error'] = str(exc)
            failed += 1
            print(f"[{record['status']}] {item['sid']}：{exc}", flush=True)
            if isinstance(exc, smtplib.SMTPAuthenticationError):
                log[key] = record
                save_json(log_path, log)
                raise
        finally:
            if record['status'] != '准备':
                log[key] = record
                save_json(log_path, log)
            if server:
                try:
                    server.quit()
                except (smtplib.SMTPException, OSError):
                    server.close()
        time.sleep(interval)
    print(f'完成：新发送 {sent} 封，跳过 {skipped} 封，失败或待核对 {failed} 封。', flush=True)
    return failed


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--week', type=int, required=True)
    parser.add_argument('--output', default='作业收集_2026秋')
    parser.add_argument('--folder', help='默认使用数据目录下的第N周作业批改')
    parser.add_argument('--send', action='store_true', help='实际发送；默认仅生成返还清单')
    parser.add_argument('--interval', type=float, default=10)
    args = parser.parse_args()
    if args.week < 1 or args.interval < 0:
        parser.error('周次须为正数，间隔不能为负数')
    root = Path(args.output).resolve()
    folder = Path(args.folder).resolve() if args.folder else root / f'第{args.week}周作业批改'
    if not folder.is_dir():
        parser.error(f'找不到批改文件夹：{folder}')
    plan, problems = make_plan(root, folder, args.week)
    save_json(root / f'第{args.week}周返还清单.json', {'messages': plan, 'needs_review': problems})
    print(f'返还清单：{len(plan)}封邮件，{sum(len(p["files"]) for p in plan)}个PDF，{len(problems)}个文件需确认。', flush=True)
    for problem in problems:
        print(f"[待确认] {Path(problem['file']).name}：{problem['reason']}", flush=True)
    if args.send:
        return 1 if send_plan(plan, root, args.week, args.interval) else 0
    print('预览完成，未发送。加 --send 执行；不会修改作业提交统计。')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
