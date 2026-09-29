"""从本地收取记录整理学生邮箱，以作业统计为名单，不连接邮箱或发送邮件。"""
import argparse
from collections import defaultdict
from email.utils import getaddresses
import json
from pathlib import Path
import re
import unicodedata

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


def match_student(record, students):
    subject = unicodedata.normalize('NFKC', record.get('subject', ''))
    if re.match(r'\s*(re\s*:|fw\s*:|fwd\s*:|回复|答复|转发|发信方已撤回)', subject, re.I):
        return '', '回复、转发或撤回通知，未自动匹配'
    ids = set(re.findall(r'(?<!\d)\d{10}(?!\d)', subject))
    if record.get('sid'):
        ids.add(record['sid'])
    if len(ids) > 1:
        return '', '出现多个学号，需人工确认'
    if ids:
        sid = next(iter(ids))
        if sid not in students:
            return '', '学号不在统计名单中'
        name = record.get('name', '').strip()
        if name and name != students[sid]:
            return sid, '按学号匹配；邮件姓名与名单不一致，需确认'
        return sid, '按学号匹配' if name else '从非标准主题提取学号，需确认'
    names = [sid for sid, name in students.items() if name and name in subject]
    if names:
        return '', '仅姓名匹配，需确认：' + '；'.join(f'{sid} {students[sid]}' for sid in names)
    return '', '未找到可匹配的学号'


def organize(root):
    root = Path(root)
    source = load_workbook(root / '作业统计.xlsx', read_only=True, data_only=True)
    try:
        rows = iter(source['提交统计'].values)
        headers = list(next(rows))
        students = {str(row[headers.index('学号')]).strip(): str(row[headers.index('姓名')] or '').strip()
                    for row in rows if row[headers.index('学号')] is not None}
    finally:
        source.close()
    state = json.loads((root / '收取记录.json').read_text(encoding='utf-8'))
    matches = defaultdict(list)
    unmatched = []
    details = []
    for record in state.values():
        sid, method = match_student(record, students)
        addresses = sorted({addr.strip().lower() for _, addr in getaddresses([record.get('sender', '')])
                            if re.fullmatch(r'[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+', addr.strip())})
        if len(addresses) != 1:
            sid = ''
            method += '；发件人地址缺失或有多个，需确认'
        line = [sid, students.get(sid, ''), '；'.join(addresses), method, record.get('subject', ''),
                record.get('sender', ''), record.get('date', ''), record.get('status', '')]
        details.append(line)
        if sid:
            matches[sid].append(line)
        else:
            unmatched.append(line)
    email_students = defaultdict(set)
    for sid, records in matches.items():
        for record in records:
            email_students[record[2]].add(sid)
    for line in details:
        if line[0] and len(email_students[line[2]]) > 1:
            line[3] += '；邮箱关联多个学号，需确认'
    unmatched = [line for line in details if not line[0] or '需确认' in line[3]]
    wb = Workbook()
    ws = wb.active
    ws.title = '学生邮箱'
    ws.append(['学号', '姓名', '邮箱（多个用分号分隔）', '邮箱数量', '匹配邮件数', '备注'])
    for sid, name in students.items():
        records = matches[sid]
        addresses = sorted({r[2] for r in records})
        notes = []
        if not addresses:
            notes.append('本地收取记录未找到邮箱')
        if len(addresses) > 1:
            notes.append('使用过多个邮箱，未自动选择主邮箱')
        if any(len(email_students[email]) > 1 for email in addresses):
            notes.append('邮箱关联多个学号，需确认')
        notes.extend(sorted({r[3] for r in records if '需确认' in r[3]}))
        ws.append([sid, name, '；'.join(addresses), len(addresses), len(records), '；'.join(notes)])
    headers = ['匹配学号', '名单姓名', '发件邮箱', '匹配说明', '邮件主题', '原始发件人', '邮件时间', '收取状态']
    for title, rows in [('来源明细', details), ('待人工核对', unmatched)]:
        sheet = wb.create_sheet(title)
        sheet.append(headers)
        for row in rows:
            sheet.append(row)
    for sheet in wb:
        sheet.freeze_panes = 'C2'
        sheet.auto_filter.ref = sheet.dimensions
        for row in sheet:
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = 's'
        for cell in sheet[1]:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='4472C4')
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = min(70, max(16, max(len(str(c.value or '')) for c in column) + 2))
    target = root / '学生邮箱.xlsx'
    temporary = target.with_name('学生邮箱.tmp.xlsx')
    wb.save(temporary)
    wb.close()
    temporary.replace(target)
    print(f'已保存：{target.resolve()}')
    print(f'名单 {len(students)} 人，找到邮箱 {sum(bool(v) for v in matches.values())} 人；待人工核对邮件 {len(unmatched)} 封。')
    return target


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', default='作业收集_2026秋', help='本地作业数据目录')
    args = parser.parse_args()
    organize(args.output)
