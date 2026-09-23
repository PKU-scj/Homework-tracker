"""从提交统计生成可直接打开、上传的独立网页，不导出邮件明细。"""
import argparse
import json
import re
from html import escape
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


def build_site(excel, output, title='高等数学 B'):
    excel, output = Path(excel), Path(output)
    wb = load_workbook(excel, read_only=True, data_only=True)
    try:
        rows = iter(wb['提交统计'].values)
        headers = list(next(rows))
        sid_col = headers.index('学号')
        weeks = [(i, int(m.group(1))) for i, h in enumerate(headers)
                 if (m := re.fullmatch(r'第(\d+)周(?:作业)?', str(h)))]
        weeks.sort(key=lambda item: item[1])
        students = []
        for row in rows:
            if row[sid_col] is None:
                continue
            students.append({'sid': str(row[sid_col]),
                             'submitted': [week for i, week in weeks if row[i] == '已交']})
    finally:
        wb.close()
    data = {'updated': datetime.fromtimestamp(excel.stat().st_mtime).astimezone().isoformat(timespec='seconds'),
            'weeks': [week for _, week in weeks], 'students': students}
    # 阻止单元格内容提前关闭 script 标签；页面用 textContent 渲染数据。
    payload = json.dumps(data, ensure_ascii=False).replace('&', '\\u0026').replace('<', '\\u003c').replace('>', '\\u003e')
    template = Path(__file__).with_name('网页模板.html').read_text(encoding='utf-8')
    html = template.replace('__COURSE_TITLE__', escape(title)).replace('__HOMEWORK_DATA__', payload)
    output.mkdir(parents=True, exist_ok=True)
    target = output / 'index.html'
    temporary = output / 'index.tmp.html'
    temporary.write_text(html, encoding='utf-8')
    temporary.replace(target)
    print(f'查询网页已生成：{target.resolve()}（{len(students)} 名学生）')
    return data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--excel', default='作业收集_2026秋/作业统计.xlsx')
    parser.add_argument('--output', default='site', help='GitHub Pages 发布目录，默认 site')
    parser.add_argument('--title', default='高等数学 B', help='网页课程名称')
    args = parser.parse_args()
    try:
        build_site(args.excel, args.output, args.title)
    except (OSError, ValueError, KeyError, StopIteration) as exc:
        parser.exit(1, f'生成失败：{exc}\n')


if __name__ == '__main__':
    main()
