"""离线验证网页数据、跨周状态及 HTML 数据转义。"""
import json
import re
import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook
from 生成网页 import build_site


class WebsiteTests(unittest.TestCase):
    def test_public_data_and_safe_embedding(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            wb = Workbook()
            ws = wb.active
            ws.title = '提交统计'
            ws.append(['姓名', '学号', '第1周', '第2周', '第3周', '备注'])
            malicious = '</script><script>alert(1)</script>'
            ws.append([malicious, '0123456789', '已交', None, '已交', 'PRIVATE_NOTE'])
            ws.append(['另一同学', '1123456789', None, None, None, ''])
            ws.append(['PRIVATE_STUDENT_NAME', malicious, None, None, None, ''])
            wb.create_sheet('邮件明细').append(['PRIVATE_EMAIL', 'PRIVATE_ATTACHMENT'])
            excel = root / 'stats.xlsx'
            wb.save(excel)
            wb.close()
            result = build_site(excel, root / 'site', '<示例课程>')
            self.assertEqual(result['weeks'], [1, 2, 3])
            self.assertEqual(result['students'][0]['submitted'], [1, 3])
            self.assertEqual(result['students'][1]['submitted'], [])
            self.assertEqual(result['students'][0]['sid'], '0123456789')
            html = (root / 'site/index.html').read_text(encoding='utf-8')
            self.assertNotIn(malicious, html)
            self.assertNotIn('PRIVATE_', html)
            self.assertNotIn('另一同学', html)
            self.assertTrue(all(set(student) == {'sid', 'submitted'} for student in result['students']))
            self.assertNotIn('__HOMEWORK_DATA__', html)
            self.assertNotIn('__COURSE_TITLE__', html)
            self.assertIn('&lt;示例课程&gt;', html)
            embedded = re.search(r'<script id="homework-data" type="application/json">(.*?)</script>', html, re.S)
            self.assertEqual(json.loads(embedded.group(1)), result)


if __name__ == '__main__':
    unittest.main()
