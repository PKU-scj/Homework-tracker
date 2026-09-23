"""本课程负责的学号范围（不含下限）。"""
import re


MIN_STUDENT_ID = 2600015446


def in_scope(sid):
    value = str(sid).strip()
    return bool(re.fullmatch(r'[0-9]{10}', value)) and int(value) > MIN_STUDENT_ID
