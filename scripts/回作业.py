import os
import re
import ssl
import time
import smtplib
import random
from collections import defaultdict
from email.message import EmailMessage
from typing import List, Dict

# ============ 可配置 ============
SMTP_SERVER = "smtp.qq.com"
SMTP_PORT_SSL = 587
SLEEP_BETWEEN_SENDS = 10   # 频控间隔（秒）
DRY_RUN = False             # 设 True 先演练，不发信
GROUP_BY_STUDENT = True     # True: 同一学号的多份PDF合并一封邮件；False: 每个PDF发一封
FOLDER_DEFAULT = "./第十五周作业批改"   # 默认PDF文件夹
SUBJECT = "第十五周作业批改"
BODY = "同学你好，附件是你本周的作业批改，请查收\n\n祝好，\n孙谌劼"
SENDER = os.getenv("HOMEWORK_EMAIL", "")
AUTH_CODE = os.getenv("HOMEWORK_SMTP_AUTH_CODE") or os.getenv("HOMEWORK_AUTH_CODE", "")
# ==============================

ID_PATTERN = re.compile(r"(\d{10})")

def extract_id(name: str) -> str:
    m = ID_PATTERN.search(name)
    return m.group(1) if m else ""

def recipient_from_id(sid: str) -> str:
    return f"{sid}@stu.pku.edu.cn"

def list_pdf_files(folder: str) -> List[str]:
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"找不到文件夹：{folder}")
    return [os.path.join(folder, f) for f in os.listdir(folder)
            if os.path.isfile(os.path.join(folder, f)) and f.lower().endswith(".pdf")]

# 放在文件顶部 import 后
from openpyxl import load_workbook

def mark_excel_cell_for_week(
    sid: str,
    excel_path: str = "./25fall作业.xlsx",
    sheet_name: str | None = None,
    id_header: str = "学号",
    week_header: str = "第十五周",
    value: str = "已交"
):
    if not os.path.exists(excel_path):
        print(f"[Excel] 未找到文件：{excel_path}（跳过标记）")
        return
    wb = load_workbook(excel_path)
    ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) else wb.active

    # 表头映射
    headers = { (cell.value or ""): idx for idx, cell in enumerate(ws[1], start=1) }
    if id_header not in headers or week_header not in headers:
        print(f"[Excel] 缺少列：{id_header} 或 {week_header}（跳过标记）")
        return
    id_col, week_col = headers[id_header], headers[week_header]

    # 找到学号所在行
    target_row = None
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=id_col).value).strip() == str(sid):
            target_row = r
            break
    if target_row is None:
        print(f"[Excel] 未找到学号 {sid}（跳过标记）")
        return

    ws.cell(row=target_row, column=week_col, value=value)
    wb.save(excel_path)
    print(f"[Excel] 标记完成：学号 {sid} → {week_header} = {value}")


def build_messages(sender: str, pdf_paths: List[str], subject: str, body: str) -> List[EmailMessage]:
    messages: List[EmailMessage] = []

    if GROUP_BY_STUDENT:
        bucket: Dict[str, List[str]] = defaultdict(list)
        for p in pdf_paths:
            sid = extract_id(os.path.basename(p))
            if not sid:
                print(f"[跳过] 未找到10位学号：{os.path.basename(p)}")
                continue
            bucket[sid].append(p)

        for i, (sid, files) in enumerate(bucket.items(), start=1):
            to_addr = recipient_from_id(sid)
            msg = EmailMessage()
            msg["From"] = sender
            msg["To"] = to_addr
            msg["Subject"] = subject
            msg.set_content(body)

            for fp in files:
                with open(fp, "rb") as f:
                    msg.add_attachment(
                        f.read(),
                        maintype="application",
                        subtype="pdf",
                        filename=os.path.basename(fp)
                    )
            msg.add_header("X-Attachment-Count", str(len(files)))
            msg.add_header("X-Student-ID", sid)
            messages.append(msg)
            print(f"[准备] {to_addr} | 附件数: {len(files)}")
            mark_excel_cell_for_week(sid)
    return messages

def _send_once(msg):
    context = ssl.create_default_context()
    server = None
    try:
        server = smtplib.SMTP("smtp.qq.com", 587, timeout=60)
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(SENDER, AUTH_CODE)
        server.send_message(msg)     # 成功发出会在这里返回；失败会在这里抛错
        return True
    finally:
        if server is not None:
            try:
                server.quit()        # QQ 有时在这里直接掐连接
            except (smtplib.SMTPServerDisconnected,
                    smtplib.SMTPResponseException,
                    OSError):
                pass                  # 安全忽略退出异常

def send_messages(sender: str, auth_code: str, messages: List[EmailMessage]):
    for i, msg in enumerate(messages, start=1):
        to_addr = msg["To"]
        success = False
        if DRY_RUN:
            print(f"[DRY-RUN] ({i}/{len(messages)}) -> {to_addr}")
            success = True
            continue
        _send_once(msg)
        print(f"[发送] ({i}/{len(messages)}) -> {to_addr}")
        time.sleep(SLEEP_BETWEEN_SENDS + random.uniform(-2, 2))  # 防止频率过快
def main():
    print("=== QQ 邮箱批量发送 PDF → 学号@stu.pku.edu.cn ===")
    sender = SENDER
    auth_code = AUTH_CODE
    folder = FOLDER_DEFAULT
    subject = SUBJECT
    body = BODY

    pdfs = list_pdf_files(FOLDER_DEFAULT)
    if not pdfs:
        print("[提示] 文件夹中没有 .pdf 文件。")
        return

    messages = build_messages(sender, pdfs, subject, body)
    if not messages:
        print("[提示] 没有可发送的邮件（可能文件名里都没有10位学号）。")
        return

    print(f"将发送 {len(messages)} 封邮件。DRY_RUN={DRY_RUN}, GROUP_BY_STUDENT={GROUP_BY_STUDENT}")
    send_messages(sender, auth_code, messages)
    print("全部完成。")

if __name__ == "__main__":
    main()
