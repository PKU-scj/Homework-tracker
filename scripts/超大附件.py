"""QQ 超大附件：只访问邮件下载入口及中转站，不执行网页脚本或提交登录信息。"""
import hashlib
import html
import http.client
import json
import re
import shutil
import time
from html.parser import HTMLParser
from http.cookiejar import CookieJar
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, quote, urlencode, urldefrag, urljoin, urlsplit
from urllib.request import HTTPCookieProcessor, HTTPRedirectHandler, Request, build_opener


CHUNK = 1024 * 1024
PAGE_LIMIT = 2 * CHUNK
DEFAULT_LIMIT = 10 * 1024 ** 3


class DownloadProblem(ValueError):
    pass


def allowed_url(url):
    """精确限定主机和下载路径，重定向也应用同一规则。"""
    try:
        parsed = urlsplit(url)
        host = (parsed.hostname or '').lower()
        if (parsed.scheme not in ('http', 'https') or parsed.username or parsed.password
                or parsed.port not in (None, 80, 443) or re.search(r'[\s\\\x00-\x1f]', url)):
            return False
        path = parsed.path.lower()
        return ((host == 'mail.qq.com' or host.endswith('.mail.qq.com'))
                and path in ('/cgi-bin/ftnexs_download', '/cgi-bin/ftndownload', '/cgi-bin/ftndownload302')
                or host == 'wx.mail.qq.com' and path == '/ftn/download'
                or host.endswith('.mail.ftn.qq.com') and bool(re.match(r'/[0-9]+/', path))
                or (host == 'ftn.qq.com' or host.endswith('.ftn.qq.com'))
                and path.startswith('/ftn_handler/'))
    except ValueError:
        return False


def normalize_url(value, base='https://mail.qq.com/'):
    value = html.unescape(value).replace('\\/', '/').replace('\\x26', '&').replace('\\u0026', '&')
    url = urldefrag(urljoin(base, value.strip()))[0]
    return url if allowed_url(url) else None


class LinkParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.links = []
        self.text = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == 'a' and attrs.get('href'):
            self.links.append(attrs['href'])
        if tag == 'meta' and attrs.get('http-equiv', '').lower() == 'refresh':
            match = re.search(r'url\s*=\s*[\"\']?(.+?)[\"\']?$', attrs.get('content', ''), re.I)
            if match:
                self.links.append(match.group(1))

    def handle_data(self, value):
        self.text.append(value)


def extract_links(text, base='https://mail.qq.com/', is_html=False):
    candidates = []
    if is_html:
        parser = LinkParser()
        parser.feed(text)
        candidates.extend(parser.links)
        text = '\n'.join(parser.text)
    text = html.unescape(text).replace('\\/', '/').replace('\\x26', '&').replace('\\u0026', '&')
    candidates.extend(re.findall(r'https?://[^\s<>\"\'）)\]，。；]+', text))
    # 静态字符串形式的跳转；不执行 JS，也不猜测或拼接私有接口参数。
    candidates.extend(re.findall(r'[\"\'](/cgi-bin/ftn[^\"\'\s]+)[\"\']', text))
    return list(dict.fromkeys(url for value in candidates if (url := normalize_url(value, base))))


def find_tasks(msg, key):
    urls, marker = [], False
    for part in msg.walk():
        if part.is_multipart() or part.get_content_type() not in ('text/plain', 'text/html'):
            continue
        payload = part.get_payload(decode=True) or b''
        if len(payload) > PAGE_LIMIT:
            continue
        try:
            text = payload.decode(part.get_content_charset() or 'utf-8', errors='replace')
        except LookupError:
            text = payload.decode('utf-8', errors='replace')
        marker = marker or '超大附件' in text
        urls.extend(extract_links(text, is_html=part.get_content_type() == 'text/html'))
    token = hashlib.sha256(key.encode()).hexdigest()[:20]
    return [dict(id=token + '_' + hashlib.sha256(url.encode()).hexdigest()[:16],
                 url=url, status='待下载', error='', file='')
            for url in dict.fromkeys(urls or ([''] if marker else []))]


class SafeRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if not allowed_url(newurl):
            raise DownloadProblem('下载跳转到登录页或未支持的地址，请在 QQ 邮箱网页下载后补录')
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def valid_pdf(path):
    """轻量检查文件头和结束标记，不把错误页面或明显截断的文件当作 PDF。"""
    path = Path(path)
    if not path.is_file():
        return False
    with path.open('rb') as stream:
        if not stream.read(1024).lstrip().startswith(b'%PDF-'):
            return False
        stream.seek(max(0, path.stat().st_size - 65536))
        return b'%%EOF' in stream.read()


def resolve_modern_download(url, opener, timeout, max_bytes):
    """与 QQ 文件分享页 getFileInfo 一致：func=3 查询，body.url 指向 func=4 文件。"""
    parsed = urlsplit(url)
    if parsed.hostname != 'wx.mail.qq.com' or parsed.path != '/ftn/download':
        return url, None
    query = parse_qs(parsed.query)
    if query.get('func') == ['4']:
        return url, None
    if not all(len(query.get(key, [])) == 1 and query[key][0] for key in ('key', 'code')):
        raise DownloadProblem('QQ 分享链接缺少 key/code，请从原邮件重新复制完整链接')
    params = urlencode(dict(f='json', func=3, key=query['key'][0], code=query['code'][0])).encode()
    request = Request('https://wx.mail.qq.com/ftn/download', data=params, headers={
        'User-Agent': 'Mozilla/5.0', 'Content-Type': 'application/x-www-form-urlencoded',
        'Accept-Encoding': 'identity', 'Referer': 'https://wx.mail.qq.com/ftn/download',
        'Origin': 'https://wx.mail.qq.com',
    })
    with opener.open(request, timeout=timeout) as response:
        if response.status != 200 or not allowed_url(response.geturl()):
            raise DownloadProblem('QQ 文件信息请求未成功，请稍后重试')
        raw = response.read(PAGE_LIMIT + 1)
    if len(raw) > PAGE_LIMIT:
        raise DownloadProblem('QQ 文件信息响应过大')
    try:
        result = json.loads(raw)
    except (ValueError, UnicodeError):
        raise DownloadProblem('QQ 文件信息未返回 JSON，可能需要在网页中验证') from None
    head = result.get('head') if isinstance(result, dict) else None
    if not isinstance(head, dict) or str(head.get('ret')) != '0':
        # 不输出服务端原始消息，以免其中包含分享密钥。
        ret = str(head.get('ret', '未知')) if isinstance(head, dict) else '未知'
        ret = ret if re.fullmatch(r'-?\d{1,10}', ret) else '未知'
        raise DownloadProblem(f'QQ 拒绝了文件信息请求（返回码 {ret}），链接可能已失效或需要网页验证')
    body = result.get('body')
    if not isinstance(body, dict):
        raise DownloadProblem('QQ 文件信息缺少文件数据')
    direct = body.get('url')
    if not isinstance(direct, str) or not allowed_url(direct):
        raise DownloadProblem('QQ 未提供受支持的文件下载地址')
    try:
        size = int(body['size'])
    except (KeyError, ValueError, TypeError):
        raise DownloadProblem('QQ 文件信息缺少有效文件大小') from None
    if size <= 0 or size > max_bytes:
        raise DownloadProblem('文件为空或超过配置的大小上限（--large-max-mb）')
    if Path(str(body.get('name', ''))).suffix.lower() != '.pdf':
        raise DownloadProblem('QQ 分享文件不是 PDF（可能为 ZIP 压缩包），当前仅自动收取 PDF')
    return direct, body


def download_once(url, target, timeout, max_bytes, progress, opener):
    if not allowed_url(url):
        raise DownloadProblem('下载地址不受支持')
    url, metadata = resolve_modern_download(url, opener, timeout, max_bytes)
    seen = set()
    for _ in range(6):
        if url in seen or not allowed_url(url):
            raise DownloadProblem('下载页面循环跳转或地址不受支持，请网页下载后补录')
        seen.add(url)
        request = Request(quote(url, safe=':/?&=%+#;,@!$()*[]~-_'), headers={
            'User-Agent': 'Mozilla/5.0',
            'Accept-Encoding': 'identity',
        })
        if metadata and urlsplit(url).hostname == 'wx.mail.qq.com':
            name, value = metadata.get('cookie_name', ''), metadata.get('cookie_value', '')
            if name or value:
                if not (isinstance(name, str) and isinstance(value, str)
                        and re.fullmatch(r'[A-Za-z0-9_-]+', name)
                        and re.fullmatch(r'[\x21-\x3a\x3c-\x7e]+', value)):
                    raise DownloadProblem('QQ 返回的下载 Cookie 格式不受支持')
                request.add_unredirected_header('Cookie', f'{name}={value}')
        with opener.open(request, timeout=timeout) as response:
            if response.status != 200:
                raise DownloadProblem('服务器未返回完整文件，请网页下载后补录')
            if not allowed_url(response.geturl()):
                raise DownloadProblem('服务器返回未支持的下载地址')
            encoding = response.headers.get('Content-Encoding', 'identity').lower()
            if encoding != 'identity':
                raise DownloadProblem('服务器返回了不支持的传输编码，请网页下载后补录')
            length = response.headers.get('Content-Length')
            expected = int(length) if length is not None else None
            if expected is not None and (expected < 0 or expected > max_bytes):
                raise DownloadProblem('附件超过配置的大小上限（--large-max-mb）')
            first = response.read(1024)
            if first.lstrip().startswith(b'%PDF-'):
                temporary = target.with_suffix('.pdf.part')
                total, last_report = 0, time.monotonic()
                digest = hashlib.sha1()
                try:
                    with temporary.open('wb') as output:
                        chunk = first
                        while chunk:
                            total += len(chunk)
                            if total > max_bytes:
                                raise DownloadProblem('附件超过配置的大小上限（--large-max-mb）')
                            output.write(chunk)
                            digest.update(chunk)
                            if progress and time.monotonic() - last_report >= 5:
                                progress(f'已下载 {total / CHUNK:.1f} MB')
                                last_report = time.monotonic()
                            chunk = response.read(CHUNK)
                    if expected is not None and total != expected:
                        raise OSError('文件传输不完整')
                    if metadata:
                        if total != int(metadata['size']):
                            raise OSError('文件大小与 QQ 文件信息不符')
                        expected_sha = str(metadata.get('sha', ''))
                        if re.fullmatch(r'[0-9a-fA-F]{40}', expected_sha) and digest.hexdigest() != expected_sha.lower():
                            raise OSError('文件校验和与 QQ 文件信息不符')
                    if not valid_pdf(temporary):
                        raise OSError('PDF 缺少结束标记，可能下载不完整')
                    temporary.replace(target)
                    filename = str(metadata.get('name', '')) if metadata else response.headers.get_filename()
                    if not filename:
                        filename = parse_qs(urlsplit(response.geturl()).query).get('fname', ['附件.pdf'])[0]
                    return filename
                finally:
                    temporary.unlink(missing_ok=True)
            content_type = response.headers.get_content_type()
            if content_type not in ('text/html', 'text/plain', 'application/xhtml+xml'):
                raise DownloadProblem('下载结果不是 PDF，当前只收取 PDF 作业')
            page = first + response.read(PAGE_LIMIT + 1)
            if len(page) > PAGE_LIMIT:
                raise DownloadProblem('下载页面过大，无法自动解析')
            charset = response.headers.get_content_charset() or 'utf-8'
            try:
                text = page.decode(charset, errors='replace')
            except LookupError:
                text = page.decode('utf-8', errors='replace')
            links = [link for link in extract_links(text, response.geturl(), True) if link not in seen]
            if not links:
                raise DownloadProblem('页面没有可直接下载的 PDF 入口，可能需要登录、验证码或链接已失效；请网页下载后补录')
            if len(links) != 1:
                raise DownloadProblem('页面存在多个下载入口，无法确定对应文件；请网页下载后补录')
            url = links[0]
    raise DownloadProblem('下载跳转次数过多，请网页下载后补录')


def download_pdf(url, target, timeout=120, max_bytes=DEFAULT_LIMIT, attempts=3, progress=print):
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    opener = build_opener(SafeRedirect(), HTTPCookieProcessor(CookieJar()))
    for attempt in range(attempts):
        try:
            return download_once(url, target, timeout, max_bytes, progress, opener)
        except HTTPError as exc:
            code = exc.code
            exc.close()
            if code not in (408, 429, 500, 502, 503, 504) or attempt == attempts - 1:
                raise DownloadProblem(f'下载服务返回 HTTP {code}，请稍后重试或网页下载后补录') from None
        except (OSError, URLError, http.client.HTTPException):
            if attempt == attempts - 1:
                raise DownloadProblem('下载连接失败或文件不完整，已重试；可再次运行或网页下载后补录') from None
        if progress:
            progress(f'下载中断，重试 {attempt + 2}/{attempts}（重新下载）')
        time.sleep(attempt + 1)


def local_file(root, relative):
    path = (root / relative).resolve()
    if not relative or not path.is_relative_to(root.resolve()):
        raise ValueError('附件记录路径无效')
    return path


def refresh_status(record):
    pending = any(task['status'] != '已下载' for task in record.get('large_attachments', []))
    record['status'] = '已交' if record['files'] else ('超大附件待下载' if pending else '无PDF附件')


def archive_pdf(root, record, task, path, filename):
    """与普通附件一致：邮件标识_附件序号_原始文件名。序号保存到任务避免重跑变化。"""
    token = task['id'].split('_', 1)[0]
    if not task.get('attachment_index'):
        indices = [int(m[1]) for relative in record['files']
                   if (m := re.match(re.escape(token) + r'_(\d+)_', Path(relative).name))]
        indices.extend(t['attachment_index'] for t in record.get('large_attachments', [])
                       if t.get('attachment_index'))
        task['attachment_index'] = max(indices, default=0) + 1
    filename = filename or '附件.pdf'
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename).strip(' .')[:100] or '附件.pdf'
    if not safe.lower().endswith('.pdf'):
        safe = safe[:96] + '.pdf'
    destination = path.parent / f"{token}_{task['attachment_index']}_{safe}"
    if destination != path:
        if destination.exists() and task.get('file') != str(destination.relative_to(root)):
            raise ValueError('同名归档文件已存在，停止改名以避免覆盖')
        path.replace(destination)
    old = task.get('file')
    relative = str(destination.relative_to(root))
    record['files'] = list(dict.fromkeys(relative if p == old else p for p in record['files']))
    if relative not in record['files']:
        record['files'].append(relative)
    task.update(status='已下载', file=relative, original_name=filename,
                size=destination.stat().st_size, error='')
    return destination


def process_tasks(root, record, folder, previous=(), timeout=120, max_bytes=DEFAULT_LIMIT):
    old = {task['id']: task for task in previous}
    for task in record.get('large_attachments', []):
        saved = old.get(task['id'], task)
        for field in ('attachment_index', 'original_name', 'file'):
            if field in saved:
                task[field] = saved[field]
        if saved.get('file'):
            path = local_file(root, saved['file'])
            if valid_pdf(path) and path.stat().st_size == saved.get('size'):
                task.update(saved)
                task['status'] = '已下载'
                if task['file'] not in record['files']:
                    record['files'].append(task['file'])
                continue
            if saved['file'] in record['files']:
                record['files'].remove(saved['file'])
        task.update(status='待下载', error='')
        if not task['url']:
            task['error'] = '邮件提到了超大附件，但未识别到下载链接；请网页下载后补录'
            continue
        target = folder / f"{task['id']}_超大附件.pdf"
        try:
            print(f"[超大附件] {task['id']} 开始下载")
            filename = download_pdf(task['url'], target, timeout=timeout, max_bytes=max_bytes)
            archive_pdf(root, record, task, target, filename)
        except (OSError, ValueError) as exc:
            # 网络异常由 download_pdf 转成不含链接/token 的说明。
            task['error'] = str(exc)
            print(f"[超大附件待处理] {task['id']}：{task['error']}")
    refresh_status(record)


def import_pdf(root, state, task_id, source, folder_for):
    matches = [(record, task) for record in state.values()
               for task in record.get('large_attachments', []) if task['id'] == task_id]
    if len(matches) != 1:
        raise ValueError('未找到唯一的超大附件任务，请从待处理清单复制完整任务编号')
    record, task = matches[0]
    source = Path(source)
    if not valid_pdf(source):
        raise ValueError('补录文件不是 PDF 或缺少 PDF 结束标记')
    folder = folder_for(root, record)
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f'{task_id}_超大附件.pdf'
    temporary = target.with_suffix('.pdf.part')
    try:
        with source.open('rb') as incoming, temporary.open('wb') as outgoing:
            shutil.copyfileobj(incoming, outgoing, CHUNK)
        if not valid_pdf(temporary):
            raise ValueError('补录文件复制后未通过检查')
        temporary.replace(target)
    finally:
        temporary.unlink(missing_ok=True)
    archive_pdf(root, record, task, target, task.get('original_name') or source.name)
    refresh_status(record)
    return record


def write_pending(root, state):
    lines = ['QQ 超大附件待处理清单（仅本地保存，链接可能含下载凭据，请勿公开）',
             '在 QQ 邮箱网页下载 PDF 后，用 --import-large 任务编号 "PDF路径" 补录。', '']
    count = 0
    for record in state.values():
        for task in record.get('large_attachments', []):
            if task['status'] == '已下载':
                continue
            count += 1
            lines.extend([f"任务编号：{task['id']}",
                          f"学生：{record['sid']} {record['name']}；第{record['number']}周",
                          f"主题：{record['subject']}", f"原因：{task['error']}",
                          f"下载入口：{task['url'] or '请在 QQ 邮箱网页打开原邮件'}", ''])
    if not count:
        lines.append('没有待处理的超大附件。')
    path = root / '超大附件待处理.txt'
    temporary = path.with_suffix('.txt.tmp')
    temporary.write_text('\n'.join(lines), encoding='utf-8-sig')
    temporary.replace(path)
    if count:
        print(f'还有 {count} 个超大附件待处理：{path}')
