# -*- coding: utf-8 -*-
"""淘宝自动登录（手机号 + 短信验证码）
验证码来源：MySQL sms_receiver.sms_codes 表（短信接收程序写入，脚本轮询读取）
登录失败时返回 False，由 main.py 负责发送提醒邮件。
"""
import imaplib
import json
import re
import time
from email.header import decode_header
from email import message_from_bytes
from email.utils import parsedate_to_datetime


def _decode_header_str(s):
    """解码 MIME 编码的邮件主题（=?utf-8?b?...?=）"""
    if not s:
        return ""
    out = ""
    for part, enc in decode_header(s):
        if isinstance(part, bytes):
            try:
                out += part.decode(enc or "utf-8", errors="replace")
            except LookupError:
                out += part.decode("utf-8", errors="replace")  # unknown-8bit 等未知编码兜底
        else:
            out += part
    return out


def _decode_email_payload(part):
    """解码单个 MIME part 的正文为 str。

    注意：不能用 get_payload(decode=True)——Python 3.13 email 库对无
    Content-Transfer-Encoding 的 UTF-8 中文会把正文 backslashreplace 成 \\uXXXX 字面。
    这里用 decode=False 拿原始内容，再按需手动解 base64 / quoted-printable。
    """
    cte = (part.get("Content-Transfer-Encoding") or "").lower().strip()
    payload = part.get_payload(decode=False)
    if payload is None:
        return ""
    if isinstance(payload, bytes):
        raw = payload
    elif cte in ("base64", "quoted-printable"):
        raw = payload.encode("ascii", errors="backslashreplace")
    else:
        return str(payload)  # 无编码：str 原样返回
    if cte == "base64":
        try:
            import base64
            raw = base64.b64decode(raw)
        except Exception:
            return str(payload)
    elif cte == "quoted-printable":
        import quopri
        raw = quopri.decodestring(raw)
    try:
        charset = part.get_content_charset() or "utf-8"
        return raw.decode(charset, errors="replace")
    except LookupError:
        return raw.decode("utf-8", errors="replace")


def _strip_html(html):
    """粗略剥离 HTML 标签，保留可见文本（提取验证码用）"""
    import html as html_mod
    html = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", html)
    html = re.sub(r"(?s)<[^>]+>", " ", html)
    return html_mod.unescape(html)


def _get_email_text(msg):
    """提取邮件正文为纯文本（递归处理 multipart）。

    text/plain 优先；无 text/plain 时退回 text/html 并剥离标签
    （第三方短信转发 App 的邮件正文常是 HTML 格式，只有 html 部分）。
    """
    if msg.is_multipart():
        plain_parts = []
        html_parts = []
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == "text/plain":
                t = _decode_email_payload(part)
                if t:
                    plain_parts.append(t)
            elif ct == "text/html":
                t = _decode_email_payload(part)
                if t:
                    html_parts.append(t)
        if plain_parts:
            return "\n".join(plain_parts)
        if html_parts:
            return "\n".join(_strip_html(h) for h in html_parts)
        return ""
    ct = msg.get_content_type()
    t = _decode_email_payload(msg)
    if ct == "text/html":
        return _strip_html(t)
    return t


def get_code_from_email(cfg, start_ts_ms, poll_seconds=120):
    """轮询邮箱获取短信验证码（第三方短信转发 App 把验证码短信转发到邮箱）。

    cfg.email: {sender, auth_code, imap_host, imap_port, code_subject_keywords}
    规则：只处理未读邮件（最多看最新 100 封），在「点获取验证码之后收到」的
    未读邮件中，筛选**主题含 code_subject_keywords** 的，取其中**时间最近的一封**
    提取验证码（主题匹配防止误取其他邮件，时间排序保证取最新的验证码邮件）；
    用正则从主题+正文提取（优先"验证码"后数字，兜底取 4~6 位完整数字块）；
    消费后标记已读，避免重复返回。返回验证码字符串（纯数字）或 None。
    """
    mail_cfg = cfg.get("email", {})
    user = mail_cfg.get("sender") or mail_cfg.get("user")
    auth = mail_cfg.get("auth_code")
    if not user or not auth:
        print("[登录] 未配置邮箱（config.json -> email.sender / email.auth_code），无法从邮箱获取验证码")
        return None
    host = mail_cfg.get("imap_host", "imap.qq.com")
    port = int(mail_cfg.get("imap_port", 993))
    keywords = mail_cfg.get("code_subject_keywords") or ["淘宝登录验证码"]
    deadline = time.time() + poll_seconds
    while time.time() < deadline:
        try:
            conn = imaplib.IMAP4_SSL(host, port, timeout=15)
            try:
                conn.login(user, auth)
                conn.select("INBOX")
                typ, data = conn.search(None, "UNSEEN")
                if typ != "OK":
                    continue
                nums = data[0].split()
                if not nums:
                    conn.logout()
                    time.sleep(3)
                    continue
                # 只看最新 100 封未读，批量取 INTERNALDATE + 主题/日期头（避免拉全文太慢）
                recent = nums[-100:]
                typ, d = conn.fetch(
                    ",".join(n.decode() for n in recent),
                    "(INTERNALDATE BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])",
                )
                if typ != "OK":
                    continue
                # 收集「点获取验证码之后」收到且主题含关键词的未读邮件
                candidates = []
                for block in reversed(d):
                    if not isinstance(block, tuple) or len(block) < 2:
                        continue
                    hdr = message_from_bytes(block[1])
                    subject = _decode_header_str(hdr.get("Subject", ""))
                    dt = None
                    try:
                        raw_block = b" ".join(
                            x if isinstance(x, bytes) else str(x).encode() for x in block
                        )
                        mm = re.search(rb'INTERNALDATE "([^"]+)"', raw_block)
                        if mm:
                            dt = parsedate_to_datetime(mm.group(1).decode("utf-8", errors="replace"))
                    except Exception:
                        dt = None
                    if dt is None:
                        try:
                            dt = parsedate_to_datetime(hdr.get("Date", ""))
                        except Exception:
                            dt = None
                    if dt is None or dt.timestamp() * 1000 < start_ts_ms:
                        continue
                    # 日志：打印每封时间窗内邮件的主题和时间（匹配/不匹配都打，便于诊断）
                    matched = any(k in subject for k in keywords)
                    print(f"[邮箱] 时间={dt.strftime('%Y-%m-%d %H:%M:%S')} "
                          f"主题={subject[:40]}{' [候选]' if matched else ' [跳过:主题不匹配]'}")
                    if not matched:
                        continue
                    num = block[0].split()[0].decode()
                    candidates.append((dt, num, subject))
                if not candidates:
                    conn.logout()
                    time.sleep(3)
                    continue
                # 取时间最近的一封（不依赖主题关键词，转发 App 主题可能变化）
                candidates.sort(key=lambda x: x[0], reverse=True)
                dt, num, subject = candidates[0]
                typ2, d2 = conn.fetch(num, "(BODY.PEEK[])")
                if typ2 != "OK":
                    continue
                msg = message_from_bytes(d2[0][1])
                text = subject + "\n" + _get_email_text(msg)
                m = re.search(r"验证码[^\d]{0,8}(\d{4,8})", text)
                if not m:
                    m = re.search(r"(?<!\d)(\d{4,6})(?!\d)", text)
                conn.store(num, "+FLAGS", "\\Seen")
                if m:
                    code = m.group(1)
                    print(f"[登录] 从邮箱获取到验证码: {code} (主题: {subject[:40]})")
                    return code
                # 提取失败：打印正文摘要，便于诊断转发邮件的实际格式
                print(f"[邮箱] 未能提取验证码: 主题={subject[:40]!r} 正文前200字={text[:200]!r}")
                conn.logout()
            except Exception as e:
                print(f"[登录] 处理邮件失败: {e}")
                try:
                    conn.logout()
                except Exception:
                    pass
        except Exception as e:
            print(f"[登录] 连接邮箱失败: {e}")
        time.sleep(3)
    print(f"[登录] 轮询 {poll_seconds}s 超时，邮箱未出现验证码邮件（请检查短信转发 App 是否正常运行）")
    return None


_AD_CLOSE_SELECTORS = [
    '[class*="close"]', '[class*="Close"]', '[class*="icon-close"]',
    '[aria-label="关闭"]', '[role="button"][aria-label*="close" i]',
    'text=关闭', 'text=×', 'text=✕', 'text=知道了', 'text=跳过', 'text=跳过广告',
]


def close_ads(page):
    """登录成功/搜索页常弹广告浮层，尝试点击关闭按钮。返回是否执行了关闭操作"""
    for sel in _AD_CLOSE_SELECTORS:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click(timeout=1500)
                page.wait_for_timeout(400)
                return True
        except Exception:
            continue
    return False


def taobao_login(page, cfg):
    """淘宝手机号+验证码登录。成功返回 True，失败返回 False
    流程：打开登录页 -> 填手机号 -> 点获取验证码 -> 轮询数据库 sms_codes 拿验证码 -> 填码 -> 提交
    -> 关闭广告弹窗后确认登录态
    """
    tb = cfg.get("taobao", {})
    phone = tb.get("phone")
    if not phone:
        print("[登录] 未配置淘宝手机号（config.json -> taobao.phone），无法自动登录")
        return False

    try:
        page.goto("https://login.taobao.com/member/login.jhtml", timeout=60000,
                  wait_until="domcontentloaded")
        page.wait_for_timeout(4000)
        # 切到短信登录 tab（探测确认存在 "短信登录" 标签）
        try:
            page.click('a.sms-login-tab-item', timeout=5000)
            page.wait_for_timeout(1000)
        except Exception:
            pass  # 可能默认就在短信登录

        # 填手机号
        page.fill('#fm-sms-login-id', phone)
        # 勾选协议
        try:
            page.check('#fm-agreement-checkbox', timeout=3000)
        except Exception as e:
            print(f"[登录] 勾选协议失败: {e}")
        # 点"获取验证码"（探测确认：a.send-btn-link）
        start_ts_ms = int(time.time() * 1000)
        try:
            page.click('a.send-btn-link', timeout=5000)
        except Exception:
            page.click('text=获取验证码', timeout=5000)
        print(f"[登录] 已请求验证码（手机号 {phone[:3]}****{phone[-4:]}），等待验证码邮件到达邮箱...")

        # 从邮箱轮询验证码（第三方短信转发 App -> 邮件）
        code = get_code_from_email(cfg, start_ts_ms, poll_seconds=cfg.get("code_wait_seconds", 120))
        if not code:
            return False
        page.fill('#fm-smscode', code)
        # 提交登录（探测确认：button.fm-button.fm-submit）
        try:
            page.click('button.fm-button.fm-submit', timeout=5000)
        except Exception:
            page.keyboard.press("Enter")
        # 等待跳转/登录生效（登录成功常弹广告浮层：先关广告再确认登录态）
        # 判定信号：unb cookie 是 HttpOnly，document.cookie 读不到——
        # 必须组合「URL 离开登录页 / nick 昵称元素」等信号，否则登录成功也会误判失败
        for _ in range(15):
            page.wait_for_timeout(2000)
            close_ads(page)
            info = page.evaluate("""() => {
                const body = document.body ? document.body.innerText : '';
                const hasUnb = document.cookie.includes('unb=');
                const nickEl = document.querySelector('.site-nav-login-info-nick');
                const hasNick = !!nickEl && nickEl.innerText.trim() !== '';
                const onLoginPage = /login\\.taobao\\.com|havana/.test(location.href);
                const errText = /验证码错误|验证码不正确|验证码已过期|验证码失效|操作频繁|请稍后重试/.test(body);
                const slider = /拖动滑块|请完成验证|安全验证/.test(body);
                return {
                    url: location.href.slice(0, 100),
                    onLoginPage,
                    hasUnb,
                    hasNick,
                    errText,
                    slider,
                    bodyHead: body.slice(0, 80)
                };
            }""")
            print("[登录] 提交后状态:", json.dumps(info, ensure_ascii=False)[:300])
            if info["errText"]:
                print("[登录] 页面提示验证码错误/操作异常，判定登录失败")
                return False
            if info["slider"]:
                print("[登录] 检测到滑块/安全验证，请在弹出的 Edge 窗口手动完成后再试")
                return False
            # 成功：已离开登录页（登录成功必然跳转）或出现已登录标志
            if not info["onLoginPage"] or info["hasUnb"] or info["hasNick"]:
                print("[登录] 淘宝登录成功")
                return True
            # 兜底：若登录后被广告/活动页截留（url 已离开淘宝域），回首页重新确认
            if "taobao.com" not in page.url:
                try:
                    page.goto("https://www.taobao.com/", timeout=30000, wait_until="domcontentloaded")
                except Exception:
                    pass
        print("[登录] 提交验证码后未确认到登录态，判定登录失败")
        return False
    except Exception as e:
        print(f"[登录] 自动登录异常: {e}")
        return False
