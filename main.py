# -*- coding: utf-8 -*-
"""
淘宝商品搜索抓取脚本
====================================
功能（仅按需求实现）：
  1. 连接本机 Edge（CDP 调试模式），打开淘宝
  2. 检查登录状态：
     - 未登录 -> 发送邮件到收件人邮箱，退出
     - 已登录 -> 继续
  3. 读取 Excel（需要搜索的商品.xlsx）的工作表名作为搜索关键词
  4. 对每个关键词在淘宝搜索并抓取：商品标题 / 商品链接 / 商品图片 / 商品价格 / 已售件数
  5. 抓取结果写回同一个 Excel 的对应工作表

运行方式：
  .venv\\Scripts\\python.exe main.py

配置：见 config.json（Excel 路径、Edge 调试端口/Profile、SMTP 邮件、抓取数量等）
"""
import json
import os
import re
import smtplib
import subprocess
import sys
import time
import urllib.request
from email.header import Header
from email.mime.text import MIMEText
from urllib.parse import quote

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright
import pymysql

import login_taobao

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):
    # 打包为 exe 后：以 exe 所在目录为基准，config.json / Excel 放在 exe 旁边
    BASE_DIR = os.path.dirname(sys.executable)


# ---------------------------------------------------------------- 工具函数

def load_config():
    with open(os.path.join(BASE_DIR, "config.json"), "r", encoding="utf-8") as f:
        return json.load(f)


def get_keywords_from_excel(path):
    """读取 Excel 的工作表名作为搜索关键词，过滤系统自动生成的表（microsoft.com: 开头）"""
    wb = load_workbook(path, read_only=True)
    sheets = [name for name in wb.sheetnames if not name.startswith("microsoft.com:")]
    wb.close()
    return sheets


def wait_for_debug_port(port, timeout=30):
    """等待 Edge 调试端口就绪"""
    for _ in range(timeout):
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=2)
            return True
        except Exception:
            time.sleep(1)
    return False


def get_edge_profile(cfg):
    """Edge 调试 Profile 路径：优先用 config.json 的 edge_profile（兼容旧配置）；
    未配置时自动取当前用户主目录下的 edge-debug-profile，不再写死用户路径"""
    p = (cfg.get("edge_profile") or "").strip()
    if p:
        return p
    return os.path.join(os.path.expanduser("~"), "edge-debug-profile")


def start_edge_debug(cfg):
    """若 9222 端口未监听，则用独立 Profile 启动 Edge 调试模式"""
    if wait_for_debug_port(cfg["debug_port"], timeout=3):
        return
    subprocess.Popen([
        cfg["edge_path"],
        f"--remote-debugging-port={cfg['debug_port']}",
        f"--user-data-dir={get_edge_profile(cfg)}",
        "--no-first-run",
        "--no-default-browser-check",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if not wait_for_debug_port(cfg["debug_port"], timeout=30):
        raise RuntimeError("Edge 调试端口启动超时，请检查 config.json 中的 edge_path")


# ---------------------------------------------------------------- 登录检查

def check_login(page):
    """检查淘宝是否已登录。返回 True=已登录 / False=未登录"""
    page.goto("https://www.taobao.com/", timeout=60000, wait_until="domcontentloaded")
    page.wait_for_timeout(4000)
    result = page.evaluate("""() => {
        const body = document.body ? document.body.innerText : '';
        const hasLoginLink = /亲，请登录|请登录|免费注册/.test(body);
        const hasNick = !!document.querySelector('.site-nav-login-info-nick');
        const hasUnb = document.cookie.includes('unb=');
        // 权威信号：导航栏出现"亲，请登录" -> 未登录；否则看登录态 cookie/昵称
        if (hasLoginLink) return { loggedIn: false, detail: '出现登录提示' };
        if (hasUnb || hasNick) return { loggedIn: true, detail: 'unb cookie 或昵称存在' };
        return { loggedIn: false, detail: '无登录态标志' };
    }""")
    print("[登录检查]", result)
    return result["loggedIn"]


# ---------------------------------------------------------------- 搜索抓取

def scrape_keyword(page, keyword, max_items):
    """在淘宝搜索 keyword，抓取商品列表。返回 [ {title, link, image, price, sales}, ... ]"""
    url = "https://s.taobao.com/search?q=" + quote(keyword)
    page.goto(url, timeout=60000, wait_until="domcontentloaded")

    items = []
    for _ in range(8):  # 最多等 8 轮（滚动 + 等待）
        page.wait_for_timeout(2000)
        items = page.evaluate(r"""() => {
            const result = { mode: null, items: [] };
            // 方案A：页面内嵌 JSON（旧版页面存在，保留兜底）
            try {
                const g = window.g_page_config;
                if (g && g.mods && g.mods.itemlist && g.mods.itemlist.data) {
                    const auctions = g.mods.itemlist.data.auctions || [];
                    result.mode = 'g_page_config';
                    result.items = auctions.map(a => ({
                        title: a.raw_title || '',
                        link: a.detail_url || ('https://item.taobao.com/item.htm?id=' + (a.nid || '')),
                        image: a.pic_url ? ('https:' + a.pic_url) : '',
                        price: a.view_price || a.price || '',
                        sales: a.view_sales || ''
                    }));
                    return result;
                }
            } catch (e) {}
            // 方案B（新版页面主路径）：商品卡片 <a class*="doubleCardWrapperAdapt">
            const cards = Array.from(document.querySelectorAll(
                'a[href*="item.htm"][class*="doubleCardWrapperAdapt"]'
            ));
            const domItems = [];
            for (const a of cards) {
                const titleEl = a.querySelector('[class*="title--"]');
                const imgEl = a.querySelector('img[class*="mainPic"], img[class*="mainImg"]');
                const unitEl = a.querySelector('[class*="unit--"]');
                const priceIntEl = a.querySelector('[class*="priceInt--"]');
                const priceFloatEl = a.querySelector('[class*="priceFloat--"]');
                const salesEl = a.querySelector('[class*="realSales--"]');
                // 价格：¥ + 整数 + 小数
                let price = '';
                if (priceIntEl) {
                    price = (unitEl ? unitEl.innerText.trim() : '') +
                            priceIntEl.innerText.trim() +
                            (priceFloatEl ? priceFloatEl.innerText.trim() : '');
                }
                // 销量
                let sales = salesEl ? salesEl.innerText.trim() : '';
                if (!sales) {
                    const m = (a.innerText || '').match(/([0-9.,]+万?\+?人付款|已售[0-9.,]+万?\+?件)/);
                    if (m) sales = m[1];
                }
                // 标题
                let title = titleEl ? titleEl.innerText.trim() : '';
                if (!title) title = (a.innerText || '').split('\n')[0].trim();
                // 链接：保留卡片原始链接（如 https://detail.tmall.com/item.htm?id=xxx&ns=1&...）
                const link = a.href;
                // 图片：优先主图，其次卡片内第一张 alicdn 商品图
                let image = imgEl ? (imgEl.currentSrc || imgEl.src || imgEl.getAttribute('data-src') || '') : '';
                if (!image) {
                    const anyImg = a.querySelector('img[src*="alicdn"]');
                    if (anyImg) image = anyImg.currentSrc || anyImg.src || '';
                }
                domItems.push({ title, link, image, price, sales });
            }
            if (domItems.length) {
                result.mode = 'dom';
                result.items = domItems;
                return result;
            }
            // 未取到数据：报告页面状态
            const body = document.body ? document.body.innerText : '';
            result.pageState = {
                verifying: /安全验证|验证码|滑块|请完成验证/.test(body),
                loading: body.includes('加载中'),
                cardCount: document.querySelectorAll('a[href*="item.htm"]').length,
                bodyHead: body.slice(0, 120)
            };
            return result;
        }""")
        if items.get("items"):
            break
        # 滚动页面触发懒加载
        page.mouse.wheel(0, 1500)
        page.wait_for_timeout(1500)

    mode = items.get("mode") or items.get("pageState") or {}
    print(f"[搜索] 关键词={keyword} 提取模式={mode} 条数={len(items.get('items', []))}")
    if isinstance(mode, dict) and mode.get("verifying"):
        print("[警告] 页面出现安全验证，请在 Edge 窗口手动完成验证后重试")
    if isinstance(mode, dict) and mode.get("loading") and not items.get("items"):
        print("[警告] 页面仍处于加载中，可能未登录导致商品列表不返回")

    out = items.get("items", [])
    return out[:max_items]


# ---------------------------------------------------------------- 数据库同步

def parse_price(price_text):
    """价格文本 -> (纯数字 DECIMAL, 原始文本)。如 '¥9.96' -> (9.96, '¥9.96')"""
    if not price_text:
        return None, None
    m = re.search(r"(\d+(?:\.\d+)?)", price_text)
    return (float(m.group(1)), price_text.strip()) if m else (None, price_text.strip())


def parse_sold_count(sales_text):
    """已售文本 -> 件数。如 '20万+人付款' -> 200000，'8000+人付款' -> 8000，'已售1.2万件' -> 12000"""
    if not sales_text:
        return None
    m = re.search(r"([\d.]+)\s*万", sales_text)
    if m:
        return int(float(m.group(1)) * 10000)
    m = re.search(r"([\d.]+)", sales_text)
    return int(float(m.group(1))) if m else None


def sync_to_db(cfg, rows, keyword):
    """把抓取的商品列表同步到 MySQL（INSERT ... ON DUPLICATE KEY UPDATE，按 source+url 去重）
    keyword: 本次搜索的关键词，写入 products.keyword 字段"""
    if not rows:
        return 0
    db = cfg["db"]
    conn = pymysql.connect(host=db["host"], port=db["port"], user=db["user"],
                           password=db["password"], database=db["database"], charset=db["charset"])
    sql = """
        INSERT INTO products (source, external_id, title, url, image_url, price, price_text, sold_count, keyword)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE
            title = VALUES(title),
            image_url = VALUES(image_url),
            price = VALUES(price),
            price_text = VALUES(price_text),
            sold_count = VALUES(sold_count),
            keyword = VALUES(keyword)
    """
    records = []
    for r in rows:
        idm = re.search(r"id=(\d+)", r["link"] or "")
        price, price_text = parse_price(r["price"])
        records.append((
            db["source"], idm.group(1) if idm else None,
            r["title"], r["link"], r["image"],
            price, price_text, parse_sold_count(r["sales"]), keyword,
        ))
    try:
        with conn.cursor() as cur:
            cur.executemany(sql, records)
        conn.commit()
        return len(records)
    finally:
        conn.close()


# ---------------------------------------------------------------- Excel 写入

def write_to_excel(path, sheet_name, rows):
    """把抓取结果写入指定工作表（表头 + 数据）。已存在时清空重写，保证每次运行结果是最新的"""
    wb = load_workbook(path)
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(title=sheet_name)
    # 清空旧内容（删除所有行，确保 append 从第 1 行开始）
    if ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)
    headers = ["标题", "商品链接", "商品图片", "商品价格", "已售件数"]
    ws.append(headers)
    for r in rows:
        ws.append([r["title"], r["link"], r["image"], r["price"], r["sales"]])
    wb.save(path)
    print(f"[Excel] 已写入 sheet「{sheet_name}」 共 {len(rows)} 条 -> {path}")


# ---------------------------------------------------------------- 邮件发送

def send_notify_email(cfg):
    """淘宝自动登录失败时发送提醒邮件。授权码未配置时打印警告并跳过"""
    mail_cfg = cfg["email"]
    if not mail_cfg.get("auth_code"):
        print("[邮件] 未配置 SMTP 授权码（config.json -> email.auth_code），跳过发送")
        print("[邮件] 提示：请在本机 Edge 手动登录淘宝，或在 config.json 填入授权码")
        return False
    subject = "【淘宝抓取】淘宝自动登录失败"
    body = (
        "淘宝商品抓取脚本尝试自动登录（手机号+验证码）失败，无法抓取商品。\n\n"
        f"检测时间：{time.strftime('%Y-%m-%d %H:%M:%S')}\n"
        "可能原因：验证码未写入 sms_receiver.sms_codes 表、验证码错误或登录超时。\n"
        "请在本机打开淘宝手动登录后重新运行脚本。"
    )
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = mail_cfg["sender"]
    msg["To"] = mail_cfg["to"]
    try:
        with smtplib.SMTP_SSL(mail_cfg["smtp_host"], mail_cfg["smtp_port"], timeout=30) as s:
            s.login(mail_cfg["sender"], mail_cfg["auth_code"])
            s.sendmail(mail_cfg["sender"], [mail_cfg["to"]], msg.as_string())
        print(f"[邮件] 登录失败提醒已发送至 {mail_cfg['to']}")
        return True
    except Exception as e:
        print(f"[邮件] 发送失败: {e}")
        return False


# ---------------------------------------------------------------- 主流程

def probe_search_page(page, keyword):
    """诊断：登录后检查搜索页的数据源（卡片 DOM / g_page_config 全局）"""
    url = "https://s.taobao.com/search?q=" + quote(keyword)
    page.goto(url, timeout=60000, wait_until="domcontentloaded")
    page.wait_for_timeout(5000)
    info = page.evaluate("""() => {
        const g = window.g_page_config;
        const auctions = (g && g.mods && g.mods.itemlist && g.mods.itemlist.data)
            ? (g.mods.itemlist.data.auctions || []) : [];
        const card = document.querySelector('a[href*="item.htm"][class*="doubleCardWrapperAdapt"]');
        return {
            url: location.href,
            hasGPageConfig: !!g,
            auctionCount: auctions.length,
            cardCount: document.querySelectorAll('a[href*="item.htm"][class*="doubleCardWrapperAdapt"]').length,
            sampleCard: card ? {
                title: (card.querySelector('[class*="title--"]') || {innerText: ''}).innerText.slice(0, 30),
                href: card.href.slice(0, 90)
            } : null,
            bodyLoading: document.body ? document.body.innerText.includes('加载中') : null
        };
    }""")
    print("[probe]", json.dumps(info, ensure_ascii=False, indent=1)[:2500])


def main():
    import sys
    cfg = load_config()
    probe_only = "--probe" in sys.argv
    excel_path = os.path.join(BASE_DIR, cfg["excel_path"])
    if not os.path.exists(excel_path):
        print(f"[错误] 找不到 Excel 文件: {excel_path}")
        return
    keywords = get_keywords_from_excel(excel_path)
    if not keywords:
        print("[错误] Excel 中没有任何可用的工作表名（搜索关键词）")
        return
    print("[Excel] 搜索关键词（工作表名）:", keywords)

    start_edge_debug(cfg)
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{cfg['debug_port']}")
        page = browser.contexts[0].new_page()

        # 1. 检查登录
        if not check_login(page):
            print("[登录] 未登录淘宝，尝试手机号+验证码自动登录")
            if not login_taobao.taobao_login(page, cfg):
                print("[登录] 自动登录失败，发送提醒邮件")
                send_notify_email(cfg)
                return
        if probe_only:
            probe_search_page(page, keywords[0])
            page.close()
            return

        # 2. 已登录 -> 逐个关键词搜索抓取
        print("[登录] 已登录淘宝，开始抓取")
        total_synced = 0
        for kw in keywords:
            rows = scrape_keyword(page, kw, cfg["max_items_per_keyword"])
            if rows:
                write_to_excel(excel_path, kw, rows)
                n = sync_to_db(cfg, rows, kw)
                total_synced += n
                print(f"[数据库] 关键词「{kw}」同步 {n} 条 -> {cfg['db']['database']}.products")
            else:
                print(f"[Excel] 关键词「{kw}」无抓取结果，跳过写入")
        if total_synced:
            print(f"[数据库] 本次共同步 {total_synced} 条商品记录")
        # page.close()


if __name__ == "__main__":
    main()
