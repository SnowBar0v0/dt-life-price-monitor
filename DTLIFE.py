# -*- coding: utf-8 -*-
"""
Author: SnowBar0v0
GitHub: https://github.com/SnowBar0v0

DT生活 商品价格监控（完整文件，已移除 Telegram 部分）

功能：
- 连接指定 PID 的 DT生活 窗口（CONFIG.target_pid）
- 首次扫描并列出所有商品，交互选择要监控的条目并设置阈值
- 循环刷新并显示已配置监控项实时价格，价格低于阈值时通过微信推送提醒（若配置）

说明（仅基于本脚本逻辑）：
- 已剔除 Telegram 相关配置与发送逻辑（send_tg 已移除）。
- 保留并实现一个轻量的微信推送方法（connect_wechat / send_to_wechat），
  实现方式：使用剪贴板粘贴（pyperclip）并发送 Ctrl+V + Enter 到已连接的微信聊天窗口。
  该实现仅为本脚本用于发送提醒的简单方案，不包含任何外部脚本/工具的扩展逻辑。
- 其余功能均基于 DT生活 窗口的 UIA 文本抓取与解析：查找 Edit 控件刷新、抓取 Text 控件、按
  固定顺序解析商品条目（title/status/price/market）、交互式选择监控项并按阈值发送提醒。
"""

import time
import logging
import pyperclip
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys

# ==============================
# 配置区（请按需修改）
# ==============================
CONFIG = {
    # 目标进程 PID（在此固定，不再通过终端输入）
    # 请将此处替换为实际的 DT生活 进程 ID（例如 1848）
    "target_pid": ,

    # 微信：如果要启用微信推送，请填写微信进程 PID（在任务管理器中查看）
    # 若不启用，置为 None
    "wechat_pid": ,
    # 微信窗口标题匹配（正则），用于定位独立聊天窗口
    "wechat_window_title": ".*微信.*",

    # 扫描间隔（秒）
    "scan_interval_seconds": 4,

    # 提醒冷却（秒）
    "alert_cooldown_seconds": 300,

    # 搜索框刷新
    "press_enter_each_scan": True,
    "refresh_wait_seconds": 0.6,
    "search_hint_text": "",
}

# ==============================
# 日志
# ==============================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

# ==============================
# 微信推送（轻量实现，仅用于发送提醒）
# 说明：
# - 使用剪贴板粘贴并发送 Ctrl+V + Enter 到已连接的微信聊天窗口
# - 优先尝试定位聊天输入 Edit 控件，失败则以窗口聚焦后发送 Ctrl+V + Enter
# - 仅为本脚本提供简易推送能力，不包含复杂转发/持久化/命令解析等功能
# ==============================
WECHAT_WIN = None
WECHAT_AVAILABLE = False

def connect_wechat():
    """
    通过 CONFIG['wechat_pid'] 连接微信应用窗口（UIA）。
    使用 Application.window(title_re=...) 定位聊天窗口，返回的对象支持 child_window。
    """
    global WECHAT_WIN, WECHAT_AVAILABLE
    pid = CONFIG.get("wechat_pid")
    title_re = CONFIG.get("wechat_window_title") or ".*微信.*"
    if not pid:
        WECHAT_WIN = None
        WECHAT_AVAILABLE = False
        logging.info("wechat_pid 未配置，微信推送未启用。")
        return None
    try:
        app = Application(backend="uia").connect(process=pid, timeout=5)
        # 尝试按标题正则获取窗口
        try:
            win = app.window(title_re=title_re)
            win.wait("visible ready", timeout=5)
        except Exception:
            # 回退：从 app.windows() 里找第一个可见窗口
            wins = app.windows()
            win = None
            for w in wins:
                try:
                    if getattr(w, "is_visible", lambda: False)():
                        win = w
                        break
                except Exception:
                    continue
            if win is None and wins:
                win = wins[0]
        try:
            win.set_focus()
        except Exception:
            pass
        WECHAT_WIN = win
        WECHAT_AVAILABLE = True
        logging.info(f"微信窗口连接成功 (pid={pid})：{win.window_text()}")
        return win
    except Exception as e:
        WECHAT_WIN = None
        WECHAT_AVAILABLE = False
        logging.warning(f"无法连接微信 (pid={pid})：{e}")
        return None

def send_to_wechat(text):
    """
    将文本通过剪贴板粘贴并发送到已连接的微信窗口（Ctrl+V + Enter）。
    简单实现：优先定位输入 Edit 控件并发送，否则以窗口聚焦后发送 Ctrl+V Enter。
    """
    global WECHAT_WIN, WECHAT_AVAILABLE
    if not text:
        return
    if not WECHAT_AVAILABLE or WECHAT_WIN is None:
        connect_wechat()
    if not WECHAT_AVAILABLE or WECHAT_WIN is None:
        logging.info("微信不可用，跳过微信推送。")
        return

    try:
        pyperclip.copy(text)

        box = None
        # 尝试通过 child_window 获取输入框（若能定位到更可靠）
        try:
            box = WECHAT_WIN.child_window(auto_id="chat_input_field", control_type="Edit")
            if not getattr(box, "exists", lambda timeout=0: False)(timeout=1):
                box = None
        except Exception:
            box = None

        if box is None:
            try:
                box = WECHAT_WIN.child_window(control_type="Edit")
                if not getattr(box, "exists", lambda timeout=0: False)(timeout=1):
                    box = None
            except Exception:
                box = None

        if box is not None:
            try:
                box.click_input()
                box.type_keys("^v{ENTER}", pause=0.05)
            except Exception:
                try:
                    WECHAT_WIN.set_focus()
                    send_keys("^v{ENTER}")
                except Exception:
                    logging.warning("微信发送失败（输入框操作异常）。")
                    WECHAT_AVAILABLE = False
                    WECHAT_WIN = None
        else:
            try:
                WECHAT_WIN.set_focus()
                send_keys("^v{ENTER}")
            except Exception:
                logging.warning("微信发送失败（无法找到输入框）。")
                WECHAT_AVAILABLE = False
                WECHAT_WIN = None

    except Exception as e:
        logging.warning(f"微信发送异常：{e}")
        WECHAT_AVAILABLE = False
        WECHAT_WIN = None

# ==============================
# 获取 DT生活 窗口（优先 Desktop 全局查找）
# ==============================
def get_window(pid):
    try:
        desktop = Desktop(backend="uia")
        for w in desktop.windows():
            try:
                elem = getattr(w, "element_info", None)
                proc = None
                if elem is not None:
                    proc = getattr(elem, "process_id", None)
                if proc is None:
                    try:
                        proc = w.process_id()
                    except Exception:
                        proc = None
                if proc == pid and "DT生活" in w.window_text():
                    logging.info(f"找到窗口 (desktop): {w.window_text()}")
                    return w
            except Exception:
                continue
    except Exception:
        pass

    try:
        app = Application(backend="uia").connect(process=pid)
        for w in app.windows():
            try:
                if "DT生活" in w.window_text():
                    logging.info(f"找到窗口 (app): {w.window_text()}")
                    return w
            except Exception:
                continue
    except Exception:
        pass

    raise Exception("未找到 DT生活 窗口，请检查 target_pid 或窗口标题。")

# ==============================
# 搜索框查找与刷新逻辑
# ==============================
def _collect_search_edits(win):
    roots = [win]
    try:
        roots.extend(win.descendants(control_type="Document", title="Page-Frame"))
    except Exception:
        pass
    try:
        roots.extend(win.descendants(control_type="Document", title="AppIndex"))
    except Exception:
        pass

    edits = []
    seen = set()
    for root in roots:
        try:
            for e in root.descendants(control_type="Edit"):
                key = None
                try:
                    key = tuple(e.element_info.runtime_id)
                except Exception:
                    try:
                        key = e.handle
                    except Exception:
                        key = None
                if key is not None:
                    if key in seen:
                        continue
                    seen.add(key)
                edits.append(e)
        except Exception:
            continue
    return edits

def find_search_edit(win):
    edits = _collect_search_edits(win)
    if not edits:
        try:
            doc = win.child_window(title="Page-Frame", control_type="Document")
            doc.set_focus()
            time.sleep(0.1)
        except Exception:
            pass
        edits = _collect_search_edits(win)

    if not edits:
        try:
            desktop = Desktop(backend="uia")
            all_edits = []
            for top in desktop.windows():
                try:
                    for e in top.descendants(control_type="Edit"):
                        all_edits.append(e)
                except Exception:
                    pass
        except Exception:
            all_edits = []

        if not all_edits:
            return None

        for e in all_edits:
            try:
                if e.has_keyboard_focus():
                    return e
            except Exception:
                pass

        hint = (CONFIG.get("search_hint_text") or "").strip()
        if hint:
            for e in all_edits:
                val = ""
                try:
                    val = e.get_value()
                except Exception:
                    try:
                        val = e.window_text()
                    except Exception:
                        val = ""
                if val and hint in val:
                    return e

        for e in all_edits:
            try:
                if e.element_info and e.element_info.is_keyboard_focusable:
                    return e
            except Exception:
                pass

        try:
            visible = []
            for e in all_edits:
                try:
                    if e.is_offscreen():
                        continue
                except Exception:
                    pass
                try:
                    r = e.rectangle()
                    visible.append((r.top, r.left, e))
                except Exception:
                    visible.append((999999, 0, e))
            if visible:
                return sorted(visible, key=lambda x: (x[0], x[1]))[0][2]
        except Exception:
            return all_edits[0]

    for e in edits:
        try:
            if e.has_keyboard_focus():
                return e
        except Exception:
            pass

    hint = (CONFIG.get("search_hint_text") or "").strip()
    if hint:
        for e in edits:
            val = ""
            try:
                val = e.get_value()
            except Exception:
                try:
                    val = e.window_text()
                except Exception:
                    val = ""
            if val and hint in val:
                return e

    for e in edits:
        try:
            if e.element_info and e.element_info.is_keyboard_focusable:
                return e
        except Exception:
            pass

    try:
        visible = []
        for e in edits:
            try:
                if e.is_offscreen():
                    continue
            except Exception:
                pass
            try:
                r = e.rectangle()
                visible.append((r.top, r.left, e))
            except Exception:
                visible.append((999999, 0, e))
        if visible:
            return sorted(visible, key=lambda x: (x[0], x[1]))[0][2]
    except Exception:
        return edits[0]

def refresh_list_by_enter(win):
    if not CONFIG.get("press_enter_each_scan", True):
        return False
    edit = find_search_edit(win)
    try:
        if edit is not None:
            try:
                edit.click_input()
            except Exception:
                try:
                    edit.set_focus()
                except Exception:
                    pass
            try:
                edit.type_keys("{ENTER}")
            except Exception:
                send_keys("{ENTER}")
        else:
            try:
                win.set_focus()
            except Exception:
                pass
            send_keys("{ENTER}")
        time.sleep(CONFIG.get("refresh_wait_seconds", 0.5))
        return True
    except Exception:
        return False

# ==============================
# 文本抓取与解析
# ==============================
def get_all_texts(win):
    texts = []
    try:
        for c in win.descendants(control_type="Text"):
            t = c.window_text()
            if t:
                texts.append(t)
    except Exception:
        pass
    return texts

def parse_items(texts):
    items = []
    for i in range(len(texts) - 3):
        title = texts[i]
        status = texts[i + 1]
        price = texts[i + 2]
        market = texts[i + 3]
        if not price.startswith("¥"):
            continue
        try:
            p = float(price.replace("¥", "").strip())
        except Exception:
            continue
        items.append({"title": title, "price": p, "status": status})
    return items

# ==============================
# 交互式选择监控项
# ==============================
def _parse_selection(sel_text, max_index):
    sel_text = (sel_text or "").strip().lower()
    if not sel_text:
        return []
    if sel_text in ("all", "a"):
        return list(range(max_index + 1))
    if sel_text in ("none", "n", "q", "quit", "exit"):
        return []
    parts = sel_text.split(",")
    idxs = set()
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if "-" in p:
            try:
                a, b = p.split("-", 1)
                a = int(a); b = int(b)
                if a > b: a, b = b, a
                for i in range(max(0, a), min(max_index, b) + 1):
                    idxs.add(i)
            except Exception:
                continue
        else:
            try:
                i = int(p)
                if 0 <= i <= max_index:
                    idxs.add(i)
            except Exception:
                continue
    return sorted(idxs)

def select_items_interactive(detected_items):
    if not detected_items:
        logging.info("未检测到任何商品，无法进入选择流程。")
        return []
    print("\n检测到以下商品：")
    for idx, it in enumerate(detected_items):
        print(f"[{idx}] {it['title']}    价格: {it['price']}    状态: {it['status']}")
    while True:
        sel = input("\n请输入要监控的商品索引（例如 0,2-4），输入 all 监控所有，直接回车退出: ").strip()
        idxs = _parse_selection(sel, len(detected_items) - 1)
        if not idxs:
            print("未选择任何商品，退出选择。")
            return []
        print(f"已选择索引: {idxs}")
        confirm = input("确认选择并继续设置阈值？(y/n): ").strip().lower()
        if confirm in ("y", "yes"):
            break
    monitors = []
    for i in idxs:
        it = detected_items[i]
        default_name = it['title'][:40]
        prompt_name = input(f"为 [{i}] 设置友好名称（回车使用默认: \"{default_name}\"]: ").strip()
        name = prompt_name if prompt_name else default_name
        default_threshold = it['price']
        while True:
            thr_in = input(f"为 [{name}] 设置提醒阈值（当前价格 {it['price']}，回车使用当前价格 {default_threshold}）: ").strip()
            if thr_in == "":
                threshold = default_threshold
                break
            try:
                threshold = float(thr_in)
                break
            except Exception:
                print("阈值输入无效，请输入数字或回车。")
        monitors.append({
            "name": name,
            "match": it['title'],
            "threshold": threshold,
            "enabled": True
        })
    logging.info("用户配置的监控项：")
    for m in monitors:
        logging.info(f"{m['name']}  匹配: {m['match']}  阈值: {m['threshold']}")
    return monitors

# ==============================
# 主程序
# ==============================
def main():
    pid = CONFIG.get("target_pid")
    if not pid or not isinstance(pid, int):
        logging.error("未在 CONFIG 中配置有效的 target_pid，程序退出。")
        return

    try:
        win = get_window(pid)
    except Exception as e:
        logging.exception("获取窗口失败：%s", e)
        return

    try:
        win.set_focus()
    except Exception:
        pass

    # 尝试连接微信（若配置了 wechat_pid）
    try:
        connect_wechat()
    except Exception:
        pass

    # 首次连接后获取并显示完整商品列表，进行交互式选择
    try:
        refresh_list_by_enter(win)
        texts = get_all_texts(win)
        detected = parse_items(texts)
        if not detected:
            logging.info("首次扫描未发现任何商品文本。程序退出。")
            return
        runtime_monitors = select_items_interactive(detected)
        if not runtime_monitors:
            logging.info("未配置监控项，程序退出。")
            return
    except Exception as e:
        logging.exception("首次扫描/选择时发生异常：%s", e)
        return

    last_alert = {}
    logging.info("进入监控循环。按 Ctrl+C 退出。")

    while True:
        try:
            refresh_list_by_enter(win)
            texts = get_all_texts(win)
            items = parse_items(texts)

            # 每次刷新后，显示已配置监控项的实时价格
            for conf in runtime_monitors:
                if not conf.get("enabled", True):
                    continue
                found = None
                for it in items:
                    if conf["match"] in it["title"]:
                        found = it
                        break
                if found:
                    logging.info(f"{conf['name']} 实时价格: {found['price']}")
                else:
                    logging.info(f"{conf['name']} 未找到匹配商品")

            # 阈值检测与提醒（仅微信推送，若已配置）
            for conf in runtime_monitors:
                if not conf.get("enabled", True):
                    continue
                for it in items:
                    if conf["match"] in it["title"]:
                        price = it["price"]
                        logging.info(f"{conf['name']} 当前价格 {price}")
                        if price <= conf["threshold"]:
                            now = time.time()
                            last = last_alert.get(conf["name"], 0)
                            if now - last > CONFIG.get("alert_cooldown_seconds", 300):
                                msg = f"""{conf['name']} 低价提醒

当前价格: {price}
阈值: {conf['threshold']}

{it['title']}
"""
                                logging.info(msg)
                                # 只使用微信推送（若已配置）
                                try:
                                    send_to_wechat(msg)
                                except Exception:
                                    logging.exception("发送微信提醒失败。")
                                last_alert[conf["name"]] = now

            time.sleep(CONFIG.get("scan_interval_seconds", 4))
        except KeyboardInterrupt:
            logging.info("用户中断，程序退出。")
            break
        except Exception:
            logging.exception("循环监控过程中发生异常，继续下一轮。")
            time.sleep(CONFIG.get("scan_interval_seconds", 4))

# ==============================
# 启动
# ==============================
if __name__ == "__main__":
    logging.info("=== price monitor 启动（Telegram 已移除，仅微信推送可用） ===")
    main()
