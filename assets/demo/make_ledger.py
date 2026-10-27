"""造一份**完全合成**的演示账本，供报表截图使用。

为什么不截真表再涂改：涂改会让数字彼此矛盾（KPI 加总 ≠ 明细、结论句里的
百分比对不上表格），一眼假；而截真表又必然带出本机的路径、会话名与真实用量。
所以从数据源头造假：生成一份符合 collect.py 解析口径的假 workbuddy.db 与
projects/**/*.jsonl，再用**原封不动的 report.py** 渲染。渲染逻辑一行没改，
出来的就是产品真实的样子，只是里面的人和数字都不存在。

用法：
    python make_ledger.py            # 生成 demo-ledger/
    python make_ledger.py --clean    # 先清空再生成
"""

from __future__ import annotations

import argparse
import json
import os
import random
import shutil
import sqlite3
from datetime import datetime, timedelta, timezone

TZ = timezone(timedelta(hours=8))
HERE = os.path.dirname(os.path.abspath(__file__))
LEDGER = os.path.join(HERE, "demo-ledger")

SEED = 20261025
# 起点必须早于「近 30 日的上期起点」（08-27），否则报表会给这一档打星号
# 「上期起始日早于本机最早记录日」——那是真话，但演示图里不该出现。
START = datetime(2026, 8, 20)
END = datetime(2026, 10, 25)       # 今天

# 假期与周末压低用量：每日趋势图需要有起伏，平的一条直线没人信
HOLIDAYS = {(10, d) for d in range(1, 9)}   # 10-01 ~ 10-08

# 工作区一律用类 Unix 的中性路径，不出现任何真实用户名
WORKSPACES = [
    "/home/dev/workspace/orbit-api",
    "/home/dev/workspace/data-pipeline",
    "/home/dev/workspace/site-docs",
]

# (会话 id, 标题, 工作区序号, 起, 止)
SESSIONS = [
    ("s01", "重构鉴权中间件",   0, "08-20", "09-12"),
    ("s02", "修复时区偏移问题", 0, "08-26", "09-15"),
    ("s03", "配置 CI 缓存",     0, "09-10", "10-25"),
    ("s04", "补齐接口单元测试", 0, "09-22", "10-20"),
    ("s05", "梳理数据库迁移脚本", 0, "10-06", "10-25"),
    ("s06", "重写清洗规则表",   1, "08-21", "09-18"),
    ("s07", "优化批量导入流程", 1, "09-05", "10-02"),
    ("s08", "排查任务重复执行", 1, "09-28", "10-22"),
    ("s09", "补齐字段命名文档", 1, "10-12", "10-25"),
    ("s10", "调整文档导航结构", 2, "08-22", "09-20"),
    ("s11", "整理季度复盘",     2, "09-18", "10-14"),
    ("s12", "补写部署说明",     2, "10-01", "10-24"),
]

# 模型权重与 token 画像；credit 与 token 大致同比例（演示用，不必是官方换算）
MODELS = [
    ("deepseek-v4.1-flash", 0.72, (4000, 26000), (300, 2200), (0.70, 0.88), 0.85),
    ("glm-5.3-flash", 0.18, (3000, 18000), (200, 1500), (0.55, 0.75), 0.55),
    ("deepseek-v4-pro", 0.10, (12000, 55000), (600, 3800), (0.60, 0.80), 3.20),
]


def _pick_day(key):
    m, d = key.split("-")
    return datetime(2026, int(m), int(d))


def _level(day):
    if (day.month, day.day) in HOLIDAYS:
        return 0.12
    if day.weekday() >= 5:
        return 0.45
    return 1.0


def _model(rng):
    roll = rng.random()
    acc = 0.0
    for name, weight, *_ in MODELS:
        acc += weight
        if roll <= acc:
            return name
    return MODELS[0][0]


def _profile(model):
    for name, _w, tin, tout, cache, k in MODELS:
        if name == model:
            return tin, tout, cache, k
    raise KeyError(model)


def build(clean=False):
    if clean and os.path.isdir(LEDGER):
        shutil.rmtree(LEDGER)
    os.makedirs(LEDGER, exist_ok=True)
    projects = os.path.join(LEDGER, "projects")
    if os.path.isdir(projects):
        shutil.rmtree(projects)

    rng = random.Random(SEED)
    sessions = []
    for sid, title, ws_i, s, e in SESSIONS:
        sessions.append({"sid": sid, "title": title, "ws": WORKSPACES[ws_i],
                         "start": _pick_day(s), "end": _pick_day(e)})

    # 每个会话一个 jsonl；子代理日志单独落到 subagents/ 下
    handles = {}
    used = {}          # sid -> 当前上下文占用
    credit_map = {}    # sid -> {req_id: credit}
    last_ts = {}       # sid -> 最近时间戳
    counter = [0]

    def handle(path):
        if path not in handles:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            handles[path] = open(path, "w", encoding="utf-8", newline="\n")
        return handles[path]

    def emit(sess, ts, model, sub=False):
        tin, tout, cache_rng, k = _profile(model)
        counter[0] += 1
        rid = "req_%06x" % (0x100000 + counter[0] * 7919)
        inp = rng.randint(*tin)
        out = rng.randint(*tout)
        hit = int(inp * rng.uniform(*cache_rng))
        think = rng.randint(0, 900) if model != "deepseek-v4-pro" else rng.randint(200, 2200)
        credit = round((inp + out) / 10000.0 * k * rng.uniform(0.9, 1.15), 2)

        sid = sess["sid"]
        name = ("sub-%s.jsonl" % sid) if sub else ("%s.jsonl" % sid)
        path = os.path.join(projects, os.path.basename(sess["ws"]),
                            "subagents" if sub else "chat", name)
        fh = handle(path)
        fh.write(json.dumps({
            "type": "assistant",
            "sessionId": sid,
            "cwd": sess["ws"],
            "timestamp": ts,
            "providerData": {
                "conversationRequestId": rid,
                "model": model,
                "rawUsage": {
                    "prompt_tokens": inp,
                    "completion_tokens": out,
                    "prompt_cache_hit_tokens": hit,
                    "prompt_cache_miss_tokens": inp - hit,
                    "completion_thinking_tokens": think,
                    "credit": credit,
                },
            },
        }, ensure_ascii=False) + "\n")
        # 同一请求在日志里常有多行；补一行只带 credit 的，token 不会被重复计
        if rng.random() < 0.35:
            fh.write(json.dumps({
                "type": "session_usage",
                "sessionId": sid,
                "timestamp": ts + rng.randint(200, 3000),
                "providerData": {
                    "conversationRequestId": rid,
                    "rawUsage": {"credit": credit},
                },
            }, ensure_ascii=False) + "\n")

        credit_map.setdefault(sid, {})[rid] = credit
        used[sid] = used.get(sid, 0) + inp // 3
        last_ts[sid] = max(last_ts.get(sid, 0), ts)

        # 联网检索：只记次数，报表里单列，不折钱
        if rng.random() < 0.06:
            web_path = os.path.join(projects, os.path.basename(sess["ws"]), "chat", "%s.jsonl" % sid)
            handle(web_path).write(json.dumps({
                "type": "function_call",
                "name": rng.choice(["WebSearch", "WebFetch"]),
                "sessionId": sid,
                "callId": "call_%06x" % (0x200000 + counter[0] * 104729),
                "timestamp": ts + rng.randint(100, 900),
                "providerData": {},
            }, ensure_ascii=False) + "\n")

    day = START
    while day <= END:
        active = [s for s in sessions if s["start"] <= day <= s["end"]]
        if day.date() == END.date():
            # 今天只用了一个上午：既能看出「进行中」，又不是空卡
            target = 22
            h0, h1 = 8, 10
        else:
            target = max(1, int(round(34 * _level(day) * rng.uniform(0.85, 1.15))))
            h0, h1 = 9, 23
        if not active:
            day += timedelta(days=1)
            continue

        picks = rng.sample(active, min(len(active), rng.choice([2, 3, 3, 4])))
        for _ in range(target):
            sess = rng.choice(picks)
            hour = rng.randint(h0, h1 - 1)
            minute = rng.randint(0, 59)
            ts = int(datetime(day.year, day.month, day.day, hour, minute,
                              rng.randint(0, 59), tzinfo=TZ).timestamp() * 1000)
            if day.date() == END.date() and hour >= h1:
                continue
            model = _model(rng)
            emit(sess, ts, model, sub=(rng.random() < 0.11))
        day += timedelta(days=1)

    for fh in handles.values():
        fh.close()

    # ── workbuddy.db ──────────────────────────────────────────────
    db = os.path.join(LEDGER, "workbuddy.db")
    if os.path.exists(db):
        os.remove(db)
    con = sqlite3.connect(db)
    con.execute("""create table sessions(
        id text primary key, title text, custom_title text, model text, cwd text,
        updated_at integer, created_at integer, deleted_at integer,
        is_playground integer, context_window integer)""")
    con.execute("""create table session_usage(
        session_id text primary key, used integer, size integer,
        credit_json text, updated_at integer)""")
    for s in sessions:
        sid = s["sid"]
        if sid not in credit_map:
            continue
        total = sum(credit_map[sid].values())
        con.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?)", (
            sid, s["title"], None, "deepseek-v4.1-flash", s["ws"],
            last_ts.get(sid), int(s["start"].replace(tzinfo=TZ).timestamp() * 1000),
            None, 0, 200000))
        con.execute("insert into session_usage values(?,?,?,?,?)", (
            sid, min(used.get(sid, 0), 190000), 200000,
            json.dumps(credit_map[sid], separators=(",", ":")), last_ts.get(sid)))
    con.commit()
    con.close()

    total_req = sum(len(v) for v in credit_map.values())
    total_credit = sum(sum(v.values()) for v in credit_map.values())
    print("账本已生成: %s" % LEDGER)
    print("会话 %d 个 · 请求 %d 条 · 积分合计 %.2f" % (
        len(credit_map), total_req, total_credit))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--clean", action="store_true", help="先清空 demo-ledger")
    build(clean=ap.parse_args().clean)
