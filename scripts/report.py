"""wb-usage-report 命令行入口。

设计目标：Agent 只需要「执行一条命令 + 读十几行摘要」，不读源码、不读明细，
从而把每次出报表的上下文成本压到最低。所有明细落在文件里。

默认行为：基准日 auto —— 本地时间 06:00 前自动退到前一天。
所以推荐 00:05 的定时任务，不加任何参数就能出「前一天」的完整账单。

用法
----
    python report.py                       # 截至基准日，输出到 ~/.workbuddy/usage-reports/<基准日>/
    python report.py --anchor prev         # 强制出前一天
    python report.py --anchor 2026-10-20   # 回溯某个历史日
    python report.py --days 7              # 只看最近 7 天
    python report.py --out D:/bill         # 指定输出目录
    python report.py --format html --open  # 只出 HTML 并打开
    python report.py --format text         # 只在终端看摘要
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import aggregate as agg  # noqa: E402
import collect  # noqa: E402
import render  # noqa: E402
import xlsx  # noqa: E402

VERSION = "1.13.0"
DEFAULT_OUT_ROOT = "~/.workbuddy/usage-reports"


# ── CSV 导出 ─────────────────────────────────────────────────────
# 只有一个导出函数：所有数据进同一个文件。旧的按表拆分（summary/requests/...）
# 已删除——七个文件互相不知道对方存在，用起来还得挨个点开找。


def _drop_legacy_csv_dir(out_dir):
    """清掉 v1.5 及以前留下的 ``csv/`` 子目录——新方案已不再引用它。

    只在目录里**全是**已知旧产物时才删：多出任何别的文件，说明用户往里放过东西，
    那就一个都不动。宁可不清理，也不误删。
    """
    legacy = os.path.join(out_dir, "csv")
    if not os.path.isdir(legacy):
        return
    known = {"summary.csv", "requests.csv", "top_requests.csv", "sessions.csv",
             "models.csv", "workspaces.csv", "days.csv"}
    names = set(os.listdir(legacy))
    if not names <= known:
        return
    for name in names:
        os.remove(os.path.join(legacy, name))
    try:
        os.rmdir(legacy)
    except OSError:
        pass


def _drop_stale_exports(out_dir, keep_name):
    """同一个归档目录里只留最新那份 ``<日期>-<时刻>.xlsx``。

    文件名带生成时刻，重跑一次就多一份只差最后千分之一的快照；归档目录是
    「一天一份」，攒着只会让人挑错版本。只删**严格匹配本报表命名规则**的文件
    （.xlsx 与 v1.6 及以前的 .csv），别的表格文件（哪怕就放在旁边）一律不碰。
    """
    pattern = re.compile(r"^\d{4}-\d{2}-\d{2}-\d{4}\.(xlsx|csv)$")
    try:
        names = os.listdir(out_dir)
    except OSError:
        return
    for name in names:
        if name == keep_name or not pattern.match(name):
            continue
        try:
            os.remove(os.path.join(out_dir, name))
        except OSError:
            pass


def _export_workbook(payload, out_dir):
    """把**全部**数据写进一个六 sheet 的 Excel 工作簿，返回 (路径, base64)。

    为什么从「一个 CSV 分六节」改成 xlsx：CSV 是同构纯文本，没有 sheet、没有
    单元格格式的概念，六个小节只能靠 `===== 标题 =====` 硬隔开，滚动时既看不到
    表头也不会停。改用 xlsx 后六张表各占一个 sheet（汇总 / 逐请求 / 会话 /
    模型 / 工作区 / 按天），全表居中，首行冻结。

    为什么自己写 OOXML 而不用 openpyxl：这份报表要在任何 python3 下直接跑，
    第三方包一旦缺失整份报表就出不来，而这里只需要「多 sheet + 居中」。

    base64 一并返回，是给 HTML 内嵌用的：报表被「单文件预览」托管时，
    同目录的文件是取不到的（相对路径会 404），内嵌一份才点得开。
    """
    stamp = payload["generated_at"].strftime("%Y-%m-%d-%H%M")
    path = os.path.join(out_dir, "%s.xlsx" % stamp)
    sheets = []   # [(sheet 名, 表头, 行)]

    # 1) 汇总：五档窗口（今日 / 昨日 / 近 7 日 / 近 30 日 / 全部）+ 环比
    rows = []
    for key, label in payload["windows"]:
        s = payload["summary"][key]
        cmp_ = s.get("compare") or {}
        prev = cmp_.get("prev") or {}

        def _p(value):
            return round(value * 100, 2) if value is not None else ""

        rows.append(
            [
                s.get("label") or label,
                s.get("since") or "全部",
                s.get("until") or "",
                s["requests"],
                s["sessions"],
                s["active_days"],
                s["input"],
                s["output"],
                s["tokens"],
                s["cache_hit"],
                s["cache_miss"],
                round(s["cache_rate"] * 100, 2) if s["cache_rate"] is not None else "",
                s["thinking"],
                s["credit"],
                round(s.get("cost") or 0, 4),
                round(s.get("cost_if_peak") or 0, 4),
                s.get("priced_tokens") or 0,
                _p(s.get("cost_coverage")),
                s["free_requests"],
                s["free_tokens"],
                s["sub_requests"],
                s["sub_tokens"],
                s["sub_credit"],
                s["no_credit_requests"],
                s["no_credit_tokens"],
                s["llm_calls"],
                (s.get("web") or {}).get("total", ""),
                (s.get("web") or {}).get("search", ""),
                (s.get("web") or {}).get("fetch", ""),
                cmp_.get("label") or "",
                prev.get("since") or "",
                prev.get("until") or "",
                prev.get("credit", ""),
                # 起步中的今日档不给环比：把 -99% 印进 Excel 只会被当成用量塌陷
                "" if s.get("gated") else _p(cmp_.get("credit_pct")),
                "" if s.get("gated") else cmp_.get("credit_abs", ""),
                prev.get("tokens", ""),
                _p(cmp_.get("tokens_pct")),
                prev.get("requests", ""),
                _p(cmp_.get("requests_pct")),
                prev.get("sessions", ""),
                _p(cmp_.get("sessions_pct")),
                (round(cmp_["cache_rate_pp"], 2)
                 if cmp_.get("cache_rate_pp") is not None else ""),
                (("今日起步中（%s），环比与归因留空"
                  % ("00:30 前" if s.get("gated_reason") == "early"
                     else "积分未达 %.0f" % (s.get("gate_threshold") or 5)))
                 if s.get("gated") else
                 (agg._attrib_text(cmp_.get("attribution"),
                                   s.get("label") or label) or "")),
            ]
        )
    sheets.append((
        "汇总",
        ["窗口", "起算日", "截至日", "请求数", "会话数", "活跃天数", "输入token",
         "输出token", "token合计", "缓存命中token", "缓存未命中token", "缓存命中率%",
         "思考token", "积分", "官方价预估元", "全高峰价元", "有价token", "费用覆盖率%",
         "免费请求数", "免费token", "子代理请求数", "子代理token",
         "子代理积分", "无积分请求数", "无积分token", "LLM调用次数",
         "联网调用次数", "联网检索次数", "联网抓取次数",
         "对比期", "上期起算日", "上期截至日", "上期积分", "积分环比%", "积分变化量",
         "上期token", "token环比%", "上期请求数", "请求环比%", "上期会话数",
         "会话环比%", "缓存命中变化pp", "环比归因"],
        rows,
    ))

    # 2) 逐请求全量明细（按时间序）
    rows = []
    for r in payload["byTurn"]:
        rows.append(
            [
                r["day"] or "",
                r["session_id"] or "",
                r["title"],
                r["workspace"],
                r["model"],
                r["turn"],
                "是" if r["is_sub"] else "否",
                "是" if r["is_free"] else "否",
                {"ledger": "结算值", "raw": "末次值补齐", "none": "无记录"}[r["credit_src"]],
                r["input"],
                r["output"],
                r["input"] + r["output"],
                r["cache_hit"],
                r["cache_miss"],
                r["thinking"],
                r["credit"],
                round(r.get("cost") or 0, 6),
                {"peak": "高峰", "offpeak": "空闲", "flat": "统一"}.get(
                    r.get("cost_tier"), "未定价"),
            ]
        )
    sheets.append((
        "逐请求",
        ["日期", "会话ID", "会话标题", "工作区", "模型", "轮次", "是否子代理",
         "是否免费", "积分来源", "输入token", "输出token", "token合计",
         "缓存命中token", "缓存未命中token", "思考token", "积分",
         "官方价预估元", "峰谷档"],
        rows,
    ))

    # 2b 段（原 top_requests.csv 的「按积分降序榜单」）已按需求删除：
    # 网页上改成「今日 / 昨日 / 近 7 天 / 近 30 天各取最贵一轮」，四行就够看，
    # 再留一份全量排名反而和 requests.csv 重复。

    # 3) 会话
    rows = []
    for s in payload["bySession"]:
        rows.append(
            [
                s["session_id"] or "",
                s["title"],
                s["workspace"],
                s["model"],
                s.get("first_day") or "",
                s.get("last_day") or "",
                s["requests"],
                s["input"],
                s["output"],
                s["tokens"],
                s["credit"],
                s["sub_requests"],
                s["sub_credit"],
                s["context_used"],
                s["context_size"],
                round(s["context_rate"] * 100, 2) if s.get("context_rate") is not None else "",
                "是" if s.get("deleted") else "否",
                "是" if s.get("playground") else "否",
                round(s.get("cost") or 0, 4),
            ]
        )
    sheets.append((
        "会话",
        ["会话ID", "标题", "工作区", "模型", "首日", "末日", "请求数", "输入token",
         "输出token", "token合计", "积分", "子代理请求数", "子代理积分",
         "上下文占用token", "上下文上限", "上下文占用%", "已删除", "游乐场",
         "官方价预估元"],
        rows,
    ))

    # 4) 模型
    rows = []
    for name, s in payload["byModel"].items():
        rows.append(
            [
                name, s["requests"], s["sessions"], s["input"], s["output"],
                s["tokens"], s["cache_hit"], s["cache_miss"],
                round(s["cache_rate"] * 100, 2) if s["cache_rate"] is not None else "",
                s["thinking"], s["credit"], s["free_requests"], s["free_tokens"],
                s["credit_per_1m"], s["llm_calls"],
                round(s.get("cost") or 0, 4),
                round(s.get("cost_if_peak") or 0, 4),
                s.get("priced_tokens") or 0,
            ]
        )
    sheets.append((
        "模型",
        ["模型", "请求数", "会话数", "输入token", "输出token", "token合计",
         "缓存命中token", "缓存未命中token", "缓存命中率%", "思考token", "积分",
         "免费请求数", "免费token", "积分每百万token", "LLM调用次数",
         "官方价预估元", "全高峰价元", "有价token"],
        rows,
    ))

    # 5) 工作区
    rows = []
    for name, s in payload["byWorkspace"].items():
        rows.append(
            [name, s["requests"], s["sessions"], s["input"], s["output"],
             s["tokens"], s["credit"]]
        )
    sheets.append((
        "工作区",
        ["工作区", "请求数", "会话数", "输入token", "输出token", "token合计", "积分"],
        rows,
    ))

    # 6) 按天
    rows = []
    for day in sorted(payload["byDay"]):
        s = payload["byDay"][day]
        rows.append(
            [
                day, s["requests"], s["sessions"], s["input"], s["output"],
                s["tokens"], s["cache_hit"], s["cache_miss"],
                round(s["cache_rate"] * 100, 2) if s["cache_rate"] is not None else "",
                s["thinking"], s["credit"], s["free_tokens"], s["sub_requests"],
                s["no_credit_requests"], round(s.get("cost") or 0, 4),
            ]
        )
    sheets.append((
        "按天",
        ["日期", "请求数", "会话数", "输入token", "输出token", "token合计",
         "缓存命中token", "缓存未命中token", "缓存命中率%", "思考token", "积分",
         "免费token", "子代理请求数", "无积分请求数", "官方价预估元"],
        rows,
    ))

    # 落盘：一个工作簿六张表
    xlsx.write_workbook(path, sheets)

    with open(path, "rb") as fh:
        b64 = base64.b64encode(fh.read()).decode("ascii")
    _drop_stale_exports(out_dir, os.path.basename(path))
    return path, b64


# ── 终端摘要（≤15 行） ──────────────────────────────────────────
def _w(text):
    """中文按 2 个显示列算，否则终端表格永远对不齐。

    半角 ¥（U+00A5）只占 1 列，按 ord>127 一刀切会让金额列整体左移一格。
    """
    total = 0
    for ch in str(text):
        if ch == "\u00a5":
            total += 1
        else:
            total += 2 if ord(ch) > 127 else 1
    return total


def _cell(text, width, right=False):
    text = str(text)
    pad = width - _w(text)
    if pad <= 0:
        return text
    return (" " * pad + text) if right else (text + " " * pad)


SUMMARY_COLS = [12, 6, 8, 14, 12, 10, 10, 10]


def _print_summary(payload, out_dir, html_path, data_files, json_path=None):
    cov = payload["coverage"]
    scope = payload["summary"]["all"]
    lines = []
    lines.append(
        "WorkBuddy 用量与积分账单 v%s · 数据截至 %s（%s）%s"
        % (VERSION, payload.get("anchor") or "—", payload.get("anchor_label") or "—",
           ("　·　聚焦最近 %d 天" % payload["trimmed_days"])
           if payload.get("trimmed_days") else "")
    )
    lines.append("  数据源  %s（只读）" % os.path.normpath(payload["config_dir"]))
    # 聚焦模式下用裁剪后的口径，避免「覆盖请求数」和表里行数对不上
    lines.append(
        "  覆盖    %s 请求 · %s 次模型调用 · %s 会话 · %s 活跃天"
        % (render.num(scope["requests"]), render.num(scope.get("llm_calls", 0)),
           render.num(scope["sessions"]), render.num(scope["active_days"]))
    )
    # 先扫一遍，把「上期覆盖不足」的标记放到表头，省掉一行脚注
    partial = any(
        (payload["summary"][k].get("compare") or {}).get("partial_prev")
        for k, _ in payload["windows"]
    )
    head = ["窗口", "请求", "会话", "token", "积分", "环比", "缓存命中", "价估"]
    if partial:
        head[-1] = "价估 *"
    lines.append("  " + "".join(
        _cell(t, w, right=(i > 0)) for i, (t, w) in enumerate(zip(head, SUMMARY_COLS))
    ))
    for key, label in payload["windows"]:
        s = payload["summary"][key]
        cmp_ = s.get("compare") or {}
        empty = not s["requests"]
        delta = "—"
        if empty:
            # 0 请求不等于断崖，别在终端上印一个 -100% 让人误判
            delta = "暂无"
        elif s.get("gated"):
            # 起步中：今日刚过零点、或积分还没到门槛，环比是 -99% 量级的噪音。
            # 数字照给，只把「和昨天比」这一列收起来。
            delta = "起步中"
        elif cmp_.get("has_prev") and cmp_.get("credit_pct") is not None:
            delta = render.delta_pct(cmp_["credit_pct"])
            if cmp_.get("partial_prev"):
                delta += "*"
        cells = [
            s.get("label") or label,
            render.num(s["requests"]),
            render.num(s["sessions"]),
            render.compact(s["tokens"]),
            render.credit_fmt(s["credit"]) if not empty else "暂无",
            delta,
            render.pct(s["cache_rate"]),
            render.money(s.get("cost")) if s.get("cost") else "—",
        ]
        # 进行中的单日（今日）打 †：紧挨着就别扭地排一行 -98%，打标是唯一
        # 能让人先看到「这天还没过完」的地方。
        if s.get("partial") and s.get("as_of"):
            cells[0] += "†"
        lines.append("  " + "".join(
            _cell(t, w, right=(i > 0))
            for i, (t, w) in enumerate(zip(cells, SUMMARY_COLS))
        ))
    # 口径与「无积分」提示合并成一行，把行数让给下面的「未计入项」
    cal = "  口径    积分=结算值 %d 条 + 末次值补齐 %d 条；子代理 %d 条已并入" % (
        cov["ledger_credit"], cov["raw_credit"], cov["subagent_requests"])
    if cov["no_credit"]:
        cal += "；%d 条无积分记录（%s token）已计入 token" % (
            cov["no_credit"], render.num(cov["no_credit_tokens"]))
    lines.append(cal)
    if scope.get("cost"):
        saved = (scope.get("cost_if_peak") or 0) - scope["cost"]
        note = "  价估    全部用量 %s（官方牌价·人民币" % render.money(scope["cost"])
        if saved > scope["cost"] * 0.005:
            note += "；全高峰 %s，错峰省 %s" % (
                render.money(scope.get("cost_if_peak")), render.money(saved))
        cov_ = scope.get("cost_coverage")
        if cov_ is not None:
            note += "；覆盖 %s token" % render.pct(cov_)
        lines.append(note + "；仅作参考，非账单）")
    # 联网调用：只报次数。WorkBuddy 侧没有公开的按次牌价，不折钱、不进合计，
    # 也不借用别家（如 DeepSeek Search/Fetch 的 US$0.75/千次）的单价。
    # --no-cost 时不提：那块整体是费用视图，说了「不算钱」又冒一行按次项会自相矛盾。
    web = payload.get("web") or {}
    if web.get("total") and not (payload.get("pricing") or {}).get("disabled"):
        lines.append(
            "  未计入  近 7 日联网检索 / 抓取 %s 次（Search %d · Fetch %d），"
            "只统计调用次数，仅作参考、不能当账单"
            % (render.num(web["total"]), web.get("search", 0), web.get("fetch", 0))
        )
    # 上下文压缩：一次都没发生时不提。压缩是平台在「经济模式」下自动做的，
    # 用户看不到也没有开关；不写这一行，Agent 通道只读终端摘要，就完全看不到
    # 「上下文被省下来了」这件事。两档同值时只报一段，免得同一个数字印两遍。
    comps = []
    for tag, key in (("近 7 日", "7d"), ("全部", "all")):
        c = (payload["summary"].get(key) or {}).get("compact") or {}
        if not c.get("count"):
            continue
        if comps and comps[-1][0] == (c["count"], c["saved"]):
            continue
        comps.append(((c["count"], c["saved"]),
                      "%s %d 次 / 省下约 %s token（占输入 %s）" % (
                          tag, c["count"], render.cn_compact(c["saved"]),
                          render.pct(c.get("share")))))
    if comps:
        lines.append("  压缩    " + "；".join(text for _, text in comps)
                     + "；毛估算，非积分节省")
    if partial:
        lines.append("  备注    * 上期起始日早于本机最早记录日，环比只能看方向")
    live_today = payload["summary"].get("today") or {}
    if live_today.get("partial") and live_today.get("as_of"):
        lines.append("  备注    † 今日还没过完（截至 %s），环比对着昨日整天，只作方向参考"
                     % live_today["as_of"])
    if data_files:
        tail = "　+　%s" % os.path.basename(data_files[0])
        if json_path:
            tail += "　+　data.json"
        lines.append("  产物    %s%s" % (html_path or out_dir, tail))
    elif html_path:
        lines.append("  产物    %s" % html_path)
    print("\n".join(lines))


def _emit_json(payload, path):
    printable = dict(payload)
    # data_b64 是给 HTML 内嵌用的整份工作簿 base64，写进 data.json 只会把文件
    # 撑大一倍、还没人用得上——要看数据的人直接开 xlsx。data_name 留着当指针。
    printable.pop("data_b64", None)
    printable.pop("generated_at", None)
    printable["generated_at"] = payload["generated_at"].isoformat()
    # byTurn 全量保留，它是「逐条对话」的原始依据
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(printable, fh, ensure_ascii=False, indent=1, default=str)


def _open_file(path):
    try:
        if os.name == "nt":
            os.startfile(path)  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception:
        pass


def run(args):
    bundle = collect.build(args.config_dir, fx=args.fx,
                           with_cost=not args.no_cost,
                           use_cache=not args.no_cache)
    try:
        anchor, anchor_mode = agg.resolve_anchor(bundle, args.anchor)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    bundle = agg.apply_anchor(bundle, anchor)
    if args.days and args.days > 0:
        bundle = agg.trim_bundle(bundle, args.days, anchor)
    payload = agg.aggregate(bundle, today_gate_mode=args.today_gate)
    payload["anchor_mode"] = anchor_mode
    # HTML 里的「完整数据」链接只在真的会写工作簿时生成，避免死链
    payload["export_data"] = args.format in ("xlsx", "all")

    if not payload["coverage"]["requests"]:
        print("未取到任何请求记录。检查 %s 下是否存在 projects/*.jsonl 与 workbuddy.db"
              % bundle["config_dir"], file=sys.stderr)
        return 2

    if args.format == "text":
        _print_summary(payload, None, None, None)
        return 0

    if args.format == "json":
        # stdout 只给纯 JSON，便于管道消费；摘要让位给 --format text
        if args.out:
            path = args.out if args.out.endswith(".json") else os.path.join(args.out, "data.json")
            os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
            _emit_json(payload, path)
            _print_summary(payload, None, None, None)
            print("  JSON    %s" % path)
        else:
            print(json.dumps(
                {k: v for k, v in payload.items() if k != "byTurn"},
                ensure_ascii=False, indent=1, default=str))
        return 0

    # 归档目录按**基准日**命名，而不是运行日：00:05 跑出来的是前一日的账单，
    # 放进前一日目录才符合「一天一份」的直觉。
    out_dir = os.path.expanduser(
        args.out or os.path.join(DEFAULT_OUT_ROOT, payload.get("anchor") or "")
    )
    os.makedirs(out_dir, exist_ok=True)
    _drop_legacy_csv_dir(out_dir)

    html_path = None
    data_files = []
    json_path = None

    # 工作簿必须先于 HTML：HTML 要把整份 xlsx 内嵌成 data URI，否则在
    # 「单文件预览」里点「完整数据」只会拿到 404——同目录的文件不在托管范围内。
    if args.format in ("xlsx", "all"):
        data_path, data_b64 = _export_workbook(payload, out_dir)
        data_files = [data_path]
        payload["export_data"] = True
        payload["data_name"] = os.path.basename(data_path)
        payload["data_b64"] = data_b64

    if args.format in ("html", "all"):
        html_path = os.path.join(out_dir, "index.html")
        with open(html_path, "w", encoding="utf-8") as fh:
            fh.write(render.render_html(payload))

    if args.format in ("json", "all"):
        json_path = os.path.join(out_dir, "data.json")
        _emit_json(payload, json_path)

    _print_summary(payload, out_dir, html_path, data_files, json_path)

    if args.open and html_path:
        _open_file(html_path)
    return 0


def main(argv=None):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(
        prog="wb-usage-report",
        description="生成 WorkBuddy 本机用量与积分账单（今日/昨日/7日/30日/全部）",
    )
    parser.add_argument("--config-dir", default=None,
                        help="WorkBuddy 配置目录，默认 ~/.workbuddy")
    parser.add_argument("--out", default=None,
                        help="输出目录，默认 ~/.workbuddy/usage-reports/<基准日>/")
    parser.add_argument("--format", default="all",
                        choices=["all", "html", "xlsx", "json", "text"],
                        help="产出格式，默认 all（HTML + Excel 工作簿 + JSON）")
    parser.add_argument("--anchor", default="auto",
                        help="数据截至日：auto（默认，06:00 前自动退到前一天）/"
                             "today/prev/YYYY-MM-DD")
    parser.add_argument("--days", type=int, default=0,
                        help="只统计最近 N 个自然日（0=全部）")
    parser.add_argument("--today-gate", default="auto", choices=["auto", "off"],
                        help="今日档的「起步中」闸门：auto（默认，00:30 前或今日积分"
                             "不足 %.0f 时只给数字、不算环比）/ off（强制照常算环比）"
                             % agg.TODAY_MIN_CREDIT)
    parser.add_argument("--fx", type=float, default=None,
                        help="美元牌价折人民币的汇率，默认 %.2f；只影响仅发布美元价的模型"
                             % 7.10)
    parser.add_argument("--no-cost", action="store_true",
                        help="不做费用预估，只看 token 与积分")
    parser.add_argument("--no-cache", action="store_true",
                        help="忽略日志解析缓存，全部 jsonl 重新解析（慢，仅在怀疑缓存陈旧时用）")
    parser.add_argument("--open", action="store_true", help="生成后立即打开 HTML")
    parser.add_argument("--version", action="version", version="wb-usage-report " + VERSION)
    return run(parser.parse_args(argv))


if __name__ == "__main__":
    sys.exit(main())
