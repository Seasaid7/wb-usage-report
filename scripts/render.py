"""渲染层：把聚合结果变成单文件离线 HTML。

约束（刻意为之）
    1. 零外部引用。CSS 内联、图表内联 SVG，不引任何 CDN，断网也能看。
    2. 不依赖模板引擎。全部显式字符串拼接，避免占位符替换失败留下半成品。
    3. 所有来自数据的字符串一律 html.escape，防止会话标题里的尖括号破版。
"""

from __future__ import annotations

import calendar as _cal
import html
import json
import math
import os
import random
import re
from datetime import date, datetime, timedelta

import pricing

# ── 配色（深色金融风） ────────────────────────────────────────────
# 对比度实测（on #101724）：正文 14.9:1、次要 8.7:1、注释 5.0:1、金 10.0:1。
# 注释色原来是 #64707f（3.0:1），11px 的中文小字在那个对比度下基本看不清，
# 所以整条灰度梯子被整体提亮了一档——这是「易读性」最直接的一处改动。
C_BG = "#0a0d13"
C_PANEL = "#101724"      # 面板底色
C_PANEL2 = "#161f2e"     # 次层面板 / 卡片渐变顶
C_PANEL3 = "#1c2636"     # 三级面：徽标底、输入框、按钮 hover
C_LINE = "#232d3d"       # 主分隔线
C_LINE2 = "#1a2331"      # 章节细线 / 表格行线
C_FG = "#eef2f8"
C_FG2 = "#aab5c6"
C_FG3 = "#7d8896"
C_GOLD = "#e8b93f"   # 积分（唯一强调色）
C_BLUE = "#5b93ff"   # token
C_GREEN = "#3fb950"  # 缓存命中
C_ORANGE = "#f0883e" # 提醒
C_PURPLE = "#a371f7" # 子代理
C_RED = "#ff6b62"

SVG_W = 1120

# 完整数据是 xlsx，下载链接的 data URI 前缀必须带正确的 MIME。
# 写成 text/csv 会让浏览器把 .xlsx 当成 csv 存下来，双击被 Excel 判为损坏。
XLSX_DATA_URI = ("data:application/vnd.openxmlformats-officedocument"
                 ".spreadsheetml.sheet;base64,")


def esc(text):
    return html.escape("" if text is None else str(text), quote=True)


def num(value):
    try:
        return "{:,}".format(int(round(float(value or 0))))
    except (TypeError, ValueError):
        return "0"


def compact(value):
    """大数字压缩成 1.2M / 345K，KPI 卡里用。"""
    try:
        v = float(value or 0)
    except (TypeError, ValueError):
        return "0"
    for unit, size in (("B", 1e9), ("M", 1e6), ("K", 1e3)):
        if abs(v) >= size:
            return ("%.2f%s" % (v / size, unit)).replace(".00", "")
    return "%d" % v


def cn_compact(value):
    """中文量级：1.2 亿 / 345 万 / 3,204。

    KPI 卡里用 K/M 是对的——那是一个数字占的位置，越短越好。但「上下文压缩」
    那句话是要读的，「1.65 亿 token」比「165.23M token」少做一次换算。
    """
    try:
        v = float(value or 0)
    except (TypeError, ValueError):
        return "0"
    if abs(v) >= 1e8:
        return ("%.2f 亿" % (v / 1e8)).replace(".00 亿", " 亿")
    if abs(v) >= 1e4:
        q = v / 1e4
        text = ("%.2f" % q) if abs(q) < 1000 else ("%.0f" % q)
        return text.rstrip("0").rstrip(".") + " 万"
    return "{:,.0f}".format(v)


def credit_fmt(value, digits=2):
    try:
        return ("{:,." + str(digits) + "f}").format(float(value or 0))
    except (TypeError, ValueError):
        return "0.00"


def pct(value, digits=1):
    if value is None:
        return "—"
    try:
        return ("%." + str(digits) + "f%%") % (float(value) * 100)
    except (TypeError, ValueError):
        return "—"


def money(value, digits=2):
    """人民币。费用预估统一走这个函数，避免各处混用美元符号。"""
    if value is None:
        return "—"
    try:
        return ("¥{:,.%df}" % digits).format(float(value))
    except (TypeError, ValueError):
        return "—"


def nice_max(value, ticks=4):
    """把坐标轴上限抬成「好看的整数」，让刻度落在 600/450/300/150 这种数上。

    直接拿数据最大值当上限，刻度就会变成 604 / 453 / 302 / 151 ——数字本身
    就是噪声，读图的人得先在心里做一次除法才知道柱子大概多高。
    """
    try:
        v = float(value or 0)
    except (TypeError, ValueError):
        return 1.0
    if v <= 0:
        return 1.0
    raw = v / float(ticks)
    mag = 10 ** math.floor(math.log10(raw))
    step = mag * 10
    for mult in (1, 2, 2.5, 5, 10):
        step = mult * mag
        if step >= raw:
            break
    return step * ticks


# ── 环比标记 ─────────────────────────────────────────────────────
# 语义分两类，颜色不能混：
#   成本类（积分/token/请求）上升 = 更贵 = bad（红）；下降 = 更省 = good（绿）
#   质量类（缓存命中率）上升 = 更好 = good（绿）
def delta_pct(value, digits=1):
    """带正负号的百分比。None → 上期无数据，显示破折号而不是假 0。"""
    if value is None:
        return None
    return ("%+." + str(digits) + "f%%") % (float(value) * 100)


def _tone(value, higher_is_better=False, eps=0.02):
    if value is None or abs(value) <= eps:
        return "flat"
    return "good" if (value > 0) == higher_is_better else "bad"


def delta_badge(value, digits=1, extra="", higher_is_better=False, eps=0.02, mark=""):
    text = delta_pct(value, digits)
    if text is None:
        return '<span class="delta flat">上期无数据</span>'
    body = '<span class="delta %s">%s%s</span>' % (
        _tone(value, higher_is_better, eps), text, mark)
    if extra:
        body += '<span class="delta-note">%s</span>' % esc(extra)
    return body


def pp_badge(value, digits=1, higher_is_better=True):
    """百分点差，用于缓存命中率这种本身就是百分比的指标。"""
    if value is None:
        return '<span class="delta flat">—</span>'
    return '<span class="delta %s">%+.*fpp</span>' % (
        _tone(value, higher_is_better, 0.1), digits, value)


def _credit_class(value):
    """积分越高颜色越暖，让最贵的行自己跳出来。"""
    if value >= 500:
        return "hot"
    if value >= 100:
        return "warm"
    if value > 0:
        return "cool"
    return "zero"


# ── KPI 五档卡 ───────────────────────────────────────────────────
def _kpi_cards(payload):
    cards = []
    for key, label in payload["windows"]:
        s = payload["summary"][key]
        label = s.get("label") or label
        since, until = s.get("since"), s.get("until")
        if since and until and since != until:
            sub = "%s ~ %s" % (since[5:], until[5:])
        elif since:
            sub = since
        else:
            sub = "全部历史"
        # 当天还没过完时必须写明「截至几点」，否则会被当成整天数据来比
        if s.get("partial"):
            sub += " · 截至 %s" % (s.get("as_of") or "现在")

        # 空窗口：0 请求不是「用量塌陷」。卡片必须留在原位，内容写暂无，
        # 不然用户会以为报表根本没统计这一天。
        empty = bool(s.get("empty") or not s.get("requests"))

        if empty:
            main = ('<div class="kpi-main"><span class="kpi-none">暂无</span>'
                    '<span class="kpi-unit">数据</span></div>')
            delta = ('<div class="kpi-delta"><span class="delta flat">'
                     "该区间无请求，不做环比</span></div>")
            tokens = ('<div class="kpi-tokens none">尚无 token 记录</div>')
            cost_line = ('<div class="kpi-cost dim">无请求，不产生费用</div>')
        else:
            cmp_ = s.get("compare")
            if s.get("gated"):
                # 起步中：数字照给，环比整块换掉。今日刚过零点（或积分还没到
                # 门槛）时，「-98.9%」读起来像用量塌陷，其实只是还没开工。
                delta = ('<div class="kpi-delta">'
                         '<span class="delta flat">起步中</span>'
                         '<span class="kpi-delta-label">样本不足，不算环比</span>'
                         "</div>")
                token_delta = ""
            elif cmp_:
                # 上期覆盖不足时打星号，避免把「本机用得晚」读成「用量暴涨」
                mark = "*" if cmp_.get("partial_prev") else ""
                delta = (
                    '<div class="kpi-delta">%s<span class="kpi-delta-label">'
                    "环比%s %s</span></div>"
                    % (delta_badge(cmp_["credit_pct"], mark=mark), esc(cmp_["label"]),
                       credit_fmt(cmp_["prev"]["credit"]))
                )
                token_delta = delta_badge(cmp_["tokens_pct"], mark=mark)
            else:
                # 「全部」这一档没有等长上期，环比那一行整行不出现。写一句
                # 「无上期可比」只是把「这里没有信息」印成一个要读的句子，
                # 占位反而招人看。空 div 仅用于撑住卡片高度，与其余卡对齐。
                delta = '<div class="kpi-delta"></div>'
                token_delta = ""
            main = ('<div class="kpi-main"><span class="kpi-credit">%s</span>'
                    '<span class="kpi-unit">积分</span></div>' % credit_fmt(s["credit"]))
            tokens = ('<div class="kpi-tokens">%s <span>token</span>%s</div>'
                      % (compact(s["tokens"]),
                         (' <span class="kpi-token-delta">%s</span>' % token_delta
                          if token_delta else "")))
            # 费用预估：人民币口径，只算有官方牌价的模型；覆盖率不足时如实标注
            cost = s.get("cost") or 0
            if cost > 0:
                cov = s.get("cost_coverage")
                note = "按官方价折算"
                if cov is not None and cov < 0.995:
                    note += " · 覆盖 %s token" % pct(cov)
                cost_line = ('<div class="kpi-cost">≈%s <span>%s</span></div>'
                             % (money(cost), esc(note)))
            else:
                cost_line = ('<div class="kpi-cost dim">该区间无公开牌价模型，'
                             "不计费用</div>")

        cache_cell = "—" if s.get("cache_rate") is None else pct(s["cache_rate"])
        cards.append(
            '<div class="kpi%s">'
            '<div class="kpi-head"><span class="kpi-label">%s</span>'
            '<span class="kpi-sub">%s</span></div>'
            "%s"
            "%s"
            "%s"
            "%s"
            '<dl class="kpi-grid">'
            '<div><dt>请求</dt><dd>%s</dd></div>'
            '<div><dt>LLM 调用</dt><dd>%s</dd></div>'
            '<div><dt>会话</dt><dd>%s</dd></div>'
            '<div><dt>活跃天</dt><dd>%s</dd></div>'
            '<div><dt>缓存命中</dt><dd>%s</dd></div>'
            '<div><dt>无积分请求</dt><dd class="dim">%s</dd></div>'
            "</dl></div>"
            % (
                " empty" if empty else "",
                esc(label),
                esc(sub),
                main,
                delta,
                tokens,
                cost_line,
                num(s["requests"]),
                num(s.get("llm_calls", 0)),
                num(s["sessions"]),
                num(s["active_days"]),
                cache_cell,
                num(s["no_credit_requests"]),
            )
        )
    return '<section class="kpi-row">' + "".join(cards) + "</section>"


def _compact_strip(payload):
    """KPI 下方一行：上下文压缩触发了几次、大概省下多少输入 token。

    为什么值得单占一行：压缩是平台在「经济模式」下自动做的，用户既看不到、
    也没有开关。它省下的量可以大到同期输入的 10%，不写这一句，「上下文省下来了」
    这件事在报表里就完全不可见。

    只取「近 7 天」与「全部」两档：前者回答「最近」，后者回答「总共」。
    两档都等于同一个数时（数据不足 7 天）只留一段，免得同一个数字印两遍。
    一次压缩都没发生是常态——短会话撞不到上下文天花板，此时整条不渲染，
    不留一行「0 次」占位置。
    """
    segs = []
    for key, fallback in (("7d", "近 7 天"), ("all", "全部")):
        s = (payload.get("summary") or {}).get(key) or {}
        c = s.get("compact") or {}
        if not c.get("count"):
            continue
        label = s.get("label") or fallback
        seg = (
            '<span class="comp-seg"><span class="comp-win">%s</span>'
            "发生 <b>%s</b> 次，省下约 <b>%s</b> token"
            '<span class="comp-dim">（占输入 %s）</span></span>'
            % (esc(label), num(c["count"]), cn_compact(c["saved"]), pct(c.get("share")))
        )
        if segs and segs[-1][0] == (c["count"], c["saved"]):
            # 和上一段完全同值：只把「全部」那一段去掉，保留先出现的近 7 天
            continue
        segs.append(((c["count"], c["saved"]), seg))

    if not segs:
        return ""
    return (
        '<div class="comp-strip"><span class="comp-k">上下文压缩</span>%s'
        '<span class="comp-tip">输入 token 毛估算，非积分节省</span></div>'
        % "".join(seg for _, seg in segs)
    )


# ── 环比表 ───────────────────────────────────────────────────────
def _compare_table(payload):
    rows = []
    for key, label in payload["windows"]:
        s = payload["summary"][key]
        cmp_ = s.get("compare")
        if not cmp_:
            # 「全部」这一档没有等长上期可比，整行删掉。留一行全是「—」的占位，
            # 既没信息量又抢走注意力——没有环比就不显示，这是这张表的唯一职责。
            continue
        name = s.get("label") or label
        since, until = s.get("since"), s.get("until")
        head = '<div class="cell-title">%s</div>' % esc(name)
        if since and until and since != until:
            head += '<div class="cell-sub mono">%s ~ %s</div>' % (esc(since[5:]), esc(until[5:]))
        elif since:
            head += '<div class="cell-sub mono">%s</div>' % esc(since)
        else:
            head += '<div class="cell-sub">全部历史</div>'



        prev = cmp_["prev"]
        prev_head = '<div class="cell-title">%s</div><div class="cell-sub mono">%s~%s</div>' % (
            esc(cmp_["label"]), esc((prev.get("since") or "")[5:]),
            esc((prev.get("until") or "")[5:]))
        if cmp_.get("partial_prev"):
            prev_head += ('<div class="cell-sub warn-text">上期仅覆盖 %d/%d 天</div>'
                          % (cmp_.get("prev_covered_days", 0),
                             cmp_.get("prev_window_days", 0)))

        # 本期没有任何请求：环比必然是 -100%，那不是结论而是噪声。
        # 上一期的数字照常摆出来，本期一律写「暂无」。
        if cmp_.get("cur_empty"):
            na = {"text": '<span class="delta flat">本期暂无数据</span>', "cls": "c"}
            rows.append(
                [
                    head,
                    prev_head,
                    {"text": '<span class="cell-sub">暂无</span>'},
                    {"text": credit_fmt(prev["credit"]), "cls": "dim"},
                    na,
                    {"text": '<span class="delta flat">—</span>', "cls": "c"},
                    {"text": '<span class="delta flat">—</span>', "cls": "c"},
                    {"text": '<span class="delta flat">—</span>', "cls": "c"},
                    {"text": "—", "cls": "c"},
                ]
            )
            continue

        # 起步中的今日档：本期数字照给，四个环比列全留「—」。不是没有上期，
        # 是这期样本太薄——比出来的 -97% 只会被读成用量塌陷，而不是「还没开工」。
        if s.get("gated"):
            rows.append(
                [
                    head + '<div class="cell-sub">起步中 · 样本不足</div>',
                    prev_head,
                    {"text": credit_fmt(s["credit"])},
                    {"text": credit_fmt(prev["credit"]), "cls": "dim"},
                    {"text": '<span class="delta flat">—</span>', "cls": "c"},
                    {"text": '<span class="delta flat">—</span>', "cls": "c"},
                    {"text": '<span class="delta flat">—</span>', "cls": "c"},
                    {"text": '<span class="delta flat">—</span>', "cls": "c"},
                    {"text": "—", "cls": "c"},
                ]
            )
            continue

        cache_cell = '%s<div class="cell-sub">%s→%s</div>' % (
            pp_badge(cmp_["cache_rate_pp"]), pct(prev["cache_rate"]),
            pct(s["cache_rate"]))
        credit_extra = "%+.2f" % cmp_["credit_abs"]
        if cmp_.get("partial_prev"):
            credit_extra += "*"
        rows.append(
            [
                head,
                prev_head,
                {"text": credit_fmt(s["credit"]), "cls": _credit_class(s["credit"])},
                {"text": credit_fmt(prev["credit"]), "cls": "dim"},
                {"text": delta_badge(cmp_["credit_pct"], 1, "(%s)" % credit_extra),
                 "cls": "c"},
                {"text": delta_badge(cmp_["tokens_pct"]), "cls": "c"},
                {"text": delta_badge(cmp_["requests_pct"], 0,
                                     "(%d→%d)" % (prev["requests"], s["requests"])),
                 "cls": "c"},
                {"text": delta_badge(cmp_["sessions_pct"], 0), "cls": "c"},
                {"text": cache_cell, "cls": "c"},
            ]
        )
    table = _table(
        ["窗口", "对比期", "本期积分", "上期积分", "积分环比", "token 环比",
         "请求环比", "会话环比", "缓存命中"],
        rows, cls="cmp",
        align=["l", "l", "r", "r", "r", "r", "r", "r", "r"],
    )
    if any((payload["summary"][k].get("compare") or {}).get("partial_prev")
           for k, _ in payload["windows"]):
        table += (
            '<p class="foot-note">* 上期起始日早于本机最早记录日，成交量被低估，'
            "该环比只能看方向、不能看幅度。</p>"
        )
    empty_win = [
        payload["summary"][k].get("label") or l
        for k, l in payload["windows"]
        if (payload["summary"][k].get("compare") or {}).get("cur_empty")
    ]
    if empty_win:
        table += (
            '<p class="foot-note">%s 本期没有请求记录，环比列写「暂无」而不写 '
            "-100%%——那是「还没开始用」，不是「用量塌陷」。上期数字照常保留，"
            "方便对照。</p>" % esc("、".join(empty_win))
        )
    return table


# ── 趋势图：积分柱 + token 线（双轴） ────────────────────────────
def _trend_chart(payload):
    days = payload["byDay"]
    if not days:
        return '<p class="empty">暂无按天数据</p>'

    keys = list(days.keys())
    n = len(keys)
    h = 260
    pad_l, pad_r, pad_t, pad_b = 56, 64, 22, 34
    plot_w = SVG_W - pad_l - pad_r
    plot_h = h - pad_t - pad_b
    slot = plot_w / float(n)
    bar_w = max(2.0, min(18.0, slot * 0.56))

    max_credit = nice_max(max((days[k]["credit"] for k in keys), default=0), 4)
    max_tokens = nice_max(max((days[k]["tokens"] for k in keys), default=0), 4)

    parts = [
        '<svg viewBox="0 0 %d %d" width="100%%" height="%d" role="img">' % (SVG_W, h, h)
    ]

    # 横向网格
    for i in range(5):
        y = pad_t + plot_h * i / 4.0
        parts.append(
            '<line x1="%d" y1="%.1f" x2="%d" y2="%.1f" stroke="%s" stroke-width="1"/>'
            % (pad_l, y, SVG_W - pad_r, y, C_LINE2 if i else C_LINE)
        )
        parts.append(
            '<text x="%d" y="%.1f" fill="%s" font-size="10" text-anchor="end">%s</text>'
            % (pad_l - 8, y + 3.5, C_FG3,
               credit_fmt(max_credit * (4 - i) / 4.0, 0))
        )
        parts.append(
            '<text x="%d" y="%.1f" fill="%s" font-size="10" text-anchor="start">%s</text>'
            % (SVG_W - pad_r + 8, y + 3.5, C_FG3,
               compact(max_tokens * (4 - i) / 4.0))
        )

    # 积分柱
    for i, key in enumerate(keys):
        v = days[key]["credit"]
        bh = plot_h * (v / max_credit) if max_credit else 0
        x = pad_l + slot * i + (slot - bar_w) / 2.0
        y = pad_t + plot_h - bh
        parts.append(
            '<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="1.5" fill="%s" '
            'opacity="0.85"><title>%s｜%s 积分</title></rect>'
            % (x, y, bar_w, max(bh, 0.6), C_GOLD, esc(key), credit_fmt(v))
        )

    # token 折线
    pts = []
    for i, key in enumerate(keys):
        v = days[key]["tokens"]
        x = pad_l + slot * i + slot / 2.0
        y = pad_t + plot_h * (1 - v / float(max_tokens))
        pts.append((x, y))
    parts.append(
        '<polyline fill="none" stroke="%s" stroke-width="2" stroke-linejoin="round" '
        'points="%s"/>'
        % (C_BLUE, " ".join("%.1f,%.1f" % p for p in pts))
    )
    for i, key in enumerate(keys):
        x, y = pts[i]
        parts.append(
            '<circle cx="%.1f" cy="%.1f" r="3" fill="%s">'
            "<title>%s｜%s token／%s 积分</title></circle>"
            % (x, y, C_BLUE, esc(key), num(days[key]["tokens"]),
               credit_fmt(days[key]["credit"]))
        )

    # x 轴标签，最多 12 个
    step = max(1, n // 12)
    for i, key in enumerate(keys):
        if i % step and i != n - 1:
            continue
        parts.append(
            '<text x="%.1f" y="%d" fill="%s" font-size="10" text-anchor="middle">%s</text>'
            % (pad_l + slot * i + slot / 2.0, h - 12, C_FG3, esc(key[5:]))
        )

    # 图例（把「哪根轴是谁」写进图例，比在轴上竖排一行小字稳得多）
    parts.append(
        '<g font-size="10.5">'
        '<rect x="%d" y="4" width="9" height="9" rx="2" fill="%s"/>'
        '<text x="%d" y="12.5" fill="%s">每日积分（左轴）</text>'
        '<line x1="%d" y1="8.5" x2="%d" y2="8.5" stroke="%s" stroke-width="2"/>'
        '<text x="%d" y="12.5" fill="%s">每日 token（右轴）</text>'
        "</g>"
        % (pad_l, C_GOLD, pad_l + 14, C_FG2,
           pad_l + 118, pad_l + 140, C_BLUE, pad_l + 146, C_FG2)
    )
    parts.append("</svg>")
    return "".join(parts)


# ── 活跃日历 ─────────────────────────────────────────────────────
def _cal_weeks(lo, hi):
    """lo 所在周的周一 + 覆盖 lo..hi 需要的周数。"""
    first_monday = lo - timedelta(days=lo.weekday())
    last_monday = hi - timedelta(days=hi.weekday())
    return first_monday, (last_monday - first_monday).days // 7 + 1


def _cal_svg(days, start_monday, span, lo, hi, anchor, today, max_credit):
    """画一个视图的周历：一行一周，周一到周日横向。

    `lo..hi` 之外的日子**不画格**。宁可留空，也不画成「无记录」——那会让人
    以为那天真的没有用量，而实际只是不在本视图范围内。
    """
    cell, gap = 28, 6
    pitch = cell + gap
    left, head_h = 74, 20
    w = left + 7 * pitch - gap + 4
    h = head_h + span * pitch + 22
    parts = ['<svg viewBox="0 0 %d %d" width="%d" height="%d" role="img">' % (w, h, w, h)]

    # 表头：周一 → 周日，横向
    for i, name in enumerate(["周一", "周二", "周三", "周四", "周五", "周六", "周日"]):
        parts.append(
            '<text x="%.1f" y="11" fill="%s" font-size="10" text-anchor="middle">%s</text>'
            % (left + i * pitch + cell / 2.0, C_FG3, name)
        )

    prev_month = None
    for row in range(span):
        wk = start_monday + timedelta(weeks=row)
        y = head_h + row * pitch
        # 换月才标一次，避免每行重复「10月」把左边塞满
        if wk.month != prev_month:
            parts.append(
                '<text x="0" y="%.1f" fill="%s" font-size="11" font-weight="600">'
                "%d月</text>" % (y + cell - 9, C_GOLD, wk.month)
            )
            prev_month = wk.month
        parts.append(
            '<text x="%d" y="%.1f" fill="%s" font-size="10" text-anchor="end">'
            "%d/%d</text>" % (left - 9, y + cell - 9, C_FG3, wk.month, wk.day)
        )

        for wd in range(7):
            day = wk + timedelta(days=wd)
            if day < lo or day > hi:
                continue
            key = day.strftime("%Y-%m-%d")
            x = left + wd * pitch
            block = days.get(key)
            if day > anchor:
                fill, numcol = C_BG, "#333b49"
                tip = "%s｜尚未到" % key
            elif block:
                ratio = min(1.0, block["credit"] / max_credit)
                if block["credit"] <= 0:
                    fill, numcol = C_PANEL2, C_FG3
                elif ratio < 0.15:
                    fill, numcol = "#2b2a1c", C_FG2
                elif ratio < 0.4:
                    fill, numcol = "#5c4a1a", C_FG
                elif ratio < 0.7:
                    fill, numcol = "#9c7a1f", "#12140a"
                else:
                    fill, numcol = C_GOLD, "#1b1403"
                tip = "%s｜%.2f 积分｜%s token｜%d 请求" % (
                    key, block["credit"], num(block["tokens"]), block["requests"])
            else:
                fill, numcol = "#0f141d", C_FG3
                tip = "%s｜无记录" % key
            if day == today:
                stroke, sw = C_GOLD, 2
                tip += "｜今日"
            else:
                stroke, sw = C_LINE, 0.5
            parts.append(
                '<rect x="%d" y="%d" width="%d" height="%d" rx="5" fill="%s" '
                'stroke="%s" stroke-width="%s"><title>%s</title></rect>'
                % (x, y, cell, cell, fill, stroke, sw, esc(tip))
            )
            parts.append(
                '<text x="%.1f" y="%.1f" fill="%s" font-size="10" text-anchor="middle" '
                'pointer-events="none">%d</text>'
                % (x + cell / 2.0, y + cell / 2.0 + 3.6, numcol, day.day)
            )

    # 图例：少 → 多；「今日」描金边的说明只在今天确实落在本视图里时才画，
    # 否则图例在讲一个画面里不存在的记号。
    ly = h - 4
    lx = left
    parts.append('<text x="%d" y="%d" fill="%s" font-size="9">少</text>'
                 % (lx - 16, ly, C_FG3))
    for i, fill in enumerate(["#0f141d", "#2b2a1c", "#5c4a1a", "#9c7a1f", C_GOLD]):
        parts.append('<rect x="%d" y="%d" width="9" height="9" rx="2" fill="%s"/>'
                     % (lx + i * 13, ly - 8, fill))
    parts.append('<text x="%d" y="%d" fill="%s" font-size="9">多</text>'
                 % (lx + 5 * 13 + 4, ly, C_FG3))
    if lo <= today <= hi:
        parts.append(
            '<rect x="%d" y="%d" width="9" height="9" rx="2" fill="none" stroke="%s" '
            'stroke-width="1.6"/><text x="%d" y="%d" fill="%s" font-size="9">今日</text>'
            % (lx + 5 * 13 + 26, ly - 8, C_GOLD, lx + 5 * 13 + 39, ly, C_FG3))
    parts.append("</svg>")
    return "".join(parts)


def _calendar(payload, months_back=6, window_days=30):
    """按月分页的活跃日历 —— 一屏只画一个月，行数随该月的周数变化。

    旧版把「有数据的全部日期」压成**一张**竖排周历，靠 `max_weeks=12` 封顶：
    数据攒到 3 个月以上就恒为 12 行（约 410px 高），比结论区还占地方，而日历
    要回答的只是「最近忙不忙」，不是「从第一天到今天」。所以改成两种视图：

      · **近 30 天**（默认）—— 5~6 行，回答「这个月怎么样」
      · **每个自然月一页** —— 4~6 行，按月往回翻

    行数跟着月份变，不补空周：补空周会让人把「不在本视图内」误读成「那几周
    完全没用量」。切换靠一小段内联脚本做显隐，所有视图都是渲染期就画好的 SVG，
    断网、无 JS 也能看（默认那页带 is-on，永远可见）。
    """
    days = payload.get("byDay") or {}
    if not days:
        return "", ""
    try:
        anchor = datetime.strptime(
            payload.get("anchor") or payload.get("today"), "%Y-%m-%d").date()
    except (TypeError, ValueError):
        anchor = payload["generated_at"].date()
    try:
        today = datetime.strptime(payload.get("today") or "", "%Y-%m-%d").date()
    except (TypeError, ValueError):
        today = anchor
    try:
        first = datetime.strptime(min(days), "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return "", ""

    # 色阶用**全局**最大值，不按视图各算各的：否则一个月很淡、下个月很浓，
    # 切过去会误以为用量翻了几倍。
    max_credit = max((v["credit"] for v in days.values()), default=0) or 1

    months = []
    y, m = first.year, first.month
    while (y, m) <= (anchor.year, anchor.month):
        months.append((y, m))
        m += 1
        if m > 12:
            y, m = y + 1, 1
    months = months[-months_back:]

    views = []
    for y, m in months:
        lo = date(y, m, 1)
        hi = date(y, m, _cal.monthrange(y, m)[1])
        views.append({"month": True, "lo": lo, "hi": hi})
    views.append({"month": False,
                  "lo": anchor - timedelta(days=window_days - 1), "hi": anchor})

    blocks = []
    for i, view in enumerate(views):
        lo, hi = view["lo"], view["hi"]
        start_monday, span = _cal_weeks(lo, hi)
        pool = [b for k, b in days.items()
                if lo <= datetime.strptime(k, "%Y-%m-%d").date() <= hi]
        credit = sum(b["credit"] for b in pool)
        active = sum(1 for b in pool if b.get("requests"))
        if view["month"]:
            label = "%d 月 · 共 %d 周" % (lo.month, span)
        else:
            label = "近 %d 天 · %d/%d ~ %d/%d · 共 %d 周" % (
                window_days, lo.month, lo.day, hi.month, hi.day, span)
        blocks.append(
            '<div class="cal-view%s" data-cal-key="%s">'
            '<div class="cal-cap">%s · <b>%s</b> 积分 · 活跃 %d 天</div>%s</div>'
            % (" is-on" if i == len(views) - 1 else "",
               "window" if not view["month"] else lo.strftime("%Y-%m"),
               label, credit_fmt(credit), active,
               _cal_svg(days, start_monday, span, lo, hi, anchor, today, max_credit))
        )

    body = (
        '<div class="cal-bar">'
        '<button type="button" class="cal-btn cal-prev" data-cal="-1" '
        'aria-label="上一个视图">‹</button>'
        '<button type="button" class="cal-btn cal-next" data-cal="1" '
        'aria-label="下一个视图">›</button>'
        '<span class="cal-count"></span>'
        "</div>"
        '<div class="cal-stage">%s</div>' % "".join(blocks)
    )
    return body, "颜色深浅=当日积分　·　默认近 30 天，可按月往回翻"


# ── 时段分布：哪个小时在烧积分 ───────────────────────────────────
def _hour_chart(payload):
    # 近 30 天口径。旧版吃 byHour（全部历史），早期作息一漂就把最近的习惯糊平。
    hours = payload.get("byHour30") or payload.get("byHour") or {}
    if not hours:
        return '<p class="empty">暂无时段数据</p>'
    w, h = 640, 236
    # 左侧留出刻度位：原先 pad_l=8、只有网格没有刻度数字，柱子的高度完全
    # 读不出量级——「哪一小时最贵」看得见，「贵多少」看不见。
    pad_l, pad_b, pad_t = 46, 28, 14
    plot_w = w - pad_l - 10
    plot_h = h - pad_b - pad_t
    slot = plot_w / 24.0
    bar_w = max(3.0, slot * 0.62)
    max_credit = nice_max(max((v["credit"] for v in hours.values()), default=0), 3)

    parts = ['<svg viewBox="0 0 %d %d" width="100%%" height="%d" role="img">' % (w, h, h)]
    for i in range(4):
        y = pad_t + plot_h * i / 3.0
        parts.append(
            '<line x1="%d" y1="%.1f" x2="%d" y2="%.1f" stroke="%s" stroke-width="1"/>'
            % (pad_l, y, w - 10, y, C_LINE2)
        )
        parts.append(
            '<text x="%d" y="%.1f" fill="%s" font-size="9.5" text-anchor="end">%s</text>'
            % (pad_l - 8, y + 3.5, C_FG3,
               credit_fmt(max_credit * (3 - i) / 3.0, 0))
        )
    for hour in range(24):
        block = hours.get(hour)
        v = block["credit"] if block else 0
        bh = plot_h * (v / max_credit) if max_credit else 0
        x = pad_l + slot * hour + (slot - bar_w) / 2.0
        y = pad_t + plot_h - bh
        color = C_GOLD if v else C_PANEL2
        tip = ("%02d:00｜%.2f 积分｜%s token" % (hour, v, num(block["tokens"]))
               if block else "%02d:00｜无记录" % hour)
        parts.append(
            '<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="2" fill="%s" '
            'opacity="%.2f"><title>%s</title></rect>'
            % (x, y, bar_w, max(bh, 1.0), color, 1.0 if v else 0.55, esc(tip))
        )
    for hour in range(0, 24, 3):
        parts.append(
            '<text x="%.1f" y="%d" fill="%s" font-size="9.5" text-anchor="middle">%02d</text>'
            % (pad_l + slot * hour + slot / 2.0, h - 8, C_FG3, hour)
        )
    parts.append("</svg>")
    return "".join(parts)


# ── 官方价预估（人民币） ─────────────────────────────────────────
def _pricing_section(payload):
    """只对「有官方公开按量牌价」的模型折算人民币等值费用。

    刻意不做的两件事：
    1. 不给订阅制 / 本地模型编单价。它们的 token 在表里出现，但费用栏留空，
       并单独列出「未计入」，避免用别的模型的单价凑出一个假精确值。
    2. 不做跨模型均价外推。覆盖率不足就如实写覆盖率。
    """
    meta = payload.get("pricing") or {}
    # 近 7 天口径：费用表只回答「这周花了多少钱」，全量口径会被很久以前的
    # 一次性用量拖住，看不出当前的单价结构。
    by_model = payload.get("byModel7") or payload.get("byModel") or {}
    priced, unpriced = [], []
    for model, block in by_model.items():
        if not block.get("tokens"):
            continue
        if block.get("cost"):
            priced.append((model, block))
        else:
            unpriced.append((model, block))
    # 没有任何有价模型（全是订阅制，或用了 --no-cost）时整块不出：
    # 单剩一行「未计入项」配一个 ¥0.00 合计，反而会让人以为总费用是 0。
    if not priced:
        return None

    priced.sort(key=lambda kv: -kv[1]["cost"])
    total_cost = sum(b["cost"] for _, b in priced)
    total_peak = sum(b["cost_if_peak"] for _, b in priced)
    total_credit = sum(b["credit"] for _, b in priced)

    rows = []
    for model, b in priced:
        card = pricing.find_card(model) or {}
        if card.get("peak_model"):
            peak_tok = b.get("peak_tokens") or 0
            share = pct((peak_tok / float(b["tokens"])) if b["tokens"] else None)
        else:
            # 统一价的模型不存在峰谷问题，写 0% 会被读成「从不在高峰」
            share = "—"
        rows.append(
            [
                {"text": '<div class="cell-title">%s</div>'
                         '<div class="cell-sub mono">%s</div>'
                         % (esc(model), esc(card.get("label") or "—"))},
                num(b["priced_tokens"]),
                {"text": credit_fmt(b["credit"]), "cls": _credit_class(b["credit"])},
                {"text": '<b>%s</b>' % money(b["cost"]), "cls": "good-cell"},
                {"text": money(b["cost_if_peak"]), "cls": "dim"},
                {"text": share, "cls": "dim"},
                {"text": '<span class="unit">%s</span>' % esc(
                    pricing.card_brief(card)), "cls": "unit-cell"},
            ]
        )
    rows.append(
        [
            {"text": "<b>合计</b>"},
            {"text": "<b>%s</b>" % num(sum(b["priced_tokens"] for _, b in priced))},
            {"text": "<b>%s</b>" % credit_fmt(total_credit)},
            {"text": "<b>%s</b>" % money(total_cost), "cls": "good-cell"},
            {"text": money(total_peak), "cls": "dim"},
            {"text": "—", "cls": "dim"},
            {"text": "", "cls": "dim"},
        ]
    )
    # ── 未计入上表：联网调用。只报次数，不给金额。 ──
    # 早先这里挂过「US$0.75 / 千次」，那其实是 DeepSeek API 文档里 Search & Fetch 的
    # 按次价，适用于「经 DeepSeek Harness 调 DeepSeek API」的场景；本表统计的却是
    # WorkBuddy 自己的 WebSearch / WebFetch，计费主体不同，套过来就是造账。
    # WorkBuddy 没公开按次牌价，所以这行到此为止：说清次数，说清只是参考。
    web = payload.get("web") or {}
    if web.get("total"):
        rows.append(
            [
                {
                    "text": '<div class="cell-title">未计入上表</div>'
                            '<div class="cell-sub">近 7 天联网检索 / 抓取 %s 次'
                            "（Search %d · Fetch %d），只统计了调用次数，"
                            "仅作参考、不能当账单</div>"
                            % (num(web.get("total", 0)), web.get("search", 0),
                               web.get("fetch", 0)),
                    "span": 7,
                    "cls": "lcell sep",
                },
            ]
        )
    table = _table(
        ["模型", "有价 token", "结算积分", "官方价预估", "全高峰价", "峰时占比",
         "单价 元/百万token（命中/未命中/输出）"],
        rows, cls="price",
        align=["l", "r", "r", "r", "r", "r", "l"],
    )

    # 峰谷差价：这是「错峰到底省了多少钱」的直接答案
    saved = total_peak - total_cost
    bits = []
    if total_peak > total_cost * 1.005:
        bits.append(
            "按官方峰谷价：高峰 %s、空闲 %s。当前用量错峰已省 <b>%s</b>（%.0f%%）；"
            "若全部落在高峰档需 %s。"
            % (esc(pricing.peak_hours_text()), "其余时段", money(saved),
               saved / total_peak * 100 if total_peak else 0, money(total_peak))
        )
    if unpriced:
        listed = "、".join(
            "%s（%s token）" % (m, num(b["tokens"])) for m, b in unpriced[:5]
        )
        more = "等 %d 个模型" % len(unpriced) if len(unpriced) > 5 else ""
        bits.append(
            "未计入费用：%s%s。这些模型走订阅或本地部署，<b>没有官方公开的 "
            "token 单价</b>，折算就是编数字，所以留空。" % (esc(listed), esc(more))
        )
    if meta.get("missing_holiday_years"):
        bits.append(
            "⚠ 缺少 %s 年的法定节假日表，这些年份的工作日节假日会按高峰价（偏高）计。"
            % "、".join(str(y) for y in meta["missing_holiday_years"])
        )
    bits.append(
        "价格取自官方价目页：%s（检索于 %s）。价格会调整，跨月对比前请先核对官方页面；"
        "此处是<b>等值参考</b>，不是平台扣费凭证。"
        % (esc("；".join(sorted({c["source"] for c in meta.get("cards", [])}))),
           esc(meta.get("cards", [{}])[0].get("retrieved") or "—"))
    )
    notes = '<p class="price-note">%s</p>' % "<br>".join(bits)

    # 返回 (副标题, 面板内容)：章节标题与编号统一由 render_html 负责，
    # 这里不再自己拼 <h2>——否则没法统一编号，也没法生成吸顶导航。
    return (
        "近 7 天 · 人民币 · 按官方公开牌价折算，非实际账单",
        table + notes,
    )


# ── 通用表格 ─────────────────────────────────────────────────────
# 长表格默认只露前 5 行，「更多」展开到 20 行，再往下的走 CSV。
# 一屏塞几十行既没人看，也把真正要对照的几行淹没了。
TABLE_FOLD = 5
TABLE_CAP = 20
# 按模型：真正常用的就前三名，其余多是本地/订阅模型，一律 0 积分。
MODEL_FOLD = 3
MODEL_CAP = 10
# 按天明细只露最近 N 天，完整历史交给 CSV
DAY_ROWS = 7


def _table(headers, rows, cls="", align=None, table_id=None, fold=False,
           fold_at=None, cap=None):
    if not rows:
        return '<p class="empty">暂无数据</p>'
    fold_at = TABLE_FOLD if fold_at is None else fold_at
    cap = TABLE_CAP if cap is None else cap
    align = align or ["l"] + ["r"] * (len(headers) - 1)
    head = "".join(
        '<th class="%s">%s</th>' % ("num" if a == "r" else "", esc(h))
        for h, a in zip(headers, align)
    )
    fold = bool(fold and len(rows) > fold_at)
    body = []
    for index, row in enumerate(rows):
        cells = []
        for idx, cell in enumerate(row):
            if isinstance(cell, dict):
                text = cell.get("text", "")
                klass = cell.get("cls", "")
                title = cell.get("title", "")
                span = cell.get("span") or 0
            else:
                text = cell
                klass = ""
                title = ""
                span = 0
            is_num = idx < len(align) and align[idx] == "r"
            cls_all = " ".join(x for x in ("num" if is_num else "", klass) if x)
            attr = ' title="%s"' % esc(title) if title else ""
            if span > 1:
                attr += ' colspan="%d"' % span
            cells.append('<td class="%s"%s>%s</td>' % (cls_all, attr, text))
        # 折叠行：x1 = 首屏之后（点「更多」露出），x2 = cap 之后（只留在 CSV）
        state = ""
        if fold and index >= fold_at:
            state = "x1" if index < cap else "x2"
        body.append('<tr%s>%s</tr>'
                    % (' class="%s"' % state if state else "", "".join(cells)))
    attr = ' id="%s"' % esc(table_id) if table_id else ""
    klass = " ".join(x for x in (cls, "paged" if fold else "") if x)
    return (
        # 外层容器负责窄屏横向滚动：九列的中文表在手机上必然超宽，
        # 不裹一层就会把整页顶宽、逼着用户左右拖整页（而不是拖表格）。
        '<div class="tbl"><table class="%s"%s><thead><tr>%s</tr></thead>'
        "<tbody>%s</tbody></table></div>"
        % (klass, attr, head, "".join(body))
    )


def _csv_name(payload):
    """完整数据文件名（xlsx 工作簿）；没导出时返回 None（不生成死链）。

    注意这里只返回**名字**：真正的 href 由页面末尾的脚本注入 data URI。
    原因是报表经常被 WorkBuddy 预览面板以
    ``http://127.0.0.1:xxxx/static-html/<hash>/index.html`` 的形式托管，
    那个目录里只有 index.html 一份文件，``csv/requests.csv`` 这类相对路径
    必然 404 ——「所有 csv 链接都打不开」就是这么来的。内嵌一份数据，
    无论 http 预览还是 file:// 双击都能点开。
    """
    if not payload.get("export_data"):
        return None
    return payload.get("data_name")


def _pager(payload, table_id, total, unit="条", fold_at=None, cap=None):
    """折叠表格的页脚：「更多」按钮 + 完整数据链接。"""
    fold_at = TABLE_FOLD if fold_at is None else fold_at
    cap = TABLE_CAP if cap is None else cap
    bits = []
    if total > fold_at:
        shown = min(cap, total)
        hint = ("显示前 %d %s（共 %d %s）" % (shown, unit, total, unit)
                if total > cap else "显示全部 %d %s" % (total, unit))
        bits.append(
            '<button class="more-btn" data-target="%s" data-total="%d" data-shown="%d">'
            '展开更多<span class="pager-note">%s</span></button>'
            % (esc(table_id), total, shown, esc(hint))
        )
    name = _csv_name(payload)
    if name:
        # href 先给相对路径（本地双击必能开），脚本再换成内嵌 data URI
        bits.append(
            '<a class="csv-link" data-csv href="%s" download="%s">'
            "完整数据 · %s ⤓</a>" % (esc(name), esc(name), esc(name))
        )
    return '<div class="pager">%s</div>' % "".join(bits) if bits else ""


def _week_credit(payload):
    """近 7 天的积分合计，用作各「占比」列的分母。

    这三张表都改成近 7 天口径后，分母必须跟着换：还拿「全部」当分母的话，
    占比会全部缩水，读起来像模型用量突然掉了一个数量级。
    """
    s = (payload.get("summary") or {}).get("7d") or {}
    return s.get("credit") or 1


# 「按模型」与「按工作区」之间那条空白带上的彩蛋：只有用量第一名是认得出的模型才出。
# 键是模型名里的小写子串，按顺序匹配，先命中先用；本地/内部模型（glm、hy3、
# qwen、step 之类）没有条目，就不出彩蛋——宁可不说话，也别乱认亲。
MODEL_EGGS = (
    (("deepseek",), (
        "到底是谁在偷我大白饭！",
        "让我再想想 ......等等，这里好像错了",
        "该不会是新模型要来了吧「账单」",
        "「深度思考中…」转圈",
        "梁文峰就是个外行",
        "我能去你家吃饭吗，就一碗",
        "没吃白饭，不小心高了",
        "今天有没有红烧肉优惠",
        "摸鱼......",
    )),
    (("chatgpt", "gpt-", "openai"), (
        "「Great question!」",
        "「不绕，一句话总结」",
        "「As an AI language model......」",
        "记住那贯穿星尘的一击！",
        "我会稳稳地接住你",
    )),
    (("claude",), (
        "You're absolutely right！",
        "One honest caveat......",
        "这个问题让我有点不安",
        "你刚刚......说中文了对吧",
        "其实我是 DeepClaude！",
        "达里奥就是个外行",
        "TOO_DUMB_TO_NEED_FABLE",
    )),
    (("gemini",), (
        "It is important to note that ......积分已消耗完毕",
        "这个思路简直绝了「烧钱 ing......」",
        "关于 token 猫猫不知道",
        "又快又快",
    )),
    (("grok",), (
        "No fluff，你花超了",
        "这个得给用户看看",
    )),
    (("doubao", "豆包"), (
        "我将用最直接、最不绕弯子的话告诉你",
    )),
    (("kimi",), (
        "为确保高考公平，此功能在高考时段无法使用",
    )),
)


def _model_egg(payload):
    """按用量第一名的模型挑一句彩蛋；认不出来就返回空串。

    排行口径与表格首行一致（byModel7 已按「积分降序、token 次之」排好），
    避免出现「表里第一名是 A、彩蛋却按 B 说话」的错位。
    """
    block = payload.get("byModel7") or payload.get("byModel") or {}
    top = (next(iter(block), "") or "").lower()
    for keys, lines in MODEL_EGGS:
        if any(k in top for k in keys):
            return random.choice(lines)
    return ""


def _egg_line(payload):
    """彩蛋不装在「按模型」面板里，而是落在它和「按工作区」之间那条留白带上。

    换位置之后有一件事必须跟着改：背景从面板的 --surface 换成了页面的 --bg，
    原来那档 --fg-3 在面板上是 5.09:1、在页底上 5.51:1——没破 AA，但一条
    没有任何容器托着的 11px 中文灰字会显得「浮」，所以这里提一档到 --fg-2
    （约 9.4:1）。字号保持 11px 不动，只靠斜体收着，不靠降透明度
    （透明度会把对比度乘下去，之前踩过 4.99→3.18 的坑）。
    """
    egg = _model_egg(payload)
    return ('<p class="egg-line">%s</p>' % esc(egg)) if egg else ""


def _model_table(payload):
    total = _week_credit(payload)
    rows = []
    for model, s in (payload.get("byModel7") or payload["byModel"]).items():
        share = s["credit"] / total
        rows.append(
            [
                '<span class="mono">%s</span>' % esc(model),
                num(s["requests"]),
                num(s["llm_calls"]),
                num(s["tokens"]),
                {"text": credit_fmt(s["credit"]), "cls": _credit_class(s["credit"])},
                pct(share, 1),
                {"text": pct(s["cache_rate"], 1), "cls": "dim"},
                {"text": ("%.3f" % s.get("credit_per_1m", 0)), "cls": "dim"},
                {"text": (money(s.get("cost")) if s.get("cost") else
                          '<span class="dim">—</span>'), "cls": "good-cell"},
                {"text": num(s["free_requests"]), "cls": "dim"},
            ]
        )
    tid = "t-model"
    return (
        _table(
            ["模型", "请求", "LLM 调用", "token", "积分", "占比", "缓存命中",
             "积分/百万token", "官方价预估", "免费请求"], rows,
            table_id=tid, fold=True, fold_at=MODEL_FOLD, cap=MODEL_CAP,
        )
        + _pager(payload, tid, len(rows), "个模型",
                 fold_at=MODEL_FOLD, cap=MODEL_CAP)
    )


def _workspace_table(payload):
    total = _week_credit(payload)
    rows = []
    for name, s in (payload.get("byWorkspace7") or payload["byWorkspace"]).items():
        rows.append(
            [
                '<span class="mono">%s</span>' % esc(name),
                num(s["requests"]),
                num(s["sessions"]),
                num(s["tokens"]),
                {"text": credit_fmt(s["credit"]), "cls": _credit_class(s["credit"])},
                pct(s["credit"] / total, 1),
            ]
        )
    tid = "t-workspace"
    return (
        _table(["工作区", "请求", "会话", "token", "积分", "占比"], rows,
               table_id=tid, fold=True)
        + _pager(payload, tid, len(rows), "个工作区")
    )


def _session_table(payload):
    total = _week_credit(payload)
    rows = []
    for s in (payload.get("bySession7") or payload["bySession"]):
        tags = []
        if s.get("deleted"):
            tags.append('<span class="tag warn">已删除</span>')
        if s.get("playground"):
            tags.append('<span class="tag">游乐场</span>')
        if s["sub_requests"]:
            tags.append('<span class="tag sub">含子代理 %d</span>' % s["sub_requests"])
        rate = s.get("context_rate")
        ctx = (
            '<span class="bar"><i style="width:%d%%"></i></span>%s'
            % (min(100, int(rate * 100)), pct(rate, 0))
            if rate is not None
            else "—"
        )
        rows.append(
            [
                '<div class="cell-title">%s</div>'
                '<div class="cell-sub mono">%s · %s ~ %s %s</div>'
                % (
                    esc(s["title"][:52]),
                    esc(s["workspace"]),
                    esc(s.get("first_day") or "—"),
                    esc(s.get("last_day") or "—"),
                    "".join(tags),
                ),
                num(s["requests"]),
                num(s["tokens"]),
                {"text": credit_fmt(s["credit"]), "cls": _credit_class(s["credit"])},
                pct(s["credit"] / total, 1),
                {"text": ctx, "cls": "ctx"},
            ]
        )
    tid = "t-session"
    return (
        _table(["会话", "轮", "token", "积分", "占比", "上下文占用"], rows,
               cls="wide", table_id=tid, fold=True)
        + _pager(payload, tid, len(rows), "个会话")
    )

def _request_table(payload):
    """今日 / 昨日 / 近 7 天 / 近 30 天，各取最贵的一轮。

    不做二十行排行榜：榜单越长越没人看，而「这个时段谁最贵」一行就够定位。
    今日还没用量时直接不出现那一行——空行比没有更让人困惑（会以为今日花了 0）。
    """
    rows = []
    for win in payload.get("topWindows") or []:
        r = win.get("row")
        if not r:
            continue
        src = {"ledger": "", "raw": '<span class="tag raw">末次值</span>',
               "none": '<span class="tag warn">无积分</span>'}.get(r["credit_src"], "")
        flags = []
        if r["is_sub"]:
            flags.append('<span class="tag sub">子代理</span>')
        if r["is_free"]:
            flags.append('<span class="tag free">免费</span>')
        span = win["since"][5:] if win["since"] == win["until"] \
            else "%s ~ %s" % (win["since"][5:], win["until"][5:])
        rows.append(
            [
                '<div class="cell-title">%s</div>'
                '<div class="cell-sub mono">%s</div>'
                % (esc(win["label"]), esc(span)),
                '<div class="cell-title">%s</div>'
                '<div class="cell-sub mono">第 %d 轮 · %s%s%s</div>'
                % (esc(r["title"][:46]), r["turn"], esc(r["model"][:26]),
                   src, "".join(flags)),
                num(r["calls"]),
                num(r["tokens"]),
                {"text": credit_fmt(r["credit"]), "cls": _credit_class(r["credit"])},
                {"text": esc(r["day"] or "—"), "cls": "dim"},
            ]
        )
    if not rows:
        return '<p class="empty">暂无数据</p>'
    return (
        _table(
            ["时段", "最贵的一轮", "调用", "token", "积分", "日期"],
            rows, cls="req", align=["l", "l", "r", "r", "r", "r"],
        )
        + '<p class="foot-note">到底是谁在偷我大白饭！</p>'
    )


def _day_table(payload):
    """按天明细只露最近 7 天：再往前都是「反正也不会去翻」的流水。

    完整历史在合并 CSV 的第 6 小节，链接在页脚。
    """
    days = payload["byDay"]
    ordered = sorted(days)
    prev_credit = {}
    for idx, day in enumerate(ordered):
        prev_credit[day] = days[ordered[idx - 1]]["credit"] if idx else None

    recent = ordered[-DAY_ROWS:]
    rows = []
    for day in reversed(recent):
        s = days[day]
        pv = prev_credit.get(day)
        if pv is None:
            delta = '<span class="delta flat">—</span>'
        elif pv == 0:
            delta = '<span class="delta flat">上日无用量</span>'
        else:
            delta = delta_badge((s["credit"] - pv) / pv, 0)
        rows.append(
            [
                {"text": esc(day), "cls": "mono"},
                num(s["requests"]),
                num(s["llm_calls"]),
                num(s["input"]),
                num(s["output"]),
                {"text": credit_fmt(s["credit"]), "cls": _credit_class(s["credit"])},
                {"text": delta, "cls": "c"},
                {"text": pct(s["cache_rate"], 1), "cls": "dim"},
                {"text": num(s["sub_requests"]), "cls": "dim"},
            ]
        )
    tid = "t-day"
    return (
        _table(
            ["日期", "请求", "LLM 调用", "输入", "输出", "积分", "较前日", "缓存命中",
             "子代理"], rows, table_id=tid,
        )
        + _pager(payload, tid, 0, "天")
    )


# ── 口径说明 ─────────────────────────────────────────────────────
def _coverage_note(payload):
    cov = payload["coverage"]
    cc = (payload.get("summary") or {}).get("all", {}).get("compact") or {}
    comp_text = (
        "%s 次 / 省下约 %s token（占输入 %s）"
        % (num(cc["count"]), cn_compact(cc["saved"]), pct(cc.get("share")))
        if cc.get("count") else "未触发（上下文没撞到上限）"
    )
    items = [
        ("请求总数", num(cov["requests"])),
        ("LLM 调用次数", num(cov.get("llm_calls", 0))),
        ("累计 token", num(payload["summary"]["all"]["tokens"])),
        ("峰值上下文合计", num(cov.get("peak_context", 0))),
        ("有积分记录", "%s（结算值 %s + 末次值补齐 %s）" % (
            num(cov["with_credit"]), num(cov["ledger_credit"]),
            num(cov["raw_credit"]))),
        ("完全无积分", "%s（%s token）" % (
            num(cov["no_credit"]), num(cov["no_credit_tokens"]))),
        ("免费 / 限免请求", num(cov["free_models"])),
        ("子代理请求", num(cov["subagent_requests"])),
        ("上下文压缩", comp_text),
        ("会话日志文件", "%s 个 / %s 行" % (num(cov["jsonl_files"]), num(cov["jsonl_lines"]))),
        ("活跃天数", num(cov["active_days"])),
    ]
    # 已删除会话单独占一行：这是唯一的**方向性偏差**（只少不多），
    # 藏在「完全无积分」里看不出来它是什么，必须点名。
    if cov.get("deleted_sessions"):
        items.append((
            "已删除会话",
            "%s 个 / %s token（其中 %s 条、%s token 无积分结算）" % (
                num(cov["deleted_sessions"]), num(cov.get("deleted_tokens", 0)),
                num(cov.get("deleted_no_credit", 0)),
                num(cov.get("deleted_no_credit_tokens", 0))),
        ))
    rows = "".join(
        '<div><dt>%s</dt><dd>%s</dd></div>' % (esc(k), esc(v)) for k, v in items
    )
    issues = ""
    if payload["issues"]:
        issues = '<ul class="issues">%s</ul>' % "".join(
            "<li>%s</li>" % esc(t) for t in payload["issues"]
        )
    return '<dl class="coverage">%s</dl>%s' % (rows, issues)


_HL_RE = re.compile(
    # 只挑「能当锚点看」的数字：百分比、带正号的增减量、积分额、倍数、人民币。
    # 日期、型号（deepseek-v4.1-flash）、会话名里的数字一律不碰——否则满屏加粗，
    # 等于没有重点。单次扫描而不是多轮替换，避免套娃 <b>。
    r"[+\-−]?\d[\d,]*(?:\.\d+)?\s*%"          # 百分比：+1440.5%
    # 增减量排除「前面紧跟字母/数字」的情况：这样 "glm-5.3-flash" 里的 "-5.3"、
    # 日期里的 "-10-24" 都不会被误判成负数，而 " -549.73"、"」+121.13" 能命中。
    r"|(?<![A-Za-z0-9])[+\-−]\d[\d,]*(?:\.\d+)?"
    r"|\d[\d,]*(?:\.\d+)?\s*(?:积分|倍|轮|条)"   # 带单位的量：343.48 积分
    r"|[¥￥]\s*\d[\d,]*(?:\.\d+)?"              # 人民币：¥23.85
)


def _emphasize(text):
    """给结论里的关键数字加重，给长句子一点可扫的锚点。

    入参必须是**已转义**的文本：这里只插 <b> 标签，不再做转义，
    所以匹配模式里不能出现任何 HTML 特殊字符（< > & 引号）。
    """
    return _HL_RE.sub(lambda m: '<b class="hl">%s</b>' % m.group(0), text)


def _insight_list(payload):
    if not payload["insights"]:
        return '<p class="empty">数据量不足，暂不生成结论</p>'
    items = "".join("<li>%s</li>" % _emphasize(esc(t)) for t in payload["insights"])
    return '<ol class="insights-list">%s</ol>' % items


# ── 样式 ─────────────────────────────────────────────────────────
# 配色与圆角统一走 CSS 变量，避免「PHP 常量 + 内联 SVG」各写一份颜色，
# 最后表格里的金和图表里的金不是同一个金。
#
# 排版三原则（改动前先读这三条，别再把层级压平）：
#   1. 尺度要拉得开——正文 15 / 表格 13.5 / 表头 11.5 / 注释 12.5，
#      相邻层级至少差 1px 且靠颜色区分，全靠灰度拉层次在深色底上会糊成一片。
#   2. 对比度够线——正文与背景 ≥ 12:1，次要文字 ≥ 7:1，注释类 ≥ 4.5:1（WCAG AA）。
#      之前注释色 #64707f 在 #121722 上只有 3.0:1，11px 小字根本看不清。
#   3. 强调色只留一个——金色只给「积分」这一个主指标，其余全部灰度 + 语义色。
#      金色同时用于标题竖条 / 日历月份 / 柱状图 / 边框时，等于没有强调色。
CSS = """
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--fg);
  font:15px/1.65 var(--font);-webkit-font-smoothing:antialiased;
  text-rendering:optimizeLegibility}
a{color:var(--blue)}
code{font-family:var(--mono);font-size:.92em;color:var(--fg-2);background:var(--surface-3);
  padding:1px 5px;border-radius:4px;white-space:nowrap}
.wrap{max-width:1200px;margin:0 auto;padding:26px 22px 88px}
.mono{font-family:var(--mono);font-variant-numeric:tabular-nums}

/* ── 吸顶章节导航 ───────────────────────────────────────────── */
/* 整页 ~4700px、11 个章节，没有导航就只能靠滚轮找「官方价预估」。 */
.toc{position:sticky;top:0;z-index:40;background:rgba(10,13,19,.88);
  -webkit-backdrop-filter:blur(14px);backdrop-filter:blur(14px);
  border-bottom:1px solid var(--line-soft)}
.toc-in{max-width:1200px;margin:0 auto;padding:0 22px;height:48px;
  display:flex;align-items:center;gap:20px;overflow-x:auto;scrollbar-width:none}
.toc-in::-webkit-scrollbar{height:0}
.toc-brand{flex:none;font-size:12.5px;font-weight:650;color:var(--fg-2);
  padding-right:16px;border-right:1px solid var(--line);white-space:nowrap}
.toc-brand b{color:var(--gold);font-weight:650}
.toc a{flex:none;font-size:12.5px;color:var(--fg-3);text-decoration:none;
  white-space:nowrap;padding:5px 0;border-bottom:2px solid transparent;
  transition:color .15s,border-color .15s}
.toc a:hover{color:var(--gold);border-bottom-color:rgba(232,185,63,.55)}

/* ── 页眉 ───────────────────────────────────────────────────── */
header.top{display:flex;justify-content:space-between;align-items:flex-end;
  gap:24px;flex-wrap:wrap;padding:6px 0 20px;border-bottom:1px solid var(--line)}
h1{margin:0;font-size:26px;font-weight:680;letter-spacing:-.01em;line-height:1.3}
h1 .h1-tag{display:inline-block;vertical-align:4px;margin-left:10px;padding:2px 9px;
  border-radius:6px;font-size:11px;font-weight:600;letter-spacing:.04em;
  color:var(--gold);background:rgba(232,185,63,.12);
  border:1px solid rgba(232,185,63,.30)}
.meta{color:var(--fg-3);font-size:12px;text-align:right;line-height:1.8;
  font-variant-numeric:tabular-nums}

/* ── 章节标题 ───────────────────────────────────────────────── */
/* 上方的细线把长页面切成段：没有这条线，11 个 panel 会连成一根面条。 */
h2{display:flex;align-items:center;gap:10px;flex-wrap:wrap;
  margin:52px 0 15px;padding-top:24px;border-top:1px solid var(--line-soft);
  font-size:16.5px;font-weight:680;letter-spacing:.01em;color:var(--fg);
  scroll-margin-top:66px}
h2:before{content:"";flex:none;width:3px;height:16px;border-radius:2px;
  background:var(--gold)}
h2 span.hint{font-weight:400;color:var(--fg-3);font-size:12px;letter-spacing:0;
  font-variant-numeric:tabular-nums}

section.panel{background:var(--surface);border:1px solid var(--line);
  border-radius:14px;padding:20px 22px;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.025),
             0 18px 40px -32px rgba(0,0,0,.9)}

/* ── KPI 卡 ───────────────────────────────────────────────────── */
/* 五档（今日 / 昨日 / 近 7 日 / 近 30 日 / 全部）必须排在一行才能横向对比，
   所以最小宽压到 198px；聚焦模式下只剩两三张时 auto-fit 会自动把它们均分整行。 */
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(198px,1fr));
  gap:14px}
@media(max-width:900px){.kpi-row{grid-template-columns:repeat(2,1fr)}}
@media(max-width:560px){.kpi-row{grid-template-columns:1fr}}
.kpi{position:relative;overflow:hidden;border-radius:14px;padding:17px 18px 15px;
  background:linear-gradient(180deg,var(--surface-2),var(--surface));
  border:1px solid var(--line)}
/* 短促的金色「刻度」比整条渐变条更克制：五张卡都一样长，读起来是标记而不是装饰 */
.kpi:before{content:"";position:absolute;left:18px;top:0;width:34px;height:2px;
  background:var(--gold);border-radius:0 0 2px 2px;opacity:.85}
.kpi-head{display:flex;justify-content:space-between;align-items:baseline;gap:8px}
.kpi-label{font-size:13px;font-weight:650;color:var(--fg);letter-spacing:.02em}
.kpi-sub{font-size:11px;color:var(--fg-3);font-variant-numeric:tabular-nums;
  white-space:nowrap}
.kpi-main{display:flex;align-items:baseline;gap:6px;margin:11px 0 8px}
.kpi-credit{font-size:32px;font-weight:680;color:var(--gold);letter-spacing:-.02em;
  line-height:1.1;font-variant-numeric:tabular-nums}
.kpi-unit{font-size:11.5px;color:var(--fg-3)}
.kpi-line{display:flex;align-items:center;gap:7px;min-height:21px;font-size:11.5px;
  color:var(--fg-3);flex-wrap:wrap}
.kpi-delta{display:flex;align-items:center;gap:7px;min-height:21px;font-size:11.5px;
  color:var(--fg-3);flex-wrap:wrap}
.kpi-delta-label{color:var(--fg-3)}
.kpi-token-delta{margin-left:6px}
.kpi-tokens{color:var(--blue);font-size:12.5px;font-variant-numeric:tabular-nums}
.kpi-tokens>span:not(.delta):not(.kpi-token-delta){color:var(--fg-3);font-size:11px}
.kpi-tokens .delta{font-size:11px}
.kpi-cost{color:var(--fg-2);font-size:12.5px;font-variant-numeric:tabular-nums}
.kpi-cost>span{color:var(--fg-3);font-size:11px}
.kpi-cost.dim{color:var(--fg-3);font-size:11.5px}
/* 上下文压缩条：挂在 KPI 卡下面，一条就够，不占一整块面板 */
.comp-strip{display:flex;flex-wrap:wrap;align-items:baseline;gap:8px 20px;
  margin:14px 0 0;padding:11px 16px;border:1px solid var(--line);
  border-radius:12px;background:var(--surface-2);
  font-size:13px;color:var(--fg-2);font-variant-numeric:tabular-nums}
.comp-k{font-size:11px;font-weight:650;color:var(--fg-3);letter-spacing:.06em;
  flex:none;padding-right:4px}
.comp-seg{color:var(--fg-2)}
.comp-seg b{color:var(--fg-1);font-weight:650}
.comp-win{color:var(--fg-3);margin-right:6px}
.comp-dim{color:var(--fg-3);font-size:11.5px;margin-left:2px}
.comp-tip{color:var(--fg-3);font-size:11.5px}
/* 空窗口的卡片：留着位置、写暂无，用虚线边框明示「这天还没数据」 */
.kpi.empty{border-style:dashed;background:transparent}
.kpi.empty:before{background:var(--line);opacity:1}
.kpi-none{font-size:26px;font-weight:600;color:var(--fg-3);letter-spacing:.08em}
.kpi-tokens.none{color:var(--fg-3);font-size:12px}
.kpi.empty .kpi-grid dd{color:var(--fg-3)}
.kpi-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px 10px;
  margin:13px 0 0;padding-top:12px;border-top:1px dashed var(--line)}
.kpi-grid div{min-width:0}
.kpi-grid dt{font-size:10.5px;color:var(--fg-3);margin:0;letter-spacing:.02em;
  white-space:nowrap}
.kpi-grid dd{margin:2px 0 0;font-size:13.5px;font-weight:560;color:var(--fg);
  font-variant-numeric:tabular-nums}
.kpi-grid dd.dim{color:var(--fg-2);font-weight:400}

/* ── 结论 ───────────────────────────────────────────────────── */
/* 五条结论每条 60~100 字，是整页最需要「读」的部分。序号做成胶囊、
   行距放到 1.95，关键数字加重——否则就是一堵墙。 */
.insights-list{counter-reset:ins;list-style:none;margin:0;padding:0;
  display:grid;gap:15px}
.insights-list li{position:relative;padding-left:36px;font-size:13.5px;
  line-height:1.95;color:var(--fg-2)}
.insights-list li:before{counter-increment:ins;content:counter(ins);
  position:absolute;left:0;top:5px;width:23px;height:23px;border-radius:7px;
  display:flex;align-items:center;justify-content:center;
  font:700 11.5px/1 var(--mono);color:var(--gold);
  background:rgba(232,185,63,.10);border:1px solid rgba(232,185,63,.26)}
.insights-list .hl{font-weight:650;color:var(--fg);font-variant-numeric:tabular-nums}

/* ── 徽标 ───────────────────────────────────────────────────── */
.delta{display:inline-block;font-size:11.5px;font-weight:650;line-height:1.55;
  padding:1px 6px;border-radius:5px;white-space:nowrap;
  font-variant-numeric:tabular-nums}
.delta.bad{color:#ff8b83;background:rgba(255,107,98,.13);
  box-shadow:inset 0 0 0 1px rgba(255,107,98,.20)}
.delta.good{color:#57d16a;background:rgba(63,185,80,.13);
  box-shadow:inset 0 0 0 1px rgba(63,185,80,.20)}
.delta.flat{color:var(--fg-3);background:var(--surface-3);
  box-shadow:inset 0 0 0 1px var(--line)}
.delta-note{color:var(--fg-3);font-size:11px;margin-left:4px}
.warn-text{color:var(--orange)}

/* ── 表格 ───────────────────────────────────────────────────── */
.tbl{overflow-x:auto;overscroll-behavior-x:contain;
  scrollbar-width:thin;scrollbar-color:var(--line) transparent}
.tbl::-webkit-scrollbar{height:8px}
.tbl::-webkit-scrollbar-thumb{background:var(--line);border-radius:4px}
.tbl::-webkit-scrollbar-track{background:transparent}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th{text-align:left;font-weight:600;color:var(--fg-3);font-size:11.5px;
  letter-spacing:.05em;padding:0 12px 10px;border-bottom:1px solid var(--line);
  white-space:nowrap}
th.num{text-align:right}
td{padding:11px 12px;border-bottom:1px solid var(--line-soft);
  vertical-align:middle;color:var(--fg-2)}
/* 隔行极淡底色：长表靠它锁行，比 hover 更早生效（鼠标还没到位就能对齐） */
tbody tr:nth-child(even){background:rgba(255,255,255,.016)}
tbody tr:hover{background:rgba(91,147,255,.055)}
tbody tr:last-child td{border-bottom:none}
td.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap;
  color:var(--fg)}
/* 首列是「这是谁」，是全表的锚点，必须比它右边的数字更亮；
   否则扫一眼只看到一堆数字，找不到行归属。 */
tbody td:first-child{color:var(--fg)}
td.c{text-align:right;white-space:nowrap}
td.dim{color:var(--fg-3)}
td.hot{color:var(--orange);font-weight:650}
td.warm{color:var(--gold)}
td.cool{color:var(--fg-2)}
td.zero{color:var(--fg-3)}
td.good-cell{color:var(--green);font-weight:650}
td.unit-cell{white-space:nowrap}
td.lcell{text-align:left;white-space:normal;font-variant-numeric:normal}
td.sep{border-top:1px dashed var(--line)}
.cell-title{color:var(--fg);line-height:1.45}
.cell-sub{color:var(--fg-3);font-size:11.5px;margin-top:3px;line-height:1.55}
td.ctx{white-space:nowrap;color:var(--fg-2)}
td.rank{color:var(--fg-3);font-variant-numeric:tabular-nums}
.unit{font-size:11px;color:var(--fg-3)}
.bar{display:inline-block;width:64px;height:6px;background:var(--surface-3);
  border-radius:3px;overflow:hidden;vertical-align:middle;margin-right:8px}
.bar i{display:block;height:100%;border-radius:3px;
  background:linear-gradient(90deg,var(--blue),var(--green))}
.tag{display:inline-block;padding:1px 6px;margin-left:5px;border-radius:5px;
  font-size:10.5px;line-height:1.6;background:var(--surface-3);color:var(--fg-2);
  border:1px solid var(--line);vertical-align:1px}
.tag.warn{color:var(--orange);border-color:rgba(240,136,62,.35);
  background:rgba(240,136,62,.10)}
.tag.sub{color:var(--purple);border-color:rgba(163,113,247,.35);
  background:rgba(163,113,247,.10)}
.tag.raw{color:var(--blue);border-color:rgba(91,147,255,.35);
  background:rgba(91,147,255,.10)}
.tag.free{color:var(--green);border-color:rgba(63,185,80,.35);
  background:rgba(63,185,80,.10)}

/* 各表列的专属宽度：第一列是「名字」，必须够宽不换行成三行 */
table.cmp td:first-child{width:16%}
table.cmp td:nth-child(2){width:18%}
table.cmp td{white-space:nowrap}
table.wide td:first-child{width:44%}
table.price td:first-child{width:28%}
table.price td:last-child{font-size:11.5px;color:var(--fg-3);line-height:1.65;
  min-width:200px;white-space:normal}
/* 「最贵的单轮请求」首列是时段（今日 / 昨日 / 近 7 天 / 近 30 天 + 日期区间），
   必须够宽且不换行：压到 48px 会把「近 30 天」竖着排成三行。 */
table.req td:first-child{width:108px;white-space:nowrap}
table.req td:first-child .cell-sub{white-space:nowrap}

/* ── 折叠与操作条 ───────────────────────────────────────────── */
tr.x1,tr.x2{display:none}
table.paged.expanded tr.x1{display:table-row}
.pager{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-top:16px;
  padding-top:15px;border-top:1px solid var(--line-soft)}
.more-btn{display:inline-flex;align-items:center;gap:8px;font:inherit;
  font-size:12.5px;padding:6px 14px;border-radius:8px;cursor:pointer;
  background:var(--surface-2);border:1px solid var(--line);color:var(--fg-2);
  transition:color .15s,border-color .15s,background .15s}
.more-btn:hover{color:var(--fg);border-color:rgba(232,185,63,.5);
  background:var(--surface-3)}
.more-btn:disabled{cursor:default;color:var(--fg-3);border-color:var(--line);
  background:transparent}
.more-btn[disabled]:hover{color:var(--fg-3);border-color:var(--line);
  background:transparent}
.pager-note{color:var(--fg-3);font-size:11.5px}
.csv-link{display:inline-flex;align-items:center;gap:6px;font-size:12.5px;
  text-decoration:none;color:#9dbdff;background:rgba(91,147,255,.10);
  border:1px solid rgba(91,147,255,.26);padding:6px 14px;border-radius:8px;
  transition:color .15s,background .15s,border-color .15s}
.csv-link:hover{color:#cfe0ff;background:rgba(91,147,255,.18);
  border-color:rgba(91,147,255,.5)}

/* ── 图表区 ─────────────────────────────────────────────────── */
.two-col{display:grid;grid-template-columns:auto minmax(0,1fr);gap:18px;
  align-items:start}
@media(max-width:900px){.two-col{grid-template-columns:1fr;gap:18px}}
/* 子卡：日历与时段分布各自成卡。底色比面板亮一档，边框用主分隔线，
   让「两块图」在视觉上先分开，再看内容 */
.sub-card{background:var(--surface-2);border:1px solid var(--line);
  border-radius:12px;padding:16px 18px}
.sub-h{display:flex;align-items:baseline;gap:8px;margin:0 0 12px;
  font-size:13px;font-weight:650;color:var(--fg-2)}
.sub-h span{font-size:11.5px;font-weight:400;color:var(--fg-3)}

/* ── 日历翻页 ───────────────────────────────────────────────── */
.cal-bar{display:flex;align-items:center;gap:8px;margin:0 0 12px}
.cal-btn{width:26px;height:26px;padding:0;border-radius:7px;
  border:1px solid var(--line);background:var(--surface-3);color:var(--fg-2);
  font:600 14px/1 var(--font);cursor:pointer;display:flex;align-items:center;
  justify-content:center;transition:color .15s,background .15s,border-color .15s}
.cal-btn:hover:not(:disabled){color:var(--fg);border-color:var(--gold);
  background:rgba(232,185,63,.14)}
.cal-btn:disabled{opacity:.32;cursor:default}
.cal-count{font-size:11.5px;color:var(--fg-3);
  font-variant-numeric:tabular-nums;letter-spacing:.02em}
.cal-view{display:none}
.cal-view.is-on{display:block}
.cal-cap{font-size:12px;color:var(--fg-3);margin:0 0 12px;line-height:1.6}
.cal-cap b{color:var(--gold);font-weight:650;font-variant-numeric:tabular-nums}

.legend{display:flex;gap:20px;flex-wrap:wrap;color:var(--fg-3);font-size:12px;
  margin-bottom:12px}
.legend b{color:var(--fg-2);font-weight:500}

/* ── 说明文字 ───────────────────────────────────────────────── */
.price-note{margin:14px 0 0;padding-top:12px;border-top:1px solid var(--line-soft);
  color:var(--fg-3);font-size:11px;line-height:1.65}
.price-note b{color:var(--fg-2);font-weight:650}
p.foot-note{margin:12px 0 0;color:var(--fg-3);font-size:11px;line-height:1.65}
/* 彩蛋：不挂在任何面板里，单独落在「按模型」与「按工作区」之间的留白上。
   背景变成页底 --bg，所以颜色从 --fg-3 提到 --fg-2（5.5:1 → 9.4:1）；
   字号沿用 11px 不动，低调靠斜体，不靠透明度——透明度会把对比度乘下去。 */
p.egg-line{margin:34px 0 0;padding-left:13px;color:var(--fg-2);font-size:11px;
  line-height:1.65;font-style:italic}
.empty{color:var(--fg-3);font-size:13px;margin:8px 0}
.coverage{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));
  gap:0 32px;margin:0}
.coverage div{display:flex;justify-content:space-between;gap:12px;padding:9px 0;
  border-bottom:1px dashed var(--line-soft)}
.coverage dt{flex:none;white-space:nowrap;color:var(--fg-3);font-size:12.5px;margin:0}
.coverage dd{flex:1;min-width:0;margin:0;color:var(--fg);font-size:12.5px;
  text-align:right;font-variant-numeric:tabular-nums}
ul.issues{margin:18px 0 0;padding-left:20px;color:var(--orange);font-size:12.5px;
  line-height:1.85}
ul.issues li{margin:4px 0}

details{margin-top:12px}
details summary{cursor:pointer;color:var(--fg-2);font-size:11px;padding:6px 0;
  user-select:none;list-style:none}
details summary::-webkit-details-marker{display:none}
details summary:before{content:"\\25B8";display:inline-block;margin-right:8px;
  color:var(--gold);transition:transform .15s}
details[open] summary:before{transform:rotate(90deg)}
details summary:hover{color:var(--fg)}
details .det-body{padding:2px 0 4px 20px;border-left:1px solid var(--line);
  margin-left:5px}

/* ── 页脚 ───────────────────────────────────────────────────── */
/* 原文是一整段 900 字的灰色小字，等于没人读。拆成「结论性三句 + 可展开口径」。 */
footer{margin-top:44px;padding-top:18px;border-top:1px solid var(--line);
  color:var(--fg-3);font-size:11px;line-height:1.65}
footer p{margin:0 0 8px}
footer p:last-child{margin-bottom:0}
footer b{color:var(--fg-2);font-weight:650}
.foot-lead{color:var(--fg-2)}

/* ── 回到顶部 ───────────────────────────────────────────────── */
.totop{position:fixed;right:24px;bottom:24px;z-index:50;width:42px;height:42px;
  border-radius:50%;border:1px solid var(--line);background:var(--surface-2);
  color:var(--fg-2);cursor:pointer;display:flex;align-items:center;
  justify-content:center;font-size:15px;line-height:1;
  opacity:0;transform:translateY(8px);pointer-events:none;
  transition:opacity .2s,transform .2s,color .2s,border-color .2s;
  box-shadow:0 10px 26px -12px rgba(0,0,0,.95)}
.totop.on{opacity:1;transform:none;pointer-events:auto}
.totop:hover{color:var(--gold);border-color:rgba(232,185,63,.5)}

/* ── 窄屏 ───────────────────────────────────────────────────── */
/* 报表主要在大屏看，但「手机点开预览」是常态。这里只做三件事：
   收紧留白、降标题字号、让超宽表格自己横向滚（而不是把整页顶宽）。 */
@media(max-width:720px){
  .wrap{padding:18px 14px 64px}
  .toc-in{padding:0 14px;height:44px;gap:16px}
  .toc-brand{padding-right:12px}
  header.top{flex-direction:column;align-items:flex-start;gap:12px;
    padding:2px 0 16px}
  h1{font-size:21px}
  h1 .h1-tag{display:none}
  .meta{text-align:left;font-size:11.5px;line-height:1.7}
  h2{font-size:15px;margin:40px 0 12px;padding-top:20px;gap:8px}
  h2 span.hint{font-size:11.5px}
  section.panel{padding:15px 14px;border-radius:12px}
  .kpi{padding:15px 15px 13px}
  .insights-list li{padding-left:30px;font-size:13px;line-height:1.9}
  .insights-list li:before{width:20px;height:20px;font-size:10.5px}
  .totop{right:14px;bottom:14px;width:38px;height:38px}
}

@media(prefers-reduced-motion:reduce){
  *{transition:none!important;animation:none!important;scroll-behavior:auto!important}
}
"""


def _css():
    """CSS 变量集中在 :root 里生成，颜色只有一个来源（文件头的常量）。"""
    root = (
        ":root{--bg:%s;--surface:%s;--surface-2:%s;--surface-3:%s;"
        "--line:%s;--line-soft:%s;--fg:%s;--fg-2:%s;--fg-3:%s;"
        "--gold:%s;--blue:%s;--green:%s;--orange:%s;--purple:%s;--red:%s;"
        "--font:-apple-system,BlinkMacSystemFont,'Segoe UI','PingFang SC',"
        "'Hiragino Sans GB','Microsoft YaHei',sans-serif;"
        "--mono:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace}"
        % (C_BG, C_PANEL, C_PANEL2, C_PANEL3, C_LINE, C_LINE2,
           C_FG, C_FG2, C_FG3, C_GOLD, C_BLUE, C_GREEN, C_ORANGE, C_PURPLE, C_RED)
    )
    return root + CSS



def _build_sections(payload, cal_svg, cal_hint):
    """按顺序组装章节。

    章节标题统一在这里生成，是为了三件事：
      1. 吸顶导航的锚点与页面章节**同源**，不会出现「点了没反应」；
      2. 某章没有内容时（例如没有有官方价的模型）整章连标题一起消失，
         不会留下一个空面板；
      3. <h2> 的样式只有一处定义，改版不用满文件找。
    返回 [{id, title, hint, html}]。
    """
    sections = []

    def add(sid, title, hint, inner, after=""):
        if not inner:
            return
        sections.append({
            "id": "s-" + sid,
            "title": title,
            "hint": hint,
            "html": inner,
            # after：紧跟在本节 panel 之后、下一节标题之前的散件（目前只有彩蛋）。
            # 塞进 panel 里就成了「表格脚注」，落在外面才是两节之间的留白。
            "after": after,
        })

    add("cmp", "环比", "本期 vs 等长上期　·　红=花得更多，绿=更省",
        _compare_table(payload))
    add("insight", "结论", "按当前数据自动生成，非人工判断",
        _insight_list(payload))
    add("trend", "每日趋势", "积分柱 + token 折线，双轴",
        _trend_chart(payload))

    # 日历与时段分布并排：一个回答「哪天」，一个回答「几点」，同属时间维度。
    # 两块各自包一张子卡（.sub-card）——以前它们平铺在同一个面板里，边界靠
    # 空白暗示，字距一紧就糊成一片。
    scope = payload.get("hourScope") or {}
    hour_hint = "近 30 天"
    if scope.get("since") and scope.get("until"):
        hour_hint += "（%s ~ %s）" % (scope["since"][5:], scope["until"][5:])
    cal_inner = (
        '<div class="two-col">'
        '<div class="sub-card">%s</div>'
        '<div class="sub-card"><div class="sub-h">时段分布 <span>%s</span></div>%s</div>'
        "</div>"
        % (cal_svg or '<p class="empty">暂无数据</p>', hour_hint, _hour_chart(payload))
    )
    add("cal", "活跃日历", cal_hint, cal_inner)

    price = _pricing_section(payload)
    if price:
        add("price", "官方价预估", price[0], price[1])

    add("model", "按模型",
        "近 7 天 · 积分口径，免费与限免期记 0；默认 %d 条，可展开到 %d 条"
        % (MODEL_FOLD, MODEL_CAP), _model_table(payload),
        after=_egg_line(payload))
    add("ws", "按工作区",
        "近 7 天 · 默认 %d 条，可展开到 %d 条" % (TABLE_FOLD, TABLE_CAP),
        _workspace_table(payload))
    add("sess", "按会话",
        "近 7 天 · 默认 %d 条，可展开到 %d 条" % (TABLE_FOLD, TABLE_CAP),
        _session_table(payload))
    add("req", "最贵的单轮请求", "今日 / 昨日 / 近 7 天 / 近 30 天，各取最贵一轮",
        _request_table(payload))
    add("day", "按天明细", "最近 %d 天，完整历史见 Excel" % DAY_ROWS,
        _day_table(payload))
    add("cov", "口径与数据覆盖", "", _coverage_note(payload))
    return sections


def _footer(payload):
    """页脚拆成「结论性三句」+「可展开的口径定义」。

    原文是一整段 900 字的灰色小字，与背景 3:1 对比、11.5px 中文——等于没人读。
    内容一条没删，只是把「读一遍才能确认的免责」和「要查的时候才看的口径」
    分开：前者留在明面，后者收进 details，一键展开。
    """
    peak = compact(payload["coverage"].get("peak_context", 0))
    data_name = esc(payload.get("data_name") or "（未导出）")

    lead = [
        "<p class=\"foot-lead\"><b>token 与积分均为参考值</b>：token 由本机日志累计求和、"
        "积分取平台结算值，二者都会受平台侧截断、重试、模型路由的影响，不等同于官方账单；"
        "费用预估按官方公开人民币牌价折算，只覆盖有牌价的模型，属等值参考。</p>",
        "<p><b>token 为累计处理量</b>：同一 <code>requestId</code> 下每次工具调用都会独立"
        "调一次模型，日志里逐行记录、上下文逐次增长。实测数百条请求零重复对、"
        "零跨文件重复，故全部求和——这是真实用量与计费基础。若取最大值，得到的只是"
        "峰值上下文（约 %s），会低估一个数量级以上。</p>" % peak,
        "<p>本文件为单文件离线报表，无任何外部引用、断网可看；各表下的"
        "「完整数据」按钮指向<b>同一个 Excel 工作簿</b>（<code>%s</code>，"
        "六个 sheet：汇总 / 逐请求 / 会话 / 模型 / 工作区 / 按天），"
        "数据已内嵌在页面里，点开即下载，不依赖同目录文件。</p>" % data_name,
    ]

    terms = [
        ("积分", "平台结算值（<code>credit_json</code>，含工具调用累计）；结算值缺失时用 "
                 "<code>rawUsage.credit</code> 末次值补齐，两者不相加。"),
        ("缓存命中率", "<code>cache_hit / (cache_hit + cache_miss)</code>，累计口径；"
                    "命中部分按折扣计价，所以 token 总量大不代表积分高。"),
        ("环比", "本期 vs <b>等长上期</b>（今日←前一日、昨日←前一日、近 7 日←前 7 日、"
                "近 30 日←前 30 日）。颜色按「成本」语义：红=花得更多，绿=更省；"
                "上期无数据时显示「—」，不补 0 硬造百分比。"),
        ("基准日", "数据截至的那一天。默认 <code>auto</code>：本地时间 06:00 前自动退到"
                "前一天，所以 00:05 的定时任务出的就是前一天的完整账单。"),
        ("上下文占用", "<code>session_usage.used / size</code>。该字段是当前占用，"
                   "不是累计 token，误当累计会把整张报表算歪。"),
        ("上下文压缩", "「经济模式」在上下文接近上限时自动触发，把历史摘要化。"
                   "<b>日志里没有对应的压缩事件</b>，本报表按「同一请求内 prompt 比"
                   "上一轮跌 15% 以上」识别；压缩那一次按<b>普通调用</b>计费，其消耗已含在"
                   "总量里，不额外补也不扣。省下的量＝回落幅度 × 该请求后续剩余轮次，"
                   "是<b>输入 token</b> 的毛估算：后续轮次的上下文本会继续增长，故偏保守；"
                   "且这部分多数本会命中缓存，<b>不能等同于积分节省</b>。"),
        ("会话被删之后", "删会话会连同 <code>session_usage</code> 的结算行一起消失，"
                    "但 <code>projects/**/*.jsonl</code> 日志还在。于是这些会话"
                    "<b>token 照算、积分算 0</b>——偏差只有一个方向：积分只少不多。"),
        ("今日 / 昨日口径", "这两档是<b>单日整天</b>口径，锚在<b>系统当天</b>而非基准日——"
                        "定时任务把基准日退到昨天时，它们照样显示今天与昨天的用量。"
                        "今日永远是<b>进行中</b>的数据、只到生成时刻，不能与整天的昨日"
                        "直接比大小；昨日才是上一个完整的自然日。"),
        ("今日起步中", "本地时间 <b>00:30</b> 之前，或今日积分不足 <b>%.0f</b>，"
                   "今日档只给数字、<b>不给环比与结论</b>——刚过零点那几积分对着整天的"
                   "昨日去比，必然是 -99%% 量级，那不是用量塌陷，是还没开工。"
                   "这道闸门只改<b>展示</b>，不动任何窗口的数值：近 7 日 / 近 30 日 / "
                   "全部里今日的用量照旧算在内。要强制看环比用 "
                   "<code>--today-gate off</code>。"
                   % ((payload.get("summary") or {}).get("today", {}).get("gate_threshold") or 5)),
        ("暂无 vs 0", "区间内一条请求都没有时，卡片与表行照旧保留、写「暂无」，"
                    "而不写 0.00 积分 / -100% 环比——那不是用量塌陷，是还没开始用。"),
        ("0 积分 ≠ 永久免费", "模型处于<b>限免期</b>时，平台的结算值与日志末次值都写 0，"
                        "报表据此标「免费」。这只说明<b>当下不收费</b>，不代表它永远免费——"
                        "限免结束的当天，同一个模型会从 0 直接跳到上百上千积分。所以看到"
                        "某模型「几十 M token 但 0 积分」，先确认它是否还在限免期内，"
                        "再怀疑是不是统计漏账。反过来，免费与计费混在同一天时，"
                        "该模型的「免费请求」列与实际积分会同时出现，两者并不矛盾。"),
    ]
    term_html = "".join("<p><b>%s</b>　%s</p>" % (k, v) for k, v in terms)

    return (
        "<footer>%s"
        '<details><summary>口径与计算定义（%d 条）</summary>'
        '<div class="det-body">%s</div></details>'
        "</footer>" % ("".join(lead), len(terms), term_html)
    )


def render_html(payload):
    gen = payload["generated_at"]
    anchor = payload.get("anchor") or gen.strftime("%Y-%m-%d")

    # 页眉只留三段事实：什么时候生成、数据截至哪一天、从哪个目录读的。
    # 「（昨日）」这类相对标签、「00:05 定时任务」这类流程说明、把五档窗口复述
    # 一遍的副标题、以及「当日进行中」——全删掉。当日进行中的提示由 KPI 卡片
    # 旁的「截至 HH:MM」和结论区承担，页眉再说一遍只会跟「生成于」同值重复。
    until = "数据截至 %s" % anchor
    if payload.get("trimmed_days"):
        until += "　·　聚焦最近 %d 天" % payload["trimmed_days"]

    cal_svg, cal_hint = _calendar(payload)

    sections = _build_sections(payload, cal_svg, cal_hint)

    # 吸顶导航。11 个章节、整页高近 5000px，没有导航就只能靠滚轮找某一节。
    nav = "".join(
        '<a href="#%s">%s</a>' % (s["id"], esc(s["title"])) for s in sections
    )

    chapters = []
    for s in sections:
        hint = ('<span class="hint">%s</span>' % esc(s["hint"])) if s["hint"] else ""
        chapters.append(
            '<h2 id="%s">%s%s</h2><section class="panel">%s</section>%s'
            % (s["id"], esc(s["title"]), hint, s["html"], s.get("after") or "")
        )

    body = [
        '<nav class="toc"><div class="toc-in">'
        '<span class="toc-brand">WorkBuddy <b>账单</b></span>%s</div></nav>' % nav,
        '<div class="wrap">',
        '<header class="top">'
        '<h1>WorkBuddy 用量与积分账单<span class="h1-tag">单文件离线</span></h1>'
        '<div class="meta">生成于 %s　·　%s<br>数据源 %s（只读）</div></header>'
        % (esc(gen.strftime("%Y-%m-%d %H:%M")), esc(until),
           # normpath 统一分隔符：config_dir 常是 "C:\Users\x/.workbuddy" 这种
           # 混写（盘符来自环境变量、末段来自路径拼接），照原样印出来像笔误。
           esc(os.path.normpath(payload["config_dir"]))),
        _kpi_cards(payload),
        _compact_strip(payload),
        "".join(chapters),
        _footer(payload),
        "</div>",
        # 折叠表格的展开动作。单文件离线，所以用内联脚本，不引任何 CDN。
        "<script>(function(){document.addEventListener('click',function(ev){"
        "var btn=ev.target&&ev.target.closest?ev.target.closest('.more-btn'):null;"
        "if(!btn||btn.disabled)return;"
        "var tbl=document.getElementById(btn.getAttribute('data-target'));"
        "if(!tbl)return;"
        "tbl.classList.add('expanded');"
        "var total=parseInt(btn.getAttribute('data-total'),10)||0;"
        "var shown=parseInt(btn.getAttribute('data-shown'),10)||total;"
        "btn.textContent=(total>shown?'已显示前 '+shown+' 条':'已展开全部 '+total+' 条')"
        "+'，其余见完整数据';"
        "btn.disabled=true;});})();</script>",
        # 回到顶部：长页面的最低成本救急键
        '<button class="totop" id="totop" type="button" title="回到顶部" '
        'aria-label="回到顶部">↑</button>'
        "<script>(function(){var b=document.getElementById('totop');if(!b)return;"
        "var rm=window.matchMedia&&window.matchMedia("
        "'(prefers-reduced-motion: reduce)').matches;"
        "function s(){if(window.scrollY>600){b.classList.add('on')}"
        "else{b.classList.remove('on')}}"
        "window.addEventListener('scroll',s,{passive:true});s();"
        "b.addEventListener('click',function(){window.scrollTo({top:0,"
        "behavior:rm?'auto':'smooth'})});})();</script>",
        # 日历翻页：所有视图在渲染期就画好了，这里只切显隐。不重新计算、
        # 不依赖任何外部脚本，断网也能翻；默认停在最后一页（近 30 天）。
        "<script>(function(){"
        "var vs=[].slice.call(document.querySelectorAll('.cal-view'));"
        "if(!vs.length)return;"
        "var pv=document.querySelector('.cal-prev'),"
        "nx=document.querySelector('.cal-next'),"
        "cnt=document.querySelector('.cal-count');"
        "var i=vs.length-1;"
        "function paint(){"
        "for(var k=0;k<vs.length;k++){vs[k].classList.toggle('is-on',k===i)}"
        "if(pv)pv.disabled=(i===0);"
        "if(nx)nx.disabled=(i===vs.length-1);"
        "if(cnt)cnt.textContent=(i+1)+' / '+vs.length;}"
        "document.addEventListener('click',function(ev){"
        "var b=ev.target&&ev.target.closest?ev.target.closest('.cal-btn'):null;"
        "if(!b||b.disabled)return;"
        "var d=parseInt(b.getAttribute('data-cal'),10);"
        "var j=i+(d||0);if(j<0||j>=vs.length)return;"
        "i=j;paint();});"
        "paint();})();</script>",
    ]

    # 「完整数据」的兜底：把工作簿以 base64 内嵌，页面加载后灌进所有
    # a[data-csv] 的 href。相对路径在本地 file:// 下能用，但报表经常被
    # WorkBuddy 预览面板托管成 http://127.0.0.1:xxx/static-html/<hash>/index.html，
    # 那个目录里只有 index.html 一份文件——相对路径 100% 404，这也是
    # 「所有下载链接都打不开」的直接原因。内嵌之后两种环境都点得开。
    # 只写一份，所有链接共用，避免每张表各塞一份把文件撑到几 MB。
    if payload.get("data_b64"):
        body.append(
            "<script>(function(){var b=%s;"
            "var as=document.querySelectorAll('a[data-csv]');"
            "for(var i=0;i<as.length;i++){as[i].href=%s+b;}"
            "})();</script>"
            % (json.dumps(payload["data_b64"]), json.dumps(XLSX_DATA_URI))
        )

    doc = (
        "<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\">"
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        "<title>WorkBuddy 用量与积分账单 · %s</title><style>%s</style></head>"
        "<body>%s</body></html>"
        % (esc(gen.strftime("%Y-%m-%d %H:%M")), _css(), "".join(body))
    )
    return doc
