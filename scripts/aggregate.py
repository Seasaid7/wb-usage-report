"""聚合层：把逐请求记录摊成用户要看的四个时间档与七个维度。

基准日（anchor）
    报表是「截至某日的快照」。默认 auto：本地时间早于 06:00 退到前一天，
    所以推荐 00:05 的定时任务天然出「前一天」完整账单；白天手动跑则取当天。

窗口定义（全部按本地时区 +08:00 的自然日，截至基准日）
    today  : 系统当天（基准日退到昨天时它仍是今天，见下文）
    prev   : 系统当天的前一天（即「昨日」），单日整天口径
    7d     : 基准日往前共 7 个自然日
    30d    : 基准日往前共 30 个自然日
    all    : 基准日及以前的全部历史

    today / prev 是**两张单日卡**，锚在系统当天而与基准日解耦：00:05 的定时任务
    把基准日退到昨天时，该看的恰恰是「昨天花了多少」和「今天到现在花了多少」。
    回溯历史某天时它们才跟着基准日走，避免「截至 10-05 的报表里塞一张 10-24 的卡」。

环比（compare）
    每个窗口都配一个等长上期：today→前一日（即昨日），prev→前一日（前日），
    7d→前 7 日，30d→前 30 日。全部窗口没有上期，compare 为 None，不硬造数字。

统计口径
    tokens      = input + output，均为**累计处理量**（同一 requestId 下多次 LLM
                  调用求和）。这是真实用量与计费基础，不是峰值上下文。
    peak_ctx    = 每个 request 峰值上下文之和，只用于观察上下文规模。
    cache_rate  = cache_hit / (cache_hit + cache_miss)，累计口径。
    credit      = 结算值优先、末次值补齐（实测与子代理无重叠，可直接并入）。
    cost        = 按官方公开牌价折算的**人民币**等值费用，仅覆盖有牌价的模型；
                  未定价模型的 token 不计入（`cost_coverage` 说明覆盖率）。
                  token 与积分都是参考值，不等于账单。
"""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime, timedelta

# 聚合层不引用 pricing：价估是渲染/导出层的事，联网调用这里只数次数、不折钱。

# 洞察区最多给出的条数，避免刷屏
# 结论区固定 5 组：今日 / 近 7 日与近 30 日 / 缓存 / 结构 / 隐性成本。
# 每组内部用「；」把相关数字串成一句，不再一条一个 bullet。费用不进结论区——
# 官方价预估单独成表，结论里复述一遍只是把同一组数字说两遍。
MAX_INSIGHTS = 5

# 缓存命中率低于此值视为异常日，需要在洞察里点名
CACHE_ALERT_RATE = 0.6

# 单日未缓存 token 占比超过此值视为成本集中
CACHE_MISS_SHARE_ALERT = 0.25

# 基准日自动回退阈值。定时任务推荐 00:05 执行，此时当天还没有任何数据，
# 真正该出的是「前一天」的完整账单，所以 06:00 之前自动把基准日退一天。
ANCHOR_AUTO_HOUR = 6

# ── 今日档「起步中」闸门 ─────────────────────────────────────────────
# 零点到 00:30 之间跑报表，多半是为了看昨天：今日那几十分钟的量对着整整一个
# 昨日去算环比，必然得出 -99% 量级的数字，看起来像用量塌陷，其实只是还没开工。
# 00:30 之后改为按样本量判：今日积分太少同样谈不上参考价值。
#
# 为什么用积分而不是请求数：请求粒度差异太大（小问答与长任务差两个数量级），
# 积分是这台机器上唯一跨模型可比的量。实测 28 个活跃日——全局 1.658 积分/百万
# token，5 积分约合 3.0M token（1~4 个请求）；00:30 时点今日积分低于 5 的有
# 23 天（82%）；今日累计积分首破 5 的中位时刻是 01:35。所以 5 这个数的实际
# 效果是「只盖住刚起步的一两个小时」，与 00:30 的时间闸基本重叠、互为兜底。
#
# 注意：这道闸门只管**展示**——把环比和结论收起来，数字照给。它不改窗口数值，
# 所以「近 7 日 / 近 30 日 / 全部」里今日的用量仍然算在内，对得上账。
TODAY_GATE_BEFORE_MIN = 30      # 00:30 之前一律降级
TODAY_MIN_CREDIT = 5.0          # 低于此积分视为样本不足


def today_gate(now, block, live_today=True, mode="auto"):
    """判定今日档是否降级成「起步中」，返回 (degraded, reason)。

    reason 取值：``"early"``（00:30 前）/ ``"thin"``（积分未达阈值）/ None。
    mode 取 ``"off"`` 时整道闸门关闭，用于「就是要看今日档」的场合。
    """
    if mode == "off" or not live_today:
        return False, None
    if now.hour * 60 + now.minute < TODAY_GATE_BEFORE_MIN:
        return True, "early"
    if float(block.get("credit") or 0) < TODAY_MIN_CREDIT:
        return True, "thin"
    return False, None

# 归因：变化幅度超过此比例才值得追责到具体模型/会话
ATTRIB_MIN_RATE = 0.15
# 单个贡献者至少占总变化的这个比例，才写进结论，避免报一堆零头
ATTRIB_SHARE_MIN = 0.05
# 每个维度最多点名几个
ATTRIB_MAX_ITEMS = 2

WINDOW_LABELS = [
    ("today", "今日"),
    ("prev", "昨日"),
    ("7d", "近 7 日"),
    ("30d", "近 30 日"),
    ("all", "全部"),
]

# 各窗口对应的「上期」，用于环比。窗口区间为闭区间 [start, end]。
# today 与 prev 都是单日窗口，上期就是各自的前一天（前者落在昨日、后者落在前日）。
PREV_LABELS = {"today": "前一日", "prev": "前一日",
               "7d": "前 7 日", "30d": "前 30 日"}


def resolve_anchor(bundle, mode=None):
    """解析「数据截至哪一天」，返回 (date, 模式说明)。

    auto（默认）：本地时间早于 06:00 退到前一天，否则取当天。
        这样 00:05 的定时任务天然出「前一天」报告，白天手动跑也不会空跑。
    prev        ：强制前一天
    today       ：强制当天
    YYYY-MM-DD  ：指定某天（回溯历史某日）
    """
    now = bundle["generated_at"]
    today = now.date()
    mode = (mode or "auto").strip().lower()
    if mode in ("", "auto"):
        if now.hour < ANCHOR_AUTO_HOUR:
            return today - timedelta(days=1), "auto"
        return today, "auto"
    if mode == "today":
        return today, "today"
    if mode in ("prev", "yesterday", "d-1"):
        return today - timedelta(days=1), "prev"
    try:
        return datetime.strptime(mode, "%Y-%m-%d").date(), "manual"
    except ValueError:
        raise ValueError(
            "--anchor 只接受 auto / today / prev / YYYY-MM-DD，收到 %r" % mode
        )


def apply_anchor(bundle, anchor):
    """把记录裁到「基准日及以前」，返回新 bundle。

    必须裁在聚合之前：报表是「截至某日的快照」，基准日之后的数据不该混进来，
    否则回溯历史某天时数字会被后续用量污染。

    唯一例外：**系统当天**的记录要留着。基准日取「昨日」（00:05 定时任务的默认口径）
    时，今天的用量如果不留，「今日」那张卡永远是空的——而用户最想看的恰恰是今天。
    现在裁剪上界取 ``max(基准日, 当天)``，由聚合层再把报表主体单独收回基准日：
    7 日 / 30 日 / 全部仍严格截止基准日，只有「今日」这一档用到当天数据。
    """
    anchor_str = anchor.strftime("%Y-%m-%d")
    today_str = bundle["generated_at"].strftime("%Y-%m-%d")
    keep_str = max(anchor_str, today_str)
    out = dict(bundle)
    out["records"] = [
        r for r in bundle["records"]
        if not r["day"] or r["day"] <= keep_str
    ]
    out["web_calls"] = [
        w for w in (bundle.get("web_calls") or [])
        if not w["day"] or w["day"] <= keep_str
    ]
    out["anchor"] = anchor
    return out


def _num(value, digits=2):
    return round(float(value or 0), digits)


def _stat_block(records):
    """一组记录的核心指标。所有窗口与维度都走这一个函数，保证口径一致。"""
    tokens_in = sum(r["input"] for r in records)
    tokens_out = sum(r["output"] for r in records)
    hit = sum(r["cache_hit"] for r in records)
    miss = sum(r["cache_miss"] for r in records)
    think = sum(r["thinking"] for r in records)
    credit = sum(r["credit"] for r in records)
    denom = hit + miss
    priced = [r for r in records if r.get("cost_priced")]
    priced_tokens = sum(r["input"] + r["output"] for r in priced)
    all_tokens = tokens_in + tokens_out
    return {
        "requests": len(records),
        # 空窗口要单独标出来：0 请求 ≠ 用量塌陷。渲染层据此显示「暂无」而不是
        # 0.00 积分 + -100% 环比，后者会被读成「今天什么都没干，昨天白花钱」。
        "empty": not records,
        "sessions": len({r["session_id"] for r in records if r["session_id"]}),
        "input": tokens_in,
        "output": tokens_out,
        "tokens": tokens_in + tokens_out,
        "cache_hit": hit,
        "cache_miss": miss,
        "cache_rate": (hit / float(denom)) if denom else None,
        "thinking": think,
        "credit": _num(credit),
        "credit_paid": _num(sum(r["credit"] for r in records if not r["is_free"])),
        "free_tokens": sum(r["input"] + r["output"] for r in records if r["is_free"]),
        "free_requests": sum(1 for r in records if r["is_free"]),
        "sub_requests": sum(1 for r in records if r["is_sub"]),
        "sub_tokens": sum(
            r["input"] + r["output"] for r in records if r["is_sub"]
        ),
        "sub_credit": _num(sum(r["credit"] for r in records if r["is_sub"])),
        "no_credit_requests": sum(1 for r in records if r["credit_src"] == "none"),
        "no_credit_tokens": sum(
            r["input"] + r["output"] for r in records if r["credit_src"] == "none"
        ),
        "active_days": len({r["day"] for r in records if r["day"]}),
        "llm_calls": sum(r["calls"] for r in records),
        "peak_context": sum(r["peak_input"] + r["peak_output"] for r in records),
        "credit_per_1m": (
            _num(credit / ((tokens_in + tokens_out) / 1000000.0), 3)
            if (tokens_in + tokens_out)
            else 0.0
        ),
        # 费用预估（人民币）：只覆盖有官方牌价的模型，未定价部分不折算
        "cost": _num(sum(r.get("cost_cny") or 0 for r in priced), 4),
        "cost_if_peak": _num(sum(r.get("cost_if_peak") or 0 for r in priced), 4),
        "priced_tokens": priced_tokens,
        "unpriced_tokens": all_tokens - priced_tokens,
        "cost_coverage": (priced_tokens / float(all_tokens)) if all_tokens else None,
        "peak_tokens": sum(
            r["input"] + r["output"] for r in records if r.get("cost_tier") == "peak"
        ),
        "priced_requests": len(priced),
    }


def _window_start(anchor, key):
    if key == "today":
        return anchor
    if key == "prev":
        return anchor - timedelta(days=1)
    if key == "7d":
        return anchor - timedelta(days=6)
    if key == "30d":
        return anchor - timedelta(days=29)
    return None


def _prev_range(start, key):
    """上期闭区间。等长前移，保证环比可比。

    参数是**本窗口的起始日**而不是基准日：今日/7 日/30 日的起点各不相同，
    拿基准日当锚点会把「今日」的上期算成「基准日的前一天」，基准日取昨日时
    就变成「前一日的前一日」，正好错开一天。

    today 与 prev 都是单日窗口，上期各自都是「本窗口起始日的前一天」：
    今日的上期落在昨日，昨日的上期落在前日。
    """
    if key in ("today", "prev"):
        day = start - timedelta(days=1)
        return day, day
    if key == "7d":
        return start - timedelta(days=7), start - timedelta(days=1)
    if key == "30d":
        return start - timedelta(days=30), start - timedelta(days=1)
    return None


def _pick(records, start, end):
    """闭区间取记录；start 为 None 表示不设下限。"""
    if start is None:
        return records
    lo = start.strftime("%Y-%m-%d")
    hi = end.strftime("%Y-%m-%d") if end else None
    return [
        r for r in records
        if r["day"] and r["day"] >= lo and (hi is None or r["day"] <= hi)
    ]


def _compact_block(events, start, end, tokens_in):
    """把上下文压缩事件按窗口分桶，并算「省下的输入 token 占比」。

    分母用**该窗口的输入 token**，不是输入+输出：压缩砍掉的是送进去的上下文，
    跟模型吐出来多少无关。拿 token 总和当分母会把占比压低，读起来像「压缩没用」。

    窗口内一次压缩都没发生是常态（短会话根本撞不到上下文天花板），
    此时返回 count=0，渲染层据此整条不显示，而不是印一行「0 次，省下 0 token」。
    """
    lo = start.strftime("%Y-%m-%d") if start else None
    hi = end.strftime("%Y-%m-%d") if end else None
    picked = [
        e for e in events
        if e.get("day") and (lo is None or e["day"] >= lo)
        and (hi is None or e["day"] <= hi)
    ]
    saved = sum(e["saved"] for e in picked)
    return {
        "count": len(picked),
        "saved": saved,
        # 累计砍掉的上下文（不含后续轮次的乘数），用于解释 saved 是怎么来的
        "drop": sum(e["drop"] for e in picked),
        "tokens_in": tokens_in,
        "share": (saved / float(tokens_in)) if tokens_in else None,
        "sessions": len({e["session_id"] for e in picked if e["session_id"]}),
    }


def _web_block(web_calls, start, end):
    """联网检索 / 抓取的调用次数。

    官方对 Search / Fetch 按**次**计费，跟 token 无关，进不了 token 单价模型，
    只能单独数、单独交代。计数取自日志里的 ``function_call`` 记录，是
    「本机记录的调用次数」——本地账本分辨不出它是否真的走了厂商的
    Search / Fetch API，所以只能当参考，不能当账单。
    """
    lo = start.strftime("%Y-%m-%d") if start else None
    hi = end.strftime("%Y-%m-%d") if end else None
    picked = [
        w for w in web_calls
        if w["day"] and (lo is None or w["day"] >= lo) and (hi is None or w["day"] <= hi)
    ]
    return {
        "total": len(picked),
        "search": sum(1 for w in picked if w["name"] == "WebSearch"),
        "fetch": sum(1 for w in picked if w["name"] == "WebFetch"),
        "days": len({w["day"] for w in picked}),
    }


def _pct_delta(cur, prev):
    if prev in (None, 0):
        return None
    return (cur - prev) / float(prev)


def _compare(cur_block, prev_block, label):
    """本期 vs 上期的变化。上期无数据时全部置 None，由渲染层显示「—」。"""
    has_prev = prev_block["requests"] > 0
    cache_pp = None
    if cur_block["cache_rate"] is not None and prev_block["cache_rate"] is not None:
        cache_pp = (cur_block["cache_rate"] - prev_block["cache_rate"]) * 100
    return {
        "label": label,
        "has_prev": has_prev,
        # 本期空窗口：环比算出来必然是 -100%，那是「还没开始用」不是「用量塌陷」，
        # 标记出来让渲染层与结论区都闭嘴。
        "cur_empty": cur_block["requests"] == 0,
        "prev": prev_block,
        "credit_abs": _num(cur_block["credit"] - prev_block["credit"]),
        "credit_pct": _pct_delta(cur_block["credit"], prev_block["credit"]),
        "tokens_abs": cur_block["tokens"] - prev_block["tokens"],
        "tokens_pct": _pct_delta(cur_block["tokens"], prev_block["tokens"]),
        "requests_abs": cur_block["requests"] - prev_block["requests"],
        "requests_pct": _pct_delta(cur_block["requests"], prev_block["requests"]),
        "sessions_abs": cur_block["sessions"] - prev_block["sessions"],
        "sessions_pct": _pct_delta(cur_block["sessions"], prev_block["sessions"]),
        "cache_rate_pp": cache_pp,
    }


def _anchor_label(anchor, today):
    if anchor == today:
        return "今日"
    if anchor == today - timedelta(days=1):
        return "昨日"
    return "指定日"


# ── 环比归因 ─────────────────────────────────────────────────────
def _delta_items(cur, prev, key_fn, label_fn, metric):
    """按某个维度算「本期 - 上期」，按绝对变化量降序。"""
    a = defaultdict(float)
    b = defaultdict(float)
    for rec in cur:
        a[key_fn(rec)] += rec[metric]
    for rec in prev:
        b[key_fn(rec)] += rec[metric]
    items = []
    for key in set(a) | set(b):
        delta = a[key] - b[key]
        if abs(delta) < 1e-9:
            continue
        items.append(
            {"label": label_fn(key), "cur": _num(a[key]), "prev": _num(b[key]),
             "delta": _num(delta)}
        )
    items.sort(key=lambda it: -abs(it["delta"]))
    return items


def _attribute(cur, prev, sessions):
    """把环比变化归到具体模型与会话。

    只在幅度足够大（|变化率| ≥ 15%）时才做，否则每天的随机波动都会被
    说成「某会话在烧」，那是噪声不是洞察。

    积分全为 0 的时段（免费模型期）自动退化到按 token 归因——否则会出现
    「用量翻倍但指数说持平」这种自相矛盾的结论。
    """
    cur_credit = sum(r["credit"] for r in cur)
    prev_credit = sum(r["credit"] for r in prev)
    metric = "credit"
    cur_v, prev_v = cur_credit, prev_credit
    if prev_credit <= 0 or cur_credit <= 0:
        metric = None
        for candidate in ("tokens",):
            c = sum(r["input"] + r["output"] for r in cur)
            p = sum(r["input"] + r["output"] for r in prev)
            if p > 0:
                metric, cur_v, prev_v = candidate, c, p
                break
    if not metric or prev_v <= 0:
        return None
    total = cur_v - prev_v
    rate = total / float(prev_v)
    if abs(rate) < ATTRIB_MIN_RATE:
        return {"direction": "flat", "rate": rate, "metric": metric,
                "total": _num(total), "models": [], "sessions": []}

    direction = "down" if total < 0 else "up"

    def token_of(rec):
        return rec["input"] + rec["output"]

    if metric == "credit":
        get = lambda rec: rec["credit"]  # noqa: E731
    else:
        get = token_of  # noqa: E731

    # 让 _delta_items 能用统一的取值方式
    shim_cur = [{"k_model": r["model"], "k_sid": r["session_id"], "v": get(r)}
                for r in cur]
    shim_prev = [{"k_model": r["model"], "k_sid": r["session_id"], "v": get(r)}
                 for r in prev]

    models = _delta_items(shim_cur, shim_prev, lambda r: r["k_model"],
                          lambda k: k or "(未知模型)", "v")
    sessions_items = _delta_items(
        shim_cur, shim_prev, lambda r: r["k_sid"],
        lambda k: ((sessions.get(k) or {}).get("title") or str(k or "(未知会话)")), "v"
    )

    def keep(items):
        """只留同向贡献者，份额按**同向合计**算，保证加起来不超过 100%。

        用净变化做分母会算出「占降幅 142%」这种荒唐值——那是因为另一侧有
        增量在抵消。反向的抵消单独报，不混进份额里。
        """
        sign = -1 if direction == "down" else 1
        picked = [it for it in items if it["delta"] * sign > 0]
        gross = sum(abs(it["delta"]) for it in picked)
        if not gross:
            return []
        kept = [it for it in picked
                if abs(it["delta"]) / gross >= ATTRIB_SHARE_MIN]
        kept.sort(key=lambda it: -abs(it["delta"]))
        out = []
        for it in kept[:ATTRIB_MAX_ITEMS]:
            item = dict(it)
            item["share"] = abs(it["delta"]) / gross
            item["label"] = str(item["label"])[:20]
            out.append(item)
        return out

    models_top = keep(models)
    sessions_top = keep(sessions_items)
    same_dir = sum(abs(m["delta"]) for m in models
                   if m["delta"] * (-1 if direction == "down" else 1) > 0)
    # 反向变动（抵消项）：净变化与同向合计之差
    offset = abs(abs(total) - same_dir) if same_dir else 0.0

    # 会话与模型经常是同一个来源，重复报两遍等于凑字数
    def _dup(item, tops):
        return any(abs(abs(item["delta"]) - abs(t["delta"])) <= abs(t["delta"]) * 0.01
                   for t in tops)

    sessions_top = [s for s in sessions_top if not _dup(s, models_top)]

    return {
        "direction": direction,
        "rate": rate,
        "metric": metric,
        "total": _num(total),
        "abs_total": _num(abs(total)),
        "same_dir": _num(same_dir),
        "offset": _num(offset),
        "models": models_top,
        "sessions": sessions_top,
    }


def _attrib_clause(attr, joiner="，", max_models=2):
    """归因内容本身，不带「XX上涨归因：」前缀和句号，供拼进长句。

    模型与会话之间用逗号分组：两个维度的占比各有各的分母，混在一个顿号序列里
    会被误读成同一份饼。
    """
    if not attr or attr.get("direction") == "flat":
        return None
    models = attr.get("models") or []
    sess = attr.get("sessions") or []
    if not models and not sess:
        return None
    down = attr["direction"] == "down"
    share_word = "降幅" if down else "涨幅"
    if attr["metric"] == "credit":
        fmt = lambda v: "%+.2f" % v  # noqa: E731
    else:
        fmt = lambda v: "%+.1fM" % (v / 1000000.0)  # noqa: E731
    model_parts = []
    for i, it in enumerate(models[:max_models]):
        # 「占涨幅/占降幅」只标在第一个模型上，重复三次反而看不清
        tag = "占%s " % share_word if i == 0 else ""
        model_parts.append("%s %s（%s%.0f%%）"
                           % (it["label"], fmt(it["delta"]), tag, it["share"] * 100))
    sess_parts = ["会话「%s」%s（%.0f%%）" % (it["label"], fmt(it["delta"]),
                                          it["share"] * 100)
                  for it in sess[:1]]
    line = "，".join(x for x in ("、".join(model_parts), "、".join(sess_parts)) if x)
    offset = attr.get("offset") or 0
    same = attr.get("same_dir") or 0
    if same and offset > same * 0.05:
        line += "%s另有 %s 的反向变动抵消了一部分" % (joiner, fmt(offset))
    return line


def _attrib_text(attr, title):
    """归因单独成句（不进合并的长句时用）。没有可点名对象时返回 None。"""
    body = _attrib_clause(attr, joiner="；")
    if not body:
        return None
    word = "回落" if attr.get("direction") == "down" else "上涨"
    return "%s%s归因：%s。" % (title, word, body)


def _by_day(records):
    buckets = defaultdict(list)
    for rec in records:
        if rec["day"]:
            buckets[rec["day"]].append(rec)
    out = {}
    for day in sorted(buckets):
        block = _stat_block(buckets[day])
        block["day"] = day
        out[day] = block
    return out


def _by_model(records):
    buckets = defaultdict(list)
    for rec in records:
        buckets[rec["model"]].append(rec)
    out = {}
    for model, items in buckets.items():
        block = _stat_block(items)
        block["model"] = model
        out[model] = block
    return dict(
        sorted(out.items(), key=lambda kv: (-kv[1]["credit"], -kv[1]["tokens"]))
    )


def _by_workspace(records):
    buckets = defaultdict(list)
    for rec in records:
        buckets[rec["workspace"]].append(rec)
    out = {}
    for name, items in buckets.items():
        block = _stat_block(items)
        block["workspace"] = name
        out[name] = block
    return dict(
        sorted(out.items(), key=lambda kv: (-kv[1]["credit"], -kv[1]["tokens"]))
    )


def _by_hour(records):
    buckets = defaultdict(list)
    for rec in records:
        if rec["hour"] is not None:
            buckets[rec["hour"]].append(rec)
    return {
        hour: _stat_block(items) for hour, items in sorted(buckets.items())
    }


def _by_session(records, sessions):
    buckets = defaultdict(list)
    for rec in records:
        buckets[rec["session_id"]].append(rec)
    out = []
    for sid, items in buckets.items():
        block = _stat_block(items)
        meta = sessions.get(sid) or {}
        days = sorted({r["day"] for r in items if r["day"]})
        block.update(
            {
                "session_id": sid,
                "title": meta.get("title") or items[0]["title"],
                "workspace": items[0]["workspace"],
                "model": (items[0]["model"] if len({r["model"] for r in items}) == 1
                          else "%d 个模型" % len({r["model"] for r in items})),
                "first_day": days[0] if days else None,
                "last_day": days[-1] if days else None,
                "context_used": meta.get("context_used") or 0,
                "context_size": meta.get("context_size") or 0,
                "context_rate": meta.get("context_rate"),
                "deleted": bool(meta.get("deleted")),
                "playground": bool(meta.get("playground")),
            }
        )
        out.append(block)
    out.sort(key=lambda r: (-r["credit"], -r["tokens"]))
    return out


def _by_turn(records):
    """会话内按时间排序编号，还原「第几轮」。"""
    grouped = defaultdict(list)
    for rec in records:
        if rec["ts"]:
            grouped[rec["session_id"]].append(rec)
    turns = []
    for items in grouped.values():
        items.sort(key=lambda r: r["ts"])
        for index, rec in enumerate(items, start=1):
            turns.append(
                {
                    "session_id": rec["session_id"],
                    "title": rec["title"],
                    "turn": index,
                    "model": rec["model"],
                    "workspace": rec["workspace"],
                    "credit": rec["credit"],
                    "credit_src": rec["credit_src"],
                    "input": rec["input"],
                    "output": rec["output"],
                    "tokens": rec["input"] + rec["output"],
                    "cache_hit": rec["cache_hit"],
                    "cache_miss": rec["cache_miss"],
                    "thinking": rec["thinking"],
                    "calls": rec["calls"],
                    "peak_input": rec["peak_input"],
                    "is_sub": rec["is_sub"],
                    "is_free": rec["is_free"],
                    "cost": rec.get("cost_cny") or 0.0,
                    "cost_tier": rec.get("cost_tier"),
                    "cost_priced": bool(rec.get("cost_priced")),
                    "day": rec["day"],
                    "ts": rec["ts"],
                }
            )
    turns.sort(key=lambda r: (r["ts"], r["session_id"]))
    return turns


def _insights(summary, by_day, by_model, by_session, coverage, turns, windows=None,
              by_model_7=None, by_session_7=None):
    out = []
    live = {k for k, _ in (windows or WINDOW_LABELS)}

    def _get(key):
        return (summary.get(key) or {}) if key in live else {}

    today = _get("today")
    week = _get("7d")
    month = _get("30d")
    total = _get("all")

    # 0) 环比：日报的头号结论，放最前面
    def _cmp_line(key):
        if key not in live:
            return None
        block = summary.get(key) or {}
        cmp_ = block.get("compare")
        if not cmp_ or not cmp_.get("has_prev") or cmp_.get("credit_pct") is None:
            return None
        # 本期一条记录都没有时不出环比：-100% 会被读成断崖式下滑。
        if cmp_.get("cur_empty"):
            return None
        # 上期被记录起点截断时不进结论区：放大倍数只是「本机用得晚」，
        # 不是真实节奏变化，写进结论会误导。
        if cmp_.get("partial_prev"):
            return None
        return "%s %.2f 积分，环比%s %.2f %+.1f%%" % (
            block.get("label") or key, block["credit"], cmp_["label"],
            cmp_["prev"]["credit"], cmp_["credit_pct"] * 100,
        )

    today_block = summary.get("today") or {}
    # 今日档被闸门降级时不主张：「今日 2.39 积分，环比前一日 -98.9%」这句话本身
    # 就是噪音来源，整条不出。降级的意思是「不下结论」，不是「换个说法下结论」。
    head = None if today_block.get("gated") else _cmp_line("today")
    prev_head = _cmp_line("prev")
    # ── 第 1 组：今日 / 昨日（环比 + 日均倍数 + 归因，一句说完）──
    day_parts = []
    if "today" in live and not today_block.get("requests"):
        day_parts.append(
            "%s暂无用量记录（数据截至 %s），环比留空而不补 0；下面是最近有数据的"
            "日期与其余窗口。" % (today_block.get("label") or "今日",
                                today_block.get("as_of") or "当前")
        )
    # 先把「近 7 日 vs 近 30 日日均」算出来，今日结论和 7 日结论都要用它：
    # 单看一个 +1230% 没有参照系，说清「是日均的几倍」才知道是不是真跑飞了。
    week_daily = (week["credit"] / float(week["active_days"])
                  if week.get("active_days") else None)
    month_daily = (month["credit"] / float(month["active_days"])
                   if month.get("active_days") else None)
    daily_note = None      # 拼进「近 7 日」那一句的日均对比
    if week_daily and month_daily:
        ratio = week_daily / month_daily
        if ratio >= 1.5:
            daily_note = ("日均 %.2f 积分，是近 30 日日均 %.2f 的 %.1f 倍，"
                          "近期用量明显加速" % (week_daily, month_daily, ratio))
        elif ratio <= 0.6:
            daily_note = ("日均 %.2f 积分，仅为近 30 日日均 %.2f 的 %.0f%%，"
                          "用量已回落" % (week_daily, month_daily, ratio * 100))
        else:
            daily_note = ("日均 %.2f 积分，与近 30 日日均 %.2f 基本持平（%.1f 倍），"
                          "节奏稳定" % (week_daily, month_daily, ratio))

    # ── 第 1 组（续）：今日环比 + 昨日环比 ──
    if head:
        cmp_today = summary["today"]["compare"]
        rate = cmp_today["credit_pct"]
        # 涨了，并且今日确实是 7 日日均的两倍以上，才值得喊「去看看是谁在烧」；
        # 否则只说涨幅，不硬凑一句悬赏。
        mult = (today["credit"] / week_daily
                if (week_daily and week.get("active_days", 0) >= 2) else None)
        if today_block.get("partial"):
            # 今日永远没过完。00:20 的 4 积分对着昨天整天的 387 积分必然显示 -98%，
            # 那不是「用量塌陷」而是「还没到中午」——这种数字不能下结论。
            verdict = "今日还没过完（截至 %s），与整日相比只作方向参考" % (
                today_block.get("as_of") or "现在")
        elif rate > 0.15 and mult and mult >= 2:
            verdict = ("是近 7 日日均 %.2f 的 %.1f 倍，看看是谁在偷吃大米饭"
                       % (week_daily, mult))
        elif rate > 0.15:
            verdict = "涨幅明显"
        elif rate < -0.15:
            verdict = "明显回落"
        else:
            verdict = "基本持平"
        if cmp_today.get("tokens_pct") is not None:
            head += "（token %+.1f%%）" % (cmp_today["tokens_pct"] * 100)
        head += "，%s" % verdict
        day_parts.append(head)
        # 今日没过完时不给归因：那个「降幅」是时间没走完造成的，把 -380 积分
        # 拆到具体模型头上，等于给一个假结论配套解释，越解释越错。
        if not today_block.get("partial"):
            body = _attrib_clause(cmp_today.get("attribution"), max_models=1)
            if body:
                day_parts.append("其中 " + body)
    # 昨日单列一句：今日永远只到生成时刻，昨日才是「上一个完整自然日」的真可比
    # 对象。归因不重复给——同一天干的活通常还是那几号模型，说两遍只是变长。
    if prev_head:
        prev_cmp = summary["prev"]["compare"]
        if prev_cmp.get("tokens_pct") is not None:
            prev_head += "（token %+.1f%%）" % (prev_cmp["tokens_pct"] * 100)
        day_parts.append(prev_head)
    if day_parts:
        out.append("；".join(day_parts) + "。")

    # ── 第 2 组：近 7 日 / 近 30 日（环比 + 日均对比 + 归因，一句说完）──
    span_parts = []
    week_line = _cmp_line("7d")
    week_cmp = (summary.get("7d") or {}).get("compare") or {}
    week_attr = week_cmp.get("attribution") or {}
    if week_line:
        if daily_note:
            week_line += "，" + daily_note
        span_parts.append(week_line)
        if (week_cmp.get("has_prev")
                and week_attr.get("direction") in ("up", "down")
                and abs(week_attr.get("rate") or 0) >= 0.3):
            body = _attrib_clause(week_attr, max_models=1)
            if body:
                span_parts.append("其中 " + body)
    month_line = _cmp_line("30d")
    if month_line:
        span_parts.append(month_line)
    if span_parts:
        out.append("；".join(span_parts) + "。")

    # ── 第 3 组：缓存与经济性（命中率 + 未缓存量 + 单 token 成本 + 异常日）──
    cache_parts = []
    rate = total.get("cache_rate")
    if rate is not None:
        if rate >= 0.85:
            verdict = "上下文复用充分，这也是 token 量很大但积分并不高的主因"
        elif rate >= 0.6:
            verdict = "中等，频繁换模型或长间隔闲置会拉低它"
        else:
            verdict = "偏低，未缓存输入按全价计费，这是最直接的降本点"
        cache_parts.append("缓存命中率 %.1f%%，%s" % (rate * 100, verdict))

        if total.get("tokens"):
            cache_parts.append(
                "累计 %s token 中未缓存 %s（%.1f%%，只有这部分按全价计费），"
                "折合每百万 token 约 %.2f 积分"
                % ("{:,}".format(total["tokens"]),
                   "{:,}".format(total["cache_miss"]),
                   total["cache_miss"] / float(total["tokens"]) * 100,
                   total.get("credit_per_1m", 0))
            )

        # 全局均值会掩盖单日塌陷。这里按「未缓存 token 集中度」找异常日，
        # 因为用量大的日子未缓存量天然多，但只占小比例就不是异常。
        if total.get("cache_miss"):
            ranked = sorted(
                ((d, b) for d, b in by_day.items() if b["cache_miss"] > 0),
                key=lambda kv: -kv[1]["cache_miss"],
            )
            if ranked:
                day, block = ranked[0]
                share = block["cache_miss"] / float(total["cache_miss"])
                low = (block["cache_rate"] is not None
                       and block["cache_rate"] < CACHE_ALERT_RATE)
                if share >= CACHE_MISS_SHARE_ALERT and low:
                    note = ""
                    if block["requests"] == 1:
                        note = "，当天只有 1 次请求，属新上下文首轮全量未命中，本身正常"
                    cache_parts.append(
                        "%s 一天占了 %.0f%% 的未缓存量（%s tok，命中率 %.1f%%%s）"
                        % (day, share * 100, "{:,}".format(block["cache_miss"]),
                           block["cache_rate"] * 100, note)
                    )
    if cache_parts:
        out.append("；".join(cache_parts) + "。")

    # 费用不进结论区：官方价预估单独成表，结论里再复述一遍只是把同一组数字
    # 说两遍。表格里看得更细（全高峰价、覆盖率、逐模型单价），结论留白。

    # ── 第 4 组：结构（积分集中在哪个模型、哪个会话）──
    # 近 7 天口径，跟网页「按模型」「按会话」表用同一个分母。若还按「全部」算，
    # 结论说 65%、表里按 7 天算出来是 43%，同一份报表两个百分比打架。
    focus_model = by_model_7 if by_model_7 else by_model
    week_credit = (week or {}).get("credit") or 0
    struct_parts = []
    if focus_model and week_credit:
        model, block = next(iter(focus_model.items()))
        share = block["credit"] / week_credit * 100
        if share >= 40:
            struct_parts.append(
                "近 7 日积分集中在 %s（%.2f 积分，占 %.0f%%）、%s token"
                % (model, block["credit"], share, "{:,}".format(block["tokens"]))
            )

    focus_session = by_session_7 if by_session_7 else by_session
    if focus_session and week_credit:
        top = focus_session[0]
        share = top["credit"] / week_credit * 100
        struct_parts.append(
            "最贵会话「%s」%.2f 积分（占 %.0f%%）、%d 轮、%s token"
            % (top["title"][:28], top["credit"], share, top["requests"],
               "{:,}".format(top["tokens"]))
        )
    if struct_parts:
        out.append("；".join(struct_parts) + "。")

    # ── 第 5 组：不在单会话主对话里的开销（子代理 / 免费模型）──
    hidden_parts = []
    if total.get("sub_requests"):
        hidden_parts.append(
            "子代理 %d 次请求、%s token、%.2f 积分"
            % (total["sub_requests"], "{:,}".format(total["sub_tokens"]),
               total["sub_credit"])
        )
    if total.get("free_requests"):
        hidden_parts.append(
            "免费/限免模型 %d 次请求、%s token 记 0 积分"
            % (total["free_requests"], "{:,}".format(total["free_tokens"]))
        )
    if hidden_parts:
        out.append("；".join(hidden_parts) + "。")

    return out[:MAX_INSIGHTS]


def aggregate(bundle, today_gate_mode="auto"):
    records_all = bundle["records"]   # 已裁到 max(基准日, 当天)
    # 裁过 --days 时，本期用裁剪后的记录，但环比必须回到未裁剪的记录上算：
    # 否则上期数据被裁掉了，环比会凭空变成暴跌。
    cmp_records = bundle.get("records_compare") or records_all
    sessions = bundle["sessions"]
    coverage = bundle["coverage"]
    web_all = bundle.get("web_calls") or []
    # 压缩事件在全量上识别，再按各窗口的 [start, until] 分桶
    comp_events = bundle.get("compactions") or []

    now = bundle["generated_at"]
    live_day = now.date()             # 系统当天，与基准日解耦
    today_str = live_day.strftime("%Y-%m-%d")
    anchor = bundle.get("anchor") or live_day
    anchor_str = anchor.strftime("%Y-%m-%d")
    shifted = anchor != live_day

    # 报表主体（7 日 / 30 日 / 全部 / 模型 / 会话 / 按天）一律截止基准日，
    # 只有「今日」这一档吃当天数据——否则「数据截至 10-23」的报表里，
    # 近 7 日会悄悄把 10-24 的量算进去。
    records = [r for r in records_all if not r["day"] or r["day"] <= anchor_str]
    live_records = [r for r in records_all if r["day"] == today_str]
    web_calls = [w for w in web_all if not w["day"] or w["day"] <= anchor_str]

    # 「今日」「昨日」是这张报表的两张单日卡，恒指**系统当天**及其前一天。
    # 基准日退到昨日（00:05 定时任务的默认口径）时，若还沿着基准日走，「今日」卡
    # 会被改名成「昨日」，用户就再也看不到今天花了多少——这正是「今日卡片不对」的
    # 根因；同理「昨日」卡也就跟着丢了。只有回溯历史某天（基准日早于今天一天以上）
    # 才让这两张卡跟着基准日走，否则「截至 10-05 的报表里塞一张 10-24 的卡」才叫错乱。
    live_today = (live_day - anchor).days <= 1
    today_start = live_day if live_today else anchor
    prev_day = today_start - timedelta(days=1)
    today_picked = live_records if live_today else _pick(records, today_start, None)

    summary = {}
    trimmed = bundle.get("trimmed_days")
    trim_start = bundle.get("trim_start")
    earliest = min((r["day"] for r in cmp_records if r["day"]), default=None)

    # 聚焦模式下，比聚焦范围更宽的窗口没有意义（数据已被裁掉，只剩重复行），
    # 展示层直接不出。否则「近 30 日」会显示成 7 天的数字，环比还会拿 7 天比 30 天。
    # 注意 summary 仍然四种窗口都算——「覆盖」「口径」等区块要用到 all。
    windows = []
    for key, label in WINDOW_LABELS:
        start = _window_start(anchor, key)
        if (trimmed and trim_start and key != "all"
                and start and start.strftime("%Y-%m-%d") < trim_start):
            continue
        windows.append((key, label))
    # 聚焦范围正好等于某个标准窗口时（如 --days 7 = 近 7 日），
    # 「聚焦」那行就是重复行，去掉，免得同一组数字印两遍。
    if trimmed and trim_start:
        focus_first = datetime.strptime(trim_start, "%Y-%m-%d").date()
        if any(_window_start(anchor, k) == focus_first for k, _ in windows if k != "all"):
            windows = [(k, l) for k, l in windows if k != "all"]
    # 「今日」「昨日」是这张报表的主语，任何过滤都不许把它们丢掉：没有数据就显示
    # 暂无，卡片和行必须留着，否则用户会以为报表根本没统计那一天。
    day_keys = ("today", "prev")
    present = {k for k, _ in windows}
    if not all(k in present for k in day_keys):
        windows = ([w for w in WINDOW_LABELS if w[0] in day_keys]
                   + [w for w in windows if w[0] not in day_keys])
    shown = {k for k, _ in windows}

    for key, label in WINDOW_LABELS:
        if key in ("today", "prev"):
            # 两张单日卡：起点即终点，锚在系统当天而不是基准日
            day = today_start if key == "today" else prev_day
            start = until = day
            # 昨日整天都在 records 里（prev_day ≤ 基准日恒成立），今日要走
            # records_all——基准日退到昨天时，今天的量本来就在 reports 之外。
            picked = (today_picked if key == "today"
                      else _pick(records_all, start, until))
        else:
            start = _window_start(anchor, key)
            picked = _pick(records, start, None)
            until = anchor
        block = _stat_block(picked)
        # 单日卡的标题：此刻确实邻近今天时才用「今日 / 昨日」，回溯历史时如实
        # 写「指定日 / 前一日」——否则「截至 10-05 的报表」里会凭空出现一个「昨日」。
        if key == "today":
            block["label"] = "今日" if live_today else _anchor_label(anchor, live_day)
        elif key == "prev":
            block["label"] = "昨日" if live_today else "前一日"
        else:
            block["label"] = label
        block["key"] = key
        block["since"] = start.strftime("%Y-%m-%d") if start else None
        block["until"] = until.strftime("%Y-%m-%d")
        block["shown"] = key in shown
        # 每个窗口按自己的 [start, until] 现筛，今日档的上界是当天、其余是基准日
        block["web"] = _web_block(web_all, start, until)
        block["compact"] = _compact_block(comp_events, start, until, block["input"])
        if key == "today":
            # 这一档就是「今天到现在为止」，永远没过完：要说清「截至几点」，
            # 否则「今日只有昨天的一半」会被当成用量下滑。
            block["partial"] = bool(live_today)
            block["as_of"] = now.strftime("%H:%M") if live_today else None
            block["live"] = bool(live_today)

        # 聚焦模式下，"全部"其实只是裁剪范围内的全部，标签必须如实反映
        if trimmed:
            block["focused"] = True
            if key == "all":
                known = sorted({r["day"] for r in records if r["day"]})
                block["label"] = "聚焦 %d 日" % trimmed
                block["since"] = known[0] if known else None

        # 环比：全部窗口没有「上期」，不硬造；被裁掉本期的窗口也不比，
        # 否则会变成「7 天的量 vs 30 天的量」这种不对等比较。
        rng = _prev_range(start, key) if (key in shown and start) else None
        if rng:
            prev_picked = _pick(cmp_records, rng[0], rng[1])
            prev_block = _stat_block(prev_picked)
            prev_block["since"] = rng[0].strftime("%Y-%m-%d")
            prev_block["until"] = rng[1].strftime("%Y-%m-%d")
            cmp_ = _compare(block, prev_block, PREV_LABELS[key])
            # 上期窗口若早于本机最早记录日，成交量天然被截断，环比会虚高。
            # 这里如实记录覆盖天数，渲染层据此打警告，而不是悄悄改数。
            span = (rng[1] - rng[0]).days + 1
            covered = span
            if earliest:
                first = datetime.strptime(earliest, "%Y-%m-%d").date()
                if first > rng[1]:
                    covered = 0
                elif first > rng[0]:
                    covered = (rng[1] - first).days + 1
            cmp_["prev_window_days"] = span
            cmp_["prev_covered_days"] = max(0, covered)
            cmp_["partial_prev"] = covered < span
            # 上期被截断时归因没有意义（所有项都会显示"下降"），直接不给
            cmp_["attribution"] = (
                None if cmp_["partial_prev"] else _attribute(picked, prev_picked, sessions)
            )
            block["compare"] = cmp_
        else:
            block["compare"] = None
        summary[key] = block

    # 今日档的「起步中」闸门：只标记，不改数值。渲染层看到 gated 就把环比与
    # 结论句收起来，其余窗口（含 7/30/全部里今日那部分）完全不受影响。
    gated, gate_reason = today_gate(
        now, summary.get("today") or {}, live_today, today_gate_mode)
    for key in ("today",):
        summary[key]["gated"] = bool(gated)
        summary[key]["gated_reason"] = gate_reason
    summary["today"]["gate_threshold"] = TODAY_MIN_CREDIT
    summary["today"]["gate_before_min"] = TODAY_GATE_BEFORE_MIN

    by_day = _by_day(records)
    by_model = _by_model(records)
    by_workspace = _by_workspace(records)
    by_hour = _by_hour(records)
    by_session = _by_session(records, sessions)
    turns = _by_turn(records)

    # 时段分布改用**近 30 天**口径。全部历史的小时分布会被早期的作息漂移糊平，
    # 而「几点在烧」问的是最近的习惯，不是简历。日历同理，见 render._calendar。
    hour_start = anchor - timedelta(days=29)
    hour_records = _pick(records, hour_start, None) if hour_start else records
    by_hour_30 = _by_hour(hour_records)
    hour_scope = {
        "since": hour_start.strftime("%Y-%m-%d"),
        "until": anchor_str,
        "requests": len(hour_records),
    }

    # 「按模型 / 按工作区 / 按会话 / 官方价预估」四块统一走**近 7 天**口径。
    # 全量口径对这些表几乎没用：历史越久，用户早就不关心的旧会话越占榜首，
    # 「这周谁在烧钱」反而被埋掉。这里复用现成的 _by_* 函数，只是换一个记录池，
    # 不需要另写一套聚合逻辑。
    week_start = _window_start(anchor, "7d")
    week_records = _pick(records, week_start, None) if week_start else records
    by_model_7 = _by_model(week_records)
    by_workspace_7 = _by_workspace(week_records)
    by_session_7 = _by_session(week_records, sessions)
    week_scope = {
        "since": week_start.strftime("%Y-%m-%d") if week_start else None,
        "until": anchor_str,
        "requests": len(week_records),
        "credit": _num(sum(r["credit"] for r in week_records)),
    }

    # 最贵的单轮请求：改成「今日 / 昨日 / 近 7 天 / 近 30 天 各取最贵一轮」。
    # 一张二十行的排行榜没人看完，四个时段各挑一个「最贵的」反而能直接定位
    # 到底是谁在吃什么。基准用**系统当天**而不是基准日——用户问的是「今天谁最贵」。
    turns_all = _by_turn(records_all)

    def _top_of(lo, hi):
        pool = [t for t in turns_all if t["day"] and lo <= t["day"] <= hi]
        return max(pool, key=lambda t: (t["credit"], t["tokens"])) if pool else None

    top_windows = []
    for key, label, back_from, back_to in (
        ("today", "今日", 0, 0),
        ("prev", "昨日", 1, 1),
        ("7d", "近 7 天", 6, 0),
        ("30d", "近 30 天", 29, 0),
    ):
        lo = live_day - timedelta(days=back_from)
        hi = live_day - timedelta(days=back_to)
        top_windows.append({
            "key": key,
            "label": label,
            "since": lo.strftime("%Y-%m-%d"),
            "until": hi.strftime("%Y-%m-%d"),
            "row": _top_of(lo.strftime("%Y-%m-%d"), hi.strftime("%Y-%m-%d")),
        })

    # 联网调用：只报次数，不折钱。WorkBuddy 侧没有公开的按次牌价，借用别家
    # （如 DeepSeek Search/Fetch 的 US$0.75/千次）等于凭空造账单，故此处无金额字段。
    # 口径跟随「官方价预估」表——它现在是近 7 天，这行也就跟着近 7 天。
    web_scope = (summary.get("7d") or {}).get("web") or {
        "total": 0, "search": 0, "fetch": 0, "days": 0,
    }

    return {
        "generated_at": now,
        "today": today_str,
        "anchor": anchor_str,
        "anchor_label": _anchor_label(anchor, now.date()),
        "anchor_shifted": shifted,
        "windows": windows,
        "trimmed_days": trimmed,
        "summary": summary,
        "web": dict(web_scope),
        "byDay": by_day,
        "byModel": by_model,
        "byWorkspace": by_workspace,
        "byHour": by_hour,
        # 近 30 天口径的小时分布（网页的「时段分布」用它，全量版保留给 JSON/Excel 消费方）
        "byHour30": by_hour_30,
        "hourScope": hour_scope,
        "bySession": by_session,
        # 近 7 天口径的三张表（网页上的按模型/按工作区/按会话与官方价预估都用它）
        "byModel7": by_model_7,
        "byWorkspace7": by_workspace_7,
        "bySession7": by_session_7,
        "weekScope": week_scope,
        "byTurn": turns,
        # 四段各取最贵一轮，渲染层按序展示；row 为 None 表示该时段没有请求
        "topWindows": top_windows,
        "insights": _insights(summary, by_day, by_model, by_session, coverage, turns,
                              windows, by_model_7, by_session_7),
        "coverage": coverage,
        "pricing": bundle.get("pricing") or {},
        "issues": bundle["issues"],
        "config_dir": bundle["config_dir"],
    }


def trim_bundle(bundle, days, anchor=None):
    """把取数结果裁到「截至基准日的最近 N 个**自然日**」，返回新 bundle。

    必须在聚合**之前**裁，而不是裁已算好的报表：窗口缩小后 KPI、模型排序、
    洞察都会跟着变；只裁明细会让「近 7 日」的总额仍然显示全量，那是错的。

    注意按自然日裁，不是按「有记录的 N 天」：后者在数据稀疏时会悄悄把 7 天
    变成 15 天，标签写着「聚焦 7 天」而表里是半个月，属于骗人。

    原记录另存在 ``records_compare`` 里，供环比使用——上期数据一旦被裁掉，
    环比就会凭空变成暴跌。
    """
    if not days or days <= 0:
        return bundle
    end = anchor or bundle.get("anchor")
    if end is None:
        return bundle
    start = end - timedelta(days=days - 1)
    start_str = start.strftime("%Y-%m-%d")
    end_str = end.strftime("%Y-%m-%d")
    # 上界同样要放行「系统当天」：基准日取昨日时，今天的量得留着给「今日」卡用，
    # 否则 --days 与 --anchor prev 一起用会把今日裁成空的。
    keep_str = max(end_str, bundle["generated_at"].strftime("%Y-%m-%d"))
    trimmed = dict(bundle)
    trimmed["records"] = [
        r for r in bundle["records"]
        if r["day"] and start_str <= r["day"] <= keep_str
    ]
    trimmed["web_calls"] = [
        w for w in (bundle.get("web_calls") or [])
        if w["day"] and start_str <= w["day"] <= keep_str
    ]
    # 压缩事件同样要裁：否则聚焦 7 天时，KPI 与表格都是 7 天，
    # 唯独「上下文压缩」还报着整段历史的次数，两个数字对不上。
    trimmed["compactions"] = [
        e for e in (bundle.get("compactions") or [])
        if e.get("day") and start_str <= e["day"] <= keep_str
    ]
    if not trimmed["records"]:
        return bundle
    trimmed["records_compare"] = bundle.get("records_compare") or bundle["records"]
    trimmed["trimmed_days"] = days
    trimmed["trim_start"] = start_str
    return trimmed
