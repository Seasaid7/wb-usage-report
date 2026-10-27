"""价格预估层：把本机统计出的 token 量，按**官方公开牌价**折算成等价 API 费用。

这不是账单，是**参考值**。三个前提必须先说清楚：

1. 只对「有官方公开按量计费牌价」的模型估算。
   WorkBuddy 走的是订阅积分制，多数模型的积分口径由平台定义、且不公开
   对应的 token 单价；这类模型一律**不计价**，在报表里明确列为「不计入」，
   绝不用别的模型的单价去顶替——那会造出一个看着精确、实则虚构的数字。

2. 假设「用量 = 你直接调官方 API 的用量」。
   你的实际 token 由本机日志统计，若平台在中间做了截断、重试、改写或
   路由到别的版本，等值费用会跟着偏。

3. 峰谷价按**每条请求的时间戳**判定，不按天均价。
   DeepSeek 的规则是：周一至周五（不含中国法定节假日）北京时间
   09:00–12:00、14:00–18:00 为高峰，其余（含周末与法定节假日全天）为空闲，
   空闲价恰为高峰价的一半。00:00 之后跑的定时任务，正好把大部分用量
   落在空闲价上，所以这一层算出来的钱明显比「全按高峰价」低——差额本身
   就是一条可执行的成本结论。

牌价来源见每张卡片的 source / retrieved 字段，全部是官方页面原文。
价格会变，`--fx` 与价目表都刻意写成可改的常量，不硬编码进渲染层。
"""

from __future__ import annotations

from datetime import date, datetime, timedelta

# 汇率只用于「官方只发了美元价」的模型折算人民币。
# DeepSeek 官方同时给了美元与人民币牌价，两者比例约 6.67；
# 这里默认按 7.10 折算（偏保守，宁可高估），可用 --fx 覆盖。
DEFAULT_FX_CNY = 7.10

PER = 1_000_000.0

# ── 中国法定节假日（国务院办公厅通知）─────────────────────────────
# 峰谷判定要用：法定节假日**全天按空闲价**，即使它落在周一至周五。
# 2026 年表来自 国办发明电〔2025〕（2025-11-04 发布）。
# 注意：DeepSeek 原文写的是「周一至周五，不含中国法定节假日」，
# 因此调休上班的周六/周日仍按周末（空闲价）处理——这是照原文执行，
# 不是漏判。缺表的年份会退回「不扣除节假日」，并在 meta 里报缺口。
CN_HOLIDAYS = {
    2026: [
        ("2026-01-01", "2026-01-03"),  # 元旦
        ("2026-02-15", "2026-02-23"),  # 春节
        ("2026-04-04", "2026-04-06"),  # 清明
        ("2026-05-01", "2026-05-05"),  # 劳动节
        ("2026-06-19", "2026-06-21"),  # 端午
        ("2026-09-25", "2026-09-27"),  # 中秋
        ("2026-10-01", "2026-10-07"),  # 国庆
    ],
}

HOLIDAY_SOURCE = (
    "国务院办公厅关于 2026 年部分节假日安排的通知 "
    "(gov.cn/yaowen/liebiao/202511/content_7047099.htm)"
)

# 高峰时段（北京时间，左闭右开）
PEAK_WINDOWS = ((9, 12), (14, 18))

_HOLIDAY_CACHE = {}


def _holidays(year):
    """展开某年的节假日日期集合；无表返回 None（调用方据此报缺口）。"""
    if year not in _HOLIDAY_CACHE:
        ranges = CN_HOLIDAYS.get(year)
        if not ranges:
            _HOLIDAY_CACHE[year] = None
        else:
            days = set()
            for lo, hi in ranges:
                cur = datetime.strptime(lo, "%Y-%m-%d").date()
                end = datetime.strptime(hi, "%Y-%m-%d").date()
                while cur <= end:
                    days.add(cur)
                    cur += timedelta(days=1)
            _HOLIDAY_CACHE[year] = days
    return _HOLIDAY_CACHE[year]


# ── 价目表 ────────────────────────────────────────────────────────
# rates 单位：每百万 token，币种见 currency。
#   hit  = 输入·缓存命中
#   miss = 输入·缓存未命中
#   out  = 输出
# peak_model=True 表示该模型有峰谷两档；否则 flat 单档。
RATE_CARDS = [
    {
        "id": "deepseek-v4.1-flash",
        "match": [
            "deepseek-v4.1-flash",
            "deepseek-flash",
            "deepseek-v4-flash",
            "deepseek-v4-flash-vision-exp",
        ],
        "label": "DeepSeek-V4.1-Flash-1017",
        "vendor": "DeepSeek",
        "currency": "CNY",
        "peak_model": True,
        "rates": {
            "peak": {"hit": 0.04, "miss": 2.00, "out": 8.00},
            "offpeak": {"hit": 0.02, "miss": 1.00, "out": 4.00},
        },
        "source": "https://api-docs.deepseek.com/zh-cn/quick_start/pricing",
        "retrieved": "2026-10-24",
        "note": "人民币牌价为官方中文口径；空闲价 = 高峰价的一半",
    },
    {
        "id": "deepseek-v4-pro",
        "match": ["deepseek-v4-pro"],
        "label": "DeepSeek-V4-Pro-0813",
        "vendor": "DeepSeek",
        "currency": "CNY",
        "peak_model": True,
        # 官方**中文**价目页直接给人民币：高峰 命中 0.30 / 未命中 9.0 / 输出 27.0，
        # 空闲为高峰的一半。这里不再走美元 × 汇率——折算会带进 4% 左右的误差，
        # 而官方本来就有人民币牌价，没有理由再折一次。
        "rates": {
            "peak": {"hit": 0.30, "miss": 9.00, "out": 27.00},
            "offpeak": {"hit": 0.15, "miss": 4.50, "out": 13.50},
        },
        "source": "https://api-docs.deepseek.com/zh-cn/quick_start/pricing",
        "retrieved": "2026-10-24",
        "note": "人民币牌价取自官方中文页；空闲价 = 高峰价的一半",
    },
    {
        "id": "glm-5.3-flash",
        "match": ["glm-5.3-flash"],
        "label": "GLM-5.3-Flash",
        "vendor": "智谱 AI",
        "currency": "CNY",
        "peak_model": False,
        "rates": {"flat": {"hit": 0.23, "miss": 0.80, "out": 2.80}},
        "usd": None,
        "source": "https://bigmodel.cn/pricing?tab=api",
        "retrieved": "2026-10-24",
        "note": "统一价，不分峰谷；1M 上下文",
    },
]

# 联网检索 / 网页抓取：**刻意不在这里放单价。**
#
# 曾用过「US$0.75 / 千次成功请求」——该数字确有其出处，但出处在 DeepSeek API 文档的
# Search & Fetch Pricing 一节，讲的是「通过 DeepSeek Harness / Claude Code / Codex 调用
# DeepSeek API 时，服务端工具调用已经打到 DeepSeek 的 Search / Fetch 接口」这一场景。
# 本报表统计的是 WorkBuddy 本机日志里的 WebSearch / WebFetch 工具调用，走的是平台自己的
# 额度体系，与那个接口不是同一个计费主体；拿厂商 A 的按次价乘厂商 B 的调用量，等于凭空
# 造一个看着精确的账单。
#
# WorkBuddy 侧没有公开的「每次联网检索收多少钱」的牌价，所以这项只报次数、不报金额。
# 真到平台公布按次报价那天，再加一张卡，别去借别家的价。


# 订阅制 / 本地模型：明确不给价，避免用别的单价硬套。
UNPRICED_HINTS = [
    (("hy4-preview-f", "hy4-preview", "hy3"), "平台订阅内模型，无对外按量牌价"),
    (("step-5-preview", "space-bunny"), "平台订阅内模型，无对外按量牌价"),
    (("qwen",), "本地部署或量化模型，不走 API 计费"),
    (("(未知模型)",), "模型元数据缺失"),
]

DISCLAIMER = (
    "token 与积分均为本机日志统计值，仅供参考；"
    "费用预估按官方公开牌价换算，非实际账单"
)


def find_card(model):
    """按模型名匹配价目卡。大小写不敏感，兼容平台别名。"""
    if not model:
        return None
    name = str(model).strip().lower()
    for card in RATE_CARDS:
        for alias in card["match"]:
            if name == alias.lower():
                return card
    # 前缀兜底：deepseek-v4.1-flash-xxxx 这类带日期的版本号
    for card in RATE_CARDS:
        for alias in card["match"]:
            if name.startswith(alias.lower()):
                return card
    return None


def unpriced_reason(model):
    name = str(model or "").strip().lower()
    for keys, reason in UNPRICED_HINTS:
        for key in keys:
            if key in name:
                return reason
    return "无官方公开 token 牌价"


def is_peak(when):
    """按北京时间判定高峰档。

    when 为带时区的 datetime（+08:00）。返回 (bool, 是否缺节假日表)。
    """
    if when is None:
        return False, False
    day = when.date()
    if day.weekday() >= 5:  # 周六、周日
        return False, False
    holidays = _holidays(day.year)
    missing = holidays is None
    if holidays and day in holidays:
        return False, False
    for lo, hi in PEAK_WINDOWS:
        if lo <= when.hour < hi:
            return True, missing
    return False, missing


def tier_of(when, card):
    if not card.get("peak_model"):
        return "flat"
    peak, _ = is_peak(when)
    return "peak" if peak else "offpeak"


def split_input(rec):
    """拆出计费用的 (缓存命中, 缓存未命中, 计费输入量)。

    上游有时只报命中数、不报未命中期，差额按未命中（全价）计——宁可高估，
    也不要凭空给出一个偏低的价格让人误判成本。
    """
    inp = rec.get("input") or 0
    hit = rec.get("cache_hit") or 0
    miss = rec.get("cache_miss") or 0
    if hit or miss:
        covered = hit + miss
        if covered < inp:
            miss += inp - covered
        return hit, miss, hit + miss
    return 0, inp, inp


def _cost(hit, miss, out, rate):
    return (hit * rate["hit"] + miss * rate["miss"] + out * rate["out"]) / PER


def record_cost(rec):
    """单条请求的等值费用（人民币）。返回 dict，未定价时 priced=False。"""
    card = find_card(rec.get("model"))
    if card is None:
        return {"priced": False, "cny": 0.0, "tier": None}
    tier = tier_of(rec.get("when"), card)
    hit, miss, _ = split_input(rec)
    out = rec.get("output") or 0
    rate = card["rates"][tier]
    cny = _cost(hit, miss, out, rate)
    peak_rate = card["rates"].get("peak") or rate
    cny_if_peak = _cost(hit, miss, out, peak_rate) if card.get("peak_model") else cny
    return {
        "priced": True,
        "cny": cny,
        "cny_if_peak": cny_if_peak,
        "tier": tier,
        "card": card["id"],
    }


def annotate(rows, fx=None):
    """给每条记录挂上费用字段，并返回汇总用的 meta。就地修改 rows。

    ``fx`` 现在只是个兼容参数：所有模型都有官方人民币牌价，不再需要折算。
    留着它是为了不让 ``--fx`` 的旧命令直接报错，同时把实际汇率回填到 meta 里备查。
    """
    missing_years = set()
    for rec in rows:
        info = record_cost(rec)
        rec["cost_cny"] = round(info["cny"], 6)
        rec["cost_if_peak"] = round(info.get("cny_if_peak", info["cny"]), 6)
        rec["cost_tier"] = info["tier"]
        rec["cost_priced"] = info["priced"]
        rec["cost_card"] = info.get("card")
        if info["priced"] and rec.get("when") is not None:
            _, missing = is_peak(rec["when"])
            if missing:
                missing_years.add(rec["when"].year)
    return {
        "fx": fx or DEFAULT_FX_CNY,
        "cards": RATE_CARDS,
        "missing_holiday_years": sorted(missing_years),
        "disclaimer": DISCLAIMER,
        "holiday_source": HOLIDAY_SOURCE,
    }


def card_brief(card):
    """渲染用：把费率压成一行「命中 / 未命中 / 输出」，单位由表头交代。"""
    if card.get("peak_model"):
        r = card["rates"]
        return "高峰 %.2f / %.2f / %.2f · 空闲 %.2f / %.2f / %.2f" % (
            r["peak"]["hit"], r["peak"]["miss"], r["peak"]["out"],
            r["offpeak"]["hit"], r["offpeak"]["miss"], r["offpeak"]["out"],
        )
    r = card["rates"]["flat"]
    return "统一 %.2f / %.2f / %.2f" % (r["hit"], r["miss"], r["out"])


def peak_hours_text():
    return "工作日 09:00–12:00、14:00–18:00（北京时间）"


if __name__ == "__main__":  # 自查：价目卡与峰谷判定
    import sys

    sys.stdout.reconfigure(encoding="utf-8")
    tz = datetime(2026, 10, 24, 10, 30)
    for d in (
        datetime(2026, 10, 24, 10, 30),   # 周六 → 空闲
        datetime(2026, 10, 23, 10, 30),   # 周五上午 → 高峰
        datetime(2026, 10, 23, 13, 00),   # 周五午间 → 空闲
        datetime(2026, 10, 5, 10, 30),    # 周一但国庆 → 空闲
        datetime(2026, 10, 9, 10, 30),    # 周五 → 高峰
    ):
        print(d.strftime("%Y-%m-%d %a %H:%M"), "peak=", is_peak(d))
    for card in RATE_CARDS:
        print(card["id"], card_brief(card))
