"""取数层：把 WorkBuddy 本机账本拼成「逐请求」记录。

只读，不写任何原始数据。

三个数据源
----------
1. ``workbuddy.db`` 的 ``session_usage.credit_json``
   形如 ``{"<requestId>": 积分}``，是**轮次累计结算值**——一次用户提问若触发
   多轮工具调用，这里记的是整轮总和。这是积分的主账本。
2. ``projects/**/*.jsonl`` 的 ``providerData``
   提供 ``conversationRequestId`` / ``timestamp`` / ``model`` / ``cwd`` / ``rawUsage``。
   ``rawUsage.credit`` 是**单次 API 调用的末次值**，实测在重叠区间 100% 小于
   等于结算值，因此只在结算值缺失时用于补齐，绝不累加（否则重复计费）。
3. ``projects/**/subagents/*.jsonl``
   子代理日志。实测其 requestId 与主日志交集为 0，可直接并入不会双计。

去重规则
--------
同一 ``requestId`` 在 jsonl 里会有多行，但**不是重复写入**——实测数百条记录中
零重复的 ``(prompt_tokens, completion_tokens)`` 对，也没有跨文件重复。每一行
对应一次真实的 LLM 调用（工具调用循环里上下文逐次增长）。

因此 token 要 **求和**，得到的是「累计处理量」，也就是真正的用量与计费基础。
若取最大值，得到的只是「峰值上下文」，会低估十几倍（实测 sum/max ≈ 14.9）。

两种口径都保留：``input``/``output`` 是累计量，``peak_input``/``peak_output``
是峰值上下文，后者用于看单轮上下文规模，不参与用量统计。
"""

from __future__ import annotations

import glob
import hashlib
import json
import os
import sqlite3
from collections import defaultdict
from datetime import datetime, timedelta, timezone

import pricing

DEFAULT_CONFIG_DIR = "~/.workbuddy"
TZ = timezone(timedelta(hours=8))

# 命中率、覆盖率等统计所需的最小样本，低于此值不做趋势判断
MIN_TREND_SAMPLES = 3

# 联网类工具调用。官方对 Search / Fetch 按**次**计费，不是按 token，
# 所以这类费用进不了 token 单价模型，必须单独计数、单独交代。
WEB_TOOLS = ("WebSearch", "WebFetch")

# 上下文压缩的识别阈值：同一 requestId 内，某一轮的 prompt 比上一轮**跌掉 15%
# 以上**才算压缩。日志里没有 compact 事件类型，压缩唯一的痕迹就是这个回落。
#
# 15% 这个数不是拍的：实测全部回落都在 prompt 撞到 ~27 万的天花板之后发生，
# 幅度 42%~85%，而正常的对话轮次只会增长、偶尔有个位数百分比的抖动。
# 取 15% 既不会漏掉真实的压缩，也不会把抖动当压缩。
COMPACT_MIN_DROP = 0.15


def _ro_connect(db_path):
    """以只读方式打开 sqlite，避免任何意外写入。"""
    uri = "file:%s?mode=ro" % db_path.replace("\\", "/")
    return sqlite3.connect(uri, uri=True, timeout=5)


def _as_float(value, default=0.0):
    try:
        if value is None:
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def _as_int(value):
    try:
        if value is None:
            return 0
        return int(value)
    except (TypeError, ValueError):
        return 0


def workspace_name(cwd):
    """从 cwd 抽出可读的工作区名。Windows / POSIX 双兼容。"""
    if not cwd:
        return "(未知)"
    cleaned = str(cwd).replace("/", "\\").rstrip("\\")
    if not cleaned:
        return "(根目录)"
    tail = cleaned.split("\\")[-1]
    return tail or cleaned


def _load_sessions(config_dir):
    """读会话元数据 + 上下文占用。"""
    db_path = os.path.join(config_dir, "workbuddy.db")
    sessions = {}
    usage = {}
    issues = []

    if not os.path.isfile(db_path):
        return sessions, usage, ["未找到 workbuddy.db，积分维度不可用"]

    try:
        con = _ro_connect(db_path)
    except sqlite3.Error as exc:
        return sessions, usage, ["workbuddy.db 打不开：%s" % exc]

    try:
        tables = {
            row[0]
            for row in con.execute(
                "select name from sqlite_master where type='table'"
            )
        }
        if "sessions" in tables:
            cols = {row[1] for row in con.execute("pragma table_info(sessions)")}
            wanted = [
                c
                for c in (
                    "id",
                    "title",
                    "custom_title",
                    "model",
                    "cwd",
                    "updated_at",
                    "created_at",
                    "deleted_at",
                    "is_playground",
                    "context_window",
                )
                if c in cols
            ]
            if "id" in wanted:
                sql = "select %s from sessions" % ", ".join(wanted)
                for row in con.execute(sql):
                    rec = dict(zip(wanted, row))
                    sid = rec.get("id")
                    if not sid:
                        continue
                    title = rec.get("custom_title") or rec.get("title") or "(无标题)"
                    sessions[sid] = {
                        "id": sid,
                        "title": str(title),
                        "model": rec.get("model"),
                        "cwd": rec.get("cwd"),
                        "updated_at": rec.get("updated_at"),
                        "created_at": rec.get("created_at"),
                        "deleted": bool(rec.get("deleted_at")),
                        "playground": str(rec.get("is_playground")) == "1",
                        "context_window": _as_int(rec.get("context_window")),
                    }
        else:
            issues.append("sessions 表缺失，会话标题与工作区信息不全")

        if "session_usage" in tables:
            cols = {
                row[1] for row in con.execute("pragma table_info(session_usage)")
            }
            wanted = [
                c
                for c in ("session_id", "used", "size", "credit_json", "updated_at")
                if c in cols
            ]
            if "session_id" in wanted:
                sql = "select %s from session_usage" % ", ".join(wanted)
                for row in con.execute(sql):
                    rec = dict(zip(wanted, row))
                    sid = rec.get("session_id")
                    if not sid:
                        continue
                    usage[sid] = {
                        # 注意：used 是「当前上下文占用 token」，不是累计积分、
                        # 也不是累计 token。误当累计量会把整张报表算歪。
                        "used": _as_int(rec.get("used")),
                        "size": _as_int(rec.get("size")),
                        "credit_json": rec.get("credit_json"),
                        "updated_at": rec.get("updated_at"),
                    }
        else:
            issues.append("session_usage 表缺失，无法取积分")
    finally:
        con.close()

    return sessions, usage, issues


def _build_ledger(usage):
    """把 credit_json 摊平成 {requestId: 积分}，并记录每个会话的合计。"""
    ledger = {}
    session_credit = defaultdict(float)
    parsed = 0
    for sid, info in usage.items():
        raw = info.get("credit_json")
        if not raw:
            continue
        try:
            mapping = json.loads(raw)
        except (TypeError, ValueError):
            continue
        if not isinstance(mapping, dict):
            continue
        parsed += 1
        for req_id, value in mapping.items():
            if not req_id:
                continue
            amount = _as_float(value, None)
            if amount is None:
                continue
            ledger[req_id] = ledger.get(req_id, 0.0) + amount
            session_credit[sid] += amount
    return ledger, dict(session_credit), parsed


def _blank_record():
    return {
        "ts": None,
        "session_id": None,
        "model": None,
        "cwd": None,
        "is_sub": True,
        "raw_credit": None,
        # 累计量：多次 LLM 调用求和，代表真实处理量
        "input": 0,
        "output": 0,
        "cache_hit": 0,
        "cache_miss": 0,
        "thinking": 0,
        # 峰值上下文：单次调用的最大值，仅用于看上下文规模
        "peak_input": 0,
        "peak_output": 0,
        "calls": 0,
        # 逐次调用的 [时间戳, 本轮 prompt]，只有压缩识别用它。
        # 不能只留峰值：压缩的信号是「这一轮比上一轮小」，峰值把顺序抹掉了。
        "prompts": [],
    }


# ── 扫描缓存 ──────────────────────────────────────────────────────
# jsonl 只增不改，这是可以缓存的前提：历史文件解析一次后内容不会变，只有今天
# 那一两个文件需要重扫。缓存按**文件**切分，所以任何单个文件被追加，都只重算
# 它自己，不会让全部 200MB 陪绑。
#
# 用 JSON 而不是 pickle：缓存内容是本机账本的派生数据，用纯数据格式就不会有
# 反序列化执行代码的隐患，出问题也只会被当作一次未命中。
_SCAN_CACHE_VER = 2
_SCAN_CACHE_SUBDIR = os.path.join("usage-reports", ".cache")

# rawUsage 里同一含义的字段在不同模型下命名不同，按顺序取第一个存在的。
_RAW_USAGE_FIELDS = (
    ("input", ("prompt_tokens", "input_tokens")),
    ("output", ("completion_tokens", "output_tokens")),
    ("cache_hit", ("prompt_cache_hit_tokens", "cache_read_input_tokens")),
    ("cache_miss", ("prompt_cache_miss_tokens",)),
    ("thinking", ("completion_thinking_tokens",)),
)


def _cache_dir(config_dir):
    return os.path.abspath(os.path.join(config_dir, _SCAN_CACHE_SUBDIR))


def _cache_file(config_dir, path):
    digest = hashlib.sha1(os.path.abspath(path).encode("utf-8", "replace")).hexdigest()
    return os.path.join(_cache_dir(config_dir), digest + ".json")


def _stat_key(path):
    """(mtime_ns, size) 作为失效依据：任何一个变了就重扫这个文件。"""
    try:
        st = os.stat(path)
    except OSError:
        return None
    return [st.st_mtime_ns, st.st_size]


def _cache_load(config_dir, path, key, is_sub):
    """命中返回局部结果，否则 None。任何异常一律当未命中。"""
    cache_path = _cache_file(config_dir, path)
    try:
        with open(cache_path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        return None
    if not isinstance(data, dict):
        return None
    if data.get("v") != _SCAN_CACHE_VER or data.get("key") != key:
        return None
    if bool(data.get("is_sub")) != is_sub:
        return None
    recs = data.get("recs")
    if not isinstance(recs, dict):
        return None
    web = {}
    for item in data.get("web") or []:
        if not isinstance(item, (list, tuple)) or len(item) != 4:
            continue
        sid, cid, name, stamp = item
        web[(sid, cid)] = (name, stamp)
    return {
        "recs": recs,
        "web": web,
        "lines": int(data.get("lines") or 0),
        "bad": int(data.get("bad") or 0),
    }


def _cache_store(config_dir, path, key, is_sub, part):
    cache_path = _cache_file(config_dir, path)
    payload = {
        "v": _SCAN_CACHE_VER,
        "key": key,
        "is_sub": is_sub,
        "lines": part["lines"],
        "bad": part["bad"],
        "recs": part["recs"],
        "web": [[k[0], k[1], v[0], v[1]] for k, v in part["web"].items()],
    }
    tmp = cache_path + ".tmp"
    try:
        os.makedirs(os.path.dirname(cache_path), exist_ok=True)
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, separators=(",", ":"))
        os.replace(tmp, cache_path)
    except OSError:
        try:
            os.remove(tmp)
        except OSError:
            pass


def _cache_sweep(config_dir, alive):
    """清掉来源已经消失的缓存，避免目录无限长。"""
    directory = _cache_dir(config_dir)
    try:
        names = os.listdir(directory)
    except OSError:
        return
    for name in names:
        if not name.endswith(".json"):
            continue
        full = os.path.join(directory, name)
        if full in alive:
            continue
        try:
            os.remove(full)
        except OSError:
            pass


def _scan_one(path, is_sub):
    """解析单个 jsonl，返回该文件自己的局部归并结果；打不开返回 None。

    逐行逻辑与旧版完全一致，只是把结果收在局部表里，方便按文件缓存与合并。
    """
    recs = {}
    web_calls = {}
    lines = 0
    bad = 0
    try:
        fh = open(path, "rb")
    except OSError:
        return None
    with fh:
        for raw in fh:
            lines += 1
            # 字节级预过滤：用 bytes 比对比先解码成 str 再比对便宜，未命中的
            # 行连解码都省掉。
            if b'"providerData"' not in raw:
                continue
            try:
                obj = json.loads(raw.decode("utf-8", "replace"))
            except (TypeError, ValueError):
                bad += 1
                continue
            if not isinstance(obj, dict):
                continue

            # 联网检索 / 网页抓取：官方按**次**计费（不是按 token），
            # 只在这里计数，不参与 token 求和，也不进积分口径。
            if obj.get("type") == "function_call" and obj.get("name") in WEB_TOOLS:
                call_key = (obj.get("sessionId"), obj.get("callId") or obj.get("id"))
                if call_key[1] and call_key not in web_calls:
                    stamp = obj.get("timestamp")
                    web_calls[call_key] = (
                        obj.get("name"),
                        stamp if isinstance(stamp, (int, float)) else None,
                    )

            provider = obj.get("providerData")
            if not isinstance(provider, dict):
                continue
            req_id = provider.get("conversationRequestId")
            if not req_id:
                continue

            rec = recs.get(req_id)
            if rec is None:
                rec = recs[req_id] = _blank_record()
                rec["is_sub"] = is_sub
            else:
                # 只有全部出现都在 subagents 目录下才算子代理
                rec["is_sub"] = rec["is_sub"] and is_sub

            ts = obj.get("timestamp")
            if isinstance(ts, (int, float)):
                rec["ts"] = ts if rec["ts"] is None else min(rec["ts"], ts)
            if not rec["session_id"] and obj.get("sessionId"):
                rec["session_id"] = obj["sessionId"]
            if not rec["cwd"] and obj.get("cwd"):
                rec["cwd"] = obj["cwd"]
            if not rec["model"] and provider.get("model"):
                rec["model"] = provider["model"]

            raw_usage = provider.get("rawUsage")
            if isinstance(raw_usage, dict) and raw_usage:
                credit = raw_usage.get("credit")
                amount = _as_float(credit, None)
                if amount is not None:
                    prev = rec["raw_credit"]
                    rec["raw_credit"] = amount if prev is None else max(prev, amount)
                # 每行都是一次独立调用，故求和得到累计处理量；
                # 同时保留峰值，用于观察上下文规模。
                values = {}
                for key, source in _RAW_USAGE_FIELDS:
                    value = 0
                    for name in source:
                        if raw_usage.get(name) is not None:
                            value = _as_int(raw_usage[name])
                            break
                    values[key] = value
                    rec[key] += value
                if values["input"] or values["output"]:
                    rec["calls"] += 1
                # 逐次 prompt 序列：压缩识别靠的是「这一轮比上一轮小」，
                # 只有按时间排好的序列才能看出来。没有时间戳的行跳过——
                # 排不了序的样本留着只会让序列乱掉。
                if values["input"] > 0 and isinstance(ts, (int, float)):
                    rec["prompts"].append([ts, values["input"]])
                rec["peak_input"] = max(rec["peak_input"], values["input"])
                rec["peak_output"] = max(rec["peak_output"], values["output"])

    return {"recs": recs, "web": web_calls, "lines": lines, "bad": bad}


def _merge_part(records, web_calls, part):
    """把一个文件的局部结果并进全局表。

    每个字段的合并规则与「跨全部文件逐行累积」严格等价：token 与调用次数求和、
    raw_credit 取最大、ts 取最小、session/cwd/model 取第一个非空、is_sub 取与。
    """
    for req_id, rec in part["recs"].items():
        merged = records.get(req_id)
        if merged is None:
            merged = records[req_id] = dict(rec)
            # dict() 是浅拷贝，prompts 会与 part 共用同一个 list；后面再 extend
            # 就会把 part 里的也改掉。缓存命中时 part 是现读的，改它没影响，
            # 但共享可变状态迟早出事，这里显式换成独立表。
            merged["prompts"] = list(rec.get("prompts") or [])
            continue
        merged["is_sub"] = merged["is_sub"] and rec["is_sub"]
        # 同一 requestId 理论上只落在一个文件里（实测 499 个全部如此），
        # 万一跨文件，就把两边的序列并起来，交给 _find_compactions 排序去重。
        if rec.get("prompts"):
            merged.setdefault("prompts", []).extend(rec["prompts"])
        if rec["ts"] is not None and (merged["ts"] is None or rec["ts"] < merged["ts"]):
            merged["ts"] = rec["ts"]
        for field in ("session_id", "cwd", "model"):
            if not merged[field] and rec[field]:
                merged[field] = rec[field]
        if rec["raw_credit"] is not None and (
                merged["raw_credit"] is None or rec["raw_credit"] > merged["raw_credit"]):
            merged["raw_credit"] = rec["raw_credit"]
        for field in ("input", "output", "cache_hit", "cache_miss", "thinking", "calls"):
            merged[field] += rec[field]
        for field in ("peak_input", "peak_output"):
            if rec[field] > merged[field]:
                merged[field] = rec[field]
    for key, value in part["web"].items():
        if key[1] and key not in web_calls:
            web_calls[key] = value


def _find_compactions(records):
    """识别上下文压缩，并估算它省下的输入 token。

    **日志里没有 compact 事件**——全量 9 种 ``type`` 里没有一种是它。压缩在数据
    里唯一的痕迹是：同一个 ``requestId`` 内，某一轮的 ``prompt_tokens`` 比上一轮
    突然大幅回落。因此这里只能按序列做启发式判定，阈值见 ``COMPACT_MIN_DROP``。

    省下多少：压缩把上下文砍掉 ``drop = 压缩前 - 压缩后`` 那么多，此后这一轮的
    每一次调用都会**少发 drop 个 token**（被砍掉的旧消息不会再传）。所以

        单次压缩省下 ≈ drop × 该 requestId 内压缩之后的剩余调用次数

    这是个**毛估算**，三点要说清：

    1. 假设「不压缩的话上下文会一直带着被砍掉的那部分」——真实情况下上下文还会
       继续增长，所以这个数偏**保守**；
    2. 压缩那次调用本身的消耗（prompt 约等于压缩前的值）**已经算在总 token 里**，
       不从这个数里扣，也不额外加；
    3. 省下的是**输入 token**，其中大部分本会命中缓存，故不能直接当积分节省看。

    返回按时间升序的事件表，供聚合层按窗口分桶。
    """
    events = []
    for req_id, rec in records.items():
        raw = rec.get("prompts") or ()
        if len(raw) < 2:
            continue
        # 排序 + 去重：跨文件并入时可能有同一 (ts, prompt) 的重复样本，
        # 不去掉会让「上一轮」变成自己，回落判不出来。
        series = sorted({(int(item[0]), int(item[1]))
                         for item in raw
                         if len(item) == 2 and item[1]})
        n = len(series)
        for i in range(1, n):
            prev, cur = series[i - 1][1], series[i][1]
            if prev <= 0 or cur >= prev * (1 - COMPACT_MIN_DROP):
                continue
            rest = n - 1 - i
            when = datetime.fromtimestamp(series[i][0] / 1000.0, TZ)
            events.append({
                "ts": series[i][0],
                "day": when.strftime("%Y-%m-%d"),
                "hour": when.hour,
                "session_id": rec.get("session_id"),
                "request_id": req_id,
                "before": prev,
                "after": cur,
                "drop": prev - cur,
                "rest": rest,
                "saved": (prev - cur) * rest,
            })
    events.sort(key=lambda e: e["ts"])
    return events


def _scan_jsonl(config_dir, use_cache=True):
    """扫全部会话日志，按 requestId 归并（token 求和）。

    返回 (records, files, total_lines, bad_lines, issues, web_list, cached_files)。
    """
    root = os.path.join(config_dir, "projects")
    records = {}
    files = []
    bad_lines = 0
    total_lines = 0
    web_calls = {}
    cached_files = 0

    if not os.path.isdir(root):
        return {}, [], 0, 0, ["未找到 projects 目录，token 维度不可用"], [], 0

    alive = set()
    paths = glob.glob(os.path.join(root, "**", "*.jsonl"), recursive=True)
    for path in paths:
        files.append(path)
        is_sub = (os.sep + "subagents" + os.sep) in path
        key = _stat_key(path) if use_cache else None
        part = None
        if key is not None:
            cache_path = _cache_file(config_dir, path)
            alive.add(cache_path)
            part = _cache_load(config_dir, path, key, is_sub)
            if part is not None:
                cached_files += 1
        if part is None:
            part = _scan_one(path, is_sub)
            if part is None:
                continue
            if key is not None:
                _cache_store(config_dir, path, key, is_sub, part)
        total_lines += part["lines"]
        bad_lines += part["bad"]
        _merge_part(records, web_calls, part)

    if use_cache and alive:
        # alive 为空说明这次一个文件都没扫到（目录临时不可读之类），
        # 这时候清理会把好缓存一起端掉，得不偿失——直接跳过。
        _cache_sweep(config_dir, alive)

    issues = []
    if bad_lines:
        issues.append("%d 行无法解析，已跳过" % bad_lines)
    if total_lines and bad_lines / float(total_lines) > 0.2:
        issues.append("跳过率超过 20%，数据可能不完整，慎用")

    web_list = []
    for name, stamp in web_calls.values():
        day = None
        if stamp:
            day = datetime.fromtimestamp(stamp / 1000.0, TZ).strftime("%Y-%m-%d")
        web_list.append({"name": name, "ts": stamp, "day": day})
    web_list.sort(key=lambda w: (w["day"] or "", w["ts"] or 0))
    return records, files, total_lines, bad_lines, issues, web_list, cached_files


def _classify_credit(ledger, records):
    """给每条记录定积分来源，并判断是否属于免费模型。

    优先级：结算值(ledger) > 末次值(rawUsage.credit) > 无记录。
    末次值只用于补齐，不与结算值相加。
    """
    for req_id, rec in records.items():
        if req_id in ledger:
            rec["credit"] = ledger[req_id]
            rec["credit_src"] = "ledger"
        elif rec["raw_credit"] is not None:
            rec["credit"] = rec["raw_credit"]
            rec["credit_src"] = "raw"
        else:
            rec["credit"] = 0.0
            rec["credit_src"] = "none"
        rec["is_free"] = rec["credit_src"] != "none" and rec["credit"] <= 0.0


def _has_tokens(rec):
    return (rec["input"] + rec["output"]) > 0


def build(config_dir=None, fx=None, with_cost=True, use_cache=True):
    """组装取数结果。返回 dict，供 aggregate 层消费。

    fx 只影响「官方仅发布美元牌价」的模型（DeepSeek-V4-Pro）折算人民币；
    有官方人民币牌价的模型（Flash、GLM）不受影响。报表口径一律人民币。
    with_cost=False 时不做任何价格折算（--no-cost），费用字段一律 0。
    use_cache=False 时忽略扫描缓存，全部日志重新解析（--no-cache）。
    """
    config_dir = os.path.expanduser(config_dir or DEFAULT_CONFIG_DIR)
    issues = []

    sessions, usage, sess_issues = _load_sessions(config_dir)
    issues.extend(sess_issues)

    ledger, session_credit, parsed_sessions = _build_ledger(usage)

    (raw_records, files, total_lines, bad_lines,
     jsonl_issues, web_rows, cached_files) = _scan_jsonl(config_dir, use_cache=use_cache)
    issues.extend(jsonl_issues)

    _classify_credit(ledger, raw_records)

    rows = []
    for req_id, rec in raw_records.items():
        sid = rec["session_id"]
        meta = sessions.get(sid) or {}
        ts = rec["ts"]
        when = datetime.fromtimestamp(ts / 1000.0, TZ) if ts else None
        rows.append(
            {
                "request_id": req_id,
                "session_id": sid,
                "title": meta.get("title") or "(会话已删除或无记录)",
                "model": rec["model"] or "(未知模型)",
                "cwd": rec["cwd"] or "",
                "workspace": workspace_name(rec["cwd"] or meta.get("cwd")),
                "ts": ts,
                "when": when,
                "day": when.strftime("%Y-%m-%d") if when else None,
                "hour": when.hour if when else None,
                "input": rec["input"],
                "output": rec["output"],
                "cache_hit": rec["cache_hit"],
                "cache_miss": rec["cache_miss"],
                "thinking": rec["thinking"],
                "peak_input": rec["peak_input"],
                "peak_output": rec["peak_output"],
                "calls": rec["calls"],
                "credit": round(rec["credit"], 4),
                "credit_src": rec["credit_src"],
                "is_sub": rec["is_sub"],
                "is_free": rec["is_free"],
                "has_tokens": _has_tokens(rec),
            }
        )

    rows.sort(key=lambda r: (r["ts"] or 0, r["request_id"]))

    # 费用预估：只给有官方公开牌价的模型算，其余明确不打价
    if with_cost:
        price_meta = pricing.annotate(rows, fx=fx)
    else:
        price_meta = {"fx": fx, "cards": [], "missing_holiday_years": [],
                      "disclaimer": "", "disabled": True}
        for rec in rows:
            rec["cost_cny"] = 0.0
            rec["cost_if_peak"] = 0.0
            rec["cost_tier"] = None
            rec["cost_priced"] = False
            rec["cost_card"] = None

    # 会话级补充：上下文占用与删除状态
    for sid, meta in sessions.items():
        info = usage.get(sid) or {}
        used = info.get("used") or 0
        size = info.get("size") or meta.get("context_window") or 0
        meta["context_used"] = used
        meta["context_size"] = size
        meta["context_rate"] = (used / float(size)) if size > 0 else None
        meta["credit"] = round(session_credit.get(sid, 0.0), 4)

    # 覆盖率统计——报表必须显式交代，不能静默丢数据
    total = len(rows)

    def _tok(items):
        return sum(r["input"] + r["output"] for r in items)

    priced = [r for r in rows if r.get("cost_priced")]
    unpriced = [r for r in rows if not r.get("cost_priced")]
    # 会话被删之后，`session_usage` 里的结算行会跟着消失，但 projects 下的 jsonl
    # 日志还在。结果就是这些会话**token 照算、积分算 0**——积分只会偏少，不会偏多。
    # 这不是 bug 而是数据源本身的残缺点，必须在报表里写明，否则「token 很多、
    # 积分很低」会被当成统计错误。
    #
    # 只认 `sessions.deleted_at` 这一个标记。**不能**用「session_id 不在 sessions
    # 表里」当删除判据：子代理会话本来就不进 sessions 表，那样会把 8 个活着的
    # 子代理会话误判成「已删除」，凭空造出一条不存在的偏差。
    lost = [
        r for r in rows
        if (sessions.get(r["session_id"]) or {}).get("deleted")
    ]
    coverage = {
        "requests": total,
        "with_credit": sum(1 for r in rows if r["credit_src"] != "none"),
        "ledger_credit": sum(1 for r in rows if r["credit_src"] == "ledger"),
        "raw_credit": sum(1 for r in rows if r["credit_src"] == "raw"),
        "no_credit": sum(1 for r in rows if r["credit_src"] == "none"),
        "free_models": sum(1 for r in rows if r["is_free"]),
        "with_tokens": sum(1 for r in rows if r["has_tokens"]),
        "no_tokens": sum(1 for r in rows if not r["has_tokens"]),
        "subagent_requests": sum(1 for r in rows if r["is_sub"]),
        "no_credit_tokens": sum(
            r["input"] + r["output"] for r in rows if r["credit_src"] == "none"
        ),
        "no_timestamp": sum(1 for r in rows if not r["ts"]),
        "llm_calls": sum(r["calls"] for r in rows),
        "peak_context": sum(r["peak_input"] + r["peak_output"] for r in rows),
        "sessions": len(sessions),
        "sessions_with_credit": parsed_sessions,
        "jsonl_files": len(files),
        "jsonl_lines": total_lines,
        "bad_lines": bad_lines,
        "active_days": len({r["day"] for r in rows if r["day"]}),
        # 费用预估覆盖情况：只报「有官方牌价的 token 占全量多少」，
        # 不把未定价模型的 token 折算进去，避免虚增
        "priced_requests": len(priced),
        "priced_tokens": _tok(priced),
        "unpriced_tokens": _tok(unpriced),
        "cost_cny": round(sum(r["cost_cny"] for r in rows), 4),
        "cost_cny_if_peak": round(sum(r["cost_if_peak"] for r in rows), 4),
        "peak_tokens": sum(
            r["input"] + r["output"] for r in rows if r.get("cost_tier") == "peak"
        ),
        # 联网调用与 token 无关，单独计数，供「未计入项」交代
        "web_calls": len(web_rows),
        "web_search": sum(1 for w in web_rows if w["name"] == "WebSearch"),
        "web_fetch": sum(1 for w in web_rows if w["name"] == "WebFetch"),
        # 已删除会话的口径偏差：token 在、积分没了
        "deleted_sessions": len({r["session_id"] for r in lost if r["session_id"]}),
        "deleted_tokens": _tok(lost),
        "deleted_requests": len(lost),
        "deleted_no_credit": sum(1 for r in lost if r["credit_src"] == "none"),
        "deleted_no_credit_tokens": sum(
            r["input"] + r["output"] for r in lost if r["credit_src"] == "none"
        ),
    }

    if coverage["no_credit"]:
        issues.append(
            "%d 条请求没有任何积分记录（涉及 %s token）"
            % (coverage["no_credit"], "{:,}".format(coverage["no_credit_tokens"]))
        )
    if coverage["no_timestamp"]:
        issues.append(
            "%d 条请求缺时间戳，不计入按天统计" % coverage["no_timestamp"]
        )
    if coverage["deleted_sessions"]:
        issues.append(
            "已删除会话 %d 个：其中 %d 条请求（%s token）只有 token、没有积分结算——"
            "偏差只有一个方向，积分偏低、不会偏高"
            % (coverage["deleted_sessions"], coverage["deleted_no_credit"],
               "{:,}".format(coverage["deleted_no_credit_tokens"]))
        )
    if not rows:
        issues.append("未取到任何请求记录，请检查 ~/.workbuddy 下是否有会话数据")

    if with_cost and coverage["priced_requests"] < coverage["requests"]:
        issues.append(
            "%d 条请求的模型无官方公开牌价，未计入费用预估"
            % (coverage["requests"] - coverage["priced_requests"])
        )

    return {
        "config_dir": config_dir,
        "records": rows,
        "web_calls": web_rows,
        "sessions": sessions,
        "coverage": coverage,
        "pricing": price_meta,
        "issues": issues,
        "generated_at": datetime.now(TZ),
        # 上下文压缩事件（含每条估算的节省量）。这里给的是**全量**，
        # 按窗口分桶交给聚合层——一张报表有五个窗口，取数层不该替它决定口径。
        "compactions": _find_compactions(raw_records),
        # 本次有多少个日志文件直接走了缓存（供 --verbose 观察，不进报表）
        "scan_cached": cached_files,
    }


if __name__ == "__main__":  # 便于单独调试取数层
    bundle = build()
    cov = bundle["coverage"]
    print("requests      :", cov["requests"])
    print("with credit   :", cov["with_credit"])
    print("  ledger      :", cov["ledger_credit"])
    print("  raw patch   :", cov["raw_credit"])
    print("no credit     :", cov["no_credit"])
    print("free models   :", cov["free_models"])
    print("subagent reqs :", cov["subagent_requests"])
    print("active days   :", cov["active_days"])
    print("issues        :", bundle["issues"])
