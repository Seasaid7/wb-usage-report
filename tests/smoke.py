#!/usr/bin/env python3
"""wb-usage-report 自检。

只做本地检查，不联网、不写账本。分三层：

1. 静态层：解释器版本、六个模块能否导入、CLI 参数是否齐全
2. 结构层：xlsx.py 能不能写出合法 OOXML 包（用临时目录，不碰报表）
3. 端到端层：本机有账本时，真实跑一次出表并校验产物

用法：python tests/smoke.py
返回码 0 = 全通过，1 = 有失败。
"""
import html as htmllib
import io
import json
import os
import re
import struct
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, "scripts")
REPORT = os.path.join(SCRIPTS, "report.py")

sys.path.insert(0, SCRIPTS)

# 开放平台（open.workbuddy.cn）对技能包的硬要求
MARKET_REQUIRED = ("name", "display_name", "display_name_en", "description",
                   "description_zh", "description_en", "version", "author")
# 顶层只允许这些目录；tools 与 tests 仅本地开发用，打市场包时排除；
# assets 只给 GitHub README 用，两个包都不带
MARKET_ALLOWED = {"references", "scripts", "templates", "tools", "tests", "assets"}

passed = 0
failed = 0


def check(name, cond, detail=""):
    global passed, failed
    if cond:
        passed += 1
        print("  ok    %s" % name)
    else:
        failed += 1
        print("  FAIL  %s%s" % (name, ("  ← " + detail) if detail else ""))
    return bool(cond)


def section(title):
    print("\n[%s]" % title)


def main():
    section("静态层")
    check("Python >= 3.8", sys.version_info >= (3, 8),
          "current %s" % sys.version.split()[0])

    modules = ["collect", "aggregate", "pricing", "render", "xlsx"]
    for name in modules:
        try:
            __import__(name)
            check("import %s" % name, True)
        except Exception as exc:                    # noqa: BLE001
            check("import %s" % name, False, repr(exc))

    # 第三方依赖必须为零
    for name in modules:
        mod = sys.modules.get(name)
        if mod is None:
            continue
        external = [k for k in vars(mod)
                    if k not in sys.builtin_module_names
                    and k in ("openpyxl", "requests", "numpy", "pandas", "yaml")]
        check("%s 无第三方依赖" % name, not external, str(external))

    proc = subprocess.run([sys.executable, REPORT, "--version"],
                          capture_output=True, text=True)
    check("CLI --version", proc.returncode == 0 and "wb-usage-report" in proc.stdout,
          proc.stdout.strip() or proc.stderr.strip()[:120])

    proc = subprocess.run([sys.executable, REPORT, "--help"],
                          capture_output=True, text=True)
    flags = ["--config-dir", "--out", "--format", "--anchor", "--days",
             "--no-cost", "--no-cache", "--open", "--version"]
    missing = [f for f in flags if f not in proc.stdout]
    check("CLI 参数齐全", not missing, "missing %s" % missing)

    section("市场包")
    skill_md = os.path.join(ROOT, "SKILL.md")
    meta, err = _parse_frontmatter(skill_md)
    if meta is None:
        check("SKILL.md 有 frontmatter", False, err or "")
    else:
        check("SKILL.md 有 frontmatter", True)
        absent = [k for k in MARKET_REQUIRED if not meta.get(k)]
        check("市场必填字段齐全", not absent,
              "缺失 %s" % absent if absent else "")
        # YAML 裸字符串里的半角冒号+空格会把值截断成嵌套映射
        bad = [k for k, v in meta.items()
               if isinstance(v, str) and ": " in v]
        check("值内无半角冒号+空格", not bad, str(bad))
        check("版本号与 report.py 一致",
              meta.get("version") == _report_version(),
              "%s vs %s" % (meta.get("version"), _report_version()))
        # 开放平台的硬限制：每边最多 3 条，且中英必须一一对应。
        # 早期只校验「至少 3 条」，结果 zh 写了 4 条、en 写了 3 条，
        # 提交时被平台打回（examples_zh 超限 + 两边数量不一致）。改成双向卡。
        EX_MAX = 3
        counts = {}
        for key in ("examples_zh", "examples_en"):
            got = meta.get(key)
            n = len(got) if isinstance(got, list) else 0
            counts[key] = n
            check("%s 有 1~%d 条" % (key, EX_MAX), 1 <= n <= EX_MAX,
                  "got %d" % n)
        check("examples 中英数量一致",
              counts["examples_zh"] == counts["examples_en"],
              "zh %d vs en %d" % (counts["examples_zh"], counts["examples_en"]))

    extra = [d for d in os.listdir(ROOT)
             if os.path.isdir(os.path.join(ROOT, d))
             and d not in MARKET_ALLOWED and not d.startswith(".")]
    check("顶层无多余目录", not extra, "多余 %s" % extra)

    # README 的图与 assets/ 双向对齐：引用了不存在的是死链，存在但没引用的是死图
    readme = io.open(os.path.join(ROOT, "README.md"), encoding="utf-8").read()
    linked = sorted(set(re.findall(r"!\[[^\]]*\]\((assets/[^)\s]+)\)", readme)))
    check("README 引用了至少一张图", bool(linked), "got %s" % linked)
    missing = [p for p in linked
               if not os.path.isfile(os.path.join(ROOT, p))]
    check("README 图片路径都存在", not missing, "缺 %s" % missing)
    # 只比图，不比目录——assets/demo/ 放的是合成账本的生成脚本，本就不该被 README 引用
    orphan = sorted("assets/" + f for f in os.listdir(os.path.join(ROOT, "assets"))
                    if "assets/" + f not in linked
                    and os.path.isfile(os.path.join(ROOT, "assets", f))
                    and f.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".svg"))
                    ) if os.path.isdir(os.path.join(ROOT, "assets")) else []
    check("assets 里没有未被引用的图", not orphan, "多余 %s" % orphan)
    # 截图统一出 2 倍像素（2880 宽），否则 Retina 上会糊
    bad_w = []
    for p in linked:
        with open(os.path.join(ROOT, p), "rb") as fh:
            head = fh.read(24)
        if head[:8] != b"\x89PNG\r\n\x1a\n":
            bad_w.append(p + ":不是 PNG")
            continue
        w, h = struct.unpack(">II", head[16:24])
        if w != 2880:
            bad_w.append("%s:%dx%d" % (p, w, h))
    check("截图都是 2880 宽的 2 倍图", not bad_w, "异常 %s" % bad_w)

    section("脱敏层")
    # 判据挂在「模型名」上而不是「数值」上：模型名是有限的、稳定的十几个，
    # 而账单数字每天都在长——挂在数值上的清单必然漏（1.12.x 连漏三轮就是这么来的）。
    # 泄露的必要条件是「可识别的模型名 + 用量数值」同行出现；
    # 单有数字不构成泄露，单有模型名也不构成，两者同框才能对上某台机器的账。
    real_model = re.compile(
        r"(glm-|deepseek-|step-\d|hy\d|qwen|space-bunny|mimo-|kimi-|gpt-|claude-)", re.I)
    # 用量数值：两位以上，且后面既不跟数字也不跟 %（百分比是份额/阈值，不含绝对量级）。
    # 尾部的 (?![\d%]) 不能写成 (?!\s*%)——否则 "100%" 会因为回溯而把 "10" 判成数值。
    usage_num = re.compile(r"(?<![\w<.])\d{2,}(?:[.,]\d+)*(?![\d%])")
    usage_unit = re.compile(r"(积分|token|Token|请求数|会话数)")

    doc_files = [os.path.join(ROOT, "SKILL.md"), os.path.join(ROOT, "README.md")]
    ref_dir = os.path.join(ROOT, "references")
    if os.path.isdir(ref_dir):
        doc_files += [os.path.join(ref_dir, n) for n in sorted(os.listdir(ref_dir))
                      if n.endswith(".md")]
    leaks = []
    for path in doc_files:
        if not os.path.exists(path):
            continue
        for no, line in enumerate(io.open(path, encoding="utf-8"), 1):
            if real_model.search(line) and usage_unit.search(line) and usage_num.search(line):
                leaks.append("%s:%d" % (os.path.basename(path), no))
    check("模型名不与用量数值同行", not leaks, "疑似泄露 %s" % leaks)

    # data-model.md 是口径推导的主文档，最容易在举例时抄进真实值。
    # 判据**不能**依赖「同行有单位词」：占位符所在的表格行往往没有单位词
    # （如 `| A | ... | <nA> | <sumA> |`），那样写会漏——本轮先写成那样，
    # 反向验证时把占位符换成 360 竟然仍然 PASS，才发现是假断言。
    # 改为：三位以上数字一律视为用量，只放行四种非用量上下文——
    #   年份 20xx / 百分比 / 模型版本号后缀（-1017） / 公式系数（× 100）
    dm = os.path.join(ROOT, "references", "data-model.md")
    raw_usage = []
    if os.path.exists(dm):
        big = re.compile(r"(?<![\w<.])\d{3,}(?:[.,]\d+)*")
        for no, line in enumerate(io.open(dm, encoding="utf-8"), 1):
            for m in big.finditer(line):
                tok = m.group()
                if re.fullmatch(r"20\d\d", tok):
                    continue                              # 年份
                if line[m.end():m.end() + 1] == "%":
                    continue                              # 百分比
                if m.start() and line[m.start() - 1] == "-":
                    continue                              # 版本号后缀 -1017
                if line[max(0, m.start() - 2):m.start()].strip().endswith("×"):
                    continue                              # 公式系数 × 100
                raw_usage.append("%d:%s" % (no, tok))
    check("data-model.md 用量只用占位符", not raw_usage, "具体数字 %s" % raw_usage[:6])

    # 图片必须来自固定种子的合成账本，不是手工截图——手工截图没法证明里面是谁的数据
    maker = os.path.join(ROOT, "assets", "demo", "make_ledger.py")
    seeded = False
    if os.path.exists(maker):
        seeded = bool(re.search(r"^SEED\s*=\s*\d+", io.open(maker, encoding="utf-8").read(), re.M))
    check("演示图由固定种子合成账本生成", seeded, "缺 %s 或 SEED" % maker)

    section("结构层")
    import xlsx  # noqa: E402

    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "wb-usage-report-smoke.xlsx")
        sheets = [
            ("汇总", ["窗口", "积分", "token"],
             [["今日", 1.5, 12000], ["近 7 日", 9.25, 78000]]),
            ("逐请求", ["时间", "模型", "积分"],
             [["2026-01-01 00:00", "example-model", 0.5]]),
            # 故意塞入 Excel 非法字符与超长名，验证清洗与截断
            ("a/b:c*d?e[f]", ["x"], [["<&>\"'"]]),
        ]
        xlsx.write_workbook(path, sheets)
        check("写出工作簿", os.path.exists(path) and os.path.getsize(path) > 0)
        check("是合法 zip", zipfile.is_zipfile(path))

        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            need = ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                    "xl/styles.xml", "xl/worksheets/sheet1.xml",
                    "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"]
            absent = [n for n in need if n not in names]
            check("OOXML 部件齐全", not absent, "missing %s" % absent)

            bad = [n for n in names if n.endswith(".xml") and
                   not zf.read(n).decode("utf-8", "replace").startswith("<?xml")]
            check("所有 XML 部件可解析", not bad, str(bad))

            wb = zf.read("xl/workbook.xml").decode("utf-8")
            check("sheet 名已清洗", "a_b_c_d_e_f_" in wb)
            # Excel 硬限制 31 字符
            check("sheet 名不超 31 字符",
                  all(len(n) <= 31 for n in _sheet_names(wb)))

    section("端到端层")
    config_dir = os.path.expanduser("~/.workbuddy")
    if not os.path.isdir(os.path.join(config_dir, "projects")):
        print("  skip  本机无 projects/ 账本，跳过真实出表")
    else:
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "out")
            proc = subprocess.run(
                [sys.executable, REPORT, "--format", "all",
                 "--out", out, "--config-dir", config_dir],
                capture_output=True, text=True)
            if proc.returncode != 0:
                check("真实出表", False, proc.stderr.strip()[:200])
            else:
                check("真实出表", True)
                html = os.path.join(out, "index.html")
                data = os.path.join(out, "data.json")
                books = [f for f in os.listdir(out) if f.endswith(".xlsx")] \
                    if os.path.isdir(out) else []
                check("index.html 已生成", os.path.exists(html))
                check("data.json 已生成", os.path.exists(data))
                check("xlsx 已生成", bool(books))
                if os.path.exists(data):
                    with open(data, encoding="utf-8") as fh:
                        payload = json.load(fh)
                    check("data.json 有五档窗口",
                          all(k in payload.get("summary", {})
                              for k in ("today", "prev", "7d", "30d", "all")))
                    # 昨日档必须真的按「系统当天前一天」取，而不是把 today 复制一份
                    s = payload.get("summary", {})
                    check("昨日档锚在今日前一天",
                          bool(s.get("today", {}).get("since"))
                          and bool(s.get("prev", {}).get("since"))
                          and s["prev"]["since"] < s["today"]["since"])
                    check("全部档无环比且卡片不留占位文案",
                          s.get("all", {}).get("compare") is None
                          and os.path.exists(html)
                          and "无上期可比" not in io.open(
                              html, encoding="utf-8").read())
                    check("data.json 不含内嵌 base64",
                          "data_b64" not in payload)
                    # 上下文压缩：日志里没有压缩事件，只能靠 prompt 回落识别。
                    # 每个窗口都必须带这个块（哪怕是 0 次），否则渲染层会 KeyError；
                    # 省下量不能为负，也不能超过该窗口的输入 token。
                    comps = {k: (s.get(k, {}).get("compact") or {})
                             for k in ("today", "prev", "7d", "30d", "all")}
                    check("每个窗口都带压缩口径块",
                          all("count" in c and "saved" in c for c in comps.values()))
                    check("压缩省下量不超过该窗口输入",
                          all(c.get("saved", 0) >= 0
                              and c.get("saved", 0) <= (c.get("tokens_in") or 0) * 1.0
                              for c in comps.values() if c.get("count")))
                    # 已删会话：删会话带走结算行、留下 token，是唯一的单向偏差
                    cov = payload.get("coverage") or {}
                    check("已删会话口径字段齐备",
                          all(k in cov for k in (
                              "deleted_sessions", "deleted_tokens",
                              "deleted_requests", "deleted_no_credit"))
                          and cov.get("deleted_no_credit", 0)
                          <= cov.get("deleted_requests", 0))
                    # 今日档「起步中」闸门：00:30 之前、或今日积分不足门槛时，
                    # 今日档只给数字、不给环比——刚过零点的几积分对着整天的昨日
                    # 去比必然得出 -99% 量级，那不是用量塌陷，是还没开工。
                    agg_mod = sys.modules["aggregate"]
                    t = s.get("today", {})
                    check("今日档带闸门字段",
                          "gated" in t and "gated_reason" in t)
                    check("闸门阈值与常量一致",
                          t.get("gate_threshold") == agg_mod.TODAY_MIN_CREDIT
                          and t.get("gate_before_min") == agg_mod.TODAY_GATE_BEFORE_MIN)
                    import datetime as _dt
                    D = _dt.datetime(2030, 1, 1)
                    cases = [
                        (D.replace(hour=0, minute=29), 999.0, (True, "early")),
                        (D.replace(hour=0, minute=30), 4.99, (True, "thin")),
                        (D.replace(hour=0, minute=30), 5.00, (False, None)),
                        (D.replace(hour=12, minute=0), 5.00, (False, None)),
                        (D.replace(hour=12, minute=0), 4.99, (True, "thin")),
                    ]
                    got = [agg_mod.today_gate(n, {"credit": c}, True)
                           for n, c, _ in cases]
                    check("闸门边界 00:29/00:30 与 4.99/5.00",
                          got == [w for _, _, w in cases],
                          "%r != %r" % (got, [w for _, _, w in cases]))
                    check("闸门可整体关闭",
                          agg_mod.today_gate(D.replace(hour=0, minute=10),
                                             {"credit": 999.0}, True, "off")
                          == (False, None))
                    check("回溯历史日不设闸门",
                          agg_mod.today_gate(D.replace(hour=0, minute=10),
                                             {"credit": 0.0}, False)
                          == (False, None))
                    page = io.open(html, encoding="utf-8").read()
                    if t.get("gated"):
                        check("起步中：页面不给今日环比",
                              "样本不足，不算环比" in page,
                              "reason=%s credit=%s" % (t.get("gated_reason"), t.get("credit")))
                    else:
                        check("正常状态：今日照常给环比",
                              "样本不足，不算环比" not in page)
                    check("起步中时结论区不提今日",
                          (not t.get("gated"))
                          or not any(str(x).startswith("今日")
                                     for x in payload.get("insights") or []))
                    check("页脚有「今日起步中」定义", "今日起步中" in page)
                    # 限免期的模型两条通道都给 0，报表会标「免费」——这是对的，
                    # 但页脚必须说清「0 ≠ 永久免费」，否则很容易被读成漏账。
                    check("页脚解释 0 积分与限免期",
                          "0 积分 ≠ 永久免费" in page and "限免期" in page)
                    # 页脚文案有两条走 % 格式化、有两条不走，写错 %% 就会把
                    # 双百分号原样印到页面上（1.12.0 的「暂无 vs 0」就是这个错）。
                    check("页面没有 %% 残留", "%%" not in page)
                    # Agent 通道只读终端摘要，压缩这件事必须也出现在那里，
                    # 否则网页上有、对话里没有，两处口径对不上
                    has_comp = any(comps[k].get("count")
                                   for k in ("7d", "all"))
                    check("终端摘要同步压缩行",
                          ("  压缩    " in proc.stdout) == has_comp,
                          "stdout 有压缩行=%s / 数据有压缩=%s"
                          % ("  压缩    " in proc.stdout, has_comp))
                    # 时段分布必须走「近 30 天」：全量版 byHour 会让早期作息漂移
                    # 把最近的习惯糊平，而图上看不出这件事
                    scope = payload.get("hourScope") or {}
                    check("时段分布取近 30 天口径",
                          bool(payload.get("byHour30")) and bool(scope.get("since"))
                          and bool(scope.get("until"))
                          and (payload.get("today") or "") >= scope["until"])
                    # 日历按月分页：多页 + 默认停在最后一页（近 30 天）
                    if os.path.exists(html):
                        page = io.open(html, encoding="utf-8").read()
                        views = page.count('class="cal-view')
                        last = page.rfind('class="cal-view')
                        frag = page[last:last + 120] if last > 0 else ""
                        check("日历按月分页且默认停在近 30 天",
                              views >= 2
                              and 'class="cal-view is-on"' in frag
                              and 'data-cal-key="window"' in frag)
                        # 按模型彩蛋：只在头部模型认得出时出现，且必须严格取自
                        # 该模型的候选池——否则会串成「表里是 A、彩蛋按 B 说话」。
                        eggs = re.findall(
                            r'<p class="egg-line">([^<]*)</p>', page)
                        top = next(iter(payload.get("byModel7")
                                        or payload.get("byModel") or {}), "")
                        low = top.lower()
                        import render                      # noqa: PLC0415
                        pool = [ln for keys, lines in render.MODEL_EGGS
                                for k in keys if k in low for ln in lines]
                        if pool:
                            check("按模型彩蛋取自该模型候选池",
                                  len(eggs) == 1
                                  and htmllib.unescape(eggs[0]) in pool,
                                  "got %r / top=%s" % (eggs, top))
                        else:
                            check("无梗模型不出彩蛋", not eggs,
                                  "got %r / top=%s" % (eggs, top))
                        # 位置：彩蛋在「按模型」面板之外、下一节「按工作区」标题
                        # 之前——塞进面板里就会退化成表格脚注，挂在后面又会被
                        # 读成「按工作区」的开场白。
                        p_model = page.find('id="s-model"')
                        p_ws = page.find('id="s-ws"')
                        p_egg = page.find('class="egg-line"')
                        if p_egg < 0:
                            check("彩蛋落在两节留白处", True)   # 无彩蛋，不适用
                        else:
                            check("彩蛋落在按模型与按工作区之间的留白",
                                  p_model < p_egg < p_ws
                                  and '</section><p class="egg-line">' in page,
                                  "model=%d egg=%d ws=%d" % (p_model, p_egg, p_ws))
                        # 压缩条：有压缩就必须有那一行，而且次数要和 data.json 对得上；
                        # 一次都没压缩时必须整条不出现（不能留一行「0 次」占位置）
                        segs = page.count('class="comp-seg"')
                        hits = [comps[k] for k in ("7d", "all") if comps[k].get("count")]
                        if hits:
                            check("页面渲染了压缩条且次数与数据一致",
                                  1 <= segs <= len(hits)
                                  and ('发生 <b>%s</b> 次' % hits[-1]["count"]) in page,
                                  "segs=%d hits=%s" % (segs, [h["count"] for h in hits]))
                        else:
                            check("无压缩时不渲染压缩条",
                                  segs == 0 and 'class="comp-strip"' not in page)
                        check("页脚解释压缩计费与删会话偏差",
                              "上下文压缩" in page and "会话被删之后" in page)

    print("\n%d passed, %d failed" % (passed, failed))
    return 0 if failed == 0 else 1


def _parse_frontmatter(path):
    """极简 frontmatter 解析：只认 `key: value` 与 `  - item`。

    够用来校验市场必填项，不追求通用 YAML（不引第三方库）。
    """
    if not os.path.exists(path):
        return None, "SKILL.md 不存在"
    text = io.open(path, encoding="utf-8").read()
    if not text.startswith("---\n"):
        return None, "首行不是 ---"
    end = text.find("\n---\n", 3)
    if end < 0:
        return None, "缺少结束标记 ---"

    data, cur = {}, None
    for raw in text[4:end].split("\n"):
        if not raw.strip():
            continue
        if raw.lstrip().startswith("- "):
            item = raw.split("- ", 1)[1].strip().strip('"').strip("'")
            if cur and isinstance(data.get(cur), list):
                data[cur].append(item)
            continue
        if ":" not in raw:
            return None, "无法解析 %r" % raw[:40]
        key, _, val = raw.partition(":")
        cur = key.strip()
        val = val.strip().strip('"').strip("'")
        data[cur] = val if val else []
    return data, ""


def _report_version():
    for line in io.open(REPORT, encoding="utf-8"):
        if line.startswith("VERSION"):
            return line.split("=", 1)[1].strip().strip('"').strip("'")
    return None


def _sheet_names(workbook_xml):
    out, pos = [], 0
    while True:
        i = workbook_xml.find('name="', pos)
        if i < 0:
            return out
        i += 6
        j = workbook_xml.find('"', i)
        out.append(workbook_xml[i:j])
        pos = j + 1


if __name__ == "__main__":
    sys.exit(main())
