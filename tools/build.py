"""构建两个分发包。

    python tools/build.py

产出（默认落在仓库的上一级目录，可用 OUT_DIR 覆盖）：
    wb-usage-report-<ver>.zip         完整包（GitHub 仓库 / 客户端本地导入）
    wb-usage-report-<ver>-market.zip  市场包（提交 open.workbuddy.cn 解析）

两个包唯一的差别是市场包裁掉了 tools/、tests/、CHANGELOG.md 与 .git*。
开放平台只认 SKILL.md / references / scripts / templates 四类，多出的目录有解析风险。

`assets/` 是两个包都不带的：那三张截图只服务于 GitHub 首页的说明，装技能的人
不需要为 1.2MB 的 PNG 付下载成本。
"""
import hashlib
import io
import os
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.dirname(HERE)                      # 仓库根 = tools/ 的上一级
PREFIX = os.path.basename(SRC)                   # zip 内的顶层目录名
OUT_DIR = os.environ.get("OUT_DIR") or os.path.dirname(SRC)

SKIP_ALWAYS = {".git", "__pycache__", ".pytest_cache"}
# 仅仓库展示用，不进任何分发包
DOC_ONLY = {"assets"}
# 市场包额外排除的顶层条目
MARKET_EXCLUDE = {"tools", "tests", "CHANGELOG.md", ".gitignore", ".gitattributes"}


def version():
    text = io.open(os.path.join(SRC, "SKILL.md"), encoding="utf-8").read()
    m = re.search(r'^version:\s*"?([0-9][^"\n]*)"?', text, re.M)
    if not m:
        raise SystemExit("SKILL.md 里没有 version 字段")
    return m.group(1).strip()


def collect(exclude=()):
    out = []
    for root, dirs, files in os.walk(SRC):
        dirs[:] = sorted(d for d in dirs if d not in SKIP_ALWAYS)
        rel_dir = os.path.relpath(root, SRC).replace(os.sep, "/")
        if rel_dir != "." and rel_dir.split("/")[0] in exclude:
            continue
        for f in sorted(files):
            if f.endswith(".pyc"):
                continue
            rel = f if rel_dir == "." else rel_dir + "/" + f
            if rel in exclude:
                continue
            out.append((os.path.join(root, f), PREFIX + "/" + rel))
    return out


def build(name, entries):
    with zipfile.ZipFile(name, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for src, arc in entries:
            zf.write(src, arc)
    raw = open(name, "rb").read()
    with zipfile.ZipFile(name) as zf:
        bad = zf.testzip()
        n = len(zf.namelist())
    print("%-34s %7d bytes  %2d files  md5 %s  %s"
          % (os.path.basename(name), len(raw), n,
             hashlib.md5(raw).hexdigest(),
             "integrity ok" if bad is None else "CORRUPT: " + bad))


def main():
    ver = version()
    print("version:", ver, "| src:", SRC)

    full = os.path.join(OUT_DIR, "wb-usage-report-%s.zip" % ver)
    market = os.path.join(OUT_DIR, "wb-usage-report-%s-market.zip" % ver)

    all_entries = collect(DOC_ONLY)
    build(full, all_entries)

    market_entries = collect(DOC_ONLY | MARKET_EXCLUDE)
    build(market, market_entries)

    # 市场包必须仍满足最小结构
    names = [arc for _, arc in market_entries]
    assert any(n.endswith("/SKILL.md") for n in names), "市场包缺 SKILL.md"
    assert any("/scripts/report.py" in n for n in names), "市场包缺 report.py"

    def tops(entries):
        return {arc.split("/", 1)[1].split("/")[0] for _, arc in entries}

    dropped = sorted(tops(all_entries) - tops(market_entries))
    print("市场包裁掉的顶层条目:", ", ".join(dropped) or "无")


if __name__ == "__main__":
    main()
