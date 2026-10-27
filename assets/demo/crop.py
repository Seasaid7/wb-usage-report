"""把整页 2 倍截图裁成两版候选图，并额外出一版「卡片」（圆角 + 描边）。

边界是**量出来的**：沿纵向扫一列像素，按颜色突变找面板起止，得到 1x 坐标后乘 2。
目测裁切的问题是留白忽多忽少、经常把表格切掉半行。
"""

import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "raw-2x.png")
OUT = os.path.join(HERE, "shots")

# 2x 版面区间（1x 值 ×2；1x 边界见 README 的边界表）
CUTS = {
    # 方案 A：导航 + 标题 + 五张 KPI + 完整环比表
    "shot-A-overview.png": (0, 1810),
    # 方案 B：结论区 + 每日趋势 + 活跃日历（含两张子卡）
    "shot-B-insight.png": (1880, 4450),
    # 方案 C：官方价预估 + 按模型（1x 2200–3200，h2 上留 37px、面板下留 26px）
    "shot-C-pricing.png": (4400, 6400),
}

# 卡片版：圆角 + 描边（2x 下 2px），圆角外透明——放在深浅两种主题的 README 上
# 都不会出现一块突兀的直角深色矩形
RADIUS = 22
BORDER = 2
# 描边要同时站在白底与深底上：半透中灰在白底读作浅灰线、在深底读作亮灰线，
# 一种颜色两边都能收边，不用出两份图
BORDER_RGBA = (150, 162, 180, 170)


def to_card(src, dst, radius=RADIUS, border=BORDER):
    im = Image.open(src).convert("RGBA")
    w, h = im.size

    # 先描边（不透明地压在图边上），再套圆角遮罩——顺序反了描边会被裁掉半根
    draw = ImageDraw.Draw(im)
    draw.rounded_rectangle(
        [border / 2.0, border / 2.0, w - 1 - border / 2.0, h - 1 - border / 2.0],
        radius=radius, outline=BORDER_RGBA, width=border)

    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, w - 1, h - 1], radius=radius, fill=255)
    im.putalpha(mask)
    im.save(dst)
    return im.size


def main():
    if not os.path.isfile(SRC):
        raise SystemExit("先截图：%s 不存在" % SRC)
    os.makedirs(OUT, exist_ok=True)
    im = Image.open(SRC)
    for name, (top, bottom) in CUTS.items():
        box = (0, top, im.size[0], bottom)
        raw = os.path.join(OUT, name)
        im.crop(box).save(raw)
        card_name = name.replace(".png", "-card.png")
        size = to_card(raw, os.path.join(OUT, card_name))
        print("%-28s %s  · 卡片版 %s" % (name, (im.size[0], bottom - top), size))


if __name__ == "__main__":
    main()
