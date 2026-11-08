"""Recreate the synthetic demo PDF and neutral placeholder banners."""
import json
from pathlib import Path
import pymupdf
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
(ROOT / "assets").mkdir(exist_ok=True)
source = ROOT / "examples/source"
source.mkdir(parents=True, exist_ok=True)
for filename, label in (("opening.png", "YOUR OPENING BANNER"), ("closing.png", "YOUR CLOSING BANNER")):
    canvas = Image.new("RGB", (1200, 360), "#e9f3ff")
    draw = ImageDraw.Draw(canvas)
    text_font = ImageFont.load_default(size=44)
    draw.text((600, 145), label, anchor="mm", font=text_font, fill="#123b66")
    draw.text((600, 220), "Replace this image with your own artwork", anchor="mm", font=ImageFont.load_default(size=25), fill="#37618c")
    canvas.save(ROOT / "assets" / filename)

pages = [
    ("学会整理书包", "分类放好物品，让准备更从容"),
    ("先分类，再检查", "把课本、文具和个人用品分类摆好。\n对照明天的课程表，检查需要携带的物品。"),
    ("给自己一个小行动", "今天放学后，独立整理一次书包。\n完成后对照清单自查，第二天分享经验。"),
]
pdf = pymupdf.open()
for title, subtitle in pages:
    page = pdf.new_page(width=960, height=540)
    page.draw_rect(page.rect, color=None, fill=(0.94, 0.97, 1))
    page.insert_textbox(pymupdf.Rect(280, 160, 680, 245), title, fontname="china-s", fontsize=42, align=1, color=(0.07, 0.23, 0.40))
    page.insert_textbox(pymupdf.Rect(290, 265, 670, 440), subtitle, fontname="china-s", fontsize=21, align=1, color=(0.2, 0.3, 0.4))
pdf.save(source / "学会整理书包.pdf")
pdf.close()
metadata = {
    "title": "小学生活习惯主题班会｜《学会整理书包》课件",
    "summary": "从分类物品到对照课程表检查，再通过一次独立整理练习，把准备工作落实到日常。配套3页示例课件。",
    "intro": "每天出门前找不到课本，或到了教室才发现忘带文具，往往与整理习惯有关。这节示例班会从整理书包这一日常任务切入，让学生先把课本、文具和个人用品分类摆好，再对照第二天的课程表逐项检查。教师可以请学生说说自己的摆放方法，帮助他们把整理步骤讲清楚。最后，学生选择在放学后独立整理一次书包，完成后对照清单自查，并在第二天分享经验。文章下方按实际顺序展示三页示例课件，方便查看整个活动的安排。",
    "resources": [{"label": "课件", "file": "学会整理书包.pdf"}],
}
(ROOT / "examples/metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
print("Created synthetic input and neutral placeholder banners")
