"""Build a WeChat courseware article package from a reviewed metadata JSON."""

import argparse
import json
from pathlib import Path
from shutil import copy2
from tempfile import TemporaryDirectory

import pymupdf
from PIL import Image, ImageOps
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


SKILL = Path(__file__).resolve().parents[1]
ASSETS = SKILL / "assets"


def read_metadata(path: Path, source: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    for key in ("title", "summary", "intro", "resources"):
        if not data.get(key):
            raise ValueError(f"metadata 缺少 {key}")
    if not isinstance(data["resources"], list):
        raise ValueError("resources 必须是列表")
    for key in ("title", "summary", "intro"):
        if not isinstance(data[key], str) or not data[key].strip():
            raise ValueError(f"{key} 必须为非空文字")
    if "\n" in data["intro"] or "下面按照上课顺序完整展示课件" in data["intro"]:
        raise ValueError("intro 须为单段导语，不重复填写固定展示提示")
    for item in data["resources"]:
        if not isinstance(item, dict) or not item.get("label") or not item.get("file"):
            raise ValueError("resources 每项须有 label 和 file")
        target = (source / item["file"]).resolve()
        if not target.is_relative_to(source.resolve()) or not target.is_file():
            raise FileNotFoundError(f"配套文件不存在或超出来源文件夹：{item['file']}")
    return data


def select_source(source: Path, explicit_pdf: Path | None, temp: Path) -> Path:
    if explicit_pdf:
        if not explicit_pdf.is_file() or explicit_pdf.suffix.lower() != ".pdf":
            raise FileNotFoundError(f"最终 PDF 不存在：{explicit_pdf}")
        return explicit_pdf
    pdfs = sorted(source.glob("*.pdf"))
    if len(pdfs) == 1:
        return pdfs[0]
    if len(pdfs) > 1:
        raise ValueError("发现多个 PDF；请用 --pdf 明确指定最终版")
    pptx = sorted(source.glob("*.pptx"))
    if len(pptx) != 1:
        raise ValueError("未找到唯一最终 PDF 或 PPTX；请整理来源文件夹或用 --pdf 指定")
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("只有 PPTX 时需要 Windows PowerPoint 和 pywin32 以导出 PDF") from exc
    app = win32com.client.DispatchEx("PowerPoint.Application")
    presentation = None
    destination = temp / "converted.pdf"
    try:
        presentation = app.Presentations.Open(str(pptx[0].resolve()), WithWindow=False)
        presentation.SaveAs(str(destination), 32)
    finally:
        if presentation is not None:
            presentation.Close()
        app.Quit()
    if not destination.is_file():
        raise RuntimeError("PowerPoint 未能导出 PDF")
    return destination


def validate_image(path: Path) -> None:
    with Image.open(path) as image:
        image.verify()


def render_pages(pdf: Path, pages_dir: Path) -> tuple[list[Path], Image.Image]:
    pages_dir.mkdir()
    page_paths = []
    first = None
    with pymupdf.open(pdf) as source:
        if source.page_count == 0:
            raise ValueError("最终 PDF 没有页面")
        digits = max(2, len(str(source.page_count)))
        for index, page in enumerate(source, 1):
            pix = page.get_pixmap(matrix=pymupdf.Matrix(1.25, 1.25), alpha=False)
            image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            if index == 1:
                first = image.copy()
            path = pages_dir / f"{index:0{digits}d}.jpg"
            image.save(path, quality=88, optimize=True)
            validate_image(path)
            page_paths.append(path)
    return page_paths, first


def make_cover(first: Image.Image, output: Path, selected: Path | None) -> None:
    if selected:
        validate_image(selected)
        with Image.open(selected) as image:
            cover = image.convert("RGB")
        ratio = cover.width / cover.height
        if not 2.30 <= ratio <= 2.40:
            raise ValueError(f"自备封面须接近 2.35:1；当前为 {ratio:.3f}:1")
    else:
        cover = ImageOps.fit(first, (1920, 817), method=Image.Resampling.LANCZOS, centering=(0.5, 0.5))
    cover.save(output, optimize=True)
    center = (cover.width - min(cover.width, cover.height)) // 2
    cover.crop((center, 0, center + cover.height, cover.height)).save(output.with_name("封面方形裁切检查.png"))


def set_font(run, size=None, bold=False, color=None) -> None:
    run.font.name = "Microsoft YaHei"
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "Microsoft YaHei")
    if size:
        run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def build_docx(output: Path, pages: list[Path], title: str, intro: str, resource_line: str) -> None:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor(37, 37, 37)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(12)

    def picture(path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(6)
        p.add_run().add_picture(str(path), width=Inches(6.5))

    def blank():
        p = doc.add_paragraph("\u00a0")
        p.paragraph_format.line_spacing = 1
        p.paragraph_format.space_after = Pt(0)

    picture(output / "开头引导图.png")
    blank()  # 名片留位
    blank()  # 组件上方空行
    heading = doc.add_paragraph()
    heading.style = doc.styles.add_style("PreviewHeading", 1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    heading.paragraph_format.space_after = Pt(0)
    for value, size, color, bold in (
        ("────  ", 13, (172, 212, 255), False),
        ("课件预览", 16, (18, 59, 102), True),
        ("  ────", 13, (172, 212, 255), False),
    ):
        run = heading.add_run(value)
        set_font(run, size, bold, color)
    blank()  # 组件下方空行
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Inches(0.24)
    set_font(p.add_run(intro))
    set_font(p.add_run(resource_line), bold=True, color=(12, 59, 112))
    set_font(p.add_run("下面按照上课顺序完整展示课件。"))
    for page in pages:
        picture(page)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    set_font(p.add_run(f"以上是最终PPT的完整预览，共{len(pages)}页"))
    blank()  # 名片留位
    picture(output / "结尾引导图.png")
    doc.core_properties.title = title
    doc.save(output / "公众号文章导入稿.docx")


def build(args) -> Path:
    source = args.source.resolve()
    output = args.output.resolve()
    if not source.is_dir():
        raise NotADirectoryError(source)
    if output.exists() and any(output.iterdir()):
        raise FileExistsError(f"输出文件夹非空，请使用新的输出位置：{output}")
    metadata = read_metadata(args.metadata, source)
    banners = {
        "开头引导图.png": args.opening_image or ASSETS / "opening.png",
        "结尾引导图.png": args.closing_image or ASSETS / "closing.png",
    }
    for path in banners.values():
        validate_image(path)
    if args.cover:
        validate_image(args.cover)
    with TemporaryDirectory() as tmp:
        pdf = select_source(source, args.pdf.resolve() if args.pdf else None, Path(tmp))
        with pymupdf.open(pdf) as checked:
            if checked.page_count < 1:
                raise ValueError("PDF 没有页面")
        output.mkdir(parents=True, exist_ok=True)
        pages, first = render_pages(pdf, output / "pages")
        for name, path in banners.items():
            # Normalize JPEG/other supported images to the declared PNG format.
            with Image.open(path) as banner:
                banner.convert("RGB").save(output / name)
        make_cover(first, output / "公众号封面图.png", args.cover)
    resource_line = "配套" + "＋".join(item["label"] for item in metadata["resources"]) + "，"
    build_docx(output, pages, metadata["title"], metadata["intro"], resource_line)
    (output / "公众号标题.txt").write_text(metadata["title"] + "\n", encoding="utf-8")
    (output / "发布信息.txt").write_text(
        f"标题：{metadata['title']}\n摘要：{metadata['summary']}\n封面图：公众号封面图.png\n"
        f"配套资料：{resource_line.rstrip('，')}\n课件页数：{len(pages)}\n", encoding="utf-8"
    )
    copy2(SKILL / "scripts/copy_docx_to_clipboard.py", output / "copy_docx_to_clipboard.py")
    (output / "一键复制到公众号.cmd").write_text(
        '@echo off\nchcp 65001 >nul\n'
        f'py -3 "%~dp0copy_docx_to_clipboard.py" "%~dp0公众号文章导入稿.docx" --expected-images {len(pages)+2}\n'
        'if errorlevel 1 pause\n', encoding="utf-8"
    )
    md = ["![开头引导图](开头引导图.png)", "", "&nbsp;", "", "&nbsp;", "", "## 课件预览", "", "&nbsp;", "",
          metadata["intro"] + f"<strong>{resource_line}</strong>下面按照上课顺序完整展示课件。", ""]
    for page in pages:
        md.extend([f"![课件第{int(page.stem)}页](pages/{page.name})", ""])
    md.extend([f"以上是最终PPT的完整预览，共{len(pages)}页", "", "&nbsp;", "", "![结尾引导图](结尾引导图.png)", ""])
    (output / "文章正文.md").write_text("\n".join(md), encoding="utf-8")
    (output / "使用说明.txt").write_text(
        "主要交付：公众号标题.txt、公众号封面图.png、公众号文章导入稿.docx。\n"
        "推荐：在公众号后台通过文档导入上传 DOCX，随后填写正式标题、摘要和封面，保存并重新打开检查。\n"
        "Obsidian：整个文件夹放入自己的库，打开文章正文.md；Markdown 用于编辑存储，固定排版以 DOCX 导入稿为准。\n"
        "一键复制为 Windows 辅助入口，嵌入图片可能被后台过滤，不能代替保存后检查。\n"
        "开头图下方和结尾图上方留给公众号名片。封面方形裁切检查.png 只用于验收。\n",
        encoding="utf-8",
    )
    (output / "生成检查.txt").write_text(
        f"来源文件夹：{source.name}\nPDF页数：{len(pages)}\n"
        f"连续页图：{pages[0].name} 至 {pages[-1].name}\n正文图片应为：{len(pages)+2} 张\n"
        "待人工检查：导语依据、封面标题和方形裁切、真实公众号草稿粘贴及手机预览。\n",
        encoding="utf-8",
    )
    return output


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True, help="单个课件产品文件夹")
    parser.add_argument("--metadata", type=Path, required=True, help="已审核的 JSON 内容配置")
    parser.add_argument("--output", type=Path, required=True, help="新的或空的交付文件夹")
    parser.add_argument("--pdf", type=Path, help="存在多个 PDF 时明确指定最终版")
    parser.add_argument("--cover", type=Path, help="已重排并审核的 2.35:1 封面；默认裁取第一页")
    parser.add_argument("--opening-image", type=Path, help="自选开头引导图；默认 assets/opening.png 示例占位图")
    parser.add_argument("--closing-image", type=Path, help="自选结尾引导图；默认 assets/closing.png 示例占位图")
    arguments = parser.parse_args()
    print(build(arguments))
