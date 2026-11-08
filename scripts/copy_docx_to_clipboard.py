"""Copy a DOCX article into the Windows HTML clipboard without opening Word."""

import argparse
import base64
from html import escape
from pathlib import Path
import time

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import win32clipboard


TEXT_STYLE = (
    "font-family:'Microsoft YaHei','微软雅黑',Arial,sans-serif;"
    "font-size:16px;line-height:1.75;letter-spacing:0;"
    "color:#252525;margin:0 0 16px 0;"
)


def paragraph_html(paragraph, doc) -> tuple[str, int]:
    bits = []
    pictures = 0
    for run in paragraph.runs:
        if run.text:
            value = escape(run.text)
            run_style = []
            if run.font.color.rgb:
                run_style.append(f"color:#{run.font.color.rgb};")
            if run.font.size:
                run_style.append(f"font-size:{run.font.size.pt:g}pt;")
            if run.bold:
                run_style.append("font-weight:700;")
            if run_style:
                value = f'<span style="{"".join(run_style)}">{value}</span>'
            bits.append(value)
        for blip in run._element.xpath(".//a:blip"):
            relation_id = blip.get(qn("r:embed"))
            if not relation_id:
                continue
            part = doc.part.related_parts[relation_id]
            encoded = base64.b64encode(part.blob).decode("ascii")
            bits.append(
                f'<img src="data:{part.content_type};base64,{encoded}" '
                'style="display:block;width:100%;max-width:100%;height:auto;'
                'margin:0 auto;border:0;" />'
            )
            pictures += 1

    content = "".join(bits) or "&nbsp;"
    alignment = "text-align:center;" if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER else ""
    if paragraph.style.name == "PreviewHeading":
        style = TEXT_STYLE + "text-align:center;white-space:nowrap;margin:12px 0 16px 0;"
    elif pictures:
        style = "margin:0 0 8px 0;" + alignment
    else:
        style = TEXT_STYLE + alignment
        if paragraph.paragraph_format.first_line_indent:
            style += "text-indent:2em;"
    return f'<p style="{style}">{content}</p>', pictures


def clipboard_html(fragment: str) -> bytes:
    prefix = b"<html><head><meta charset=\"utf-8\"></head><body><!--StartFragment-->"
    suffix = b"<!--EndFragment--></body></html>"
    body = prefix + fragment.encode("utf-8") + suffix
    header_template = (
        "Version:1.0\r\n"
        "StartHTML:{:010d}\r\n"
        "EndHTML:{:010d}\r\n"
        "StartFragment:{:010d}\r\n"
        "EndFragment:{:010d}\r\n"
    )
    header_size = len(header_template.format(0, 0, 0, 0).encode("ascii"))
    start_fragment = header_size + len(prefix)
    end_fragment = start_fragment + len(fragment.encode("utf-8"))
    header = header_template.format(header_size, header_size + len(body), start_fragment, end_fragment)
    return header.encode("ascii") + body


def copy_document(path: Path, expected_images: int | None = None) -> tuple[int, int]:
    doc = Document(path)
    html_parts = []
    image_count = 0
    for paragraph in doc.paragraphs:
        block, count = paragraph_html(paragraph, doc)
        html_parts.append(block)
        image_count += count
    if expected_images is not None and image_count != expected_images:
        raise ValueError(f"正文图片数不符：应为 {expected_images}，实际 {image_count}；未覆盖原剪贴板")
    html = clipboard_html("".join(html_parts))
    plain = "\n".join(paragraph.text for paragraph in doc.paragraphs)
    html_format = win32clipboard.RegisterClipboardFormat("HTML Format")
    for attempt in range(5):
        try:
            win32clipboard.OpenClipboard()
            break
        except win32clipboard.error:
            if attempt == 4:
                raise RuntimeError("剪贴板正被其他程序占用，请关闭占用程序后重试")
            time.sleep(0.2)
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(html_format, html)
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, plain)
    finally:
        win32clipboard.CloseClipboard()
    return image_count, len(html)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--expected-images", type=int)
    args = parser.parse_args()
    images, html_bytes = copy_document(args.docx, args.expected_images)
    print(f"已复制到 Windows 剪贴板：{images} 张图片，{html_bytes:,} 字节富文本。")
