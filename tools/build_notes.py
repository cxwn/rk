# -*- coding: utf-8 -*-
"""复习笔记统一生成器（方案 C：单一数据源，双格式输出）。

用法：
    python tools/build_notes.py <章节目录> [--html-only] [--docx-only]

章节目录约定：
    README.md   唯一数据源（Markdown 子集）
    figs/*.svg  插图源文件，README 中以 ![图注](figs/figNN.svg) 引用

产物（覆盖生成，按章节目录名命名）：
    <章节目录名>.html  A4 打印版（SVG 内联、单文件、微软雅黑）
    <章节目录名>.docx  A4 Word 版（SVG 经 PDF 链路渲染为 3 倍 PNG 嵌入）

依赖（建议虚拟环境）：
    pip install python-docx svglib reportlab "pymupdf==1.26.7"
    仅生成 HTML 时无需上述依赖（纯标准库）。

微软雅黑字体自动探测（msyh.ttc / msyhbd.ttc），也可用环境变量覆盖：
    MSYH_FONT=/path/msyh.ttc  MSYH_BOLD_FONT=/path/msyhbd.ttc

支持的 Markdown 子集：
    # / ## / ### 标题、**加粗**、`代码`、| 表格 |、- 与 1. 列表、> 引用、![图注](路径)
样式约定：
    以「举例/类比」开头的段落或引用 → 橙色实例卡片
    含「助记」的引用               → 绿色助记卡片
    其他引用                       → 浅蓝原文/要点卡片
"""
import html as html_mod
import io
import os
import re
import sys
import tempfile
import xml.etree.ElementTree as ET

# ---------------------------------------------------------------- 解析

def parse_markdown(text):
    lines = text.split("\n")
    blocks = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        m = re.match(r"^\s*!\[(.*?)\]\((.*?)\)\s*$", line)
        if m:  # 图片
            blocks.append(("fig", m.group(2).strip(), m.group(1).strip()))
            i += 1
            continue
        m = re.match(r"^(#{1,4})\s+(.*)$", line)
        if m:  # 标题
            blocks.append((f"h{min(len(m.group(1)), 3)}", m.group(2).strip()))
            i += 1
            continue
        if line.lstrip().startswith("|") and i + 1 < n and re.match(r"^\s*\|?[\s:|-]+\|?\s*$", lines[i + 1]) and "---" in lines[i + 1]:
            def split_row(s):
                return [c.strip() for c in s.strip().strip("|").split("|")]
            headers = split_row(line)
            rows = []
            i += 2
            while i < n and lines[i].lstrip().startswith("|"):
                rows.append(split_row(lines[i]))
                i += 1
            blocks.append(("table", headers, rows))
            continue
        if line.lstrip().startswith(">"):
            buf = []
            while i < n and lines[i].lstrip().startswith(">"):
                buf.append(lines[i].lstrip()[1:].strip())
                i += 1
            textq = " ".join(x for x in buf if x)
            kind = classify(textq, quote=True)
            blocks.append(("p", textq, kind))
            continue
        m = re.match(r"^\s*[-*]\s+(.*)$", line)
        if m:
            items = []
            while i < n:
                mm = re.match(r"^\s*[-*]\s+(.*)$", lines[i])
                if not mm:
                    break
                items.append(mm.group(1).strip())
                i += 1
            if len(items) == 1 and classify(items[0]) in ("example", "mnemonic"):
                blocks.append(("p", items[0], classify(items[0])))
            else:
                blocks.append(("list", False, items))
            continue
        m = re.match(r"^\s*\d+[.、]\s+(.*)$", line)
        if m:
            items = []
            while i < n:
                mm = re.match(r"^\s*\d+[.、]\s+(.*)$", lines[i])
                if not mm:
                    break
                items.append(mm.group(1).strip())
                i += 1
            if len(items) == 1 and classify(items[0]) in ("example", "mnemonic"):
                blocks.append(("p", items[0], classify(items[0])))
            else:
                blocks.append(("list", True, items))
            continue
        # 普通段落（单行即一段）
        blocks.append(("p", line.strip(), classify(line.strip())))
        i += 1
    return blocks


def classify(text, quote=False):
    plain = re.sub(r"^[\s>*]+", "", text)
    plain = re.sub(r"^\*\*", "", plain)
    if plain.startswith(("举例", "类比")):
        return "example"
    if "助记" in plain[:12]:
        return "mnemonic"
    if quote:
        return "quote"
    return "normal"


def split_bold(text):
    """拆为 [(片段, 是否加粗)]，并去掉行内代码反引号。"""
    parts = []
    for seg in re.split(r"(\*\*.+?\*\*)", text):
        if not seg:
            continue
        if seg.startswith("**") and seg.endswith("**"):
            parts.append((seg[2:-2].replace("`", ""), True))
        else:
            parts.append((seg.replace("`", ""), False))
    return parts


def strip_md(text):
    return "".join(seg for seg, _ in split_bold(text))


def disp_width(text):
    """估算文本显示宽度：CJK/全角字符按 2，其余按 1。"""
    return sum(2 if ord(ch) > 0x2E7F else 1 for ch in text)


def analyze_columns(headers, rows):
    """按内容适配表格：返回 (各列宽度比例, 各列是否居中)。

    列宽权重 = 该列最大显示宽度 + 8（留白），限制在 [12, 88]，
    长文本列更宽、短内容列更窄；最大显示宽度 ≤16（约 8 个汉字）
    的短内容列单元格居中，长文本列保持左对齐。
    """
    n = len(headers)
    maxw = []
    for j in range(n):
        w = disp_width(strip_md(headers[j]))
        for row in rows:
            if j < len(row):
                w = max(w, disp_width(strip_md(row[j])))
        maxw.append(w)
    weights = [min(max(w + 8, 12), 88) for w in maxw]
    total = sum(weights)
    ratios = [w / total for w in weights]
    center = [w <= 16 for w in maxw]
    return ratios, center


# ---------------------------------------------------------------- HTML

CSS = """
  :root{ --ink:#22303f; --main:#1f4e79; --main-light:#eaf1f8; --accent:#c05621;
    --accent-light:#fdf0e6; --line:#d5dde5; --muted:#6b7a8a; }
  *{ box-sizing:border-box; print-color-adjust:exact; -webkit-print-color-adjust:exact; }
  html,body{ margin:0; padding:0; }
  body{ font-family:"Microsoft YaHei","微软雅黑","PingFang SC","Noto Sans CJK SC",sans-serif;
    color:var(--ink); background:#e9edf1; font-size:10.5pt; line-height:1.75; }
  .sheet{ max-width:190mm; margin:8mm auto; background:#fff; padding:10mm 11mm;
    box-shadow:0 2px 14px rgba(0,0,0,.12); }
  h1{ font-size:21pt; color:var(--main); margin:0 0 2mm; border-bottom:3px solid var(--main); padding-bottom:3mm; }
  h2{ font-size:14pt; color:var(--main); margin:8mm 0 3mm; padding:1.2mm 0 1.2mm 3mm;
    border-left:5px solid var(--main); background:linear-gradient(to right,var(--main-light),#fff); break-after:avoid; }
  h3{ font-size:11.5pt; color:var(--main); margin:5mm 0 2mm; break-after:avoid; }
  p{ margin:1.5mm 0; }
  code{ background:#f0f3f6; border-radius:1mm; padding:0 1mm; font-size:9.5pt; }
  .def,.example,.mnemonic{ border-radius:2mm; padding:2.5mm 3.5mm; margin:2.5mm 0; break-inside:avoid; }
  .def{ border:1.5px solid var(--main); background:#fbfdff; }
  .example{ background:var(--accent-light); border-left:4px solid var(--accent); }
  .mnemonic{ background:#f3f6ee; border:1px dashed #8aa15c; color:#4a6130; }
  table{ border-collapse:collapse; width:100%; margin:2.5mm 0; font-size:9.5pt; table-layout:fixed; }
  th,td{ border:1px solid var(--line); padding:1.6mm 2mm; text-align:left; vertical-align:top; overflow-wrap:break-word; }
  th{ background:var(--main); color:#fff; text-align:center; }
  thead{ display:table-header-group; }
  tr{ break-inside:avoid; }
  td.tc{ text-align:center; }
  tr:nth-child(even) td{ background:#f6f9fc; }
  .fig{ margin:3mm 0; padding:2.5mm; border:1px solid var(--line); border-radius:2mm;
    background:#fdfefe; break-inside:avoid; }
  .fig .cap{ text-align:center; color:var(--muted); font-size:9pt; margin-top:1mm; }
  svg{ display:block; width:100%; height:auto; }
  svg text{ font-family:inherit; }
  ul,ol{ margin:1.5mm 0 1.5mm 5mm; padding-left:4mm; }
  li{ margin:0.8mm 0; }
  .footer{ margin-top:8mm; padding-top:2mm; border-top:1px solid var(--line);
    color:var(--muted); font-size:8.5pt; text-align:center; }
  @page{ size:A4; margin:12mm; }
  @media print{ body{ background:#fff; font-size:10pt; }
    .sheet{ max-width:none; margin:0; padding:0; box-shadow:none; } h2{ margin-top:6mm; } }
"""


def inline_rich(text):
    """**加粗** / `代码` / [链接](url) -> HTML（先转义）。"""
    out = []
    for seg, bold in split_bold(text):
        esc = html_mod.escape(seg)
        esc = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", r'<a href="\2" style="color:#1f4e79">\1</a>', esc)
        out.append(f"<b>{esc}</b>" if bold else esc)
    return "".join(out)


def svg_viewbox_width(svg_text):
    m = re.search(r'viewBox="[\d.]+\s+[\d.]+\s+([\d.]+)', svg_text)
    return float(m.group(1)) if m else 720.0


def render_html(blocks, chapter_dir, title):
    parts = [
        "<!DOCTYPE html>\n<html lang=\"zh-CN\">\n<head>\n<meta charset=\"UTF-8\">",
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">",
        f"<title>{html_mod.escape(title)}</title>",
        f"<style>{CSS}</style>\n</head>\n<body>\n<div class=\"sheet\">",
    ]
    for b in blocks:
        kind = b[0]
        if kind in ("h1", "h2", "h3"):
            parts.append(f"<{kind}>{inline_rich(b[1])}</{kind}>")
        elif kind == "p":
            cls = {"normal": "p", "example": "example", "mnemonic": "mnemonic", "quote": "def"}[b[2]]
            if cls == "p":
                parts.append(f"<p>{inline_rich(b[1])}</p>")
            else:
                parts.append(f"<div class=\"{cls}\">{inline_rich(b[1])}</div>")
        elif kind == "list":
            tag = "ol" if b[1] else "ul"
            items = "".join(f"<li>{inline_rich(t)}</li>" for t in b[2])
            parts.append(f"<{tag}>{items}</{tag}>")
        elif kind == "table":
            headers, rows = b[1], b[2]
            ratios, center = analyze_columns(headers, rows)
            cols = "".join(f'<col style="width:{r * 100:.1f}%">' for r in ratios)
            head = "".join(f"<th>{inline_rich(h)}</th>" for h in headers)
            body = []
            for row in rows:
                tds = []
                for j in range(len(headers)):
                    val = row[j] if j < len(row) else ""
                    cls = ' class="tc"' if center[j] else ""
                    tds.append(f"<td{cls}>{inline_rich(val)}</td>")
                body.append("<tr>" + "".join(tds) + "</tr>")
            parts.append(
                f"<table><colgroup>{cols}</colgroup>"
                f"<thead><tr>{head}</tr></thead><tbody>{''.join(body)}</tbody></table>"
            )
        elif kind == "fig":
            svg_path = os.path.join(chapter_dir, b[1])
            with open(svg_path, encoding="utf-8") as f:
                svg = f.read().strip()
            w = svg_viewbox_width(svg)
            style = "" if w >= 500 else " style=\"max-width:105mm;margin-left:auto;margin-right:auto\""
            cap = html_mod.escape(b[2])
            parts.append(
                f"<div class=\"fig\"><div{style}>{svg}</div><div class=\"cap\">{cap}</div></div>"
            )
    parts.append(
        f"<div class=\"footer\">{html_mod.escape(title)} · A4 打印版 · 由 README.md 统一生成</div>"
    )
    parts.append("</div>\n</body>\n</html>\n")
    return "\n".join(parts)


# ---------------------------------------------------------------- DOCX

def find_fonts():
    regular = os.environ.get("MSYH_FONT")
    bold = os.environ.get("MSYH_BOLD_FONT")
    candidates_r = ["/mnt/c/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/msyh.ttc"]
    candidates_b = ["/mnt/c/Windows/Fonts/msyhbd.ttc", "C:/Windows/Fonts/msyhbd.ttc"]
    for c in candidates_r:
        if regular is None and os.path.exists(c):
            regular = c
    for c in candidates_b:
        if bold is None and os.path.exists(c):
            bold = c
    if regular is None:
        raise RuntimeError("未找到微软雅黑字体，请设置 MSYH_FONT 环境变量")
    if bold is None:
        bold = regular
    return regular, bold


def svg_to_png(svg_path, tmpdir, cache={}):
    """svglib -> PDF（内嵌雅黑） -> pymupdf 3 倍 PNG。"""
    if svg_path in cache:
        return cache[svg_path]
    from reportlab.graphics import renderPDF
    from svglib.svglib import svg2rlg
    import pymupdf

    with open(svg_path, encoding="utf-8") as f:
        root = ET.fromstring(f.read())
    ns = "{http://www.w3.org/2000/svg}"
    for el in root.iter(f"{ns}text"):
        el.set("font-family", "Microsoft YaHei")
    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    drawing = svg2rlg(io.BytesIO(data))
    pdf_bytes = renderPDF.drawToString(drawing)
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    pix = doc[0].get_pixmap(matrix=pymupdf.Matrix(3, 3))
    out = os.path.join(tmpdir, os.path.basename(svg_path) + ".png")
    pix.save(out)
    doc.close()
    cache[svg_path] = out
    return out


def render_docx(blocks, chapter_dir, out_path):
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Mm, Pt, RGBColor

    MAIN = RGBColor(0x1F, 0x4E, 0x79)
    MUTED = RGBColor(0x6B, 0x7A, 0x8A)
    INK = RGBColor(0x22, 0x30, 0x3F)
    FONT = "微软雅黑"

    regular, boldfont = find_fonts()
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    if "Microsoft YaHei" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("Microsoft YaHei", regular, subfontIndex=0))
        pdfmetrics.registerFont(TTFont("Microsoft YaHei-Bold", boldfont, subfontIndex=0))
        pdfmetrics.registerFontFamily(
            "Microsoft YaHei", normal="Microsoft YaHei",
            bold="Microsoft YaHei-Bold", italic="Microsoft YaHei", boldItalic="Microsoft YaHei-Bold",
        )

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(18)
    sec.top_margin, sec.bottom_margin = Mm(16), Mm(16)
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    def style_run(r, size=10.5, bold=False, color=INK):
        r.font.name = FONT
        rPr = r._element.get_or_add_rPr()
        rf = rPr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rPr.append(rf)
        rf.set(qn("w:eastAsia"), FONT)
        r.font.size = Pt(size)
        r.bold = bold
        r.font.color.rgb = color

    def add_runs(p, text, size=10.5, bold=False, color=INK):
        for seg, is_bold in split_bold(text):
            seg = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", r"\1", seg)  # docx 中链接只保留显示文本
            r = p.add_run(seg)
            style_run(r, size, bold or is_bold, color)

    def para(text, size=10.5, bold=False, color=INK, align=None, after=4, before=0, shade=None):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.space_before = Pt(before)
        if align is not None:
            p.alignment = align
        if shade:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:fill"), shade)
            p._p.get_or_add_pPr().append(shd)
        add_runs(p, text, size, bold, color)
        return p

    def h1(text):
        p = para(text, size=20, bold=True, color=MAIN, after=6)
        pPr = p._p.get_or_add_pPr()
        pbdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "18")
        bottom.set(qn("w:color"), "1F4E79")
        pbdr.append(bottom)
        pPr.append(pbdr)

    def shade_cell(cell, hexcolor):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), hexcolor)
        tcPr.append(shd)

    tmpdir = tempfile.mkdtemp(prefix="notes_figs_")

    # 页脚：居中 x/y 页码（PAGE/NUMPAGES 域，当前页/总页数）
    def add_field(paragraph, instr):
        r = paragraph.add_run()
        style_run(r, 9, False, MUTED)
        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        instr_text = OxmlElement("w:instrText")
        instr_text.set(qn("xml:space"), "preserve")
        instr_text.text = instr
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end")
        r._r.append(fld_begin)
        r._r.append(instr_text)
        r._r.append(fld_end)

    footer_p = sec.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field(footer_p, "PAGE")
    sep_run = footer_p.add_run("/")
    style_run(sep_run, 9, False, MUTED)
    add_field(footer_p, "NUMPAGES")

    for b in blocks:
        kind = b[0]
        if kind == "h1":
            h1(b[1])
        elif kind == "h2":
            para(b[1], size=14, bold=True, color=MAIN, before=12, after=5)
        elif kind == "h3":
            para(b[1], size=11.5, bold=True, color=MAIN, before=7, after=3)
        elif kind == "p":
            style = b[2]
            shade = {"example": "FDF0E6", "mnemonic": "F3F6EE", "quote": "EAF1F8"}.get(style)
            para(b[1], shade=shade, after=5 if shade else 4)
        elif kind == "list":
            for idx, item in enumerate(b[2], 1):
                prefix = f"{idx}. " if b[1] else "• "
                para(prefix + item, after=2)
        elif kind == "table":
            headers, rows = b[1], b[2]
            ratios, center = analyze_columns(headers, rows)
            t = doc.add_table(rows=1 + len(rows), cols=len(headers))
            t.style = "Table Grid"
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            t.autofit = False
            widths = [Mm(174 * r) for r in ratios]  # A4 正文宽 174mm，按内容比例分配
            # 跨页表格在各页顶端重复表头行（Word「标题行重复」）
            trPr = t.rows[0]._tr.get_or_add_trPr()
            tbl_header = OxmlElement("w:tblHeader")
            tbl_header.set(qn("w:val"), "true")
            trPr.append(tbl_header)
            for j, htxt in enumerate(headers):
                cell = t.rows[0].cells[j]
                cell.width = widths[j]
                cell.text = ""
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_runs(cell.paragraphs[0], htxt, size=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
                shade_cell(cell, "1F4E79")
            for i, row in enumerate(rows):
                for j in range(len(headers)):
                    val = row[j] if j < len(row) else ""
                    cell = t.rows[i + 1].cells[j]
                    cell.width = widths[j]
                    cell.text = ""
                    if center[j]:
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    add_runs(cell.paragraphs[0], val, size=9.5)
                    if i % 2 == 1:
                        shade_cell(cell, "F6F9FC")
            para("", after=2, size=4)
        elif kind == "fig":
            svg_path = os.path.join(chapter_dir, b[1])
            with open(svg_path, encoding="utf-8") as f:
                w = svg_viewbox_width(f.read())
            png = svg_to_png(svg_path, tmpdir)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(png, width=Cm(16.5 if w >= 500 else 9.5))
            para(b[2], size=9, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, after=6)

    doc.save(out_path)


# ---------------------------------------------------------------- 入口

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    chapter_dir = sys.argv[1]
    md_path = os.path.join(chapter_dir, "README.md")
    with open(md_path, encoding="utf-8") as f:
        text = f.read()
    blocks = parse_markdown(text)
    title = strip_md(blocks[0][1]) if blocks and blocks[0][0] == "h1" else os.path.basename(chapter_dir)
    html_only = "--html-only" in sys.argv
    docx_only = "--docx-only" in sys.argv

    chapter_name = os.path.basename(os.path.normpath(chapter_dir))
    if not docx_only:
        out_html = os.path.join(chapter_dir, chapter_name + ".html")
        with open(out_html, "w", encoding="utf-8") as f:
            f.write(render_html(blocks, chapter_dir, title))
        print("HTML:", out_html)
    if not html_only:
        out_docx = os.path.join(chapter_dir, chapter_name + ".docx")
        render_docx(blocks, chapter_dir, out_docx)
        print("DOCX:", out_docx, os.path.getsize(out_docx), "bytes")
    figs = sum(1 for b in blocks if b[0] == "fig")
    print(f"解析块数: {len(blocks)}，插图: {figs} 张")


if __name__ == "__main__":
    main()
