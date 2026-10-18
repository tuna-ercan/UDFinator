# -*- coding: utf-8 -*-
"""
UDFinator donusturme cekirdegi.

Desteklenen donusumler:
  * UDF  -> DOCX   (udf_to_docx)
  * UDF  -> PDF    (udf_to_pdf)
  * PDF  -> UDF    (pdf_to_udf)

Yaklasim:
  - Bir UDF (UYAP Dokuman Formati) dosyasi, icinde content.xml bulunan bir ZIP
    arsividir. Tum metin tek bir CDATA havuzunda tutulur; bicim parcalari bu
    havuza karakter tabanli, mutlak startOffset/length ile referans verir.
  - UDF once ortak bir belge modeline cozumlenir (_parse_udf), sonra DOCX ya da
    PDF olarak islenir. Boylece iki cikti da ayni icerigi kullanir.
  - PDF -> UDF, UDF-Toolkit'in (saidsurucu/UDF-Toolkit) yontemini bagimsiz olarak
    yeniden uygular: her sayfanin metni get_text ile alinir, gomulu resimler PNG
    olarak base64'lenir, U+FFFC yer tutucusuyla content.xml olusturulur.
"""

import base64
import io
import os
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as _xml_escape


class UdfError(Exception):
    """Girdi dosyasi okunamadiginda/cozumlenemediginde firlatilir."""


# ---------------------------------------------------------------------------
# Belge modeli
# ---------------------------------------------------------------------------

class Run:
    __slots__ = ("text", "family", "size", "bold", "italic", "underline", "color")

    def __init__(self, text, family=None, size=None, bold=False,
                 italic=False, underline=False, color=None):
        self.text = text
        self.family = family
        self.size = size
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.color = color  # 'RRGGBB' ya da None


class Para:
    __slots__ = ("alignment", "left_indent", "right_indent", "first_line_indent", "runs")

    def __init__(self):
        self.alignment = None  # 0 sol, 1 orta, 2 sag, 3 iki yana yasli
        self.left_indent = None
        self.right_indent = None
        self.first_line_indent = None
        self.runs = []


class Img:
    __slots__ = ("data", "width_pt", "height_pt")

    def __init__(self, data, width_pt=None, height_pt=None):
        self.data = data
        self.width_pt = width_pt
        self.height_pt = height_pt


class Tbl:
    __slots__ = ("rows", "col_count")

    def __init__(self, col_count):
        self.rows = []       # list[list[list[Para]]] -> satir -> hucre -> paragraflar
        self.col_count = col_count


# ---------------------------------------------------------------------------
# UDF cozumleme yardimcilari
# ---------------------------------------------------------------------------

_IMAGE_PLACEHOLDER = "\uFFFC"
_ZERO_WIDTH = "\u200B"


def _attr(el, *names, default=None):
    lowered = {k.lower(): v for k, v in el.attrib.items()}
    for name in names:
        if name in el.attrib:
            return el.attrib[name]
        if name.lower() in lowered:
            return lowered[name.lower()]
    return default


def _as_bool(value):
    return str(value).strip().lower() in ("true", "1", "yes")


def _as_float(value, default=None):
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return default


def _as_int(value, default=None):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


def _argb_to_hex(value):
    """Isaretli ARGB tam sayiyi 'RRGGBB' hex dizgesine cevirir."""
    num = _as_int(value)
    if num is None:
        return None
    num &= 0xFFFFFF
    return "%06X" % num


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _read_content_xml(udf_path):
    if zipfile.is_zipfile(udf_path):
        with zipfile.ZipFile(udf_path) as zf:
            name = next((n for n in zf.namelist()
                         if n.lower().endswith("content.xml")), None)
            if name is None:
                raise UdfError("UDF icinde content.xml bulunamadi.")
            return zf.read(name).decode("utf-8", errors="replace")
    with open(udf_path, "rb") as fh:
        raw = fh.read()
    if b"<template" in raw:
        return raw.decode("utf-8", errors="replace")
    raise UdfError("Dosya gecerli bir UDF arsivi degil.")


def _run_text(el, pool):
    start = _attr(el, "startOffset")
    length = _attr(el, "length")
    if start is not None and length is not None:
        s = _as_int(start)
        n = _as_int(length)
        if s is not None and n is not None:
            return pool[s:s + n]
        return ""
    return el.text or ""


def _make_run(el, pool):
    text = _run_text(el, pool).replace(_ZERO_WIDTH, "").replace(_IMAGE_PLACEHOLDER, "")
    run = Run(text)
    run.family = _attr(el, "family", "fontName")
    run.size = _as_float(_attr(el, "size", "fontSize"))
    if _attr(el, "bold") is not None:
        run.bold = _as_bool(_attr(el, "bold"))
    if _attr(el, "italic") is not None:
        run.italic = _as_bool(_attr(el, "italic"))
    if _attr(el, "underline") is not None:
        run.underline = _as_bool(_attr(el, "underline"))
    fg = _attr(el, "foreground")
    if fg is not None:
        run.color = _argb_to_hex(fg)
    return run


def _make_image(el):
    data = _attr(el, "imageData", "data")
    if not data:
        return None
    try:
        raw = base64.b64decode(data)
    except Exception:
        return None
    return Img(raw, _as_float(_attr(el, "width")), _as_float(_attr(el, "height")))


def _paragraph_props(para, el):
    align = _as_int(_attr(el, "Alignment"))
    if align is not None:
        para.alignment = align
    para.left_indent = _as_float(_attr(el, "LeftIndent"))
    para.right_indent = _as_float(_attr(el, "RightIndent"))
    para.first_line_indent = _as_float(_attr(el, "FirstLineIndent"))


def _collect_paragraph(el, pool):
    """Bir <paragraph> elemanini blok listesine cevirir (Para ve Img karisik)."""
    blocks = []
    current = Para()
    _paragraph_props(current, el)
    children = list(el)

    if not children:
        # Cocuk yoksa paragrafin kendi metin araligini dene
        run = _make_run(el, pool)
        if run.text:
            current.runs.append(run)
        blocks.append(current)
        return blocks

    for child in children:
        tag = _local(child.tag)
        if tag == "content":
            run = _make_run(child, pool)
            if run.text:
                current.runs.append(run)
        elif tag == "tab":
            current.runs.append(Run("\t"))
        elif tag == "space":
            current.runs.append(Run(" "))
        elif tag == "image":
            img = _make_image(child)
            if img is not None:
                if current.runs:
                    blocks.append(current)
                    current = Para()
                    _paragraph_props(current, el)
                blocks.append(img)
    blocks.append(current)
    return blocks


def _collect_table(el, pool):
    rows_el = [r for r in el if _local(r.tag) == "row"]
    if not rows_el:
        return None
    col_count = _as_int(_attr(el, "columnCount"))
    if not col_count:
        col_count = max((len([c for c in r if _local(c.tag) == "cell"])
                         for r in rows_el), default=1)
    col_count = max(col_count, 1)
    table = Tbl(col_count)
    for row_el in rows_el:
        cells = []
        for cell_el in [c for c in row_el if _local(c.tag) == "cell"]:
            cell_paras = []
            for p in cell_el:
                if _local(p.tag) == "paragraph":
                    for blk in _collect_paragraph(p, pool):
                        if isinstance(blk, Para):
                            cell_paras.append(blk)
            if not cell_paras:
                cell_paras.append(Para())
            cells.append(cell_paras)
        table.rows.append(cells)
    return table


def _parse_udf(udf_path):
    """UDF dosyasini blok listesine (Para/Img/Tbl) cozumler."""
    xml_text = _read_content_xml(udf_path)
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise UdfError("content.xml cozumlenemedi: %s" % exc)

    pool = ""
    elements_el = None
    for child in root:
        tag = _local(child.tag)
        if tag == "content" and child.text is not None:
            pool = child.text
        elif tag == "elements":
            elements_el = child

    blocks = []
    if elements_el is None:
        for line in pool.split("\n"):
            p = Para()
            if line:
                p.runs.append(Run(line))
            blocks.append(p)
        return blocks

    for el in elements_el:
        tag = _local(el.tag)
        if tag == "paragraph":
            blocks.extend(_collect_paragraph(el, pool))
        elif tag == "table":
            tbl = _collect_table(el, pool)
            if tbl is not None:
                blocks.append(tbl)
        # header/footer/pageBreak ve bilinmeyenler atlanir
    return blocks


# ---------------------------------------------------------------------------
# DOCX cikti
# ---------------------------------------------------------------------------

def _docx_add_runs(paragraph, para):
    from docx.shared import Pt, RGBColor
    for r in para.runs:
        # \n ve \t'yi koruyarak isle
        segments = r.text.split("\n")
        for si, seg in enumerate(segments):
            if si > 0:
                paragraph.add_run().add_break()
            parts = seg.split("\t")
            for pi, part in enumerate(parts):
                if pi > 0:
                    paragraph.add_run().add_tab()
                if part == "":
                    continue
                run = paragraph.add_run(part)
                if r.family:
                    run.font.name = r.family
                if r.size:
                    run.font.size = Pt(r.size)
                run.bold = r.bold
                run.italic = r.italic
                run.underline = r.underline
                if r.color:
                    try:
                        run.font.color.rgb = RGBColor.from_string(r.color)
                    except Exception:
                        pass


def _docx_set_align(paragraph, align):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    amap = {
        0: WD_ALIGN_PARAGRAPH.LEFT,
        1: WD_ALIGN_PARAGRAPH.CENTER,
        2: WD_ALIGN_PARAGRAPH.RIGHT,
        3: WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    if align in amap:
        paragraph.alignment = amap[align]


def _docx_fill_para(paragraph, para):
    from docx.shared import Pt
    _docx_set_align(paragraph, para.alignment)
    if para.left_indent:
        paragraph.paragraph_format.left_indent = Pt(para.left_indent)
    if para.right_indent:
        paragraph.paragraph_format.right_indent = Pt(para.right_indent)
    if para.first_line_indent:
        paragraph.paragraph_format.first_line_indent = Pt(para.first_line_indent)
    _docx_add_runs(paragraph, para)


def udf_to_docx(udf_path, docx_path):
    from docx import Document
    from docx.shared import Emu

    blocks = _parse_udf(udf_path)
    document = Document()

    for blk in blocks:
        if isinstance(blk, Para):
            _docx_fill_para(document.add_paragraph(), blk)
        elif isinstance(blk, Img):
            try:
                kw = {}
                if blk.width_pt:
                    kw["width"] = Emu(int(blk.width_pt * 12700))
                document.add_picture(io.BytesIO(blk.data), **kw)
            except Exception:
                pass
        elif isinstance(blk, Tbl):
            try:
                table = document.add_table(rows=0, cols=blk.col_count)
                try:
                    table.style = "Table Grid"
                except Exception:
                    pass
                for row in blk.rows:
                    cells = table.add_row().cells
                    for i, cell_paras in enumerate(row):
                        if i >= blk.col_count:
                            break
                        first = True
                        for cp in cell_paras:
                            para = cells[i].paragraphs[0] if first else cells[i].add_paragraph()
                            first = False
                            _docx_fill_para(para, cp)
            except Exception:
                pass

    document.save(docx_path)


# ---------------------------------------------------------------------------
# PDF cikti (reportlab)
# ---------------------------------------------------------------------------

_FONTS_READY = False
_TNR = "Helvetica"   # varsayilan; _ensure_fonts gercek fontlari kaydeder
_ARL = "Helvetica"


def _ensure_fonts():
    """Windows sistem fontlarini (Turkce karakterli) reportlab'e kaydeder."""
    global _FONTS_READY, _TNR, _ARL
    if _FONTS_READY:
        return
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    win = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")

    def _register(family, regular, bold, italic, bolditalic):
        path = lambda f: os.path.join(win, f)
        try:
            pdfmetrics.registerFont(TTFont(family, path(regular)))
        except Exception:
            return False
        names = {"normal": family}
        for face, fname in (("Bold", bold), ("Italic", italic), ("BoldItalic", bolditalic)):
            tag = family + "-" + face
            try:
                pdfmetrics.registerFont(TTFont(tag, path(fname)))
                names[{"Bold": "bold", "Italic": "italic", "BoldItalic": "boldItalic"}[face]] = tag
            except Exception:
                names.setdefault({"Bold": "bold", "Italic": "italic", "BoldItalic": "boldItalic"}[face], family)
        pdfmetrics.registerFontFamily(
            family,
            normal=names.get("normal", family),
            bold=names.get("bold", family),
            italic=names.get("italic", family),
            boldItalic=names.get("boldItalic", family),
        )
        return True

    if _register("Times", "times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"):
        _TNR = "Times"
    if _register("Arial", "arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"):
        _ARL = "Arial"
    # Times kaydedilemediyse Arial'e, o da olmadiysa Helvetica'ya dus
    if _TNR == "Helvetica" and _ARL != "Helvetica":
        _TNR = _ARL
    _FONTS_READY = True


def _pdf_face(family):
    if not family:
        return _TNR
    f = family.lower()
    if "times" in f or "serif" in f or "roman" in f:
        return _TNR
    return _ARL


def _run_markup(run):
    text = run.text.replace("\r", "")
    text = _xml_escape(text).replace("\n", "<br/>").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;")
    face = _pdf_face(run.family)
    attrs = 'face="%s"' % face
    if run.size:
        attrs += ' size="%s"' % run.size
    if run.color:
        attrs += ' color="#%s"' % run.color
    s = "<font %s>%s</font>" % (attrs, text)
    if run.bold:
        s = "<b>%s</b>" % s
    if run.italic:
        s = "<i>%s</i>" % s
    if run.underline:
        s = "<u>%s</u>" % s
    return s


def _pdf_paragraph(para, default_size=12):
    from reportlab.platypus import Paragraph as RLParagraph, Spacer
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY

    amap = {0: TA_LEFT, 1: TA_CENTER, 2: TA_RIGHT, 3: TA_JUSTIFY}
    markup = "".join(_run_markup(r) for r in para.runs).strip()
    if not markup:
        return Spacer(1, default_size)

    size = default_size
    face = _TNR
    for r in para.runs:
        if r.size:
            size = r.size
        if r.family:
            face = _pdf_face(r.family)
            break
    style = ParagraphStyle(
        "p",
        fontName=face,
        fontSize=size,
        leading=size * 1.25,
        alignment=amap.get(para.alignment, TA_LEFT),
        leftIndent=para.left_indent or 0,
        rightIndent=para.right_indent or 0,
        firstLineIndent=para.first_line_indent or 0,
        spaceAfter=4,
    )
    return RLParagraph(markup, style)


def udf_to_pdf(udf_path, pdf_path):
    _ensure_fonts()
    from reportlab.platypus import (SimpleDocTemplate, Table as RLTable,
                                    TableStyle, Image as RLImage, Spacer)
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors

    blocks = _parse_udf(udf_path)

    doc = SimpleDocTemplate(
        pdf_path, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
    )
    usable_w = doc.width
    story = []

    for blk in blocks:
        try:
            if isinstance(blk, Para):
                story.append(_pdf_paragraph(blk))
            elif isinstance(blk, Img):
                stream = io.BytesIO(blk.data)
                w = blk.width_pt
                h = blk.height_pt
                try:
                    img = RLImage(stream)
                except Exception:
                    continue
                iw = w or img.imageWidth
                ih = h or img.imageHeight
                if iw and ih:
                    scale = min(1.0, usable_w / float(iw))
                    img.drawWidth = iw * scale
                    img.drawHeight = ih * scale
                story.append(img)
                story.append(Spacer(1, 6))
            elif isinstance(blk, Tbl):
                data = []
                for row in blk.rows:
                    cells = []
                    for cell_paras in row:
                        flow = [_pdf_paragraph(cp, default_size=10) for cp in cell_paras]
                        cells.append(flow if flow else "")
                    while len(cells) < blk.col_count:
                        cells.append("")
                    data.append(cells[:blk.col_count])
                if data:
                    table = RLTable(data, colWidths=usable_w / float(blk.col_count))
                    table.setStyle(TableStyle([
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ]))
                    story.append(table)
                    story.append(Spacer(1, 6))
        except Exception:
            continue

    if not story:
        story.append(Spacer(1, 1))
    doc.build(story)


# ---------------------------------------------------------------------------
# PDF -> UDF  (UDF-Toolkit yontemi: metin + gomulu resimler)
# ---------------------------------------------------------------------------

_UDF_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<template format_id="1.8">\n'
    '<content><![CDATA[{content}]]></content>\n'
    '<properties><pageFormat mediaSizeName="1" leftMargin="42.51968479156494" '
    'rightMargin="42.51968479156494" topMargin="42.51968479156494" '
    'bottomMargin="42.51968479156494" paperOrientation="1" '
    'headerFOffset="20.0" footerFOffset="20.0"/></properties>\n'
    '<elements resolver="hvl-default">\n{elements}\n</elements>\n'
    '<styles>\n'
    '<style name="default" description="Geis Stded Stil" family="Times New Roman" '
    'size="12" bold="false" italic="false" foreground="-13421773" '
    'FONT_ATTRIBUTE_KEY="Dialog" />\n'
    '<style name="hvl-default" family="Times New Roman" size="12" '
    'description="Govde" />\n'
    '</styles>\n'
    '</template>'
)


def _xml_valid(text):
    """XML 1.0'da gecersiz kontrol karakterlerini atar (offset tutarliligi icin
    once temizlenir, sonra uzunluk olculur)."""
    out = []
    for ch in text:
        o = ord(ch)
        if ch in ("\t", "\n", "\r") or 0x20 <= o <= 0xD7FF or 0xE000 <= o <= 0xFFFD or o >= 0x10000:
            out.append(ch)
    return "".join(out)


def _cdata_safe(text):
    """CDATA bloklarini bozan ]]> dizgesini kacis yapar."""
    return text.replace("]]>", "]]]]><![CDATA[>")


def pdf_to_udf(pdf_path, udf_path):
    try:
        import pymupdf
    except Exception:
        try:
            import fitz as pymupdf  # eski surum adi
        except Exception:
            raise UdfError("PDF okuma kutuphanesi (PyMuPDF) yuklenemedi.")
    try:
        from PIL import Image as PILImage
        _has_pil = True
    except Exception:
        _has_pil = False

    try:
        pdf = pymupdf.open(pdf_path)
    except Exception as exc:
        raise UdfError("PDF acilamadi: %s" % exc)

    content_parts = []
    elements = []
    offset = 0

    def add_paragraph(text, align=0):
        nonlocal offset
        n = len(text)
        content_parts.append(text)
        elements.append(
            '<paragraph Alignment="%d" LeftIndent="0.0" RightIndent="0.0">'
            '<content startOffset="%d" length="%d" family="Times New Roman" '
            'size="12" bold="false" italic="false" foreground="-13421773"/>'
            '</paragraph>' % (align, offset, n)
        )
        offset += n

    try:
        for page in pdf:
            text = _xml_valid(page.get_text() or "")
            if text.strip():
                add_paragraph(text)

            # Gomulu resimler
            for img in page.get_images(full=True):
                xref = img[0]
                try:
                    base = pdf.extract_image(xref)
                    image_bytes = base["image"]
                except Exception:
                    continue
                # PNG'ye cevir
                b64 = None
                w_pt = h_pt = None
                if _has_pil:
                    try:
                        im = PILImage.open(io.BytesIO(image_bytes))
                        if im.mode in ("P", "RGBA", "LA"):
                            im = im.convert("RGBA")
                        else:
                            im = im.convert("RGB")
                        buf = io.BytesIO()
                        im.save(buf, format="PNG")
                        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                    except Exception:
                        b64 = None
                if b64 is None:
                    b64 = base64.b64encode(image_bytes).decode("ascii")
                # Goruntulenecek boyut (sayfadaki yerlesim dikdortgeninden, pt)
                try:
                    rects = page.get_image_rects(xref)
                    if rects:
                        w_pt = round(rects[0].width, 2)
                        h_pt = round(rects[0].height, 2)
                except Exception:
                    pass
                size_attr = ""
                if w_pt and h_pt:
                    size_attr = ' width="%s" height="%s"' % (w_pt, h_pt)

                content_parts.append(_IMAGE_PLACEHOLDER)
                elements.append(
                    '<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0">'
                    '<image family="Times New Roman" size="10" imageData="%s" '
                    'startOffset="%d" length="1"%s/></paragraph>'
                    % (b64, offset, size_attr)
                )
                offset += 1

            # Sayfa ayraci
            content_parts.append("\n")
            elements.append(
                '<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0">'
                '<content startOffset="%d" length="1" family="Times New Roman" '
                'size="12"/></paragraph>' % offset
            )
            offset += 1
    finally:
        pdf.close()

    content = _cdata_safe("".join(content_parts))
    xml = _UDF_TEMPLATE.format(content=content, elements="\n".join(elements))

    try:
        with zipfile.ZipFile(udf_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("content.xml", xml.encode("utf-8"))
    except Exception as exc:
        raise UdfError("UDF yazilamadi: %s" % exc)


# ---------------------------------------------------------------------------
# Komut satiri (hizli test)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 4:
        print("Kullanim: python converters.py [udf2docx|udf2pdf|pdf2udf] girdi cikti")
        sys.exit(1)
    mode, src, dst = sys.argv[1], sys.argv[2], sys.argv[3]
    {"udf2docx": udf_to_docx, "udf2pdf": udf_to_pdf, "pdf2udf": pdf_to_udf}[mode](src, dst)
    print("Tamam:", dst)
