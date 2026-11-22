# -*- coding: utf-8 -*-
"""Convert a Word FAQ with blue/teal questions into Question / Answer Excel rows.

HOW TO USE
    Run this file with Python 3.9 or newer, then paste the .docx file path.
    Or paste the path into WORD_PATH below. No extra Python packages are needed.

OUTPUT
    One .xlsx file beside the Word document, one worksheet, two columns.
    Existing output files are kept; later runs use _2, _3, etc.

ASSUMPTIONS
    Questions are blue/teal editable text; answers are ordinary body text.
    Word headers/footers and recognized section headings are excluded.
    This is a formatting-based extractor, not OCR or an AI classifier.
    Text boxes and embedded subdocuments need a simpler Word layout first.
    Original wording is kept; question numbers and layout formatting are removed.
"""

from pathlib import Path
from dataclasses import dataclass, field
import colorsys
import math
import os
import re
import sys
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED, BadZipFile


# OPTIONAL: paste your path between the quotes, or leave blank to be prompted.
WORD_PATH = r""

# Leave empty to recognize blue/teal automatically. For another question color,
# enter its six-digit RGB code, for example: QUESTION_COLORS = ("54899C",)
QUESTION_COLORS = ()
COLUMN_HEADERS = ("Question", "Answer")
REMOVE_QUESTION_NUMBERS = True

# Section headings that should not become part of an answer.
# Add another exact heading to SECTION_TITLES if your file needs it.
SECTION_PREFIXES = ("استفسارات ", "الاستفسارات ")
SECTION_TITLES = ()

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
STRICT_W = "{http://purl.oclc.org/ooxml/wordprocessingml/main}"
STRICT_A = "{http://purl.oclc.org/ooxml/drawingml/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
BIDI = re.compile("[\u200e\u200f\u202a-\u202e\u2066-\u2069\ufeff]")
NUMBER = re.compile(r"^\s*[(\[]?\d{1,4}\s*[-\u2010-\u2015\u2212.)\]:\u060c]\s*")
LABEL = re.compile(r"^(?:السؤال\s*[:：]|س\s*[:：]|Question\s*[:：]|Q\s*[:：])\s*", re.I)


def tidy(text):
    return BIDI.sub("", text).strip()


def is_blue(color):
    if not color or not re.fullmatch(r"[0-9a-fA-F]{6}", color):
        return False
    if QUESTION_COLORS:
        return color.upper() in {c.lstrip("#").upper() for c in QUESTION_COLORS}
    r, g, b = (int(color[i:i + 2], 16) / 255 for i in (0, 2, 4))
    hue, saturation, value = colorsys.rgb_to_hsv(r, g, b)
    return 0.44 <= hue <= 0.72 and saturation >= 0.15 and value >= 0.20


def read_xml(package, name):
    try:
        root = ET.fromstring(package.read(name))
    except KeyError:
        return None
    # Support both ordinary and Strict OOXML namespaces.
    for node in root.iter():
        node.tag = node.tag.replace(STRICT_W, W).replace(STRICT_A, A)
        node.attrib = {k.replace(STRICT_W, W).replace(STRICT_A, A): v
                       for k, v in node.attrib.items()}
    return root


def val(element, default=None):
    return element.get(W + "val", default) if element is not None else default


class Formatting:
    """Resolve question colors through direct formatting, styles, and themes."""

    def __init__(self, package):
        self.theme = {}
        theme = read_xml(package, "word/theme/theme1.xml")
        if theme is not None:
            scheme = theme.find(".//" + A + "clrScheme")
            if scheme is not None:
                for entry in scheme:
                    if len(entry):
                        c = entry[0]
                        self.theme[entry.tag.split("}")[-1]] = (
                            c.get("lastClr") or c.get("val", "000000"))
        self.styles = {}
        self.cache = {}
        self.default_style = "Normal"
        self.defaults = {}
        styles = read_xml(package, "word/styles.xml")
        if styles is not None:
            self.defaults = self.properties(styles.find(
                W + "docDefaults/" + W + "rPrDefault/" + W + "rPr"))
            for style in styles.findall(W + "style"):
                self.styles[style.get(W + "styleId")] = style
                if style.get(W + "type") == "paragraph" and style.get(W + "default") == "1":
                    self.default_style = style.get(W + "styleId")

    def color(self, element):
        if element is None:
            return None
        name = element.get(W + "themeColor")
        aliases = {"dark1": "dk1", "light1": "lt1", "dark2": "dk2",
                   "light2": "lt2", "hyperlink": "hlink", "followedHyperlink": "folHlink",
                   "text1": "dk1", "background1": "lt1", "text2": "dk2", "background2": "lt2"}
        color = self.theme.get(aliases.get(name, name)) if name else None
        from_theme = color is not None
        color = color or val(element, "000000")
        if color == "auto" or not re.fullmatch(r"[0-9a-fA-F]{6}", color):
            return "000000"
        rgb = [int(color[i:i + 2], 16) for i in (0, 2, 4)]
        if from_theme:
            for attribute in ("themeShade", "themeTint"):
                modifier = element.get(W + attribute)
                if modifier and re.fullmatch(r"[0-9a-fA-F]{2}", modifier):
                    factor = int(modifier, 16) / 255
                    if attribute == "themeShade":
                        rgb = [round(c * factor) for c in rgb]
                    else:
                        rgb = [round(c + (255 - c) * (1 - factor)) for c in rgb]
        return "".join(f"{max(0, min(255, c)):02X}" for c in rgb)

    def properties(self, rpr):
        result = {}
        if rpr is not None:
            color = rpr.find(W + "color")
            if color is not None:
                result["color"] = self.color(color)
            for tag, key in (("vanish", "hidden"), ("webHidden", "web_hidden")):
                el = rpr.find(W + tag)
                if el is not None:
                    result[key] = val(el, "1").lower() not in ("0", "false", "off")
        return result

    def style(self, style_id, seen=None):
        if style_id in self.cache:
            return self.cache[style_id].copy()
        seen = set() if seen is None else set(seen)
        if not style_id or style_id in seen or style_id not in self.styles:
            return {}
        seen.add(style_id)
        style = self.styles[style_id]
        result = self.style(val(style.find(W + "basedOn")), seen)
        result.update(self.properties(style.find(W + "rPr")))
        ppr = style.find(W + "pPr")
        result.update(self.paragraph_properties(ppr))
        result["style_name"] = val(style.find(W + "name"), style_id)
        self.cache[style_id] = result.copy()
        return result

    @staticmethod
    def paragraph_properties(ppr):
        result = {}
        if ppr is not None:
            outline = ppr.find(W + "outlineLvl")
            if outline is not None:
                result["outline"] = int(val(outline, "9"))
            numpr = ppr.find(W + "numPr")
            if numpr is not None:
                for tag in ("numId", "ilvl"):
                    el = numpr.find(W + tag)
                    if el is not None:
                        result[tag] = val(el)
        return result

    def paragraph(self, p):
        ppr = p.find(W + "pPr")
        sid = val(ppr.find(W + "pStyle")) if ppr is not None else None
        result = self.defaults.copy()
        result.update(self.style(sid or self.default_style))
        result.update(self.paragraph_properties(ppr))
        return result


class Numbering:
    """Keep common automatic answer-list markers instead of dropping them."""

    def __init__(self, package):
        self.levels, self.counts = {}, {}
        root = read_xml(package, "word/numbering.xml")
        if root is None:
            return
        abstracts = {n.get(W + "abstractNumId"): n for n in root.findall(W + "abstractNum")}
        for num in root.findall(W + "num"):
            source = abstracts.get(val(num.find(W + "abstractNumId")))
            levels = {}
            if source is not None:
                levels = {int(n.get(W + "ilvl", "0")): n for n in source.findall(W + "lvl")}
            overrides = {}
            for override in num.findall(W + "lvlOverride"):
                level = int(override.get(W + "ilvl", "0"))
                if override.find(W + "lvl") is not None:
                    levels[level] = override.find(W + "lvl")
                start = override.find(W + "startOverride")
                if start is not None:
                    overrides[level] = int(val(start, "1"))
            self.levels[num.get(W + "numId")] = (levels, overrides)

    def marker(self, props):
        numid = props.get("numId")
        if not numid or numid == "0":
            return ""
        level = int(props.get("ilvl", "0"))
        levels, overrides = self.levels.get(numid, ({}, {}))
        info = levels.get(level)
        if info is None:
            return ""
        counters = self.counts.setdefault(numid, {})
        start = overrides.get(level, int(val(info.find(W + "start"), "1")))
        counters[level] = counters.get(level, start - 1) + 1
        for deeper in list(counters):
            if deeper > level:
                restart = val(levels[deeper].find(W + "lvlRestart")) if deeper in levels else None
                if restart != "0":
                    del counters[deeper]
        kind = val(info.find(W + "numFmt"), "decimal")
        if kind == "none":
            return ""
        if kind == "bullet":
            return "• "
        template = val(info.find(W + "lvlText"), f"%{level + 1}.")

        def replace(match):
            index = int(match.group(1)) - 1
            entry = levels.get(index)
            fmt = val(entry.find(W + "numFmt"), "decimal") if entry is not None else "decimal"
            n = counters.get(index, overrides.get(index, int(val(entry.find(W + "start"), "1"))
                                                if entry is not None else 1))
            if fmt in ("lowerLetter", "upperLetter") and n > 0:
                letters = ""
                while n:
                    n, digit = divmod(n - 1, 26)
                    letters = chr(65 + digit) + letters
                return letters.lower() if fmt == "lowerLetter" else letters
            if fmt == "decimalZero":
                return f"{n:02d}"
            return str(n)

        return re.sub(r"%([1-9])", replace, template) + " "


@dataclass
class Span:
    text: str
    color: str = "000000"
    hyperlink: bool = False


@dataclass
class Paragraph:
    spans: list = field(default_factory=list)
    props: dict = field(default_factory=dict)
    marker: str = ""


def body_paragraphs(node):
    # Do not count the fallback and modern representation of the same content twice.
    if node.tag in (W + "del", W + "moveFrom"):
        return
    if node.tag == MC + "AlternateContent":
        selected = node.find(MC + "Choice")
        if selected is None:
            selected = node.find(MC + "Fallback")
        if selected is not None:
            yield from body_paragraphs(selected)
        return
    if node.tag == W + "p":
        yield node
        return
    for child in node:
        yield from body_paragraphs(child)


def read_spans(p, formatting, props):
    def walk(node, hyperlink=False):
        if node.tag in (W + "del", W + "moveFrom", W + "txbxContent", W + "pPr"):
            return
        if node.tag == MC + "AlternateContent":
            selected = node.find(MC + "Choice")
            if selected is None:
                selected = node.find(MC + "Fallback")
            if selected is not None:
                yield from walk(selected, hyperlink)
            return
        hyperlink = hyperlink or node.tag == W + "hyperlink"
        if node.tag == W + "r":
            rp = node.find(W + "rPr")
            sid = val(rp.find(W + "rStyle")) if rp is not None else None
            effective = props.copy()
            effective.update(formatting.style(sid))
            effective.update(formatting.properties(rp))
            if effective.get("hidden"):
                return
            hyperlink = hyperlink or "hyperlink" in (sid or "").lower()
            text = []
            for item in node:
                if item.tag == W + "t":
                    text.append(item.text or "")
                elif item.tag == W + "tab":
                    text.append("\t")
                elif item.tag in (W + "br", W + "cr"):
                    text.append("\n" if item.get(W + "type", "textWrapping") == "textWrapping" else " ")
                elif item.tag == W + "noBreakHyphen":
                    text.append("‑")
            if text:
                yield Span("".join(text), effective.get("color", "000000"), hyperlink)
            return
        for child in node:
            yield from walk(child, hyperlink)

    return list(walk(p))


def lines(spans):
    result, current = [], []
    for span in spans:
        pieces = span.text.split("\n")
        for index, piece in enumerate(pieces):
            if index:
                result.append(current)
                current = []
            if piece:
                current.append(Span(piece, span.color, span.hyperlink))
    result.append(current)
    return result


def blue_line(spans):
    letters = [(c, is_blue(s.color) and not s.hyperlink)
               for s in spans for c in s.text if c.isalpha()]
    return bool(letters) and letters[0][1] and sum(blue for _, blue in letters) / len(letters) >= 0.60


def split_inline(spans):
    """Handle a blue question ending in ? followed by a black answer on one line."""
    chars = [(c, s.color, s.hyperlink) for s in spans for c in s.text]
    first = next((x for x in chars if x[0].isalpha()), None)
    if first is None or not is_blue(first[1]) or first[2]:
        return [spans]
    for index, (char, _, _) in enumerate(chars):
        if char not in "?؟":
            continue
        after = next((x for x in chars[index + 1:] if x[0].isalpha()), None)
        before_spans = [Span(c, col, link) for c, col, link in chars[:index + 1]]
        if after and not is_blue(after[1]) and blue_line(before_spans):
            return [before_spans, [Span(c, col, link) for c, col, link in chars[index + 1:]]]
    return [spans]


def is_heading(paragraph):
    text = tidy("".join(s.text for s in paragraph.spans))
    clean = NUMBER.sub("", text).strip()
    if clean in SECTION_TITLES or clean.startswith(SECTION_PREFIXES):
        return True
    props = paragraph.props
    name = props.get("style_name", "").lower()
    styled = props.get("outline", 9) < 9 or bool(re.match(r"^(heading|title|subtitle|عنوان)", name))
    return styled and not any(blue_line(line) for line in lines(paragraph.spans))


def extract_docx(path):
    """Return ([(question, answer), ...], [review notes, ...])."""
    path = Path(path)
    if path.suffix.lower() != ".docx":
        raise ValueError("Use a .docx Word document. For an old .doc file, open it in Word and use Save As > Word Document (.docx).")
    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")
    notes = []
    try:
        with ZipFile(path) as package:
            root = read_xml(package, "word/document.xml")
            body = root.find(W + "body") if root is not None else None
            if body is None:
                raise ValueError("This file does not contain a readable Word document body.")
            if any(t.text for box in body.iter(W + "txbxContent") for t in box.iter(W + "t")):
                raise ValueError("The document contains text boxes. Move their question/answer text into normal Word paragraphs before running this script.")
            if body.find(".//" + W + "altChunk") is not None:
                raise ValueError("The document contains embedded subdocuments. Open it in Word and save a normal .docx copy first.")
            formatting, numbering = Formatting(package), Numbering(package)
            paragraphs = []
            for p in body_paragraphs(body):
                props = formatting.paragraph(p)
                marker = numbering.marker(props)
                spans = read_spans(p, formatting, props)
                if tidy("".join(s.text for s in spans)):
                    paragraphs.append(Paragraph(spans, props, marker))
            if body.find(".//" + W + "drawing") is not None or body.find(".//" + W + "pict") is not None:
                notes.append("The document body contains images/shapes. Only editable text was extracted; check any text inside images separately.")
    except BadZipFile as exc:
        raise ValueError("This is not a readable .docx file. If it is password-protected, save an unprotected .docx copy in Word first.") from exc
    except ET.ParseError as exc:
        raise ValueError("The Word XML could not be read. Open the document in Word and save a new .docx copy.") from exc
    if not paragraphs:
        raise ValueError("No editable text was found. A scanned/image-only document needs OCR first.")

    pairs, question, answer = [], [], []
    ignored = 0

    def finish():
        if question:
            pairs.append(("\n".join(question).strip(), "\n".join(answer).strip()))
        question.clear()
        answer.clear()

    for paragraph in paragraphs:
        if is_heading(paragraph):
            finish()
            continue
        first_content = True
        for line in lines(paragraph.spans):
            for fragment in split_inline(line):
                text = tidy("".join(s.text for s in fragment))
                if not text:
                    continue
                if blue_line(fragment):
                    explicit_number = bool(NUMBER.match(text)) or (
                        first_content and paragraph.props.get("numId", "0") != "0")
                    starts_new = answer or explicit_number or (question and question[-1].rstrip().endswith(("?", "؟")))
                    if question and starts_new:
                        finish()
                    if REMOVE_QUESTION_NUMBERS:
                        text = NUMBER.sub("", text)
                    text = LABEL.sub("", text).strip()
                    if text:
                        question.append(text)
                elif question:
                    # Numbering prefixes are metadata in Word, not literal text.
                    marker = paragraph.marker if first_content and not NUMBER.match(text) else ""
                    answer.append(marker + text)
                else:
                    ignored += 1
                first_content = False
    finish()
    if not pairs:
        raise ValueError("No blue/teal questions were detected. This script expects the formatting shown in your screenshots. Check that the questions are editable blue text, or set QUESTION_COLORS at the top of the script to their actual RGB color.")
    empty = [str(i + 2) for i, (_, a) in enumerate(pairs) if not a]
    if empty:
        notes.append(f"{len(empty)} question(s) have an empty answer. Review Excel row(s): {', '.join(empty[:12])}" + (" ..." if len(empty) > 12 else ""))
    if ignored:
        notes.append(f"Ignored {ignored} non-question text line(s) outside question/answer groups (for example, introductory text).")
    return pairs, notes


def write_xlsx(path, pairs):
    """Write a real .xlsx file using only Python's standard library."""
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise ValueError("The output filename must end with .xlsx.")
    rows = [COLUMN_HEADERS] + list(pairs)
    if len(rows) > 1048576:
        raise ValueError("There are too many rows for one Excel worksheet.")
    for row_number, row in enumerate(rows, 1):
        for column, text in zip(("A", "B"), row):
            if len(text.encode("utf-16-le")) // 2 > 32767:
                raise ValueError(f"Cell {column}{row_number} exceeds Excel's 32,767-character limit. No workbook was written; shorten or split that entry in Word first.")

    rtl = any(re.search(r"[\u0600-\u06ff]", q + a) for q, a in pairs)
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    worksheet = [f'<worksheet xmlns="{ns}"><dimension ref="A1:B{len(rows)}"/>',
                 f'<sheetViews><sheetView workbookViewId="0" rightToLeft="{int(rtl)}" showGridLines="0">',
                 '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
                 '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>',
                 '<sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="65" customWidth="1"/>',
                 '<col min="2" max="2" width="105" customWidth="1"/></cols><sheetData>']
    for row_number, row in enumerate(rows, 1):
        line_count = max(sum(max(1, math.ceil(len(line) / width)) for line in text.split("\n"))
                         for text, width in zip(row, (58, 96)))
        height = 28 if row_number == 1 else min(409, max(40, 18 * line_count + 12))
        worksheet.append(f'<row r="{row_number}" ht="{height}" customHeight="1">')
        for column, text in zip(("A", "B"), row):
            style = 1 if row_number == 1 else (2 if row_number % 2 == 0 else 3)
            # inlineStr preserves Arabic, line breaks, and literal text starting '='.
            worksheet.append(f'<c r="{column}{row_number}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c>')
        worksheet.append('</row>')
    worksheet.append(f'</sheetData><autoFilter ref="A1:B{len(rows)}"/></worksheet>')
    align = 'right' if rtl else 'left'
    reading = '2' if rtl else '1'
    styles = f'''<styleSheet xmlns="{ns}">
      <fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts>
      <fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF28596B"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF1F6F8"/><bgColor indexed="64"/></patternFill></fill></fills>
      <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="4">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="49" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
        <xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="{align}" vertical="top" wrapText="1" readingOrder="{reading}"/></xf>
        <xf numFmtId="49" fontId="0" fillId="3" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="{align}" vertical="top" wrapText="1" readingOrder="{reading}"/></xf>
      </cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
    parts = {
        '[Content_Types].xml': '''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>''',
        '_rels/.rels': '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>''',
        'xl/workbook.xml': f'''<workbook xmlns="{ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Questions and Answers" sheetId="1" r:id="rId1"/></sheets></workbook>''',
        'xl/_rels/workbook.xml.rels': '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>''',
        'xl/styles.xml': styles,
        'xl/worksheets/sheet1.xml': ''.join(worksheet),
    }
    # Exclusive creation prevents accidental overwriting of an existing workbook.
    with path.open("xb") as stream:
        try:
            with ZipFile(stream, "w", ZIP_DEFLATED) as package:
                for name, text in parts.items():
                    package.writestr(name, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + text)
        except Exception:
            stream.close()
            path.unlink(missing_ok=True)
            raise


def clean_path(raw):
    raw = tidy(raw)
    if len(raw) >= 3 and raw[0] in "rR" and raw[1] in "\"'" and raw[-1] == raw[1]:
        raw = raw[1:]
    if len(raw) >= 2 and (raw[0], raw[-1]) in (("\"", "\""), ("'", "'"), ("“", "”")):
        raw = raw[1:-1]
    return Path(os.path.expandvars(raw)).expanduser()


def main():
    print("Word questions and answers -> Excel")
    print("Input: editable .docx with blue/teal questions. No extra packages needed.\n")
    raw = WORD_PATH.strip()
    if not raw and len(sys.argv) == 2 and sys.argv[1] not in ("-h", "--help"):
        raw = sys.argv[1]
    if not raw:
        raw = input("Paste the Word document path, then press Enter:\n> ").strip()
    if not raw:
        raise ValueError("No file path was entered.")
    source = clean_path(raw)
    pairs, notes = extract_docx(source)
    index = 1
    while True:
        suffix = "" if index == 1 else f"_{index}"
        output = source.with_name(f"{source.stem}_QA{suffix}.xlsx")
        try:
            write_xlsx(output, pairs)
            break
        except FileExistsError:
            index += 1
    print(f"\nCreated {len(pairs)} question/answer row(s).")
    print(f"Excel file: {output.resolve()}")
    for note in notes:
        print(f"Review note: {note}")
    print("\nSpot-check several rows against the Word document, especially long answers.")


if __name__ == "__main__":
    status = 0
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled.")
        status = 1
    except PermissionError:
        print("\nERROR: The file or folder is not accessible. Check permissions or copy the Word file to a writable folder and run again.")
        status = 1
    except (ValueError, OSError) as error:
        print(f"\nERROR: {error}")
        status = 1
    # Keep a double-clicked Windows terminal open so the result is readable.
    if os.name == "nt" and len(sys.argv) == 1:
        try:
            input("\nPress Enter to close...")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(status)
