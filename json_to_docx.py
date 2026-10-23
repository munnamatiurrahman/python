#!/usr/bin/env python3
"""Create editable, two-column DOCX question-and-answer books from JSON files.

Place this script beside one or more JSON files and double-click it. Each JSON
file creates or replaces output/<json-filename>.docx. The JSON format used by
this folder is a list of question records with sub, chap, top, qus, opt and ans.
"""
from __future__ import annotations

import html
import json
import os
import posixpath
import re
import shutil
import sys
import subprocess
import tempfile
import time
from collections import OrderedDict
from copy import deepcopy
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

if os.name == "nt":
    for _stream in (sys.stdout, sys.stderr):
        if hasattr(_stream, "reconfigure"):
            try:
                _stream.reconfigure(encoding="utf-8", errors="replace")
            except OSError:
                pass

try:
    from docx import Document
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.section import WD_SECTION
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Mm, Pt, RGBColor
except ImportError as exc:
    print("python-docx ইনস্টল করা নেই। এই ফোল্ডারে PowerShell/Command Prompt খুলে চালান:")
    print("  python -m pip install -r requirements.txt")
    print(f"বিস্তারিত: {exc}")
    if os.name == "nt" and sys.stdin.isatty():
        try:
            input("বন্ধ করতে Enter চাপুন...")
        except EOFError:
            pass
    raise SystemExit(1)

LATIN_FONT = "Times New Roman"
BENGALI_FONT = "SolaimanLipi"
BODY_SIZE_PT = 9.5
COLUMN_GAP_IN = 0.25
PAGE_MARGIN_IN = 0.42


class _VisibleText(HTMLParser):
    """Extract readable text from HTML fragments in question/option values."""

    _BREAK_TAGS = {"br", "p", "div", "li", "tr", "td", "th", "section"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag in {"script", "style"}:
            self._skip_depth += 1
        elif not self._skip_depth and tag in self._BREAK_TAGS:
            self.parts.append(" ")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in {"script", "style"} and self._skip_depth:
            self._skip_depth -= 1
        elif not self._skip_depth and tag in self._BREAK_TAGS:
            self.parts.append(" ")

    def handle_data(self, data: str) -> None:
        if not self._skip_depth:
            self.parts.append(data)


def clean_text(value: Any) -> str:
    """Return plain, whitespace-normalized text for a JSON value."""
    if value is None:
        return ""
    if not isinstance(value, str):
        value = str(value)
    parser = _VisibleText()
    try:
        parser.feed(html.unescape(value))
        parser.close()
        text = "".join(parser.parts)
    except Exception:
        text = value
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def _descriptor(value: Any, fallback: str) -> tuple[str, str]:
    """Convert a JSON grouping field to a stable (code, display name) pair."""
    if isinstance(value, (list, tuple)):
        code = clean_text(value[0]) if value else ""
        name = clean_text(value[1]) if len(value) > 1 else ""
    elif isinstance(value, dict):
        code = clean_text(value.get("code", value.get("id", value.get("key", ""))))
        name = clean_text(value.get("name", value.get("title", value.get("label", ""))))
    else:
        code = ""
        name = clean_text(value)
    if not name:
        name = code or fallback
    return code or name, name


def _question_records(parsed: Any) -> list[dict[str, Any]]:
    """Accept the current top-level list and common list wrappers."""
    if isinstance(parsed, list):
        records = parsed
    elif isinstance(parsed, dict):
        records = None
        for key in ("questions", "items", "records", "data"):
            if isinstance(parsed.get(key), list):
                records = parsed[key]
                break
        if records is None and ("qus" in parsed or "question" in parsed):
            records = [parsed]
        if records is None:
            raise ValueError("JSON object-এ questions/items/records/data তালিকা পাওয়া যায়নি")
    else:
        raise ValueError("JSON-এর মূল অংশটি list অথবা question-record object হতে হবে")
    return [item for item in records if isinstance(item, dict)]


def _option_list(record: dict[str, Any]) -> list[Any]:
    options = record.get("opt", record.get("options", []))
    if isinstance(options, dict):
        return list(options.values())
    return options if isinstance(options, list) else []


def _resolve_one_answer(answer: Any, options: list[Any]) -> str:
    """Resolve the folder's zero-based ans index to its answer text."""
    if isinstance(answer, dict):
        for key in ("index", "option", "value", "text", "answer"):
            if key in answer:
                return _resolve_one_answer(answer[key], options)
        return ""
    if isinstance(answer, bool):
        return clean_text(answer)
    if isinstance(answer, int):
        if 0 <= answer < len(options):
            return clean_text(options[answer])
        return ""
    if isinstance(answer, float) and answer.is_integer():
        return _resolve_one_answer(int(answer), options)
    if isinstance(answer, str):
        value = clean_text(answer)
        if value.isdecimal() and options:
            index = int(value)
            if 0 <= index < len(options):
                return clean_text(options[index])
        # Some JSON exports store a letter such as ক/খ/গ/ঘ instead of an index.
        letter_codes = {"a": 0, "b": 1, "c": 2, "d": 3, "ক": 0, "খ": 1, "গ": 2, "ঘ": 3}
        if value.casefold() in letter_codes:
            index = letter_codes[value.casefold()]
            if index < len(options):
                return clean_text(options[index])
        return value
    return ""


def answer_text(record: dict[str, Any]) -> str:
    answer = record.get("ans", record.get("answer"))
    options = _option_list(record)
    if isinstance(answer, list):
        values = [_resolve_one_answer(item, options) for item in answer]
        return "; ".join(value for value in values if value)
    return _resolve_one_answer(answer, options)


def _alpha_label(number: int) -> str:
    """Spreadsheet-style lower-case labels: a, b, ..., z, aa, ab, ..."""
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("a") + remainder) + result
    return result or "a"


def _roman_label(number: int) -> str:
    values = ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
              (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
              (5, "v"), (4, "iv"), (1, "i"))
    if number <= 0:
        return str(number)
    if number > 3999:
        return str(number)
    output = []
    for value, numeral in values:
        count, number = divmod(number, value)
        output.append(numeral * count)
    return "".join(output)


def build_tree(records: list[dict[str, Any]], fallback_subject: str) -> tuple[OrderedDict, int, int]:
    """Group records by subject, chapter and topic in their source order."""
    subjects: OrderedDict = OrderedDict()
    warnings = 0
    question_count = 0
    for record in records:
        if "sub" in record:
            sub_code, sub_name = _descriptor(record.get("sub"), fallback_subject)
        else:
            sub_code, sub_name = fallback_subject, fallback_subject
        chap_code, chap_name = _descriptor(record.get("chap"), "অধ্যায়")
        topic_code, topic_name = _descriptor(record.get("top"), "বিষয়")
        subject_key = (sub_code, sub_name)
        chapter_key = (chap_code, chap_name)
        topic_key = (topic_code, topic_name)
        subject = subjects.setdefault(subject_key, {"name": sub_name, "chapters": OrderedDict()})
        chapter = subject["chapters"].setdefault(chapter_key, {"name": chap_name, "topics": OrderedDict()})
        topic = chapter["topics"].setdefault(topic_key, {"name": topic_name, "questions": []})

        question = clean_text(record.get("qus", record.get("question", "")))
        if not question:
            warnings += 1
            continue
        answer = answer_text(record)
        if not answer:
            answer = "উত্তর অনুপস্থিত"
            warnings += 1
        topic["questions"].append((question, answer))
        question_count += 1
    return subjects, question_count, warnings


def set_script_fonts(target: Any, size_pt: float | None = None, bold: bool | None = None) -> None:
    """Set Latin and Bengali font slots in a style or a python-docx run."""
    if hasattr(target, "font"):
        target.font.name = LATIN_FONT
        if size_pt is not None:
            target.font.size = Pt(size_pt)
        if bold is not None:
            target.font.bold = bold
        element = target._element
        rpr = element.get_or_add_rPr()
    else:
        element = target._element
        rpr = element.get_or_add_rPr()
        target.font.name = LATIN_FONT
        if size_pt is not None:
            target.font.size = Pt(size_pt)
        if bold is not None:
            target.font.bold = bold
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    fonts.set(qn("w:ascii"), LATIN_FONT)
    fonts.set(qn("w:hAnsi"), LATIN_FONT)
    fonts.set(qn("w:eastAsia"), BENGALI_FONT)
    fonts.set(qn("w:cs"), BENGALI_FONT)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "bn-BD")
    lang.set(qn("w:bidi"), "bn-BD")


def _ensure_paragraph_style(doc: Any, name: str) -> Any:
    try:
        return doc.styles[name]
    except KeyError:
        return doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)


def _set_section_columns(section: Any, count: int) -> None:
    columns = section._sectPr.find(qn("w:cols"))
    if columns is None:
        columns = OxmlElement("w:cols")
        section._sectPr.append(columns)
    columns.set(qn("w:num"), str(count))
    columns.set(qn("w:space"), str(round(COLUMN_GAP_IN * 1440)))
    columns.set(qn("w:equalWidth"), "1")


def _format_section(section: Any, columns: int) -> None:
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Inches(PAGE_MARGIN_IN)
    section.bottom_margin = Inches(PAGE_MARGIN_IN)
    section.left_margin = Inches(PAGE_MARGIN_IN)
    section.right_margin = Inches(PAGE_MARGIN_IN)
    section.header_distance = Inches(0.18)
    section.footer_distance = Inches(0.18)
    _set_section_columns(section, columns)
    _set_page_number_start(section, None)


def _clear_paragraph(paragraph: Any) -> None:
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def _add_external_hyperlink(paragraph: Any, text: str, url: str, part: Any) -> None:
    relationship_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), relationship_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), LATIN_FONT)
    fonts.set(qn("w:hAnsi"), LATIN_FONT)
    fonts.set(qn("w:eastAsia"), BENGALI_FONT)
    fonts.set(qn("w:cs"), BENGALI_FONT)
    rpr.append(fonts)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)
    run.append(rpr)
    value = OxmlElement("w:t")
    value.text = text
    run.append(value)
    link.append(run)
    paragraph._p.append(link)


def _set_section_header(section: Any, chapter_name: str, subject_name: str, *, unlink: bool = False) -> None:
    header = section.header
    if unlink:
        header.is_linked_to_previous = False
    paragraph = header.paragraphs[0]
    paragraph.style = "Normal"
    _clear_paragraph(paragraph)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(7.30), WD_TAB_ALIGNMENT.RIGHT)
    left = paragraph.add_run(chapter_name)
    set_script_fonts(left, 8.5)
    paragraph.add_run("\t")
    right = paragraph.add_run(subject_name)
    set_script_fonts(right, 8.5)


def _set_section_footer(section: Any, *, unlink: bool = False) -> None:
    footer = section.footer
    if unlink:
        footer.is_linked_to_previous = False
    paragraph = footer.paragraphs[0]
    paragraph.style = "Normal"
    _clear_paragraph(paragraph)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(7.30), WD_TAB_ALIGNMENT.RIGHT)
    _add_external_hyperlink(paragraph, "pretestbd.web.app", "https://pretestbd.web.app", footer.part)
    paragraph.add_run("\t")
    add_complex_field(paragraph, "PAGE", "1")


def _set_page_number_start(section: Any, start: int | None) -> None:
    page_number = section._sectPr.find(qn("w:pgNumType"))
    if start is None:
        if page_number is not None:
            section._sectPr.remove(page_number)
        return
    if page_number is None:
        page_number = OxmlElement("w:pgNumType")
        columns = section._sectPr.find(qn("w:cols"))
        if columns is None:
            section._sectPr.append(page_number)
        else:
            section._sectPr.insert(section._sectPr.index(columns), page_number)
    page_number.set(qn("w:start"), str(start))


def configure_document(doc: Any) -> None:
    """Set compact A4 pages, editable styles and automatic field refresh."""
    _format_section(doc.sections[0], 1)

    normal = doc.styles["Normal"]
    set_script_fonts(normal, BODY_SIZE_PT)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(2)
    normal.paragraph_format.line_spacing = 1.0
    normal.paragraph_format.widow_control = True

    title = doc.styles["Title"]
    set_script_fonts(title, 15, True)
    title.font.color.rgb = RGBColor(0, 0, 0)
    title.font.underline = False
    title_ppr = title._element.find(qn("w:pPr"))
    if title_ppr is not None:
        for border in title_ppr.findall(qn("w:pBdr")):
            title_ppr.remove(border)
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(5)
    title.paragraph_format.keep_with_next = True

    for name, size, before, after in (("Heading 1", 12, 4, 2), ("Heading 2", 11, 3, 1.5), ("Heading 3", 10, 2, 1)):
        style = doc.styles[name]
        set_script_fonts(style, size, True)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.line_spacing = 1.0
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.keep_together = True
        style.paragraph_format.widow_control = True

    question_style = _ensure_paragraph_style(doc, "Question")
    question_style.base_style = normal
    set_script_fonts(question_style, BODY_SIZE_PT)
    question_style.paragraph_format.left_indent = Inches(0.19)
    question_style.paragraph_format.first_line_indent = Inches(-0.19)
    question_style.paragraph_format.space_before = Pt(1)
    question_style.paragraph_format.space_after = Pt(2)
    question_style.paragraph_format.line_spacing = 1.0
    question_style.paragraph_format.keep_with_next = False
    question_style.paragraph_format.keep_together = True

    answer_style = _ensure_paragraph_style(doc, "Answer")
    answer_style.base_style = normal
    set_script_fonts(answer_style, BODY_SIZE_PT)
    answer_style.paragraph_format.left_indent = Inches(0.19)
    answer_style.paragraph_format.space_before = Pt(0)
    answer_style.paragraph_format.space_after = Pt(2)
    answer_style.paragraph_format.line_spacing = 1.0
    answer_style.paragraph_format.keep_together = True

    toc_title = _ensure_paragraph_style(doc, "TOC Title")
    toc_title.base_style = normal
    set_script_fonts(toc_title, 12, True)
    toc_title.font.color.rgb = RGBColor(0, 0, 0)
    toc_title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    toc_title.paragraph_format.space_before = Pt(1)
    toc_title.paragraph_format.space_after = Pt(2)
    toc_title.paragraph_format.keep_with_next = True

    for level in (1, 2, 3):
        style = _ensure_paragraph_style(doc, f"TOC {level}")
        style.base_style = normal
        set_script_fonts(style, 8.5 if level == 3 else 9)
        style.paragraph_format.left_indent = Inches((level - 1) * 0.09)
        style.paragraph_format.first_line_indent = Inches(0)
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(0.5)
        style.paragraph_format.line_spacing = 1.0
        style.paragraph_format.keep_together = True
        style.paragraph_format.tab_stops.add_tab_stop(Inches(7.3), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

    settings = doc.settings.element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")


def add_plain_text(paragraph: Any, text: str, *, bold: bool = False) -> None:
    """Add text in editable Word runs and preserve any line breaks."""
    lines = text.splitlines() or [text]
    for index, line in enumerate(lines):
        if index:
            paragraph.add_run().add_break()
        if line:
            run = paragraph.add_run(line)
            run.bold = bold
            set_script_fonts(run)


def add_bookmark(paragraph: Any, name: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    p = paragraph._p
    ppr = p.find(qn("w:pPr"))
    p.insert(1 if ppr is not None else 0, start)
    p.append(end)


def add_column_break(paragraph: Any) -> None:
    """Start this centered topic heading in the next column."""
    run = OxmlElement("w:r")
    column_break = OxmlElement("w:br")
    column_break.set(qn("w:type"), "column")
    run.append(column_break)
    p = paragraph._p
    insert_at = 1 if p.find(qn("w:pPr")) is not None else 0
    while insert_at < len(p) and p[insert_at].tag == qn("w:bookmarkStart"):
        insert_at += 1
    p.insert(insert_at, run)


def _question_column_split(questions: list[tuple[str, str]]) -> int:
    """Balance one topic's editable questions between its two page columns."""
    if len(questions) < 2:
        return len(questions)
    weights = [len(question) + len(answer) + 80 for question, answer in questions]
    total = sum(weights)
    running = 0
    candidates: list[tuple[int, int]] = []
    for split_after, weight in enumerate(weights[:-1], 1):
        running += weight
        candidates.append((abs(total - 2 * running), split_after))
    return min(candidates)[1]


def _set_table_cell_margins(cell: Any, *, side_twips: int = 110) -> None:
    """Set a small central gutter while keeping table text away from page edges."""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side in ("top", "left", "bottom", "right"):
        item = tc_mar.find(qn(f"w:{side}"))
        if item is None:
            item = OxmlElement(f"w:{side}")
            tc_mar.append(item)
        item.set(qn("w:w"), str(side_twips if side in {"left", "right"} else 0))
        item.set(qn("w:type"), "dxa")


def _add_question_columns(doc: Any, questions: list[tuple[str, str]]) -> None:
    """Add editable Q/A paragraphs that flow naturally across two columns."""
    for local_index, (question, answer) in enumerate(questions, 1):
        q_paragraph = doc.add_paragraph(style="Question")
        number_run = q_paragraph.add_run(f"{local_index}. ")
        number_run.bold = True
        set_script_fonts(number_run)
        add_plain_text(q_paragraph, question)

        arrow_run = q_paragraph.add_run(" ↪ ")
        arrow_run.bold = True
        set_script_fonts(arrow_run)
        add_plain_text(q_paragraph, answer)


def _flatten_pdf_outline(items: list[Any]) -> list[Any]:
    flattened: list[Any] = []
    for item in items:
        if isinstance(item, list):
            flattened.extend(_flatten_pdf_outline(item))
        else:
            flattened.append(item)
    return flattened


def _update_toc_page_cache(docx_path: Path, heading_specs: list[tuple[int, str, str]]) -> bool:
    """Use LibreOffice's PDF outline to cache correct, still-editable PAGEREF fields."""
    try:
        from pypdf import PdfReader
    except ImportError:
        print("  সূচির পৃষ্ঠা নম্বর গণনায় pypdf নেই; requirements.txt ইনস্টল করুন।")
        return False

    if os.name == "nt":
        soffice = shutil.which("soffice.com") or shutil.which("soffice.exe")
        if not soffice:
            for candidate in (
                Path(r"C:\Program Files\LibreOffice\program\soffice.com"),
                Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.com"),
            ):
                if candidate.is_file():
                    soffice = str(candidate)
                    break
    else:
        soffice = shutil.which("soffice")
    if not soffice:
        print("  LibreOffice পাওয়া যায়নি; Word-এ Ctrl+A, F9 দিয়ে সূচি আপডেট করুন।")
        return False

    try:
        with tempfile.TemporaryDirectory(prefix="toc_refresh_", dir=docx_path.parent) as temp_dir:
            temp_root = Path(temp_dir)
            pdf_dir = temp_root / "pdf"
            profile_dir = temp_root / "lo_profile"
            pdf_dir.mkdir()
            profile_dir.mkdir()
            command = [
                soffice,
                "--headless",
                f"-env:UserInstallation={profile_dir.as_uri()}",
                "--norestore",
                "--nodefault",
                "--nofirststartwizard",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_dir),
                str(docx_path),
            ]
            completed = subprocess.run(command, capture_output=True, text=True, timeout=900)
            pdf_path = pdf_dir / (docx_path.stem + ".pdf")
            if completed.returncode != 0 or not pdf_path.is_file():
                details = (completed.stderr or completed.stdout or "").strip()
                raise RuntimeError(details or f"LibreOffice exit code {completed.returncode}")

            reader = PdfReader(str(pdf_path))
            outline = _flatten_pdf_outline(reader.outline)
            if len(outline) != len(heading_specs):
                raise RuntimeError(
                    f"PDF outline-এ {len(outline)}টি heading, DOCX-এ {len(heading_specs)}টি"
                )
            physical_pages = [reader.get_destination_page_number(item) for item in outline]
            if not physical_pages or any(page is None or page < 0 for page in physical_pages):
                raise RuntimeError("PDF outline থেকে heading-এর page number পাওয়া যায়নি")

            body_first_page = physical_pages[0]
            page_by_bookmark = {
                spec[2]: physical_page - body_first_page + 1
                for spec, physical_page in zip(heading_specs, physical_pages)
            }

        with ZipFile(docx_path, "r") as source_zip:
            package = {name: source_zip.read(name) for name in source_zip.namelist()}
        document_xml = etree.fromstring(package["word/document.xml"])
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        updated = 0
        for field in document_xml.xpath('.//w:fldSimple[contains(@w:instr, "PAGEREF")]', namespaces=ns):
            instruction = field.get(qn("w:instr"), "")
            match = re.search(r"PAGEREF\s+(\S+)", instruction)
            page_number = page_by_bookmark.get(match.group(1)) if match else None
            if page_number is None:
                continue
            text_nodes = field.xpath(".//w:t", namespaces=ns)
            if not text_nodes:
                run = etree.SubElement(field, qn("w:r"))
                text_nodes = [etree.SubElement(run, qn("w:t"))]
            text_nodes[0].text = str(page_number)
            for extra in text_nodes[1:]:
                extra.getparent().remove(extra)
            field.set(qn("w:dirty"), "true")
            updated += 1
        if updated != len(heading_specs):
            raise RuntimeError(f"{updated}/{len(heading_specs)}টি সূচি page reference মেলানো গেছে")
        package["word/document.xml"] = etree.tostring(
            document_xml, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        temp_docx = docx_path.with_suffix(".toc.tmp.docx")
        with ZipFile(temp_docx, "w", ZIP_DEFLATED) as output_zip:
            for name, data in package.items():
                output_zip.writestr(name, data)
        os.replace(temp_docx, docx_path)
        print(f"  সূচির {updated}টি page reference পৃষ্ঠা বিন্যাস থেকে আপডেট হয়েছে।")
        return True
    except Exception as exc:
        print(f"  সূচির পৃষ্ঠা নম্বর স্বয়ংক্রিয়ভাবে গণনা করা যায়নি: {exc}")
        print("  Word-এ Ctrl+A, F9 দিলে PAGEREF field আপডেট হবে।")
        return False


def _preserve_existing_cover(existing_path: Path, generated_path: Path) -> bool:
    """Carry forward an image-based first section from an edited output DOCX."""
    if not existing_path.is_file():
        return False
    with ZipFile(existing_path, "r") as old_zip:
        old_files = {name: old_zip.read(name) for name in old_zip.namelist()}
    with ZipFile(generated_path, "r") as new_zip:
        new_files = {name: new_zip.read(name) for name in new_zip.namelist()}

    document_xml = etree.fromstring(old_files["word/document.xml"])
    old_body = document_xml.find(qn("w:body"))
    cover_elements = []
    for child in old_body:
        cover_elements.append(deepcopy(child))
        if child.tag == qn("w:p") and child.find("./w:pPr/w:sectPr", namespaces={"w": qn("w:body").split("}")[0][1:]}) is not None:
            break
    else:
        return False

    cover_text = " ".join(
        text.text or "" for element in cover_elements for text in element.iter(qn("w:t"))
    )
    cover_text = re.sub(r"\s+", " ", cover_text).strip()
    has_picture = any(element.find(".//" + qn("w:drawing")) is not None for element in cover_elements)
    if not has_picture or not cover_text.strip():
        return False

    # The copied cover already has a visible title; remove the generated duplicate.
    if "প্রশ্নোত্তর সংকলন" in cover_text:
        new_root = etree.fromstring(new_files["word/document.xml"])
        new_body = new_root.find(qn("w:body"))
        for child in list(new_body):
            if child.tag != qn("w:p"):
                continue
            style = child.find("./w:pPr/w:pStyle", namespaces={"w": qn("w:body").split("}")[0][1:]})
            if style is not None and style.get(qn("w:val")) == "Title":
                new_body.remove(child)
                break
        new_files["word/document.xml"] = etree.tostring(
            new_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    # A cover section must not show the body header/footer or page numbering.
    cover_section = next(
        child.find("./w:pPr/w:sectPr", namespaces={"w": qn("w:body").split("}")[0][1:]})
        for child in cover_elements
        if child.tag == qn("w:p")
        and child.find("./w:pPr/w:sectPr", namespaces={"w": qn("w:body").split("}")[0][1:]}) is not None
    )
    for tag in ("headerReference", "footerReference", "pgNumType"):
        for node in cover_section.findall(qn("w:" + tag)):
            cover_section.remove(node)

    # Merge only the relationships used by the copied cover (usually its image and link).
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rel_root = etree.fromstring(new_files["word/_rels/document.xml.rels"])
    old_rel_root = etree.fromstring(old_files["word/_rels/document.xml.rels"])
    old_rels = {rel.get("Id"): rel for rel in old_rel_root}
    rel_id_map: dict[str, str] = {}
    next_id = 1
    for rel in rel_root:
        match = re.fullmatch(r"rId(\d+)", rel.get("Id", ""))
        if match:
            next_id = max(next_id, int(match.group(1)) + 1)

    rel_attr_prefix = "{" + "http://schemas.openxmlformats.org/officeDocument/2006/relationships" + "}"
    referenced_ids = {
        value
        for element in cover_elements
        for node in element.iter()
        for attr, value in node.attrib.items()
        if attr.startswith(rel_attr_prefix)
    }
    for old_id in referenced_ids:
        old_rel = old_rels.get(old_id)
        if old_rel is None:
            continue
        old_target = old_rel.get("Target", "")
        relationship_type = old_rel.get("Type", "")
        target_mode = old_rel.get("TargetMode")
        new_target = old_target
        if not target_mode:
            old_part = posixpath.normpath(posixpath.join("word", old_target))
            if old_part not in old_files:
                continue
            ext = posixpath.splitext(old_part)[1]
            new_name = "cover_" + posixpath.basename(old_part)
            new_part = "word/media/" + new_name
            suffix = 2
            while new_part in new_files:
                new_name = f"cover_{suffix}_" + posixpath.basename(old_part)
                new_part = "word/media/" + new_name
                suffix += 1
            new_files[new_part] = old_files[old_part]
            new_target = "media/" + new_name
        new_id = f"rId{next_id}"
        next_id += 1
        attrs = {"Id": new_id, "Type": relationship_type, "Target": new_target}
        if target_mode:
            attrs["TargetMode"] = target_mode
        etree.SubElement(rel_root, f"{{{rel_ns}}}Relationship", **attrs)
        rel_id_map[old_id] = new_id

    for element in cover_elements:
        for node in element.iter():
            for attr, value in list(node.attrib.items()):
                if attr.startswith(rel_attr_prefix) and value in rel_id_map:
                    node.set(attr, rel_id_map[value])

    # Insert the cover before the generated title/contents, then save the merged package.
    new_root = etree.fromstring(new_files["word/document.xml"])
    new_body = new_root.find(qn("w:body"))
    for index, element in enumerate(cover_elements):
        new_body.insert(index, element)
    new_files["word/document.xml"] = etree.tostring(
        new_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    new_files["word/_rels/document.xml.rels"] = etree.tostring(
        rel_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    # The generated package normally already declares PNG; add its media type if needed.
    content_types = etree.fromstring(new_files["[Content_Types].xml"])
    default_png = next(
        (node for node in content_types if node.get("Extension", "").lower() == "png"), None
    )
    if any(name.lower().endswith(".png") and name.startswith("word/media/cover_") for name in new_files) and default_png is None:
        etree.SubElement(
            content_types,
            "{http://schemas.openxmlformats.org/package/2006/content-types}Default",
            Extension="png",
            ContentType="image/png",
        )
    new_files["[Content_Types].xml"] = etree.tostring(
        content_types, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    temp_path = generated_path.with_suffix(".merged.docx")
    with ZipFile(temp_path, "w", ZIP_DEFLATED) as merged_zip:
        for name, data in new_files.items():
            merged_zip.writestr(name, data)
    os.replace(temp_path, generated_path)
    return True


def _append_xml_field_run(paragraph: Any, instruction: str, cached_text: str = "") -> None:
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), LATIN_FONT)
    fonts.set(qn("w:hAnsi"), LATIN_FONT)
    fonts.set(qn("w:eastAsia"), BENGALI_FONT)
    fonts.set(qn("w:cs"), BENGALI_FONT)
    rpr.append(fonts)
    run.append(rpr)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    run.append(begin)
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = f" {instruction} "
    run.append(instr)
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    run.append(separate)
    if cached_text:
        text = OxmlElement("w:t")
        text.text = cached_text
        run.append(text)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run.append(end)
    paragraph._p.append(run)


def add_complex_field(paragraph: Any, instruction: str, cached_text: str = "1") -> None:
    _append_xml_field_run(paragraph, instruction, cached_text)


def _start_toc_field(paragraph: Any) -> None:
    run = OxmlElement("w:r")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    run.append(begin)
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    run.append(instr)
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    run.append(separate)
    paragraph._p.append(run)


def _end_toc_field(paragraph: Any) -> None:
    run = OxmlElement("w:r")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run.append(end)
    paragraph._p.append(run)


def _add_internal_hyperlink(paragraph: Any, text: str, anchor: str) -> None:
    link = OxmlElement("w:hyperlink")
    link.set(qn("w:anchor"), anchor)
    link.set(qn("w:history"), "1")
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), LATIN_FONT)
    fonts.set(qn("w:hAnsi"), LATIN_FONT)
    fonts.set(qn("w:eastAsia"), BENGALI_FONT)
    fonts.set(qn("w:cs"), BENGALI_FONT)
    rpr.append(fonts)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "none")
    rpr.append(underline)
    run.append(rpr)
    value = OxmlElement("w:t")
    value.text = text
    run.append(value)
    link.append(run)
    paragraph._p.append(link)


def _add_page_ref(paragraph: Any, bookmark: str) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), f" PAGEREF {bookmark} \\h ")
    field.set(qn("w:dirty"), "true")
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), LATIN_FONT)
    fonts.set(qn("w:hAnsi"), LATIN_FONT)
    rpr.append(fonts)
    run.append(rpr)
    value = OxmlElement("w:t")
    value.text = "1"
    run.append(value)
    field.append(run)
    paragraph._p.append(field)


def _toc_paragraph(doc: Any, level: int, text: str, bookmark: str, first: bool = False) -> Any:
    paragraph = doc.add_paragraph(style=f"TOC {level}")
    if first:
        _start_toc_field(paragraph)
    _add_internal_hyperlink(paragraph, text, bookmark)
    paragraph.add_run("\t")
    _add_page_ref(paragraph, bookmark)
    return paragraph


def build_document(input_path: Path, output_path: Path) -> tuple[int, int, int]:
    """Build one editable DOCX. Returns (questions, topics, warnings)."""
    parsed = json.loads(input_path.read_text(encoding="utf-8-sig"))
    records = _question_records(parsed)
    subjects, question_count, warnings = build_tree(records, input_path.stem)
    if not subjects:
        raise ValueError("কোনো ব্যবহারযোগ্য question record পাওয়া যায়নি")

    doc = Document()
    configure_document(doc)
    doc.core_properties.title = "প্রশ্নোত্তর সংকলন - " + next(iter(subjects.values()))["name"]
    doc.core_properties.subject = "JSON থেকে তৈরি সম্পাদনাযোগ্য প্রশ্নোত্তর বই"
    doc.core_properties.author = ""

    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_plain_text(title, "প্রশ্নোত্তর সংকলন", bold=True)
    toc_heading = doc.add_paragraph(style="TOC Title")
    add_plain_text(toc_heading, "সূচিপত্র", bold=True)

    toc_entries: list[tuple[int, str, str]] = []
    heading_specs: list[tuple[int, str, str]] = []
    bookmark_id = 1
    topic_total = 0

    for subject_number, subject_data in enumerate(subjects.values(), 1):
        subject_text = f"{_alpha_label(subject_number)}. {subject_data['name']}"
        subject_bookmark = f"subject_{subject_number}"
        toc_entries.append((1, subject_text, subject_bookmark))
        heading_specs.append((1, subject_text, subject_bookmark))
        for chapter_number, chapter_data in enumerate(subject_data["chapters"].values(), 1):
            chapter_text = f"{_alpha_label(chapter_number)}. {chapter_data['name']}"
            chapter_bookmark = f"chapter_{subject_number}_{chapter_number}"
            toc_entries.append((2, chapter_text, chapter_bookmark))
            heading_specs.append((2, chapter_text, chapter_bookmark))
            for topic_number, topic_data in enumerate(chapter_data["topics"].values(), 1):
                topic_text = f"{_roman_label(topic_number)}. {topic_data['name']}"
                topic_bookmark = f"topic_{subject_number}_{chapter_number}_{topic_number}"
                toc_entries.append((3, topic_text, topic_bookmark))
                heading_specs.append((3, topic_text, topic_bookmark))
                topic_total += 1

    toc_paragraphs = []
    for index, (level, text, bookmark) in enumerate(toc_entries):
        toc_paragraphs.append(_toc_paragraph(doc, level, text, bookmark, first=(index == 0)))
    if toc_paragraphs:
        _end_toc_field(toc_paragraphs[-1])

    bookmarks: dict[str, int] = {}
    for _, _, bookmark in heading_specs:
        bookmarks[bookmark] = bookmark_id
        bookmark_id += 1

    first_subject = next(iter(subjects.values()))
    first_chapter = next(iter(first_subject["chapters"].values()))
    body_section = doc.add_section(WD_SECTION.NEW_PAGE)
    _format_section(body_section, 1)
    _set_page_number_start(body_section, 1)
    _set_section_header(body_section, first_chapter["name"], first_subject["name"], unlink=True)
    _set_section_footer(body_section, unlink=True)

    for subject_number, subject_data in enumerate(subjects.values(), 1):
        first_chapter_for_subject = next(iter(subject_data["chapters"].values()))
        if subject_number > 1:
            body_section = doc.add_section(WD_SECTION.NEW_PAGE)
            _format_section(body_section, 1)
            _set_section_header(body_section, first_chapter_for_subject["name"], subject_data["name"], unlink=True)

        subject_text = f"{_alpha_label(subject_number)}. {subject_data['name']}"
        subject_bookmark = f"subject_{subject_number}"
        paragraph = doc.add_paragraph(style="Heading 1")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_plain_text(paragraph, subject_text, bold=True)
        add_bookmark(paragraph, subject_bookmark, bookmarks[subject_bookmark])

        for chapter_number, chapter_data in enumerate(subject_data["chapters"].values(), 1):
            if chapter_number > 1:
                body_section = doc.add_section(WD_SECTION.NEW_PAGE)
                _format_section(body_section, 1)
                _set_section_header(body_section, chapter_data["name"], subject_data["name"], unlink=True)

            chapter_text = f"{_alpha_label(chapter_number)}. {chapter_data['name']}"
            chapter_bookmark = f"chapter_{subject_number}_{chapter_number}"
            paragraph = doc.add_paragraph(style="Heading 2")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_plain_text(paragraph, chapter_text, bold=True)
            add_bookmark(paragraph, chapter_bookmark, bookmarks[chapter_bookmark])

            # Switch to 2-column continuous section for questions
            q_section = doc.add_section(WD_SECTION.CONTINUOUS)
            _format_section(q_section, 2)
            q_section.header.is_linked_to_previous = True
            q_section.footer.is_linked_to_previous = True

            for topic_number, topic_data in enumerate(chapter_data["topics"].values(), 1):
                topic_text = f"{_roman_label(topic_number)}. {topic_data['name']}"
                topic_bookmark = f"topic_{subject_number}_{chapter_number}_{topic_number}"
                paragraph = doc.add_paragraph(style="Heading 3")
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_plain_text(paragraph, topic_text, bold=True)
                add_bookmark(paragraph, topic_bookmark, bookmarks[topic_bookmark])

                # Full-page-width headings remain in normal document flow;
                # each topic's editable Q/A is balanced into two table cells.
                _add_question_columns(doc, topic_data["questions"])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    # Generate to a temporary file first so an edited image cover in an existing
    # output book can be carried forward without risking the current DOCX.
    temp_handle = tempfile.NamedTemporaryFile(
        prefix=output_path.stem + "_",
        suffix=".docx",
        dir=output_path.parent,
        delete=False,
    )
    temp_path = Path(temp_handle.name)
    temp_handle.close()
    try:
        doc.save(temp_path)
        if output_path.exists():
            _preserve_existing_cover(output_path, temp_path)
        os.replace(temp_path, output_path)
        _update_toc_page_cache(output_path, heading_specs)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return question_count, topic_total, warnings


def _safe_stem(stem: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem).strip(" .")
    return cleaned or "book"


def _format_duration(seconds: float) -> str:
    """সেকেন্ডকে মানুষের পড়ার উপযোগী বাংলা বয়ানে (ঘণ্টা, মিনিট, সেকেন্ড) রূপান্তর করে।"""
    seconds = max(0, int(round(seconds)))
    mins, secs = divmod(seconds, 60)
    hours, mins = divmod(mins, 60)
    parts = []
    if hours > 0:
        parts.append(f"{hours} ঘণ্টা")
    if mins > 0:
        parts.append(f"{mins} মিনিট")
    if secs > 0 or not parts:
        parts.append(f"{secs} সেকেন্ড")
    return " ".join(parts)


def run_batch(folder: Path) -> int:
    input_files = sorted(
        (path for path in folder.iterdir() if path.is_file() and path.suffix.lower() == ".json"),
        key=lambda path: path.name.casefold(),
    )
    if not input_files:
        print(f"এই ফোল্ডারে কোনো JSON ফাইল পাওয়া যায়নি: {folder}")
        return 1

    total_files = len(input_files)
    output_dir = folder / "output"
    output_dir.mkdir(exist_ok=True)

    print("=" * 65)
    print(f"মোট JSON ফাইল পাওয়া গেছে: {total_files} টি")
    print(f"আউটপুট ফোল্ডার: {output_dir}")
    print("=" * 65)

    successes = 0
    total_questions = 0
    failures = []
    batch_start_time = time.time()

    for idx, input_path in enumerate(input_files, start=1):
        elapsed_so_far = time.time() - batch_start_time
        processed_count = idx - 1

        if processed_count > 0:
            avg_per_file = elapsed_so_far / processed_count
            remaining_files = total_files - processed_count
            eta_seconds = avg_per_file * remaining_files
            eta_str = f"বাকি আনুমানিক: {_format_duration(eta_seconds)}"
        else:
            eta_str = "বাকি আনুমানিক: হিসাব করা হচ্ছে..."

        print(f"\n[{idx}/{total_files}] কাজ চলছে: {input_path.name} ({eta_str})")
        file_start_time = time.time()
        output_path = output_dir / (_safe_stem(input_path.stem) + ".docx")
        try:
            questions, topics, warnings = build_document(input_path, output_path)
            file_duration = time.time() - file_start_time
            successes += 1
            total_questions += questions
            print(f"  ✔ তৈরি: {output_path.name} | {questions}টি প্রশ্ন | {topics}টি topic | সময় লেগেছে: {_format_duration(file_duration)}")
            if warnings:
                print(f"  ⚠ সতর্কতা: {warnings}টি প্রশ্নের উত্তর/প্রশ্ন অনুপস্থিত; বিস্তারিত বইয়ে চিহ্নিত আছে।")
        except Exception as exc:
            file_duration = time.time() - file_start_time
            failures.append((input_path.name, exc))
            print(f"  ✖ ব্যর্থ: {input_path.name} (সময়: {_format_duration(file_duration)}) | {exc}")

    total_elapsed = time.time() - batch_start_time
    print("\n" + "=" * 65)
    print(f"সম্পন্ন: {successes}/{total_files}টি বই | মোট {total_questions}টি প্রশ্ন")
    print(f"সর্বমোট সময় লেগেছে: {_format_duration(total_elapsed)}")
    print(f"আউটপুট ফোল্ডার: {output_dir}")
    if failures:
        print("ব্যর্থ ফাইল:")
        for filename, error in failures:
            print(f"  - {filename}: {error}")
        return 1
    return 0


def main() -> int:
    folder = Path(__file__).resolve().parent
    return run_batch(folder)


if __name__ == "__main__":
    exit_code = main()
    if os.name == "nt" and sys.stdin.isatty():
        try:
            input("কাজ শেষ। বন্ধ করতে Enter চাপুন...")
        except EOFError:
            pass
    raise SystemExit(exit_code)
