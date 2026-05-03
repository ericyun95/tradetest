import io
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


def _set_korean_font(run, font_name="Malgun Gothic"):
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:cs"), font_name)


def _apply_inline_bold(paragraph, text: str):
    # 헤딩에서 넘어온 ** ** 마커 제거 후 처리
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
            _set_korean_font(run)
        elif part:
            run = paragraph.add_run(part)
            _set_korean_font(run)


def _strip_bold(text: str) -> str:
    return re.sub(r"\*\*([^*]+)\*\*", r"\1", text)


def _add_hr(doc: Document):
    para = doc.add_paragraph()
    pPr = para._p.get_or_add_pPr()
    pBdr = pPr.get_or_insert_element_before(
        "w:pBdr",
        ["w:jc", "w:textAlignment", "w:outlineLvl"],
    ) if False else None
    # 간단하게 빈 줄로 대체
    return


def _add_table(doc: Document, lines: list[str]):
    rows = [l for l in lines if l.strip().startswith("|")]
    if len(rows) < 2:
        return

    header_cells = [c.strip() for c in rows[0].strip("|").split("|")]
    data_rows = []
    for r in rows[2:]:
        if re.fullmatch(r"[\|\s\-:]+", r):
            continue
        data_rows.append([c.strip() for c in r.strip("|").split("|")])

    col_count = len(header_cells)
    table = doc.add_table(rows=1 + len(data_rows), cols=col_count)
    table.style = "Table Grid"

    hdr_cells = table.rows[0].cells
    for i, cell_text in enumerate(header_cells):
        if i < len(hdr_cells):
            p = hdr_cells[i].paragraphs[0]
            run = p.add_run(_strip_bold(cell_text))
            run.bold = True
            _set_korean_font(run)

    for ri, row_data in enumerate(data_rows):
        cells = table.rows[ri + 1].cells
        for ci, cell_text in enumerate(row_data):
            if ci < len(cells):
                p = cells[ci].paragraphs[0]
                run = p.add_run(_strip_bold(cell_text))
                _set_korean_font(run)

    doc.add_paragraph()


def _set_doc_default_font(doc: Document):
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
    rFonts.set(qn("w:cs"), "Malgun Gothic")


def markdown_to_docx(
    markdown: str,
    product_name_ko: str,
    hs_code_display: str,
) -> bytes:
    doc = Document()
    _set_doc_default_font(doc)

    # 문서 제목
    title = doc.add_heading("무역 조사 보고서", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = sub.add_run(f"{product_name_ko}  |  HS Code {hs_code_display}")
    run.bold = True
    _set_korean_font(run)
    doc.add_paragraph()

    lines = markdown.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # 구분선(---) → 스킵
        if re.fullmatch(r"-{3,}", stripped):
            i += 1
            continue

        # 빈 줄
        if not stripped:
            i += 1
            continue

        # 표: 연속된 | 라인 묶음
        if stripped.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i])
                i += 1
            _add_table(doc, table_lines)
            continue

        # 제목 처리 (** 마커 제거 후 헤딩으로)
        if line.startswith("#### "):
            h = doc.add_heading(_strip_bold(line[5:].strip()), level=4)
            for run in h.runs:
                _set_korean_font(run)
        elif line.startswith("### "):
            h = doc.add_heading(_strip_bold(line[4:].strip()), level=3)
            for run in h.runs:
                _set_korean_font(run)
        elif line.startswith("## "):
            h = doc.add_heading(_strip_bold(line[3:].strip()), level=2)
            for run in h.runs:
                _set_korean_font(run)
        elif line.startswith("# "):
            h = doc.add_heading(_strip_bold(line[2:].strip()), level=1)
            for run in h.runs:
                _set_korean_font(run)

        # 불릿 리스트
        elif stripped.startswith("- "):
            para = doc.add_paragraph(style="List Bullet")
            _apply_inline_bold(para, stripped[2:])

        # 번호 리스트
        elif re.match(r"^\d+[.)]\s", stripped):
            para = doc.add_paragraph(style="List Number")
            _apply_inline_bold(para, re.sub(r"^\d+[.)]\s*", "", stripped))

        # 일반 문단
        else:
            para = doc.add_paragraph()
            _apply_inline_bold(para, stripped)

        i += 1

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
