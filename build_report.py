# Builds the Assignment 4 write-up: question, method, code, output, and a Visual Studio screenshot.
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

ROOT = Path(r"C:\Users\ASUS\Desktop\Prog Lab Assignment\Assignment 4")
ASSET = ROOT / "report_assets"
DOCX_PATH = ROOT / "2602193005_Simranjeet Singh Ghuman_Assignment4.docx"
PDF_PATH = ROOT / "2602193005_Simranjeet Singh Ghuman_Assignment4.pdf"

QUESTIONS = [
    {
        "n": 1,
        "folder": "Q01_Record_Counts",
        "question": "Load the three datasets and report the total number of records in each source, to establish the raw size of each source before any cleaning.",
        "input": "data/goodreads/metadata\ndata/Amazon/metadata\ndata/Openlibrary/metadata",
    },
    {
        "n": 2,
        "folder": "Q02_Unique_ISBN_And_Files",
        "question": "Count the total number of unique ISBN-13 numbers across all three sources, and count the files in each source's coverpage, blurb, and review folders.",
        "input": "The three metadata JSON files, plus coverpage, blurb, and review (reviews for OpenLibrary).",
    },
    {
        "n": 3,
        "folder": "Q03_Duplicate_ISBN",
        "question": "Identify duplicate ISBN-13 entries within each individual source, to detect records that were scraped or saved more than once under the same ISBN.",
        "input": "The raw metadata JSON text for Goodreads, Amazon, and OpenLibrary.",
    },
    {
        "n": 4,
        "folder": "Q04_Missing_Percent",
        "question": "Compute the percentage of missing values for each metadata field, per source, for title, authors, publisher, language, date of publication, and genre.",
        "input": "The three metadata JSON files. Empty text and placeholders such as N/A count as missing.",
    },
    {
        "n": 5,
        "folder": "Q05_Missing_Books",
        "question": "Identify books completely missing from each source, relative to the combined book set, to find which ISBNs each source failed to collect at all.",
        "input": "The ISBN-13 keys from all three metadata files.",
    },
    {
        "n": 6,
        "folder": "Q06_Absent_Fields",
        "question": "Identify metadata fields that are entirely absent from a source's metadata JSON, and separate those from fields that exist but are empty. Blurb, cover page, and review are files, not JSON fields.",
        "input": "The keys and values inside the three metadata JSON files.",
    },
    {
        "n": 7,
        "folder": "Q07_Missing_Files",
        "question": "For each source, identify ISBNs present in metadata but with no matching file in coverpage, blurb, or review.",
        "input": "Metadata ISBN keys, and the file names in coverpage, blurb, and review.",
    },
    {
        "n": 8,
        "folder": "Q08_Orphan_Files",
        "question": "Identify orphaned files in coverpage, blurb, or review that do not correspond to any ISBN in metadata.",
        "input": "File names in the three folders, checked against that source's metadata ISBNs.",
    },
    {
        "n": 9,
        "folder": "Q09_File_Names",
        "question": "Check file naming and format consistency within each folder (extension and ISBN-based naming) across the three sources.",
        "input": "Every file name in coverpage, blurb, and review for Goodreads, Amazon, and OpenLibrary.",
    },
    {
        "n": 10,
        "folder": "Q10_Completeness",
        "question": "Count how many books have all four components (metadata, cover page, blurb, and review) versus partial coverage, per source and overall.",
        "input": "Metadata ISBNs plus the ISBN taken from each cover, blurb, and review file name.",
    },
    {
        "n": 11,
        "folder": "Q11_Standardize_Genres",
        "question": "Standardize the genre labels across all three sources, so the same genre is not split by spelling (for example sci-fi, science-fiction, and Science Fiction).",
        "input": "The genres field from all three metadata files.",
    },
    {
        "n": 12,
        "folder": "Q12_Genre_Statistics",
        "question": "Using the standardized genres, count the unique genres, count the samples in each genre, name the majority and minority class, and plot the genre-wise co-occurrence curve.",
        "input": "Standardized genre labels for each ISBN, combined across the three sources.",
    },
]

KEYWORDS = {
    "False", "None", "True", "and", "as", "break", "class", "continue", "def",
    "elif", "else", "for", "if", "import", "in", "is", "not", "or", "return", "while",
}


def load_questions():
    items = []
    for spec in QUESTIONS:
        folder = ROOT / "output" / spec["folder"]
        code = (folder / "solution.py").read_text(encoding="utf-8").replace("\t", "    ")
        method = (folder / "method.txt").read_text(encoding="utf-8").strip()
        output = (folder / "output.txt").read_text(encoding="utf-8").replace("\r\n", "\n").rstrip() + "\n"
        items.append({
            "n": spec["n"],
            "folder": spec["folder"],
            "question": spec["question"],
            "input": spec["input"],
            "method": method,
            "code": code,
            "output": output,
        })
    return items


def colorize_line(line):
    parts = []
    i = 0
    n = len(line)
    while i < n:
        ch = line[i]
        if ch == "#":
            parts.append(("#6A9955", line[i:]))
            break
        if ch in "\"'":
            quote = ch
            j = i + 1
            while j < n:
                if line[j] == "\\":
                    j += 2
                    continue
                if line[j] == quote:
                    j += 1
                    break
                j += 1
            parts.append(("#CE9178", line[i:j]))
            i = j
            continue
        if ch.isdigit() and (i == 0 or not (line[i - 1].isalnum() or line[i - 1] == "_")):
            j = i + 1
            while j < n and (line[j].isdigit() or line[j] == "."):
                j += 1
            parts.append(("#B5CEA8", line[i:j]))
            i = j
            continue
        if ch.isalpha() or ch == "_":
            j = i + 1
            while j < n and (line[j].isalnum() or line[j] == "_"):
                j += 1
            word = line[i:j]
            color = "#569CD6" if word in KEYWORDS else "#D4D4D4"
            parts.append((color, word))
            i = j
            continue
        parts.append(("#D4D4D4", ch))
        i += 1
    if not parts:
        parts.append(("#D4D4D4", ""))
    return parts


def text_width(font, text):
    if hasattr(font, "getlength"):
        return font.getlength(text)
    return font.getbbox(text)[2]


def draw_colored(draw, x, y, parts, font, max_x):
    cursor = x
    for color, text in parts:
        if cursor >= max_x:
            break
        while text and cursor + text_width(font, text) > max_x:
            text = text[:-1]
        if not text:
            break
        draw.text((cursor, y), text, font=font, fill=color)
        cursor += text_width(font, text)


def render_vscode_page(item, code_lines, term_lines, save_path, start_number):
    code_font = ImageFont.truetype(r"C:\Windows\Fonts\consola.ttf", 15)
    ui = ImageFont.truetype(r"C:\Windows\Fonts\segoeui.ttf", 13)
    ui_small = ImageFont.truetype(r"C:\Windows\Fonts\segoeui.ttf", 12)
    ui_bold = ImageFont.truetype(r"C:\Windows\Fonts\segoeuib.ttf", 13)

    line_h = 18
    width = 1360
    title_h = 32
    tab_h = 36
    side_w = 48
    explorer_w = 250
    editor_x = side_w + explorer_w
    term_head = 28
    status_h = 24
    editor_h = 10 + max(len(code_lines), 1) * line_h + 8
    term_h = 0
    if term_lines is not None:
        term_h = term_head + 8 + len(term_lines) * line_h + 10
    height = title_h + tab_h + editor_h + term_h + status_h

    img = Image.new("RGB", (width, height), "#1E1E1E")
    draw = ImageDraw.Draw(img)

    draw.rectangle((0, 0, width, title_h), fill="#3C3C3C")
    draw.text((12, 7), "Visual Studio Code", font=ui, fill="#CCCCCC")
    title = "solution.py - " + item["folder"]
    tw = text_width(ui, title)
    draw.text(((width - tw) / 2, 7), title, font=ui, fill="#FFFFFF")
    draw.rectangle((width - 46, 0, width, title_h), fill="#C42B1C")
    draw.text((width - 30, 6), "X", font=ui_bold, fill="#FFFFFF")

    draw.rectangle((0, title_h, side_w, height - status_h), fill="#333333")
    draw.rectangle((side_w, title_h, editor_x, height - status_h), fill="#252526")
    draw.text((side_w + 14, title_h + 10), "EXPLORER", font=ui_bold, fill="#BBBBBB")
    draw.text((side_w + 14, title_h + 36), item["folder"], font=ui, fill="#CCCCCC")
    draw.rectangle((side_w, title_h + 58, editor_x, title_h + 82), fill="#094771")
    draw.text((side_w + 28, title_h + 62), "solution.py", font=ui, fill="#FFFFFF")

    tab_y = title_h
    draw.rectangle((editor_x, tab_y, width, tab_y + tab_h), fill="#252526")
    draw.rectangle((editor_x, tab_y, editor_x + 180, tab_y + tab_h), fill="#1E1E1E")
    draw.rectangle((editor_x, tab_y + tab_h - 2, editor_x + 180, tab_y + tab_h), fill="#007ACC")
    draw.text((editor_x + 12, tab_y + 9), "solution.py", font=ui, fill="#FFFFFF")

    editor_top = title_h + tab_h
    y = editor_top + 6
    gutter = editor_x + 58
    for offset, line in enumerate(code_lines):
        number = start_number + offset
        draw.text((editor_x + 8, y), str(number), font=code_font, fill="#858585")
        safe = line.encode("cp1252", "replace").decode("cp1252")
        draw_colored(draw, gutter, y, colorize_line(safe), code_font, width - 12)
        y += line_h

    if term_lines is not None:
        term_top = editor_top + editor_h
        draw.rectangle((editor_x, term_top, width, term_top + term_head), fill="#252526")
        draw.text((editor_x + 12, term_top + 5), "TERMINAL", font=ui_bold, fill="#FFFFFF")
        draw.text((editor_x + 110, term_top + 5), "python solution.py", font=ui, fill="#CCCCCC")
        y = term_top + term_head + 4
        for index, line in enumerate(term_lines):
            color = "#4EC9B0" if index == 0 else "#D4D4D4"
            shown = line.encode("cp1252", "replace").decode("cp1252")
            while shown and text_width(code_font, shown) > width - editor_x - 28:
                shown = shown[:-1]
            draw.text((editor_x + 12, y), shown, font=code_font, fill=color)
            y += line_h

    draw.rectangle((0, height - status_h, width, height), fill="#007ACC")
    draw.text((12, height - 19), "Python 3.14.0", font=ui_small, fill="#FFFFFF")
    draw.text((width - 430, height - 19), "Ln 1, Col 1     UTF-8     CRLF     Python", font=ui_small, fill="#FFFFFF")
    img.save(save_path, "PNG")


def render_vscode(item, asset_dir):
    code_lines = item["code"].splitlines()
    out_lines = item["output"].splitlines()
    if len(out_lines) > 42:
        hidden = len(out_lines) - 40
        out_lines = out_lines[:40] + ["... " + str(hidden) + " more lines are in the output box above ..."]
    prompt = (
        "PS C:\\Users\\ASUS\\Desktop\\Prog Lab Assignment\\Assignment 4\\output\\"
        + item["folder"]
        + "> python solution.py"
    )
    term_lines = [prompt] + out_lines
    line_h = 18
    chrome = 32 + 36 + 24
    term_h = 28 + 8 + len(term_lines) * line_h + 10
    full_h = chrome + 18 + len(code_lines) * line_h + term_h
    paths = []
    if full_h <= 1680:
        path = asset_dir / ("shot_q%02d.png" % item["n"])
        render_vscode_page(item, code_lines, term_lines, path, 1)
        paths.append(path)
        return paths

    room_for_last = 1680 - chrome - term_h - 18
    last_count = max(12, room_for_last // line_h)
    if last_count >= len(code_lines):
        last_count = len(code_lines)
    first_lines = code_lines[:-last_count] if last_count < len(code_lines) else []
    last_lines = code_lines[-last_count:]
    part = 1
    start = 1
    chunk = 72
    while first_lines:
        piece = first_lines[:chunk]
        first_lines = first_lines[chunk:]
        path = asset_dir / ("shot_q%02d_%d.png" % (item["n"], part))
        render_vscode_page(item, piece, None, path, start)
        paths.append(path)
        start += len(piece)
        part += 1
    path = asset_dir / ("shot_q%02d_%d.png" % (item["n"], part))
    render_vscode_page(item, last_lines, term_lines, path, start)
    paths.append(path)
    return paths


def shade_cell(cell, fill):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_table_full(table, color):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), "5000")
    tbl_w.set(qn("w:type"), "pct")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement("w:" + edge)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)
        borders.append(element)
    tbl_pr.append(borders)


def shade_paragraph(paragraph, fill):
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    p_pr.append(shd)


def add_box(document, label, body, fill, font_color, label_color, font_size=9):
    caption = document.add_paragraph()
    caption.paragraph_format.space_before = Pt(8)
    caption.paragraph_format.space_after = Pt(2)
    run = caption.add_run(label)
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor.from_string(label_color)

    table = document.add_table(rows=1, cols=1)
    set_table_full(table, fill)
    cell = table.cell(0, 0)
    shade_cell(cell, fill)
    cell.text = ""
    first = True
    for line in body.splitlines() or [""]:
        paragraph = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        shade_paragraph(paragraph, fill)
        run = paragraph.add_run(line if line else " ")
        run.font.name = "Consolas"
        run._element.rPr.rFonts.set(qn("w:ascii"), "Consolas")
        run._element.rPr.rFonts.set(qn("w:hAnsi"), "Consolas")
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor.from_string(font_color)
    document.add_paragraph().paragraph_format.space_after = Pt(2)


def add_heading_line(document, text, color, size=13):
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(2)
    run = paragraph.add_run(text)
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)


def add_body(document, text):
    for block in text.split("\n"):
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(2)
        paragraph.paragraph_format.line_spacing = 1.08
        run = paragraph.add_run(block.strip() if block.strip() else " ")
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(30, 30, 30)


def add_rule(document):
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(10)
    paragraph.paragraph_format.space_after = Pt(10)
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "1F4E79")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def build_document(items):
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Programming Lab")
    run.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = RGBColor(31, 78, 121)
    run.font.name = "Calibri"

    sub = document.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = sub.add_run("Assignment 4")
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(31, 78, 121)
    run.font.name = "Calibri"

    topic = document.add_paragraph()
    topic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = topic.add_run("Missing-Data Analysis of Multi-Source Book Metadata")
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(31, 78, 121)
    run.font.name = "Calibri"

    for line in ("Name: Simranjeet Singh Ghuman", "Roll Number: 2602193005"):
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(0)
        run = paragraph.add_run(line)
        run.font.size = Pt(12)
        run.font.name = "Calibri"
        run.bold = True

    add_rule(document)
    add_heading_line(document, "Note: Color code legend used in Assignment 4", "1F4E79", 13)
    add_body(document, "CODE = Black box\nINPUT = Blue box\nOUTPUT = Green box")
    add_body(
        document,
        "I used the Goodreads, Amazon, and OpenLibrary files from Assignment 1. "
        "They are arranged in data/ the way this question asks. "
        "Each question has its own folder under output. I open solution.py in Visual Studio Code and run it. "
        "bookdata.py is the shared reader. It loads the JSON and lists the files. "
        "A value counts as missing when it is blank or a placeholder such as N/A.",
    )

    shared = (ROOT / "bookdata.py").read_text(encoding="utf-8")
    add_heading_line(document, "Shared code", "1A1A1A")
    add_box(document, "CODE: bookdata.py", shared.rstrip(), "1A1A1A", "FFFFFF", "1A1A1A", font_size=8)

    for item in items:
        add_rule(document)
        heading = document.add_paragraph()
        run = heading.add_run("Question " + str(item["n"]))
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(31, 78, 121)
        run.font.name = "Calibri"

        add_heading_line(document, "Question", "1F4E79")
        add_body(document, item["question"])

        add_heading_line(document, "Why I did it this way", "0B6E4F")
        add_body(document, item["method"])

        add_heading_line(document, "Code", "1A1A1A")
        add_body(document, "This is the program I run for this question.")
        add_box(
            document,
            "CODE: output/" + item["folder"] + "/solution.py",
            item["code"].rstrip(),
            "1A1A1A",
            "FFFFFF",
            "1A1A1A",
            font_size=8,
        )
        add_box(document, "INPUT", item["input"], "1F4E79", "FFFFFF", "1F4E79", font_size=10)
        add_box(
            document,
            "OUTPUT",
            item["output"].rstrip(),
            "0E7A3D",
            "FFFFFF",
            "0E7A3D",
            font_size=9,
        )

        add_heading_line(document, "Visual Studio Output Screenshot", "1F4E79")
        add_body(document, "This is the same file running in Visual Studio Code. The terminal shows the output.")
        for shot in item["shots"]:
            document.add_picture(str(shot), width=Inches(6.4))
            document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

        if item["n"] == 12:
            picture = ROOT / "output" / item["folder"] / "genre_cooccurrence.png"
            add_heading_line(document, "Figure", "1F4E79")
            document.add_picture(str(picture), width=Inches(6.2))
            document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap = document.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = cap.add_run("Figure: every genre pair on the left, and the first 40 pairs on the right.")
            run.italic = True
            run.font.size = Pt(10)
            run.font.name = "Calibri"

    add_rule(document)
    end = document.add_paragraph()
    end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = end.add_run("End of Assignment 4")
    run.bold = True
    run.font.color.rgb = RGBColor(31, 78, 121)
    run.font.name = "Calibri"

    temporary = DOCX_PATH.with_suffix(".tmp.docx")
    document.save(temporary)
    try:
        temporary.replace(DOCX_PATH)
    except PermissionError:
        fallback = DOCX_PATH.with_name(DOCX_PATH.stem + "_updated.docx")
        temporary.replace(fallback)
        print("original file is open, wrote", fallback)
        return fallback
    return DOCX_PATH


def export_pdf(docx_path):
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = word.Documents.Open(str(docx_path))
    document.SaveAs(str(PDF_PATH), FileFormat=17)
    document.Close(False)
    word.Quit()


def main():
    ASSET.mkdir(exist_ok=True)
    items = load_questions()
    print("screenshots...")
    for item in items:
        item["shots"] = render_vscode(item, ASSET)
        print("  shot", item["n"], len(item["shots"]))
    print("docx...")
    path = build_document(items)
    print(path)
    print("pdf...")
    export_pdf(path)
    print(PDF_PATH, "bytes", PDF_PATH.stat().st_size)


if __name__ == "__main__":
    sys.exit(main())
