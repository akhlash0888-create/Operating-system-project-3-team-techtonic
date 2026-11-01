from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "reports"
ASSETS = ROOT / "tmp" / "diagram_assets"
OUTPUT = OUT / "Services_Monitor_Activity_and_Use_Case_Diagrams_Bengali_Submission.docx"
FONT = r"C:\Windows\Fonts\Nirmala.ttc"

NAVY = (31, 78, 120)
BLUE = (68, 114, 148)
PALE = (234, 242, 248)
LIGHT = (247, 249, 251)
DARK = (30, 30, 30)
GRAY = (110, 110, 110)
BORDER = (170, 180, 190)
WHITE = (255, 255, 255)


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    # Nirmala UI TTC index 1 is the semibold face on standard Windows builds.
    try:
        return ImageFont.truetype(FONT, size=size, index=1 if bold else 0)
    except OSError:
        return ImageFont.truetype(FONT, size=size)


F_TITLE = font(54, True)
F_HEAD = font(38, True)
F_BODY = font(31)
F_SMALL = font(26)
F_TINY = font(23)


def centered_text(draw: ImageDraw.ImageDraw, box, text: str, fnt, fill=DARK, spacing=7) -> None:
    x1, y1, x2, y2 = box
    bbox = draw.multiline_textbbox((0, 0), text, font=fnt, spacing=spacing, align="center")
    width = bbox[2] - bbox[0]
    height = bbox[3] - bbox[1]
    draw.multiline_text(
        ((x1 + x2 - width) / 2, (y1 + y2 - height) / 2 - bbox[1]),
        text,
        font=fnt,
        fill=fill,
        spacing=spacing,
        align="center",
    )


def rounded(draw, box, text, fill=PALE, outline=NAVY, radius=28, fnt=F_BODY) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=5)
    centered_text(draw, box, text, fnt)


def ellipse(draw, box, text, fill=WHITE, outline=NAVY, fnt=F_BODY) -> None:
    draw.ellipse(box, fill=fill, outline=outline, width=5)
    centered_text(draw, box, text, fnt)


def diamond(draw, center, width, height, text, fill=PALE, outline=NAVY, fnt=F_BODY) -> None:
    cx, cy = center
    points = [(cx, cy - height // 2), (cx + width // 2, cy), (cx, cy + height // 2), (cx - width // 2, cy)]
    draw.polygon(points, fill=fill)
    draw.line(points + [points[0]], fill=outline, width=5, joint="curve")
    centered_text(draw, (cx - width // 2 + 30, cy - height // 2 + 20, cx + width // 2 - 30, cy + height // 2 - 20), text, fnt)


def arrow(draw, start, end, fill=NAVY, width=6) -> None:
    draw.line([start, end], fill=fill, width=width)
    import math

    angle = math.atan2(end[1] - start[1], end[0] - start[0])
    length = 24
    spread = 0.55
    p1 = (end[0] - length * math.cos(angle - spread), end[1] - length * math.sin(angle - spread))
    p2 = (end[0] - length * math.cos(angle + spread), end[1] - length * math.sin(angle + spread))
    draw.polygon([end, p1, p2], fill=fill)


def poly_arrow(draw, points, fill=NAVY, width=6) -> None:
    draw.line(points, fill=fill, width=width, joint="curve")
    arrow(draw, points[-2], points[-1], fill=fill, width=width)


def actor(draw, x, y, label: str) -> None:
    draw.ellipse((x - 40, y - 120, x + 40, y - 40), outline=DARK, width=6)
    draw.line((x, y - 40, x, y + 90), fill=DARK, width=6)
    draw.line((x - 75, y + 10, x + 75, y + 10), fill=DARK, width=6)
    draw.line((x, y + 90, x - 65, y + 175), fill=DARK, width=6)
    draw.line((x, y + 90, x + 65, y + 175), fill=DARK, width=6)
    centered_text(draw, (x - 135, y + 190, x + 135, y + 260), label, F_BODY)


def database(draw, x1, y1, x2, y2, label: str) -> None:
    cap = 55
    draw.rectangle((x1, y1 + cap // 2, x2, y2 - cap // 2), fill=LIGHT, outline=NAVY, width=5)
    draw.ellipse((x1, y1, x2, y1 + cap), fill=PALE, outline=NAVY, width=5)
    draw.arc((x1, y2 - cap, x2, y2), 0, 180, fill=NAVY, width=5)
    centered_text(draw, (x1 + 20, y1 + 50, x2 - 20, y2 - 25), label, F_BODY)


def save_use_case_diagram(path: Path) -> None:
    image = Image.new("RGB", (1800, 2150), WHITE)
    draw = ImageDraw.Draw(image)
    centered_text(draw, (80, 35, 1720, 135), "Windows Services Monitor Use Case Diagram", F_TITLE, NAVY)

    boundary = (390, 210, 1420, 1960)
    draw.rounded_rectangle(boundary, radius=35, fill=(252, 253, 254), outline=NAVY, width=6)
    centered_text(draw, (500, 225, 1310, 300), "System Boundary  Windows Services Monitor", F_HEAD, NAVY)

    cases = [
        (520, 355, 1290, 540, "Live service তালিকা দেখা"),
        (520, 610, 1290, 795, "Snapshot সংরক্ষণ করা"),
        (520, 865, 1290, 1050, "সংরক্ষিত snapshot তালিকা দেখা"),
        (520, 1120, 1290, 1305, "দুই snapshot তুলনা করা"),
        (520, 1375, 1290, 1560, "একই PID এর service group দেখা"),
        (520, 1630, 1290, 1815, "পরিবর্তন ধারাবাহিকভাবে পর্যবেক্ষণ করা"),
    ]
    for box in cases:
        ellipse(draw, box[:4], box[4], fill=PALE)

    actor(draw, 190, 990, "ব্যবহারকারী")
    actor(draw, 1600, 620, "Windows SCM")
    database(draw, 1500, 1275, 1730, 1555, "SQLite\nDatabase")

    user_anchor = (325, 1000)
    for box in cases:
        target = (box[0], (box[1] + box[3]) // 2)
        draw.line([user_anchor, target], fill=GRAY, width=4)

    # SCM supplies live service data to live, snapshot, group and watch use cases.
    scm_anchor = (1465, 630)
    for index in (0, 1, 4, 5):
        box = cases[index]
        target = (box[2], (box[1] + box[3]) // 2)
        draw.line([scm_anchor, target], fill=BLUE, width=4)

    # SQLite serves persistence-oriented use cases.
    db_anchor = (1490, 1410)
    for index in (1, 2, 3, 4):
        box = cases[index]
        target = (box[2], (box[1] + box[3]) // 2)
        draw.line([db_anchor, target], fill=(93, 122, 77), width=4)

    draw.line((105, 2040, 305, 2040), fill=GRAY, width=5)
    draw.text((325, 2020), "ব্যবহারকারীর interaction", font=F_TINY, fill=DARK)
    draw.line((695, 2040, 895, 2040), fill=BLUE, width=5)
    draw.text((915, 2020), "SCM data flow", font=F_TINY, fill=DARK)
    draw.line((1240, 2040, 1440, 2040), fill=(93, 122, 77), width=5)
    draw.text((1460, 2020), "Database flow", font=F_TINY, fill=DARK)
    image.save(path, quality=95)


def save_activity_diagram(path: Path) -> None:
    image = Image.new("RGB", (1800, 2700), WHITE)
    draw = ImageDraw.Draw(image)
    centered_text(draw, (80, 30, 1720, 125), "Windows Services Monitor Activity Diagram", F_TITLE, NAVY)

    draw.ellipse((825, 150, 975, 300), fill=DARK)
    draw.text((995, 200), "শুরু", font=F_SMALL, fill=DARK)
    rounded(draw, (565, 360, 1235, 515), "Command ও option গ্রহণ করা")
    diamond(draw, (900, 700), 660, 260, "Command valid কি না", fnt=F_BODY)
    arrow(draw, (900, 300), (900, 360))
    arrow(draw, (900, 515), (900, 570))

    # Invalid-command branch.
    rounded(draw, (1280, 615, 1710, 785), "Error message দেখানো", fill=(255, 242, 242), outline=(156, 58, 58), fnt=F_SMALL)
    poly_arrow(draw, [(1230, 700), (1280, 700)], fill=(156, 58, 58))
    draw.text((1195, 650), "না", font=F_TINY, fill=(156, 58, 58))

    # One selected command handler is executed. The rows summarize every possible handler.
    panel = (300, 910, 1500, 1590)
    draw.rounded_rectangle(panel, radius=35, fill=(252, 253, 254), outline=NAVY, width=6)
    centered_text(draw, (380, 930, 1420, 1010), "নির্বাচিত command handler চালানো", F_HEAD, NAVY)
    rows = [
        ("live", "SCM থেকে live data নিয়ে table দেখানো"),
        ("snapshot", "Live data নিয়ে SQLite এ সংরক্ষণ"),
        ("snapshots", "Database থেকে snapshot তালিকা দেখানো"),
        ("diff", "দুই snapshot load করে পরিবর্তন তুলনা"),
        ("groups", "PID অনুযায়ী service group তৈরি"),
        ("watch", "Interval পরপর live data নিয়ে diff করা"),
    ]
    row_top = 1030
    for index, (command, action) in enumerate(rows):
        y1 = row_top + index * 88
        y2 = y1 + 76
        fill = PALE if index % 2 == 0 else LIGHT
        draw.rounded_rectangle((370, y1, 1430, y2), radius=18, fill=fill, outline=BORDER, width=3)
        draw.text((405, y1 + 19), command, font=F_SMALL, fill=NAVY)
        draw.text((650, y1 + 19), action, font=F_SMALL, fill=DARK)
    arrow(draw, (900, 830), (900, 910))
    draw.text((925, 840), "হ্যাঁ", font=F_TINY, fill=NAVY)

    diamond(draw, (900, 1790), 650, 250, "watch command কি না", fnt=F_BODY)
    arrow(draw, (900, 1590), (900, 1665))

    # Normal one-shot commands end after displaying their result.
    rounded(draw, (180, 1940, 760, 2100), "Result terminal এ দেখানো", fill=PALE, fnt=F_SMALL)
    poly_arrow(draw, [(575, 1790), (470, 1790), (470, 1940)])
    draw.text((490, 1810), "না", font=F_TINY, fill=NAVY)

    # Watch loop.
    rounded(draw, (1040, 1930, 1640, 2110), "Interval অপেক্ষা করে\nনতুন live data সংগ্রহ ও diff", fill=PALE, fnt=F_SMALL)
    poly_arrow(draw, [(1225, 1790), (1340, 1790), (1340, 1930)])
    draw.text((1250, 1810), "হ্যাঁ", font=F_TINY, fill=NAVY)
    diamond(draw, (1340, 2280), 540, 220, "Change পাওয়া গেছে কি", fnt=F_SMALL)
    arrow(draw, (1340, 2110), (1340, 2170))
    rounded(draw, (1040, 2450, 1640, 2590), "Change থাকলে output দেখানো", fill=LIGHT, fnt=F_SMALL)
    arrow(draw, (1340, 2390), (1340, 2450))
    draw.text((1365, 2400), "হ্যাঁ", font=F_TINY, fill=NAVY)
    poly_arrow(draw, [(1640, 2520), (1720, 2520), (1720, 2020), (1640, 2020)], fill=BLUE)
    poly_arrow(draw, [(1610, 2280), (1755, 2280), (1755, 2020), (1640, 2020)], fill=BLUE)
    draw.text((1590, 2215), "না", font=F_TINY, fill=BLUE)
    draw.text((1580, 2160), "পরবর্তী interval", font=F_TINY, fill=BLUE)

    # End state for non-watch commands and Ctrl+C from watch mode.
    draw.ellipse((410, 2225, 550, 2365), fill=DARK)
    arrow(draw, (470, 2100), (470, 2225))
    draw.text((570, 2270), "শেষ", font=F_SMALL, fill=DARK)
    draw.text((1060, 2640), "Ctrl+C দিলে watch loop শেষ হয়", font=F_TINY, fill=GRAY)
    image.save(path, quality=95)


def set_font(run, name="Nirmala UI", size=None, bold=False) -> None:
    run.font.name = name
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    if size is not None:
        run.font.size = Pt(size)
    rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{key}"), name)


def set_style(style, size, bold=False) -> None:
    style.font.name = "Nirmala UI"
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = RGBColor(0, 0, 0)
    rfonts = style._element.get_or_add_rPr().get_or_add_rFonts()
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{key}"), "Nirmala UI")


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("পৃষ্ঠা ")
    set_font(run, size=9)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    text = OxmlElement("w:instrText")
    text.set(qn("xml:space"), "preserve")
    text.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, text, end])


def configure(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.62)
    section.bottom_margin = Inches(0.62)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    set_style(doc.styles["Normal"], 10.7)
    doc.styles["Normal"].paragraph_format.space_after = Pt(6)
    doc.styles["Normal"].paragraph_format.line_spacing = 1.12
    set_style(doc.styles["Title"], 23, True)
    doc.styles["Title"].paragraph_format.space_after = Pt(12)
    ppr = doc.styles["Title"]._element.get_or_add_pPr()
    border = ppr.find(qn("w:pBdr"))
    if border is not None:
        ppr.remove(border)
    for name, size in (("Heading 1", 16), ("Heading 2", 12.5)):
        set_style(doc.styles[name], size, True)
        doc.styles[name].paragraph_format.keep_with_next = True
        doc.styles[name].paragraph_format.space_before = Pt(10)
        doc.styles[name].paragraph_format.space_after = Pt(5)
    add_page_number(section.footer.paragraphs[0])
    doc.core_properties.title = "উইন্ডোজ সার্ভিসেস মনিটর কার্যপ্রবাহ ও ব্যবহার ক্ষেত্র চিত্র"
    doc.core_properties.author = "Team Techtonic"


def add_bullets(doc: Document, items: list[str]) -> None:
    for item in items:
        p = doc.add_paragraph(item, style="List Bullet")
        p.paragraph_format.space_after = Pt(3)


def add_picture(doc: Document, path: Path, width: float, alt_text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    shape = run.add_picture(str(path), width=Inches(width))
    doc_pr = shape._inline.docPr
    doc_pr.set("descr", alt_text)


def shade(cell, color: str) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    node = tcpr.find(qn("w:shd"))
    if node is None:
        node = OxmlElement("w:shd")
        tcpr.append(node)
    node.set(qn("w:fill"), color)


def add_table(doc: Document, headers: list[str], rows: list[list[str]]) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for index, value in enumerate(headers):
        cell = table.rows[0].cells[index]
        cell.text = value
        shade(cell, "1F4E78")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for run in cell.paragraphs[0].runs:
            set_font(run, size=9.5, bold=True)
            run.font.color.rgb = RGBColor(255, 255, 255)
    for row_index, values in enumerate(rows):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cells[index].text = value
            cells[index].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if row_index % 2:
                shade(cells[index], "F5F7F9")
            for run in cells[index].paragraphs[0].runs:
                set_font(run, size=9.1)


def build() -> Path:
    OUT.mkdir(parents=True, exist_ok=True)
    ASSETS.mkdir(parents=True, exist_ok=True)
    use_case = ASSETS / "use_case_diagram.png"
    activity = ASSETS / "activity_diagram.png"
    save_use_case_diagram(use_case)
    save_activity_diagram(activity)

    doc = Document()
    configure(doc)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(90)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("অপারেটিং সিস্টেম প্রজেক্ট ৩")
    set_font(r, size=13, bold=True)
    r.font.color.rgb = RGBColor(31, 78, 120)
    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("উইন্ডোজ সার্ভিসেস মনিটর\nActivity Diagram এবং Use Case Diagram")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Project flow actor interaction এবং database সম্পর্কের বাংলা ব্যাখ্যা")
    p.paragraph_format.space_after = Pt(40)
    add_table(doc, ["বিষয়", "তথ্য"], [
        ["প্রস্তুতকারক", "Team Techtonic"],
        ["System", "Windows Services Monitor"],
        ["Primary actor", "ব্যবহারকারী"],
        ["External systems", "Windows SCM এবং SQLite Database"],
        ["প্রধান commands", "live snapshot snapshots diff groups watch"],
    ])
    doc.add_page_break()

    doc.add_heading("১ Use Case Diagram", level=1)
    doc.add_paragraph(
        "এই চিত্রে ব্যবহারকারী Windows Services Monitor-এর কোন কোন সুবিধা ব্যবহার করে এবং Windows SCM ও SQLite database কোন কাজে data দেয় তা দেখানো হয়েছে।"
    )
    add_picture(doc, use_case, 6.75, "Windows Services Monitor use case diagram with user, SCM and SQLite actors")
    doc.add_page_break()

    doc.add_heading("২ Use Case এর ব্যাখ্যা", level=1)
    add_table(doc, ["Use case", "Actor ও কাজ"], [
        ["Live service তালিকা দেখা", "ব্যবহারকারী command দেয়; SCM state startup type ও PID দেয়"],
        ["Snapshot সংরক্ষণ", "SCM থেকে live data এনে SQLite database-এ timestamp সহ রাখা হয়"],
        ["Snapshot তালিকা দেখা", "SQLite থেকে ID captured time ও label পড়া হয়"],
        ["দুই snapshot তুলনা", "SQLite থেকে দুই set data নিয়ে state start type PID ও service presence তুলনা হয়"],
        ["Shared PID group দেখা", "Live SCM data অথবা saved snapshot নিয়ে একই PID-এর services group করা হয়"],
        ["পরিবর্তন পর্যবেক্ষণ", "SCM থেকে interval পরপর data নিয়ে আগের list-এর সঙ্গে diff করা হয়"],
    ])
    doc.add_heading("Actor সম্পর্ক", level=2)
    add_bullets(doc, [
        "ব্যবহারকারী primary actor; সে command নির্বাচন করে এবং output পড়ে।",
        "Windows SCM external system; এটি live service status ও configuration সরবরাহ করে।",
        "SQLite Database external storage; এটি snapshot সংরক্ষণ ও পুনরুদ্ধার করে।",
        "Monitor query-only হওয়ায় কোনো use case service start stop বা reconfigure করে না।",
    ])
    doc.add_page_break()

    doc.add_heading("৩ Activity Diagram", level=1)
    doc.add_paragraph(
        "এই activity diagram program শুরু থেকে command parse branch execution result এবং সমাপ্তি পর্যন্ত সম্পূর্ণ control flow দেখায়। Watch command-এর loop-ও আলাদাভাবে দেখানো হয়েছে।"
    )
    # Keep the tall activity diagram, its heading, and caption on one Word page.
    add_picture(doc, activity, 5.45, "Windows Services Monitor activity diagram showing all CLI command branches and watch loop")
    doc.add_page_break()

    doc.add_heading("৪ Activity Diagram এর ধাপ", level=1)
    add_bullets(doc, [
        "Program শুরু হয়ে argparse দিয়ে command ও option গ্রহণ করে।",
        "Decision node command অনুযায়ী live snapshot snapshots diff groups অথবা watch branch নির্বাচন করে।",
        "live branch SCM থেকে বর্তমান service list এনে terminal table দেখায়।",
        "snapshot branch live data নিয়ে SQLite database-এ নতুন snapshot রাখে।",
        "snapshots branch database থেকে saved snapshot metadata দেখায়।",
        "diff branch দুই snapshot load করে added removed এবং changed services বের করে।",
        "groups branch live অথবা saved data PID অনুযায়ী group করে।",
        "watch branch interval পরপর নতুন live list নেয় এবং change থাকলে output দেখিয়ে আবার loop করে।",
        "Ctrl+C দিলে watch loop থামে এবং program শেষ হয়।",
    ])
    doc.add_heading("Diagram symbol এর মানে", level=2)
    add_table(doc, ["Symbol", "অর্থ"], [
        ["কালো বৃত্ত", "Activity শুরু বা শেষ"],
        ["Rounded rectangle", "একটি action বা processing step"],
        ["Diamond", "Condition বা branch decision"],
        ["Arrow", "Control flow কোন দিকে যাচ্ছে"],
        ["Ellipse", "Use case বা system capability"],
        ["Stick actor", "System-এর বাইরে থাকা ব্যক্তি বা external system"],
        ["Cylinder", "Persistent database storage"],
    ])
    doc.add_heading("ভাইভায় এক বাক্যে বলবে", level=2)
    doc.add_paragraph(
        "Use Case Diagram দেখায় কে system-এর কোন সুবিধা ব্যবহার করে; Activity Diagram দেখায় command পাওয়ার পর system কোন ধাপে কাজ সম্পন্ন করে।"
    )

    doc.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    print(build())
