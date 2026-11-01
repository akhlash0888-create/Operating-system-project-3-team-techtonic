from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "reports"
OUTPUT = OUT / "Services_Monitor_Complete_Bengali_Guide_Submission.docx"

NAVY = "1F4E78"
PALE = "EAF2F8"
ALT = "F5F7F9"
BORDER = "D9D9D9"
BLACK = RGBColor(0, 0, 0)
BLUE = RGBColor(31, 78, 120)


def set_run_font(run, name: str = "Nirmala UI", size: float | None = None) -> None:
    run.font.name = name
    if size is not None:
        run.font.size = Pt(size)
    fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{key}"), name)


def set_style_font(style, name: str, size: float, bold: bool = False) -> None:
    style.font.name = name
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = BLACK
    fonts = style._element.get_or_add_rPr().get_or_add_rFonts()
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{key}"), name)


def remove_borders(style_or_paragraph) -> None:
    ppr = style_or_paragraph._element.get_or_add_pPr()
    border = ppr.find(qn("w:pBdr"))
    if border is not None:
        ppr.remove(border)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("পৃষ্ঠা ")
    set_run_font(run, size=9)
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
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.78)
    section.right_margin = Inches(0.78)

    normal = doc.styles["Normal"]
    set_style_font(normal, "Nirmala UI", 10.8)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.16

    title = doc.styles["Title"]
    set_style_font(title, "Nirmala UI", 24, True)
    title.paragraph_format.space_after = Pt(12)
    remove_borders(title)

    for name, size, before, after in (
        ("Heading 1", 16, 14, 7),
        ("Heading 2", 13, 11, 5),
        ("Heading 3", 11.5, 9, 4),
    ):
        style = doc.styles[name]
        set_style_font(style, "Nirmala UI", size, True)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    doc.core_properties.title = "উইন্ডোজ সার্ভিসেস মনিটর সম্পূর্ণ বাংলা নির্দেশিকা"
    doc.core_properties.subject = "Operating Systems Project 3"
    doc.core_properties.author = "Team Techtonic"
    add_page_number(section.footer.paragraphs[0])


def set_cell_border(cell) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    borders = tcpr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcpr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "6")
        node.set(qn("w:color"), BORDER)


def shade(cell, color: str) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    node = tcpr.find(qn("w:shd"))
    if node is None:
        node = OxmlElement("w:shd")
        tcpr.append(node)
    node.set(qn("w:fill"), color)


def cell_margins(cell, value: int = 110) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    mar = tcpr.first_child_found_in("w:tcMar")
    if mar is None:
        mar = OxmlElement("w:tcMar")
        tcpr.append(mar)
    for edge in ("top", "start", "bottom", "end"):
        node = mar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def add_table(doc: Document, headers: list[str], rows: list[list[str]], widths=None) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    if widths:
        for index, width in enumerate(widths):
            table.columns[index].width = Inches(width)
    header = table.rows[0]
    trpr = header._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    trpr.append(repeat)
    for index, value in enumerate(headers):
        cell = header.cells[index]
        cell.text = value
        shade(cell, NAVY)
        set_cell_border(cell)
        cell_margins(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            set_run_font(run, size=9.3)
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
    for row_index, values in enumerate(rows):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cells[index].text = str(value)
            cells[index].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_border(cells[index])
            cell_margins(cells[index])
            if row_index % 2:
                shade(cells[index], ALT)
            for paragraph in cells[index].paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    set_run_font(run, size=9.2)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def add_bullets(doc: Document, items: list[str], numbered: bool = False) -> None:
    style = "List Number" if numbered else "List Bullet"
    for item in items:
        paragraph = doc.add_paragraph(item, style=style)
        paragraph.paragraph_format.space_after = Pt(3)
        for run in paragraph.runs:
            set_run_font(run)


def add_code(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.left_indent = Inches(0.22)
    paragraph.paragraph_format.right_indent = Inches(0.15)
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(8)
    paragraph.paragraph_format.line_spacing = 1.0
    ppr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "F2F2F2")
    ppr.append(shd)
    run = paragraph.add_run(text)
    set_run_font(run, "Consolas", 8.6)


def add_term(doc: Document, term: str, explanation: str) -> None:
    paragraph = doc.add_paragraph()
    first = paragraph.add_run(f"{term}: ")
    set_run_font(first)
    first.bold = True
    second = paragraph.add_run(explanation)
    set_run_font(second)


def add_cover(doc: Document) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(70)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("অপারেটিং সিস্টেম প্রজেক্ট ৩")
    set_run_font(run, size=13)
    run.bold = True
    run.font.color.rgb = BLUE

    paragraph = doc.add_paragraph(style="Title")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run("উইন্ডোজ সার্ভিসেস মনিটর\nসম্পূর্ণ বাংলা নির্দেশিকা")
    remove_borders(paragraph)

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(25)
    run = paragraph.add_run("একেবারে শুরু থেকে কোড বোঝা চালানো পরীক্ষা করা এবং উপস্থাপনের সহায়িকা")
    set_run_font(run, size=12)

    table = doc.add_table(rows=5, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(1.8)
    table.columns[1].width = Inches(4.2)
    rows = [
        ("প্রস্তুতকারক", "Team Techtonic"),
        ("কোর্স", "Operating Systems"),
        ("প্রজেক্ট", "Windows Services Monitor"),
        ("প্রযুক্তি", "Python ctypes Win32 API SQLite"),
        ("ব্যবহার", "নিজে বোঝা অনুশীলন এবং ভাইভা প্রস্তুতি"),
    ]
    for row, values in zip(table.rows, rows):
        for index, value in enumerate(values):
            row.cells[index].text = value
            row.cells[index].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_border(row.cells[index])
            cell_margins(row.cells[index], 130)
            if index == 0:
                shade(row.cells[index], PALE)
            for run in row.cells[index].paragraphs[0].runs:
                set_run_font(run, size=10)
                if index == 0:
                    run.bold = True
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(28)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Windows 10 বা 11 | Python 3.10 বা নতুন | কেবল Standard Library")
    set_run_font(run, size=9.5)
    run.italic = True
    doc.add_page_break()


def add_contents(doc: Document) -> None:
    doc.add_heading("সূচিপত্র", level=1)
    entries = [
        "এই নির্দেশিকা কীভাবে পড়বে",
        "প্রজেক্টের পরিচয় ও উদ্দেশ্য",
        "Windows Service SCM এবং PID",
        "প্রজেক্টের সামগ্রিক নকশা",
        "models.py সম্পূর্ণ ব্যাখ্যা",
        "scm.py সম্পূর্ণ ব্যাখ্যা",
        "db.py সম্পূর্ণ ব্যাখ্যা",
        "analysis.py সম্পূর্ণ ব্যাখ্যা",
        "main.py সম্পূর্ণ ব্যাখ্যা",
        "প্রতিটি command চালানোর নিয়ম",
        "Snapshot এবং Diff হাতে কলমে",
        "Database কীভাবে তথ্য রাখে",
        "পরীক্ষা এবং ফলাফল",
        "ত্রুটি সমাধান ও নিরাপত্তা",
        "ভাইভা প্রশ্ন ও উত্তর",
        "শিক্ষকের সামনে প্রদর্শনের ধাপ",
        "এক নজরে পুনরাবৃত্তি",
    ]
    for index, entry in enumerate(entries, 1):
        paragraph = doc.add_paragraph(f"{index}.  {entry}")
        paragraph.paragraph_format.space_after = Pt(4)
    doc.add_page_break()


def build() -> Path:
    OUT.mkdir(parents=True, exist_ok=True)
    doc = Document()
    configure(doc)
    add_cover(doc)
    add_contents(doc)

    doc.add_heading("১ এই নির্দেশিকা কীভাবে পড়বে", level=1)
    doc.add_paragraph(
        "এই নথিটি এমনভাবে লেখা হয়েছে যেন পাঠক আগে Windows Service, Win32 API, ctypes বা SQLite সম্পর্কে কিছুই জানে না। "
        "প্রথমে ধারণাগুলো পড়বে, তারপর file ধরে ব্যাখ্যা পড়বে, শেষে command নিজে চালাবে। Python-এর keyword, function-এর নাম, "
        "file-এর নাম এবং terminal command ইংরেজিতেই রাখা হয়েছে, কারণ এগুলো অনুবাদ করলে code আর চলবে না। বাকি ব্যাখ্যা সম্পূর্ণ বাংলায়।"
    )
    add_bullets(doc, [
        "প্রথম পাঠে ২ থেকে ৪ নম্বর অধ্যায় পড়ে project-এর বড় ছবি বোঝো।",
        "দ্বিতীয় পাঠে ৫ থেকে ৯ নম্বর অধ্যায় খুলে পাশে আসল code মিলিয়ে পড়ো।",
        "তৃতীয় পাঠে ১০ ও ১১ নম্বর অধ্যায়ের command নিজে চালাও।",
        "ভাইভার আগে ১৫ থেকে ১৭ নম্বর অধ্যায় অন্তত দুইবার পড়ো।",
    ])

    doc.add_heading("২ প্রজেক্টের পরিচয় ও উদ্দেশ্য", level=1)
    doc.add_heading("Projectটি আসলে কী", level=2)
    doc.add_paragraph(
        "Windows Services Monitor একটি command-line program। এটি Windows-এর Service Control Manager থেকে সব registered service-এর "
        "বর্তমান অবস্থা, startup type এবং process ID সংগ্রহ করে। তারপর চাইলে সেই তথ্য SQLite database-এ snapshot হিসেবে সংরক্ষণ করে, "
        "দুই সময়ের snapshot তুলনা করে পরিবর্তন দেখায় এবং একই process-এর ভেতরে চলা একাধিক service খুঁজে বের করে।"
    )
    doc.add_heading("কেন এই project করা হয়েছে", level=2)
    add_bullets(doc, [
        "Operating System কীভাবে background service পরিচালনা করে তা শেখা।",
        "Python থেকে সরাসরি native Windows API ব্যবহার করা।",
        "Runtime state এবং configured startup type-এর পার্থক্য বোঝা।",
        "একই PID-এর অধীনে একাধিক service চলার ধারণা বোঝা।",
        "সময়ভিত্তিক snapshot রেখে পরিবর্তন বিশ্লেষণ করা।",
    ])
    doc.add_heading("Program কী করে এবং কী করে না", level=2)
    add_table(doc, ["করে", "করে না"], [
        ["সব Windows service-এর তথ্য পড়ে", "কোনো service start বা stop করে না"],
        ["Snapshot database-এ রাখে", "Startup type পরিবর্তন করে না"],
        ["দুই snapshot-এর পরিবর্তন দেখায়", "Registry বা system file বদলায় না"],
        ["একই PID-এর service group দেখায়", "Administrator privilege জোর করে নেয় না"],
        ["Watch mode-এ পরিবর্তন পর্যবেক্ষণ করে", "ব্যবহারকারীর অজান্তে কোনো system change করে না"],
    ], [3.25, 3.25])

    doc.add_heading("৩ Windows Service SCM এবং PID", level=1)
    add_term(doc, "Windows Service", "এমন background program যা user login না থাকলেও চলতে পারে এবং operating system বা অন্য software-কে প্রয়োজনীয় সুবিধা দেয়।")
    add_term(doc, "Service Control Manager বা SCM", "Windows-এর কেন্দ্রীয় service database ও controller। services.exe process এই ব্যবস্থা পরিচালনা করে।")
    add_term(doc, "Runtime state", "এই মুহূর্তে service চলমান, বন্ধ, শুরু হচ্ছে, বন্ধ হচ্ছে নাকি paused—তার অবস্থা।")
    add_term(doc, "Startup type", "Windows boot বা প্রয়োজনের সময় service কীভাবে শুরু হবে তার স্থায়ী configuration।")
    add_term(doc, "PID", "Process Identifier। চলমান প্রতিটি process-কে Windows যে আলাদা সংখ্যা দেয়। বন্ধ service-এর PID সাধারণত 0।")
    doc.add_heading("Runtime state-এর মানে", level=2)
    add_table(doc, ["Code-এ নাম", "বাংলা অর্থ"], [
        ["STOPPED", "Service চলছে না"],
        ["START_PENDING", "শুরু হওয়ার প্রক্রিয়ায় আছে"],
        ["RUNNING", "সক্রিয়ভাবে চলছে"],
        ["STOP_PENDING", "বন্ধ হওয়ার প্রক্রিয়ায় আছে"],
        ["CONTINUE_PENDING", "Paused অবস্থা থেকে আবার চলার প্রস্তুতি নিচ্ছে"],
        ["PAUSE_PENDING", "Paused হওয়ার প্রস্তুতি নিচ্ছে"],
        ["PAUSED", "সাময়িকভাবে স্থগিত আছে"],
    ], [2.2, 4.3])
    doc.add_heading("Startup type-এর মানে", level=2)
    add_table(doc, ["Code-এ নাম", "বাংলা অর্থ"], [
        ["BOOT", "Boot loader পর্যায়ে driver শুরু হয়"],
        ["SYSTEM", "Kernel initialization-এর সময় driver শুরু হয়"],
        ["AUTOMATIC", "Windows চালুর সময় স্বয়ংক্রিয়ভাবে শুরু হয়"],
        ["AUTOMATIC_DELAYED", "Windows চালুর একটু পরে স্বয়ংক্রিয়ভাবে শুরু হয়"],
        ["MANUAL", "প্রয়োজন হলে Windows বা user শুরু করে"],
        ["DISABLED", "Service শুরু করা নিষিদ্ধ"],
        ["ACCESS_DENIED", "Status পাওয়া গেছে কিন্তু configuration পড়ার অনুমতি পাওয়া যায়নি"],
    ], [2.2, 4.3])
    doc.add_paragraph(
        "গুরুত্বপূর্ণ: Runtime state এবং startup type এক জিনিস নয়। একটি AUTOMATIC service বর্তমানে STOPPED থাকতে পারে। আবার MANUAL service কোনো application-এর প্রয়োজনের কারণে RUNNING থাকতে পারে।"
    )
    doc.add_heading("এক PID-এ একাধিক service কেন", level=2)
    doc.add_paragraph(
        "Windows অনেক service-কে আলাদা executable না দিয়ে svchost.exe নামের shared host process-এ চালায়। এতে memory ও process overhead কমে। "
        "তাই ভিন্ন service name একই PID দেখাতে পারে। Project-এর groups command এই সম্পর্কটি দেখায়।"
    )

    doc.add_heading("৪ প্রজেক্টের সামগ্রিক নকশা", level=1)
    add_table(doc, ["File", "মূল দায়িত্ব"], [
        ["models.py", "Program-এর data object নির্ধারণ করে"],
        ["scm.py", "Win32 SCM API ডেকে live service data আনে"],
        ["db.py", "SQLite database-এ snapshot সংরক্ষণ ও পড়া পরিচালনা করে"],
        ["analysis.py", "Snapshot difference এবং PID grouping হিসাব করে"],
        ["main.py", "Command-line interface ও সব module-এর সমন্বয় করে"],
        ["tests/test_project.py", "মূল logic ঠিক আছে কি না automated test করে"],
    ], [1.65, 4.85])
    doc.add_heading("Data flow", level=2)
    add_code(doc, "User command\n    -> main.py\n        -> scm.py -> Windows SCM API -> ServiceInfo list\n        -> db.py -> SQLite snapshot\n        -> analysis.py -> diff or PID groups\n    -> formatted terminal output")
    doc.add_paragraph(
        "এই separation-এর লাভ হলো operating-system access, database logic, analysis এবং user interface আলাদা থাকে। ফলে code বোঝা, পরীক্ষা করা এবং ভুল খুঁজে বের করা সহজ হয়।"
    )

    doc.add_heading("৫ models.py সম্পূর্ণ ব্যাখ্যা", level=1)
    doc.add_paragraph(
        "এই file-এ dataclass ব্যবহার করে তিন ধরনের data object বানানো হয়েছে। frozen=True object পরিবর্তন ঠেকায় এবং slots=True অপ্রয়োজনীয় attribute creation বন্ধ করে memory কম ব্যবহার করে।"
    )
    add_table(doc, ["Class", "Field", "কাজ"], [
        ["ServiceInfo", "name", "SCM-এ থাকা service-এর internal নাম"],
        ["ServiceInfo", "display_name", "মানুষের পড়ার উপযোগী নাম"],
        ["ServiceInfo", "state", "বর্তমান runtime state"],
        ["ServiceInfo", "start_type", "Configured startup type"],
        ["ServiceInfo", "pid", "Hosting process ID; বন্ধ থাকলে সাধারণত 0"],
        ["Snapshot", "id", "Database-এর unique snapshot number"],
        ["Snapshot", "captured_at", "Snapshot নেওয়ার UTC সময়"],
        ["Snapshot", "label", "User দেওয়া optional নাম"],
        ["ServiceChange", "old এবং new fields", "দুই snapshot-এর আগের ও পরের মান"],
    ], [1.35, 1.55, 3.6])
    doc.add_heading("ServiceChange.kind property", level=2)
    add_bullets(doc, [
        "old_state None হলে serviceটি আগের snapshot-এ ছিল না; ফল ADDED।",
        "new_state None হলে serviceটি পরের snapshot-এ নেই; ফল REMOVED।",
        "দুই পাশেই service থাকলে কিন্তু গুরুত্বপূর্ণ field বদলালে ফল CHANGED।",
    ])

    doc.add_heading("৬ scm.py সম্পূর্ণ ব্যাখ্যা", level=1)
    doc.add_paragraph(
        "এটি project-এর সবচেয়ে গুরুত্বপূর্ণ low-level file। ctypes ব্যবহার করে Python থেকে advapi32.dll-এর native function ডাকা হয়। os.name nt না হলে program থামে, কারণ এই API শুধু Windows-এ আছে।"
    )
    doc.add_heading("ctypes কেন ব্যবহার করা হয়েছে", level=2)
    doc.add_paragraph(
        "ctypes Python standard library-এর অংশ। এটি C data type, pointer, structure এবং DLL function-এর সঙ্গে কাজ করতে দেয়। Assignment-এর raw Win32 API requirement পূরণ করতে কোনো third-party service library ব্যবহার করা হয়নি।"
    )
    doc.add_heading("Constants", level=2)
    add_table(doc, ["Constant", "ব্যবহার"], [
        ["SC_MANAGER_CONNECT", "SCM-এর সঙ্গে connection নেওয়ার permission"],
        ["SC_MANAGER_ENUMERATE_SERVICE", "Service list পড়ার permission"],
        ["SERVICE_QUERY_CONFIG", "একটি service-এর configuration পড়ার permission"],
        ["SERVICE_WIN32", "Win32 service filter"],
        ["SERVICE_STATE_ALL", "চলমান ও বন্ধ সব service চাওয়া"],
        ["ERROR_INSUFFICIENT_BUFFER", "দেওয়া buffer ছোট; প্রয়োজনীয় size নিতে হবে"],
        ["ERROR_MORE_DATA", "আরও data আছে বা buffer বড় দরকার"],
        ["ERROR_ACCESS_DENIED", "Configuration পড়ার অনুমতি নেই"],
    ], [2.65, 3.85])
    doc.add_heading("ctypes structure", level=2)
    add_table(doc, ["Structure", "কী রাখে"], [
        ["SERVICE_STATUS_PROCESS", "State, accepted control, exit code, PID এবং status metadata"],
        ["ENUM_SERVICE_STATUS_PROCESSW", "Service name, display name এবং SERVICE_STATUS_PROCESS"],
        ["QUERY_SERVICE_CONFIGW", "Startup type, executable path, dependencies, account এবং configuration"],
        ["SERVICE_DELAYED_AUTO_START_INFO", "Automatic startup delayed কি না"],
    ], [2.7, 3.8])
    doc.add_heading("argtypes এবং restype", level=2)
    doc.add_paragraph(
        "প্রতিটি DLL function-এর parameter type argtypes-এ এবং return type restype-এ ঘোষণা করা হয়েছে। এতে ctypes pointer ও integer সঠিকভাবে পাঠায় এবং ভুল type ব্যবহারের ঝুঁকি কমে। W suffix function-এর Unicode wide-character version বোঝায়।"
    )
    doc.add_heading("_raise_last_error", level=2)
    doc.add_paragraph("ctypes.get_last_error() দিয়ে সর্বশেষ Win32 error code নেয়, readable message বানায় এবং OSError তোলে। ফলে caller ব্যর্থতার কারণ দেখতে পায়।")
    doc.add_heading("_open_scm", level=2)
    add_bullets(doc, [
        "OpenSCManagerW-কে local machine ও active database-এর জন্য None পাঠায়।",
        "শুধু CONNECT এবং ENUMERATE_SERVICE permission চায়; পরিবর্তনের permission চায় না।",
        "Handle না পেলে _raise_last_error ডাকে। সফল হলে SCM handle ফেরত দেয়।",
    ], numbered=True)
    doc.add_heading("_enumerate_statuses", level=2)
    add_bullets(doc, [
        "needed, count ও resume নামে DWORD variable তৈরি করে।",
        "প্রথম EnumServicesStatusExW call-এ buffer None এবং size 0 দেয়। এটি sizing probe।",
        "Windows ERROR_MORE_DATA বা ERROR_INSUFFICIENT_BUFFER দিয়ে কত byte দরকার তা needed-এ লিখে দেয়।",
        "Sizing probe-এর পরে resume 0 করা হয়, যাতে আসল enumeration শুরু থেকে হয়।",
        "needed size-এর byte buffer তৈরি করে দ্বিতীয় call করা হয়। আরও data থাকলে loop নতুন size নিয়ে আবার চেষ্টা করে।",
        "Raw buffer-কে ENUM_SERVICE_STATUS_PROCESSW array হিসেবে cast করা হয়।",
        "প্রতিটি entry থেকে internal name, display name, mapped state এবং PID নিয়ে tuple list বানানো হয়।",
    ], numbered=True)
    doc.add_heading("Two call sizing pattern", level=3)
    add_code(doc, "প্রথম call: buffer নেই -> Windows প্রয়োজনীয় byte count জানায়\nদ্বিতীয় call: ঠিক size-এর buffer -> Windows আসল records লিখে দেয়")
    doc.add_heading("_query_start_type", level=2)
    add_bullets(doc, [
        "OpenServiceW দিয়ে নির্দিষ্ট service query-only mode-এ খোলে।",
        "Access denied হলে service বাদ না দিয়ে ACCESS_DENIED text ফেরত দেয়।",
        "QueryServiceConfigW-এর প্রথম call প্রয়োজনীয় buffer size নেয়।",
        "দ্বিতীয় call আসল QUERY_SERVICE_CONFIGW data নেয়।",
        "dwStartType-কে START_TYPE_NAMES dictionary দিয়ে readable text-এ রূপান্তর করে।",
        "AUTOMATIC হলে _is_delayed_auto_start দিয়ে delayed flag পরীক্ষা করে।",
        "finally block-এ service handle সব অবস্থায় বন্ধ করে।",
    ], numbered=True)
    doc.add_heading("_is_delayed_auto_start", level=2)
    doc.add_paragraph(
        "QueryServiceConfig2W-কে SERVICE_CONFIG_DELAYED_AUTO_START_INFO level দিয়ে ডাকে। Call সফল এবং fDelayedAutostart true হলে True ফেরত দেয়।"
    )
    doc.add_heading("enumerate_services", level=2)
    add_bullets(doc, [
        "SCM handle খোলে।",
        "সব service-এর runtime status নেয়।",
        "প্রতিটি status-এর জন্য startup type query করে ServiceInfo object বানায়।",
        "Service name case-insensitiveভাবে সাজিয়ে list ফেরত দেয়।",
        "finally block-এ SCM handle বন্ধ করে resource leak প্রতিরোধ করে।",
    ], numbered=True)

    doc.add_heading("৭ db.py সম্পূর্ণ ব্যাখ্যা", level=1)
    doc.add_paragraph("এই file SQLite ব্যবহার করে সময়ভিত্তিক snapshot স্থায়ীভাবে রাখে। Python-এর sqlite3 standard library ব্যবহৃত হয়েছে।")
    doc.add_heading("Database schema", level=2)
    add_table(doc, ["Table", "Column", "কাজ"], [
        ["snapshots", "id", "Auto-increment primary key"],
        ["snapshots", "captured_at", "UTC timestamp"],
        ["snapshots", "label", "Optional পরিচিতিমূলক নাম"],
        ["services", "snapshot_id", "কোন snapshot-এর record"],
        ["services", "name", "Service internal name"],
        ["services", "display_name", "Service display name"],
        ["services", "state", "Captured runtime state"],
        ["services", "start_type", "Captured startup configuration"],
        ["services", "pid", "Captured PID"],
    ], [1.3, 1.7, 3.5])
    doc.add_paragraph(
        "services table-এর composite primary key হলো snapshot_id এবং name। তাই একই snapshot-এ একই service দুইবার insert হতে পারে না। snapshot_id delete হলে ON DELETE CASCADE সংশ্লিষ্ট service rows-ও মুছে দেয়। PID index group query দ্রুত করতে সাহায্য করে।"
    )
    doc.add_heading("SnapshotStore-এর method", level=2)
    add_table(doc, ["Method", "ধাপে ধাপে কাজ"], [
        ["__init__", "Path নেয়, database connection খোলে, foreign key চালু করে এবং schema তৈরি করে"],
        ["__enter__", "with statement-এ store object ফেরত দেয়"],
        ["__exit__", "with block শেষ হলে connection বন্ধ করে"],
        ["close", "SQLite connection বন্ধ করে"],
        ["save_snapshot", "Timestamp ও label insert করে ID নেয়; তারপর সব ServiceInfo row bulk insert করে"],
        ["list_snapshots", "সব snapshot ID অনুযায়ী সাজিয়ে Snapshot object list দেয়"],
        ["get_services", "নির্দিষ্ট ID-এর service rows পড়ে ServiceInfo list বানায়; ID না থাকলে ValueError দেয়"],
        ["latest_ids", "সবচেয়ে নতুন নির্দিষ্ট সংখ্যক ID নিয়ে chronological order-এ ফেরত দেয়"],
    ], [1.65, 4.85])

    doc.add_heading("৮ analysis.py সম্পূর্ণ ব্যাখ্যা", level=1)
    doc.add_heading("diff_services", level=2)
    add_bullets(doc, [
        "Before ও after list-কে service name casefold করে dictionary বানায়।",
        "দুই dictionary-এর সব key union করে alphabetically পরীক্ষা করে।",
        "সম্পূর্ণ object একই হলে skip করে। Display name ছাড়া state, start type ও PID একই থাকলেও skip করে।",
        "এক পাশে না থাকলে added বা removed change তৈরি হয়। দুই পাশে থাকলে changed হয়।",
        "Old এবং new field রেখে ServiceChange list ফেরত দেয়।",
    ], numbered=True)
    doc.add_heading("group_by_pid", level=2)
    add_bullets(doc, [
        "শুধু pid > 0 service নেয়; stopped service-এর PID 0 বাদ পড়ে।",
        "defaultdict list-এ একই PID-এর service জমা করে।",
        "minimum_services-এর কম service থাকা group বাদ দেয়। Default minimum 2।",
        "প্রতিটি group-এর service name sort করে।",
        "বড় group আগে এবং সমান হলে ছোট PID আগে রেখে final list sort করে।",
    ], numbered=True)

    doc.add_heading("৯ main.py সম্পূর্ণ ব্যাখ্যা", level=1)
    doc.add_paragraph("এই file user command গ্রহণ করে এবং অন্য module-এর function ডেকে ফল terminal-এ দেখায়।")
    add_table(doc, ["Function", "কাজ"], [
        ["print_table", "Column width হিসাব করে aligned text table print করে"],
        ["print_services", "ServiceInfo list-এর name, state, start type, PID ও display name দেখায়"],
        ["print_changes", "Change না থাকলে message দেয়; থাকলে old থেকে new মান দেখায়"],
        ["command_live", "Live enumerate করে সম্পূর্ণ service table দেখায়"],
        ["command_snapshot", "Live data নিয়ে database-এ snapshot সংরক্ষণ করে"],
        ["command_snapshots", "সংরক্ষিত snapshot metadata দেখায়"],
        ["command_diff", "নির্বাচিত বা latest দুই snapshot তুলনা করে"],
        ["command_groups", "Live বা saved data থেকে shared PID groups দেখায়"],
        ["command_watch", "নির্দিষ্ট interval-এ বারবার data নিয়ে পরিবর্তন print করে"],
        ["build_parser", "argparse দিয়ে subcommand ও option নির্ধারণ করে"],
        ["main", "Argument parse, validation, handler call এবং error handling করে"],
    ], [1.8, 4.7])
    doc.add_heading("main function-এর error handling", level=2)
    doc.add_paragraph(
        "diff command-এ একটি ID দিলে parser error হয়; দুটি ID অথবা কোনো ID দিতে হবে। OSError বা ValueError হলে traceback না দেখিয়ে পরিষ্কার Error message stderr-এ print করে এবং exit code 1 দেয়। সফল হলে 0 দেয়।"
    )

    doc.add_heading("১০ প্রতিটি command চালানোর নিয়ম", level=1)
    doc.add_heading("প্রস্তুতি", level=2)
    add_bullets(doc, [
        "Windows 10 বা 11 এবং Python 3.10 বা নতুন version ব্যবহার করো।",
        "VS Code-এ project folder খোলো।",
        "Terminal menu থেকে New Terminal খোলো।",
        "Terminal path project root কি না দেখো।",
    ], numbered=True)
    add_code(doc, "cd E:\\Operating-system-project-3-team-techtonic\npython --version")
    add_table(doc, ["Command", "কখন ব্যবহার করবে"], [
        ["python main.py live", "এই মুহূর্তের সব service দেখতে"],
        ["python main.py snapshot --label before", "বর্তমান অবস্থার named snapshot নিতে"],
        ["python main.py snapshots", "Snapshot ID, সময় ও label দেখতে"],
        ["python main.py diff", "সর্বশেষ দুই snapshot তুলনা করতে"],
        ["python main.py diff 1 2", "নির্দিষ্ট দুই snapshot তুলনা করতে"],
        ["python main.py groups --limit 10", "সবচেয়ে বড় দশটি shared PID group দেখতে"],
        ["python main.py watch --interval 3", "প্রতি তিন সেকেন্ডে change দেখতে"],
        ["Ctrl+C", "Watch mode বন্ধ করতে"],
    ], [3.15, 3.35])
    doc.add_heading("আলাদা database path", level=2)
    add_code(doc, "python main.py snapshot --db tmp\\services.db --label before\npython main.py snapshots --db tmp\\services.db\npython main.py diff --db tmp\\services.db")
    doc.add_paragraph("যে command-এ database প্রয়োজন, প্রতিবার একই --db path ব্যবহার করবে। না হলে program অন্য database খুলে snapshot খুঁজে পাবে না।")

    doc.add_heading("১১ Snapshot এবং Diff হাতে কলমে", level=1)
    add_bullets(doc, [
        "প্রথম snapshot নাও: python main.py snapshot --label before",
        "Permission আছে এমন নিরাপদ একটি test service-এর state বাইরে থেকে পরিবর্তন করো। Project নিজে service পরিবর্তন করে না।",
        "State স্থির হওয়া পর্যন্ত অপেক্ষা করো।",
        "দ্বিতীয় snapshot নাও: python main.py snapshot --label after",
        "Snapshot ID দেখো: python main.py snapshots",
        "Difference দেখো: python main.py diff",
        "Demo শেষে service পরিবর্তন করে থাকলে আগের অবস্থায় ফেরত দাও।",
    ], numbered=True)
    doc.add_heading("Diff output কীভাবে পড়বে", level=2)
    add_code(doc, "SERVICE   KIND      STATE                START TYPE          PID\nAlpha     CHANGED   RUNNING -> STOPPED   AUTOMATIC -> MANUAL  100 -> 0")
    add_bullets(doc, [
        "KIND CHANGED অর্থ service দুই snapshot-এই আছে কিন্তু field বদলেছে।",
        "RUNNING থেকে STOPPED হলে PID সাধারণত valid number থেকে 0 হয়।",
        "ADDED অর্থ দ্বিতীয় snapshot-এ নতুন; REMOVED অর্থ দ্বিতীয় snapshot-এ নেই।",
        "কোনো পরিবর্তন না হলে program পরিষ্কার no changes message দেখায়।",
    ])

    doc.add_heading("১২ Database কীভাবে তথ্য রাখে", level=1)
    doc.add_paragraph(
        "একটি snapshot-এ প্রথমে snapshots table-এ একটি metadata row তৈরি হয়। তারপর সেই ID ব্যবহার করে প্রতিটি service-এর জন্য services table-এ row তৈরি হয়। "
        "দুই snapshot-এ একই service থাকলেও snapshot_id আলাদা হওয়ায় দুটো আলাদা historical record থাকে।"
    )
    add_code(doc, "snapshots\n  id=1  label=before\n  id=2  label=after\n\nservices\n  snapshot_id=1  name=Alpha  state=RUNNING\n  snapshot_id=2  name=Alpha  state=STOPPED")
    doc.add_paragraph(
        "এই নকশায় live system বদলে গেলেও পুরোনো snapshot অপরিবর্তিত থাকে। তাই নির্ভরযোগ্যভাবে আগে ও পরে তুলনা করা যায়।"
    )

    doc.add_heading("১৩ পরীক্ষা এবং ফলাফল", level=1)
    doc.add_heading("Automated test চালানো", level=2)
    add_code(doc, "python -m unittest discover -s tests -v")
    add_table(doc, ["Test", "যা যাচাই করে"], [
        ["diff test", "State, start type, PID, added এবং removed change সঠিক ধরা হয়"],
        ["group test", "PID 0 বাদ পড়ে এবং বড় group আগে আসে"],
        ["database round trip", "Save করার পর একই data পড়া যায় এবং latest IDs ঠিক থাকে"],
        ["missing snapshot", "না থাকা ID চাইলে ValueError হয়"],
    ], [2.15, 4.35])
    doc.add_heading("Live run-এর পর্যবেক্ষণ", level=2)
    add_table(doc, ["পরিমাপ", "ফল"], [
        ["মোট service", "314"],
        ["Running", "139"],
        ["Stopped", "175"],
        ["Automatic", "92"],
        ["Automatic delayed", "14"],
        ["Manual", "197"],
        ["Disabled", "10"],
        ["Configuration access denied", "1"],
    ], [3.5, 3.0])
    doc.add_paragraph(
        "এই সংখ্যা ১ নভেম্বর ২০২৬ তারিখে ব্যবহৃত computer-এর snapshot থেকে পাওয়া। অন্য computer বা অন্য সময়ে সংখ্যা ভিন্ন হওয়া স্বাভাবিক। দুটি এক সেকেন্ড ব্যবধানে নেওয়া snapshot-এ বাস্তব change পাওয়া যায়নি।"
    )
    doc.add_heading("Shared PID উদাহরণ", level=2)
    doc.add_paragraph(
        "Captured result-এ PID 1556 পাঁচটি service host করেছিল: BrokerInfrastructure, DcomLaunch, PlugPlay, Power এবং SystemEventsBroker। এটি প্রমাণ করে service এবং process এক জিনিস নয়।"
    )

    doc.add_heading("১৪ ত্রুটি সমাধান ও নিরাপত্তা", level=1)
    add_table(doc, ["সমস্যা", "সম্ভাব্য কারণ ও সমাধান"], [
        ["scm.py can only run on Windows", "Linux বা macOS-এ চালানো হয়েছে; Windows-এ চালাও"],
        ["Python was not found", "Python install অথবা PATH ঠিক করো"],
        ["At least two snapshots are required", "একই database-এ অন্তত দুইটি snapshot নাও"],
        ["Snapshot N does not exist", "python main.py snapshots দিয়ে valid ID দেখো"],
        ["ACCESS_DENIED", "একটি service-এর configuration permission নেই; program তবুও record রাখে"],
        ["No changes detected", "দুই snapshot-এর monitored fields একই; এটি error নয়"],
        ["Database অন্য folder-এ তৈরি", "প্রতিবার একই --db path দাও অথবা project root থেকে command চালাও"],
        ["Watch বন্ধ হচ্ছে না", "Terminal-এ Ctrl+C চাপো"],
    ], [2.4, 4.1])
    doc.add_heading("নিরাপত্তা নিয়ম", level=2)
    add_bullets(doc, [
        "অচেনা বা critical Windows service বন্ধ করবে না।",
        "Project-এর demonstration-এর জন্য নিজের permission থাকা service ব্যবহার করবে।",
        "পরিবর্তনের আগে original state ও startup type লিখে রাখবে।",
        "Demo শেষে পরিবর্তন আগের অবস্থায় ফিরিয়ে দেবে।",
        "এই monitor read-only; code-এ service control permission যোগ করার প্রয়োজন নেই।",
    ])

    doc.add_heading("১৫ ভাইভা প্রশ্ন ও উত্তর", level=1)
    qa = [
        ("এই project-এর প্রধান কাজ কী", "Windows SCM থেকে service state, startup type ও PID পড়া; snapshot সংরক্ষণ; পরিবর্তন ও shared PID বিশ্লেষণ করা।"),
        ("SCM কী", "Service Control Manager হলো Windows-এর installed service database ও lifecycle controller।"),
        ("ctypes কেন", "Third-party wrapper ছাড়া Python থেকে raw Win32 DLL function, C structure ও pointer ব্যবহার করার জন্য।"),
        ("দুইবার API call কেন", "প্রথম call প্রয়োজনীয় buffer size জানে; দ্বিতীয় call সেই size-এর buffer-এ আসল data নেয়।"),
        ("State ও startup type-এর পার্থক্য কী", "State বর্তমান অবস্থা; startup type ভবিষ্যতে কীভাবে শুরু হবে তার configuration।"),
        ("PID 0 কেন", "Stopped service-এর active hosting process নেই।"),
        ("এক PID-এ অনেক service কেন", "Windows shared svchost process-এর মধ্যে একাধিক service চালাতে পারে।"),
        ("Delayed automatic কী", "Boot-এর চাপ কমাতে automatic service Windows startup-এর একটু পরে শুরু হয়।"),
        ("Snapshot কী", "একটি নির্দিষ্ট সময়ে সব monitored service field-এর সংরক্ষিত copy।"),
        ("Diff কী কী দেখে", "Service যোগ বা বাদ, state, startup type এবং PID পরিবর্তন।"),
        ("casefold কেন", "Service name case-insensitiveভাবে match করার জন্য।"),
        ("finally block কেন", "Error হলেও native handle বন্ধ করে resource leak ঠেকাতে।"),
        ("SQLite কেন", "Standard library-তে আছে, আলাদা server লাগে না এবং ছোট local snapshot data-এর জন্য উপযুক্ত।"),
        ("Project কি service পরিবর্তন করে", "না। এটি query-only permission ব্যবহার করে এবং read-only monitor হিসেবে কাজ করে।"),
        ("Test-এ live SCM ব্যবহার না করার লাভ কী", "Synthetic ServiceInfo দিয়ে deterministicভাবে analysis ও database logic পরীক্ষা করা যায়।"),
        ("ACCESS_DENIED হলে কী হয়", "Service বাদ না দিয়ে start_type field-এ ACCESS_DENIED রাখা হয়, তাই dataset সম্পূর্ণ থাকে।"),
    ]
    for question, answer in qa:
        add_term(doc, f"প্রশ্ন  {question}", f"উত্তর  {answer}")

    doc.add_heading("১৬ শিক্ষকের সামনে প্রদর্শনের ধাপ", level=1)
    doc.add_heading("দুই মিনিটের পরিচিতি", level=2)
    doc.add_paragraph(
        "বলবে: আমাদের project Windows Services Monitor। এটি Python ctypes দিয়ে সরাসরি Win32 Service Control Manager API ব্যবহার করে সব service-এর live state, configured startup type এবং PID সংগ্রহ করে। আমরা SQLite-এ snapshot রাখতে, দুই snapshot তুলনা করতে এবং একই PID-এর shared services খুঁজতে পারি। Programটি read-only।"
    )
    doc.add_heading("Live demonstration", level=2)
    add_bullets(doc, [
        "python main.py live চালিয়ে columns ব্যাখ্যা করো।",
        "python main.py snapshot --label demo-before চালাও।",
        "python main.py snapshots চালিয়ে নতুন ID দেখাও।",
        "python main.py groups --limit 5 চালিয়ে shared PID ব্যাখ্যা করো।",
        "আগে থেকেই দুই snapshot থাকলে python main.py diff ID1 ID2 চালাও।",
        "python -m unittest discover -s tests -v চালিয়ে tests pass দেখাও।",
    ], numbered=True)
    doc.add_heading("Code দেখাতে বললে", level=2)
    add_bullets(doc, [
        "প্রথমে models.py দেখিয়ে data fields বোঝাও।",
        "scm.py-তে structures, API signatures ও two-call pattern দেখাও।",
        "db.py-তে দুই table ও save_snapshot দেখাও।",
        "analysis.py-তে diff_services ও group_by_pid দেখাও।",
        "main.py-তে argparse subcommands এবং main error handling দেখাও।",
    ], numbered=True)

    doc.add_heading("১৭ এক নজরে পুনরাবৃত্তি", level=1)
    add_table(doc, ["বিষয়", "মনে রাখার কথা"], [
        ["Input", "Windows SCM-এর live service records"],
        ["Core API", "OpenSCManagerW, EnumServicesStatusExW, OpenServiceW, QueryServiceConfigW, QueryServiceConfig2W"],
        ["Main output", "Service name, state, startup type, PID, display name"],
        ["Storage", "SQLite snapshots এবং services tables"],
        ["Analysis", "Snapshot diff এবং shared PID grouping"],
        ["Safety", "Read-only query; service start, stop বা reconfigure করে না"],
        ["Extension", "Delayed automatic detection এবং continuous watch mode"],
    ], [1.8, 4.7])
    doc.add_heading("শেষ checklist", level=2)
    add_bullets(doc, [
        "আমি project-এর উদ্দেশ্য এক মিনিটে বলতে পারি।",
        "আমি state, startup type ও PID-এর পার্থক্য বুঝি।",
        "আমি five Python file-এর দায়িত্ব বলতে পারি।",
        "আমি live, snapshot, snapshots, diff, groups এবং watch command চালাতে পারি।",
        "আমি two-call buffer pattern ও handle closing ব্যাখ্যা করতে পারি।",
        "আমি database schema ও diff logic ব্যাখ্যা করতে পারি।",
        "আমি tests চালিয়ে result দেখাতে পারি।",
        "আমি নিরাপদ demo করার নিয়ম জানি।",
    ])
    doc.add_paragraph(
        "এই checklist-এর প্রতিটি বাক্যে আত্মবিশ্বাসের সঙ্গে হ্যাঁ বলতে পারলে projectটি বোঝা, চালানো এবং ভাইভায় ব্যাখ্যা করার জন্য তুমি প্রস্তুত।"
    )

    doc.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    print(build())
