from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "reports" / "Services_Monitor_Bengali_Run_and_Output_Guide.docx"
BLUE = "1F4E78"
PALE = "EAF2F8"
LIGHT = "F3F4F6"
GRID = "D9D9D9"


def font(run, name="Nirmala UI", size=12, bold=False, color="000000"):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{key}"), name)


def shade(element, fill):
    shd = element.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        element.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_borders(cell, color=GRID, size="6"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        node = borders.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:color"), color)


def set_cell_margin(cell, top=100, start=120, bottom=100, end=120):
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def keep_with_next(paragraph):
    paragraph.paragraph_format.keep_with_next = True


def heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    keep_with_next(p)
    r = p.add_run(text)
    font(r, size=17 if level == 1 else 14, bold=True)
    return p


def para(doc, text="", bold_lead=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    if bold_lead and text.startswith(bold_lead):
        r = p.add_run(bold_lead)
        font(r, bold=True)
        text = text[len(bold_lead):]
    r = p.add_run(text)
    font(r)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.1
    font(p.add_run(text))
    return p


def step(doc, index, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.28)
    p.paragraph_format.first_line_indent = Inches(-0.28)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.1
    font(p.add_run(f"{index}.  "))
    font(p.add_run(text))
    return p


def checklist(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.28)
    p.paragraph_format.first_line_indent = Inches(-0.28)
    p.paragraph_format.space_after = Pt(3)
    font(p.add_run("□  " + text))
    return p


def code(doc, command):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.18)
    p.paragraph_format.right_indent = Inches(0.18)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_together = True
    shade(p._p.get_or_add_pPr(), LIGHT)
    r = p.add_run(command)
    font(r, name="Consolas", size=10.5, bold=True)
    return p


def sample(doc, lines):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.18)
    p.paragraph_format.right_indent = Inches(0.18)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_together = True
    shade(p._p.get_or_add_pPr(), PALE)
    r = p.add_run(lines)
    font(r, name="Consolas", size=9.5)
    return p


def table(doc, headers, rows, widths):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    t.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.width = Inches(widths[i])
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade(c._tc.get_or_add_tcPr(), BLUE)
        set_cell_borders(c)
        set_cell_margin(c)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        font(p.add_run(h), size=10.5, bold=True, color="FFFFFF")
    for ridx, row in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(row):
            c = cells[i]
            c.width = Inches(widths[i])
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if ridx % 2:
                shade(c._tc.get_or_add_tcPr(), "F7FAFC")
            set_cell_borders(c)
            set_cell_margin(c)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            font(p.add_run(str(val)), size=10.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_begin, instr, fld_end])
    font(run, size=9, color="666666")


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = Inches(0.65)
    sec.bottom_margin = Inches(0.65)
    sec.left_margin = Inches(0.75)
    sec.right_margin = Inches(0.75)
    page_number(sec.footer.paragraphs[0])

    styles = doc.styles
    styles["Normal"].font.name = "Nirmala UI"
    styles["Normal"].font.size = Pt(12)
    for name in ("Title", "Heading 1", "Heading 2"):
        styles[name].font.name = "Nirmala UI"
        styles[name].font.color.rgb = RGBColor(0, 0, 0)

    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(90)
    title.paragraph_format.space_after = Pt(14)
    title_ppr = title._p.get_or_add_pPr()
    title_border = title_ppr.find(qn("w:pBdr"))
    if title_border is not None:
        title_ppr.remove(title_border)
    title_style_ppr = styles["Title"]._element.get_or_add_pPr()
    style_border = title_style_ppr.find(qn("w:pBdr"))
    if style_border is not None:
        title_style_ppr.remove(style_border)
    font(title.add_run("উইন্ডোজ সার্ভিসেস মনিটর চালানো ও আউটপুট দেখার পূর্ণ নির্দেশিকা"), size=25, bold=True)
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(sub.add_run("কমান্ড কোথায় লিখবেন থেকে project বন্ধ করা পর্যন্ত"), size=15, bold=True, color=BLUE)
    intro = doc.add_paragraph()
    intro.alignment = WD_ALIGN_PARAGRAPH.CENTER
    intro.paragraph_format.left_indent = Inches(0.65)
    intro.paragraph_format.right_indent = Inches(0.65)
    intro.paragraph_format.space_before = Pt(40)
    intro.paragraph_format.line_spacing = 1.25
    font(intro.add_run(
        "এই নির্দেশিকাটি এমনভাবে লেখা হয়েছে যাতে project সম্পর্কে আগে কিছু না জানলেও তুমি command কপি করে ধাপে ধাপে চালাতে পারো। "
        "প্রতিটি command-এর কাজ, সম্ভাব্য output, সঠিক ক্রম এবং বন্ধ করার নিয়ম এখানে দেওয়া আছে।"
    ), size=13)
    doc.add_page_break()

    heading(doc, "১ প্রথমে যা বুঝতে হবে")
    para(doc, "এই project একটি command line program। এর output VS Code-এর নিচের Terminal-এ দেখা যায়। output ফোল্ডারে program result আসে না; সেখানে শুধু report-এর DOCX file আছে।")
    para(doc, "বেশির ভাগ command কাজ শেষ করে নিজে থেকেই বন্ধ হয়ে যায়। শুধু watch command চলতে থাকে; সেটি Ctrl+C দিয়ে থামাতে হয়।", bold_lead="বেশির ভাগ command")

    heading(doc, "২ VS Code এ project খোলা ও Terminal চালু করা")
    step(doc, 1, "VS Code খুলে File মেনু থেকে Open Folder নির্বাচন করো।")
    step(doc, 2, "এই folder নির্বাচন করো: E:\\Operating-system-project-3-team-techtonic")
    step(doc, 3, "VS Code-এর উপরের Terminal মেনু থেকে New Terminal চাপো। Keyboard shortcut হলো Ctrl+`।")
    step(doc, 4, "Terminal-এর prompt-এ project path না থাকলে নিচের command কপি করে Enter চাপো।")
    code(doc, 'cd "E:\\Operating-system-project-3-team-techtonic"')
    para(doc, "ঠিক folder-এ এলে Terminal-এর শুরুতে E:\\Operating-system-project-3-team-techtonic দেখা যাবে।")

    heading(doc, "৩ Python ঠিক আছে কি না পরীক্ষা")
    para(doc, "প্রথমে Python version দেখো।")
    code(doc, "python --version")
    sample(doc, "Python 3.10.0\nঅথবা এর চেয়ে নতুন version")
    para(doc, "যদি python is not recognized দেখায়, Python install বা PATH ঠিক করতে হবে।")
    para(doc, "তারপর project-এর available command দেখো।")
    code(doc, "python main.py --help")
    sample(doc, "{live,snapshot,snapshots,diff,groups,watch}\nQuery Windows services through the raw Service Control Manager API.")
    doc.add_page_break()

    heading(doc, "৪ Demo করার সঠিক command ক্রম")
    para(doc, "নিচের commandগুলো এই ক্রমেই চালালে project-এর প্রধান সব feature দেখানো যাবে। একটির output শেষ হলে পরেরটি চালাবে।")
    table(doc, ["ধাপ", "Command", "কী দেখাবে"], [
        ("১", "python main.py live", "বর্তমান Windows services"),
        ("২", "python main.py snapshot --label before", "প্রথম snapshot সংরক্ষণ"),
        ("৩", "python main.py snapshot --label after", "দ্বিতীয় snapshot সংরক্ষণ"),
        ("৪", "python main.py snapshots", "সব saved snapshot-এর তালিকা"),
        ("৫", "python main.py diff", "সর্বশেষ দুই snapshot-এর পার্থক্য"),
        ("৬", "python main.py groups --limit 10", "একই PID-এর service group"),
        ("৭", "python main.py watch --interval 3", "প্রতি ৩ সেকেন্ডে পরিবর্তন দেখা"),
        ("৮", "python -m unittest discover -s tests -v", "project test result"),
    ], [0.55, 3.15, 3.05])

    heading(doc, "৫ ধাপ ১ বর্তমান services দেখা")
    code(doc, "python main.py live")
    para(doc, "Terminal-এ SERVICE, STATE, START TYPE, PID এবং DISPLAY NAME column দেখা যাবে। একদম শেষে মোট service সংখ্যা দেখাবে।")
    sample(doc, "SERVICE   STATE    START TYPE   PID    DISPLAY NAME\nSpooler   RUNNING  AUTO         1234   Print Spooler\n\nTotal services: 250")
    bullet(doc, "STATE বলে service RUNNING না STOPPED।")
    bullet(doc, "START TYPE বলে service AUTO, MANUAL বা DISABLED কি না।")
    bullet(doc, "PID হলো চলমান service-এর process ID।")

    heading(doc, "৬ ধাপ ২ প্রথম snapshot সংরক্ষণ")
    code(doc, "python main.py snapshot --label before")
    para(doc, "এই command বর্তমান service তালিকা services.db নামে SQLite database file-এ সংরক্ষণ করবে।")
    sample(doc, "Saved snapshot 1 with 250 services at 2026-11-01T15:30:00.")

    heading(doc, "৭ ধাপ ৩ দ্বিতীয় snapshot সংরক্ষণ")
    para(doc, "কিছুক্ষণ অপেক্ষা করে দ্বিতীয় snapshot নাও। চাইলে নিরাপদ demo-র জন্য কোনো অনুমোদিত service-এর অবস্থা পরিবর্তন করার পরে এটি চালাতে পারো। কিছু পরিবর্তন না করলেও command ঠিকভাবে কাজ করবে।")
    code(doc, "python main.py snapshot --label after")
    sample(doc, "Saved snapshot 2 with 250 services at 2026-11-01T15:35:00.")

    heading(doc, "৮ ধাপ ৪ saved snapshot তালিকা দেখা")
    code(doc, "python main.py snapshots")
    sample(doc, "ID   CAPTURED AT                  LABEL\n1    2026-11-01T15:30:00          before\n2    2026-11-01T15:35:00          after")
    para(doc, "ID column-এর সংখ্যা পরে নির্দিষ্ট snapshot তুলনা করতে ব্যবহার করা যায়।")

    heading(doc, "৯ ধাপ ৫ snapshot তুলনা করা")
    para(doc, "সর্বশেষ দুইটি snapshot স্বয়ংক্রিয়ভাবে তুলনা করতে নিচের command চালাও।")
    code(doc, "python main.py diff")
    sample(doc, "Snapshot 1 -> snapshot 2\n\nNo service state, startup type, or PID changes were detected.")
    para(doc, "No changes লেখা মানে error নয়; দুই snapshot-এর মধ্যে service state, startup type বা PID পরিবর্তন হয়নি। পরিবর্তন থাকলে SERVICE, KIND, STATE, START TYPE এবং PID table দেখা যাবে।")
    para(doc, "নির্দিষ্ট ID তুলনা করতে, 1 ও 2-এর জায়গায় snapshots command-এ দেখা আসল ID বসাও।")
    code(doc, "python main.py diff 1 2")

    heading(doc, "১০ ধাপ ৬ একই PID এর service group দেখা")
    code(doc, "python main.py groups --limit 10")
    sample(doc, "Shared service hosts from live data\n\nPID 1234: 3 services\n  ServiceA, ServiceB, ServiceC")
    para(doc, "এটি একই Windows process বা PID-এর মধ্যে চলা একাধিক service group করে দেখায়। সর্বোচ্চ ১০টি group দেখানো হবে।")

    doc.add_page_break()
    heading(doc, "১১ ধাপ ৭ live change watch করা")
    code(doc, "python main.py watch --interval 3")
    sample(doc, "Watching 250 services. Press Ctrl+C to stop.")
    para(doc, "এই command প্রতি ৩ সেকেন্ড পর service list আবার পড়ে। পরিবর্তন না হলে নতুন output না আসা স্বাভাবিক। পরিবর্তন হলে সময়সহ changes table দেখাবে।")
    para(doc, "Watch বন্ধ করতে keyboard-এ Ctrl এবং C একসঙ্গে চাপো। তখন নিচের message আসবে।", bold_lead="Watch বন্ধ করতে")
    sample(doc, "Watch stopped.")

    heading(doc, "১২ ধাপ ৮ automated test চালানো")
    code(doc, "python -m unittest discover -s tests -v")
    para(doc, "সব test সফল হলে শেষের দিকে নিচের result দেখা যাবে।")
    sample(doc, "Ran 4 tests in 0.10s\n\nOK")
    para(doc, "OK দেখানো মানে snapshot database, diff logic এবং PID grouping-এর test pass করেছে।")

    heading(doc, "১৩ একসঙ্গে কপি করার সম্পূর্ণ command sequence")
    para(doc, "নিচের block একবারে কপি না করে এক লাইন করে চালানো সবচেয়ে ভালো। প্রতিটি command শেষ হওয়ার পরে পরের line চালাবে। watch command-এর পরে Ctrl+C চাপার আগে test command চালানো যাবে না।")
    code(doc, 'cd "E:\\Operating-system-project-3-team-techtonic"\npython --version\npython main.py --help\npython main.py live\npython main.py snapshot --label before\npython main.py snapshot --label after\npython main.py snapshots\npython main.py diff\npython main.py groups --limit 10\npython main.py watch --interval 3\npython -m unittest discover -s tests -v')

    heading(doc, "১৪ Project এবং VS Code বন্ধ করার নিয়ম")
    step(doc, 1, "watch চললে প্রথমে Ctrl+C চাপো। Watch stopped message না আসা পর্যন্ত অপেক্ষা করো।")
    step(doc, 2, "অন্য command হলে আলাদা করে বন্ধ করতে হয় না; output দেখানোর পর command নিজেই শেষ হয়।")
    close_terminal_step = step(doc, 3, "Terminal session বন্ধ করতে নিচের command লিখে Enter চাপতে পারো।")
    keep_with_next(close_terminal_step)
    code(doc, "exit")
    step(doc, 4, "শুধু Terminal panel লুকাতে Ctrl+` চাপতে পারো।")
    step(doc, 5, "VS Code পুরোপুরি বন্ধ করতে File মেনু থেকে Exit চাপো অথবা Alt+F4 চাপো।")
    para(doc, "এই project কোনো background server চালায় না। watch বন্ধ করার পরে project-এর আর কোনো process চলতে থাকে না।")

    heading(doc, "১৫ সাধারণ সমস্যা ও সমাধান")
    table(doc, ["সমস্যা", "কারণ ও সমাধান"], [
        ("python is not recognized", "Python install নেই বা PATH-এ নেই। Python 3.10 বা নতুন version install করে VS Code আবার খোলো।"),
        ("main.py পাওয়া যাচ্ছে না", "Terminal ভুল folder-এ আছে। প্রথমে cd command চালাও।"),
        ("At least two snapshots are required", "diff চালানোর আগে কমপক্ষে দুইবার snapshot command চালাও।"),
        ("diff-এ change নেই", "এটি স্বাভাবিক। দুই snapshot-এর মধ্যে service পরিবর্তন হয়নি।"),
        ("watch-এ output আসছে না", "পরিবর্তন না হলে watch শুধু অপেক্ষা করে। বন্ধ করতে Ctrl+C চাপো।"),
        ("Access denied বা OSError", "VS Code Run as administrator করে আবার চেষ্টা করা যেতে পারে; তবে service পরিবর্তন করো না।"),
    ], [2.05, 4.7])

    heading(doc, "১৬ দ্রুত checklist")
    for item in [
        "Terminal-এ সঠিক project path দেখা যাচ্ছে।",
        "python --version কাজ করছে।",
        "live command service table দেখিয়েছে।",
        "before এবং after নামে দুইটি snapshot তৈরি হয়েছে।",
        "snapshots command ID দেখিয়েছে।",
        "diff command comparison result দেখিয়েছে।",
        "groups command PID group দেখিয়েছে।",
        "watch command Ctrl+C দিয়ে বন্ধ করেছি।",
        "test-এর শেষে OK দেখেছি।",
        "কাজ শেষে Terminal বা VS Code বন্ধ করেছি।",
    ]:
        checklist(doc, item)

    doc.core_properties.title = "উইন্ডোজ সার্ভিসেস মনিটর চালানো ও আউটপুট দেখার পূর্ণ নির্দেশিকা"
    doc.core_properties.subject = "Windows Services Monitor run sequence and output guide in Bengali"
    doc.core_properties.author = "Team Techtonic"
    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
