from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "reports"
OUT.mkdir(parents=True, exist_ok=True)

NAVY = "1F4E78"
PALE = "EAF2F8"
LIGHT = "F5F7F9"
BORDER = "D9D9D9"
BLACK = RGBColor(0, 0, 0)


def shade(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_border(cell, color: str = BORDER, size: str = "6") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:color"), color)


def set_cell_margins(cell, top=100, start=120, bottom=100, end=120) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def keep_with_next(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    keep = OxmlElement("w:keepNext")
    p_pr.append(keep)


def remove_paragraph_borders(paragraph_or_style) -> None:
    element = paragraph_or_style._element
    p_pr = element.get_or_add_pPr()
    borders = p_pr.find(qn("w:pBdr"))
    if borders is not None:
        p_pr.remove(borders)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Page ")
    run.font.size = Pt(9)
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = " PAGE "
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_char1, instr_text, fld_char2])


def configure_document(doc: Document, title: str) -> None:
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.78)
    section.right_margin = Inches(0.78)

    normal = doc.styles["Normal"]
    normal.font.name = "Aptos"
    normal._element.rPr.rFonts.set(qn("w:ascii"), "Aptos")
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos")
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = BLACK
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.08

    title_style = doc.styles["Title"]
    title_style.font.name = "Aptos Display"
    title_style._element.rPr.rFonts.set(qn("w:ascii"), "Aptos Display")
    title_style._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos Display")
    title_style.font.size = Pt(25)
    title_style.font.bold = True
    title_style.font.color.rgb = BLACK
    title_style.paragraph_format.space_after = Pt(12)
    remove_paragraph_borders(title_style)

    for name, size, before, after in (
        ("Heading 1", 16, 13, 6),
        ("Heading 2", 12.5, 10, 4),
        ("Heading 3", 11, 8, 3),
    ):
        style = doc.styles[name]
        style.font.name = "Aptos Display"
        style._element.rPr.rFonts.set(qn("w:ascii"), "Aptos Display")
        style._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos Display")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = BLACK
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    doc.core_properties.title = title
    doc.core_properties.subject = "Operating Systems Project 3"
    doc.core_properties.author = "Team Techtonic"
    add_page_number(section.footer.paragraphs[0])


def add_cover(doc: Document, title: str, subtitle: str, language_note: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(72)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("OPERATING SYSTEMS PROJECT 3")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 78, 120)

    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(title)
    remove_paragraph_borders(p)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(28)
    run = p.add_run(subtitle)
    run.font.size = Pt(13)

    table = doc.add_table(rows=5, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(1.8)
    table.columns[1].width = Inches(4.2)
    fields = [
        ("Prepared by", "Team Techtonic"),
        ("Student ID", "____________________________"),
        ("Course", "Operating Systems"),
        ("Submission date", "1 November 2026"),
        ("Document", language_note),
    ]
    for row, (label, value) in zip(table.rows, fields):
        row.cells[0].text = label
        row.cells[1].text = value
        row.cells[0].paragraphs[0].runs[0].bold = True
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, 110, 140, 110, 140)
            set_cell_border(cell)
        shade(row.cells[0], PALE)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(34)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Windows 10 or 11   |   Python 3.10 or newer   |   Win32 SCM API")
    run.font.size = Pt(10)
    run.font.italic = True
    doc.add_page_break()


def add_table(doc: Document, headers: list[str], rows: list[list[str]], widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    if widths:
        for index, width in enumerate(widths):
            table.columns[index].width = Inches(width)
    hdr = table.rows[0]
    tr_pr = hdr._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    tr_pr.append(repeat)
    for index, text in enumerate(headers):
        cell = hdr.cells[index]
        cell.text = text
        shade(cell, NAVY)
        set_cell_border(cell)
        set_cell_margins(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9.5)
    for row_index, values in enumerate(rows):
        cells = table.add_row().cells
        for index, text in enumerate(values):
            cells[index].text = str(text)
            cells[index].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_border(cells[index])
            set_cell_margins(cells[index])
            if row_index % 2:
                shade(cells[index], LIGHT)
            for paragraph in cells[index].paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.size = Pt(9.25)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_bullets(doc: Document, items: list[str], numbered: bool = False) -> None:
    style = "List Number" if numbered else "List Bullet"
    for item in items:
        p = doc.add_paragraph(item, style=style)
        p.paragraph_format.space_after = Pt(3)


def add_code(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.24)
    p.paragraph_format.right_indent = Inches(0.15)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(7)
    p.paragraph_format.line_spacing = 1.0
    p_pr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "F2F2F2")
    p_pr.append(shd)
    run = p.add_run(text)
    run.font.name = "Consolas"
    run._element.rPr.rFonts.set(qn("w:ascii"), "Consolas")
    run._element.rPr.rFonts.set(qn("w:hAnsi"), "Consolas")
    run.font.size = Pt(8.7)


def add_contents(doc: Document, entries: list[str]) -> None:
    doc.add_heading("Contents", level=1)
    for index, entry in enumerate(entries, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(f"{index}.  {entry}")
        run.font.size = Pt(10.5)
    doc.add_page_break()


def build_english_report() -> Path:
    doc = Document()
    title = "Windows Services Monitor Lab Report"
    configure_document(doc, title)
    add_cover(
        doc,
        title,
        "Querying runtime state configured startup and shared hosting",
        "English laboratory report",
    )
    add_contents(
        doc,
        [
            "Abstract",
            "Objectives",
            "Background and operating system concepts",
            "System design",
            "Implementation",
            "Experimental procedure",
            "Results and analysis",
            "Testing",
            "Limitations and safety",
            "Conclusion",
            "References",
        ],
    )

    doc.add_heading("Abstract", level=1)
    doc.add_paragraph(
        "This project implements a command-line monitor for the Windows Service Control Manager. "
        "The program calls the native Win32 API through Python ctypes, enumerates every registered "
        "Win32 service, reads each live service state and configured startup type, and records the "
        "hosting process ID. It stores timestamped snapshots in SQLite, compares two snapshots, "
        "groups services that share a process, and can poll continuously for changes. A live run on "
        "1 November 2026 returned 314 services. Of those services, 139 were running and 175 were "
        "stopped. The largest shared host in the captured snapshot contained five services under one PID."
    )

    doc.add_heading("Objectives", level=1)
    add_bullets(
        doc,
        [
            "Open the Service Control Manager database with query-only access.",
            "Use the two-call sizing pattern with EnumServicesStatusExW to enumerate services safely.",
            "Read configured startup information with QueryServiceConfigW.",
            "Distinguish delayed automatic startup with QueryServiceConfig2W.",
            "Combine runtime and configured information in typed ServiceInfo objects.",
            "Store snapshots in SQLite and report service changes between two points in time.",
            "Identify processes that host more than one Windows service.",
        ],
    )

    doc.add_heading("Background and operating system concepts", level=1)
    doc.add_heading("Service Control Manager", level=2)
    doc.add_paragraph(
        "The Service Control Manager is the Windows subsystem that maintains the installed-service "
        "database and controls service lifecycles. The services.exe process implements the manager. "
        "Tools such as services.msc and sc.exe are clients of the same API used by this project."
    )
    doc.add_heading("Runtime state and configured state", level=2)
    doc.add_paragraph(
        "Runtime state describes what a service is doing at the instant of the query. Configured "
        "startup type describes how Windows should start the service. The values are independent. "
        "For example, an automatic service can be stopped during maintenance, while a manual service "
        "can be running because another component requested it."
    )
    add_table(
        doc,
        ["Runtime state", "Meaning"],
        [
            ["STOPPED", "The service is not running"],
            ["START PENDING", "The service is starting"],
            ["RUNNING", "The service is operating"],
            ["STOP PENDING", "The service is stopping"],
            ["PAUSE PENDING", "The service is entering the paused state"],
            ["PAUSED", "The service is temporarily suspended"],
            ["CONTINUE PENDING", "The service is returning from paused to running"],
        ],
        [2.0, 4.6],
    )
    doc.add_heading("Services and processes", level=2)
    doc.add_paragraph(
        "A service is a managed operating-system component, while a process is an executing program. "
        "Several services may run inside one svchost.exe process. Therefore, different service names "
        "can report the same PID even though their configuration and state remain separate."
    )

    doc.add_heading("System design", level=1)
    add_table(
        doc,
        ["Module", "Responsibility"],
        [
            ["scm.py", "Native SCM connection, enumeration, and configuration queries"],
            ["models.py", "Immutable typed models for services, snapshots, and changes"],
            ["db.py", "SQLite schema and snapshot persistence"],
            ["analysis.py", "Snapshot comparison and service grouping by PID"],
            ["main.py", "Command parsing, tables, live view, diff, groups, and watch mode"],
        ],
        [1.35, 5.25],
    )
    doc.add_heading("Data flow", level=2)
    doc.add_paragraph(
        "The CLI opens the SCM, requests the required enumeration buffer size, retrieves the service "
        "status records, and queries configuration for each service. The resulting ServiceInfo objects "
        "are printed directly or inserted into SQLite. Analysis functions receive ordinary ServiceInfo "
        "lists, which keeps operating-system access separate from comparison logic."
    )
    add_code(
        doc,
        "SCM database -> EnumServicesStatusExW -> live state and PID\n"
        "             -> QueryServiceConfigW  -> startup type\n"
        "             -> ServiceInfo objects  -> CLI / SQLite / analysis"
    )

    doc.add_heading("Implementation", level=1)
    doc.add_heading("SCM connection and enumeration", level=2)
    doc.add_paragraph(
        "OpenSCManagerW requests SC_MANAGER_CONNECT and SC_MANAGER_ENUMERATE_SERVICE only. The first "
        "EnumServicesStatusExW call passes a zero-length buffer. Windows returns the required byte "
        "count. The second call uses a correctly sized byte array and returns service names, display "
        "names, runtime states, and PIDs. Every native handle is closed in a finally block."
    )
    doc.add_heading("Configuration query", level=2)
    doc.add_paragraph(
        "For each service, OpenServiceW requests SERVICE_QUERY_CONFIG. QueryServiceConfigW uses the "
        "same two-call sizing method and reads the dwStartType field. QueryServiceConfig2W checks "
        "SERVICE_CONFIG_DELAYED_AUTO_START_INFO so delayed automatic startup is reported separately. "
        "If configuration access is denied, the monitor records ACCESS_DENIED and continues."
    )
    doc.add_heading("Persistence and comparison", level=2)
    doc.add_paragraph(
        "SQLite stores one row per snapshot and one row per service within that snapshot. The composite "
        "key of snapshot ID and service name prevents duplicates. Diff analysis matches names without "
        "case sensitivity and reports changes in state, startup type, or PID, as well as added or removed services."
    )
    doc.add_heading("Command interface", level=2)
    add_table(
        doc,
        ["Command", "Purpose"],
        [
            ["python main.py live", "Print the complete live service table"],
            ["python main.py snapshot --label before", "Save a timestamped snapshot"],
            ["python main.py snapshots", "List saved snapshot IDs"],
            ["python main.py diff", "Compare the latest two snapshots"],
            ["python main.py diff 1 2", "Compare two selected snapshot IDs"],
            ["python main.py groups --limit 10", "List the largest shared service hosts"],
            ["python main.py watch --interval 3", "Poll and print only detected changes"],
        ],
        [2.8, 3.8],
    )

    doc.add_heading("Experimental procedure", level=1)
    add_bullets(
        doc,
        [
            "Opened PowerShell in the project directory.",
            "Ran the automated test suite before using live operating-system data.",
            "Saved a snapshot labelled initial to tmp\\services.db.",
            "Saved a second snapshot labelled follow-up.",
            "Listed the snapshots and compared the latest pair.",
            "Ran process grouping against snapshot 2 and displayed the five largest shared hosts.",
        ],
        numbered=True,
    )
    add_code(
        doc,
        "python -m unittest discover -s tests -v\n"
        "python main.py snapshot --db tmp\\services.db --label initial\n"
        "python main.py snapshot --db tmp\\services.db --label follow-up\n"
        "python main.py diff --db tmp\\services.db\n"
        "python main.py groups --snapshot 2 --db tmp\\services.db --limit 5"
    )

    doc.add_heading("Results and analysis", level=1)
    doc.add_heading("Live snapshot summary", level=2)
    add_table(
        doc,
        ["Measurement", "Observed value"],
        [
            ["Total registered services", "314"],
            ["Running", "139"],
            ["Stopped", "175"],
            ["Automatic", "92"],
            ["Automatic delayed", "14"],
            ["Manual", "197"],
            ["Disabled", "10"],
            ["Configuration access denied", "1"],
        ],
        [3.5, 3.1],
    )
    doc.add_paragraph(
        "The total startup-type counts equal 314, so every enumerated service received a recorded "
        "configuration result. One service allowed status enumeration but denied its configuration "
        "query; the program preserved the service in the output and marked that field explicitly."
    )
    doc.add_heading("Shared process findings", level=2)
    add_table(
        doc,
        ["PID", "Count", "Service names"],
        [
            ["1556", "5", "BrokerInfrastructure, DcomLaunch, PlugPlay, Power, SystemEventsBroker"],
            ["1408", "3", "KeyIso, SamSs, VaultSvc"],
            ["1732", "2", "RpcEptMapper, RpcSs"],
            ["2648", "2", "BFE, mpssvc"],
            ["3396", "2", "ekrn, ekrnEpfw"],
        ],
        [0.7, 0.7, 5.2],
    )
    doc.add_paragraph(
        "PID 1556 hosted five independently named services. This directly demonstrates that a service "
        "is not equivalent to a process. The SCM tracks each service separately even when several "
        "services execute in the same host process."
    )
    doc.add_heading("Snapshot difference", level=2)
    doc.add_paragraph(
        "The two live snapshots were captured one second apart and contained 314 services each. No "
        "state, startup-type, or PID changes occurred during that interval. The comparison path was "
        "also tested with controlled in-memory data that changed a service from RUNNING and AUTOMATIC "
        "to STOPPED and MANUAL while changing its PID from 100 to 0. The test passed."
    )

    doc.add_heading("Testing", level=1)
    add_table(
        doc,
        ["Test", "Expected behavior", "Result"],
        [
            ["State and configuration diff", "Detect state, start type, and PID changes", "Pass"],
            ["Added and removed services", "Return ADDED and REMOVED records", "Pass"],
            ["PID grouping", "Exclude PID 0 and order largest groups first", "Pass"],
            ["SQLite round trip", "Save and retrieve identical typed data", "Pass"],
            ["Missing snapshot", "Raise a clear ValueError", "Pass"],
            ["Live SCM enumeration", "Read all accessible Windows services", "Pass, 314 services"],
        ],
        [2.0, 3.45, 1.15],
    )
    doc.add_paragraph(
        "The unittest run completed four test methods with no failures in 0.118 seconds. Compilation "
        "also completed without syntax errors."
    )

    doc.add_heading("Limitations and safety", level=1)
    add_bullets(
        doc,
        [
            "The tool observes services but never starts, stops, pauses, or reconfigures them.",
            "A real state-change demonstration must use a service the operator is authorized to control.",
            "Short pending states may disappear between polling intervals.",
            "PID values can change after a service restarts.",
            "Protected services may expose runtime status while denying configuration access.",
        ],
    )
    doc.add_paragraph(
        "No Windows service was modified during verification. This avoided disrupting networking, "
        "security software, audio, update, or other system functions. The watch command and documented "
        "before-and-after procedure support a supervised classroom demonstration."
    )

    doc.add_heading("Conclusion", level=1)
    doc.add_paragraph(
        "The completed monitor satisfies the required architecture and deliverables. It queries the raw "
        "SCM API, combines live and configured state, persists snapshots, compares two moments, and "
        "identifies shared service hosts. The live results show both core operating-system distinctions: "
        "configured startup can differ from runtime state, and several services can share one process."
    )

    doc.add_heading("References", level=1)
    add_bullets(
        doc,
        [
            "Project 3 Services Monitor Windows assignment brief, pages 1 to 5.",
            "Microsoft Learn, OpenSCManagerW function.",
            "Microsoft Learn, EnumServicesStatusExW function.",
            "Microsoft Learn, QueryServiceConfigW function and QUERY_SERVICE_CONFIGW structure.",
            "Microsoft Learn, QueryServiceConfig2W function and delayed auto-start information.",
            "Microsoft Learn, Service State Transitions.",
        ],
    )

    path = OUT / "Services_Monitor_English_Lab_Report_Submission.docx"
    doc.save(path)
    return path


def build_banglish_guide() -> Path:
    doc = Document()
    title = "Windows Services Monitor Banglish Guide"
    configure_document(doc, title)
    add_cover(
        doc,
        title,
        "Shuru theke project bujha run kora ebong viva preparation",
        "Beginner friendly Banglish documentation",
    )
    add_contents(
        doc,
        [
            "Project ta ashole ki",
            "Age je concept gulo bujhte hobe",
            "Project folder ebong prottek file",
            "Program er full flow",
            "SCM API kibhabe kaj kore",
            "Database snapshot diff ebong grouping",
            "Install ebong run korar niyom",
            "Nijer hate demo korar steps",
            "Output kibhabe bujhbe",
            "Common problem ebong solution",
            "Viva preparation",
            "Short revision sheet",
        ],
    )

    doc.add_heading("Project ta ashole ki", level=1)
    doc.add_paragraph(
        "Ei project e amra Windows er shob registered service er information ber kori. Service mane "
        "background e chola ekta managed component, jemon Windows Audio, Update, Firewall ba Bluetooth. "
        "Program ta service er current obostha, startup setting, ebong kon process ID ba PID er vitore "
        "service ta cholche seta dekhay. Tarpor ei information snapshot hisebe SQLite database e save kore."
    )
    doc.add_paragraph(
        "Simple vabe bolle, services.msc je information dekhay, amader Python program Windows er raw "
        "Service Control Manager API call kore oi information nijer moto collect ebong analyze kore."
    )

    doc.add_heading("Age je concept gulo bujhte hobe", level=1)
    doc.add_heading("Service ki", level=2)
    doc.add_paragraph(
        "Service holo emon software component ja usually user interface chara background e kaj kore. "
        "Windows boot hole kichu service automatic start hoy, kichu demand ashle start hoy, ar kichu disabled thake."
    )
    doc.add_heading("Process ki", level=2)
    doc.add_paragraph(
        "Process holo running program er ekta instance. Prottek process er ekta PID thake. Kintu service "
        "ebong process same jinis na. Ekta svchost.exe process er vitore onek service ekshathe thakte pare."
    )
    doc.add_heading("Runtime state ebong startup type", level=2)
    add_table(
        doc,
        ["Jinis", "Banglish explanation", "Example"],
        [
            ["Runtime state", "Ekhon ei muhurte service ki korche", "RUNNING ba STOPPED"],
            ["Startup type", "Windows kobe service start korbe tar setting", "AUTOMATIC ba MANUAL"],
        ],
        [1.4, 3.5, 1.7],
    )
    doc.add_paragraph(
        "Important kotha holo ei duita independent. AUTOMATIC service ekhon STOPPED hote pare. Abar MANUAL "
        "service kono app request korle RUNNING hote pare. Ei difference tai project er main learning point."
    )
    doc.add_heading("Service state machine", level=2)
    add_table(
        doc,
        ["State", "Mane"],
        [
            ["STOPPED", "Service cholche na"],
            ["START_PENDING", "Start howar process e ache"],
            ["RUNNING", "Puro vabe cholche"],
            ["STOP_PENDING", "Bondho howar process e ache"],
            ["PAUSED", "Temporary pause kora"],
            ["PAUSE_PENDING", "Pause howar process e ache"],
            ["CONTINUE_PENDING", "Pause theke abar run howar process e ache"],
        ],
        [2.2, 4.4],
    )

    doc.add_heading("Project folder ebong prottek file", level=1)
    add_table(
        doc,
        ["File", "Ki kaj kore"],
        [
            ["scm.py", "Windows er SCM API call kore live service data ney"],
            ["models.py", "ServiceInfo Snapshot ebong ServiceChange data structure rakhe"],
            ["db.py", "SQLite database e snapshot save ebong load kore"],
            ["analysis.py", "Duita snapshot compare kore ebong same PID group kore"],
            ["main.py", "User command ney ebong output print kore"],
            ["tests/test_project.py", "Program er important logic automatically test kore"],
            ["README.md", "Quick command reference"],
            ["output/reports", "English report ebong ei Banglish guide"],
        ],
        [2.0, 4.6],
    )
    doc.add_paragraph(
        "Ei separation ke modular design bole. Ek file e shob code na rekhe kaj onujayi vag kora hoyeche. "
        "Tai bug khuja, test kora, ebong porobortite feature add kora easy hoy."
    )

    doc.add_heading("Program er full flow", level=1)
    add_bullets(
        doc,
        [
            "User PowerShell e command dey.",
            "main.py command bujhe thik function call kore.",
            "scm.py Windows SCM database open kore.",
            "EnumServicesStatusExW shob service er name state ebong PID dey.",
            "QueryServiceConfigW prottek service er startup type dey.",
            "Data ServiceInfo object e rakha hoy.",
            "Command onujayi data screen e print hoy database e save hoy ba analyze hoy.",
        ],
        numbered=True,
    )
    add_code(
        doc,
        "PowerShell command\n"
        "        |\n"
        "        v\n"
        "main.py -> scm.py -> Windows SCM\n"
        "        |\n"
        "        +-> db.py -> SQLite snapshot\n"
        "        |\n"
        "        +-> analysis.py -> diff / PID group"
    )

    doc.add_heading("SCM API kibhabe kaj kore", level=1)
    doc.add_heading("OpenSCManagerW", level=2)
    doc.add_paragraph(
        "Prothome SCM database er ekta handle nite hoy. Handle ke ekta permission deya ticket er moto "
        "vabte paro. Amra shudhu connect ebong enumerate permission chai. Program kono service change korar permission chai na."
    )
    doc.add_heading("Two call sizing pattern", level=2)
    doc.add_paragraph(
        "Windows age theke bole na buffer koto boro lagbe. Tai prothom call e zero size buffer deya hoy. "
        "API required byte count bole dey. Tarpor oi size er buffer baniye second call kora hoy. Eita two-call sizing pattern."
    )
    add_code(
        doc,
        "First call:  buffer size = 0  -> Windows bole kotogulo byte lagbe\n"
        "Second call: thik size buffer -> Windows actual service records dey"
    )
    doc.add_heading("QueryServiceConfigW", level=2)
    doc.add_paragraph(
        "Enumeration theke runtime state ar PID pai. Kintu startup type alada configuration. Tai prottek "
        "service open kore QueryServiceConfigW call korte hoy. AUTOMATIC service delayed kina seta "
        "QueryServiceConfig2W diye check kora hoy."
    )
    doc.add_heading("Handle close kora keno dorkar", level=2)
    doc.add_paragraph(
        "Native handle use shesh hole CloseServiceHandle call korte hoy. Na korle resource leak hote pare. "
        "Code e finally block use kora hoyeche, tai error holeo handle close hoy."
    )

    doc.add_heading("Database snapshot diff ebong grouping", level=1)
    doc.add_heading("Snapshot mane ki", level=2)
    doc.add_paragraph(
        "Snapshot holo ekta specific time e shob service er chobi, kintu text data akare. snapshots table e "
        "time ebong label thake. services table e oi snapshot er prottek service er details thake."
    )
    add_table(
        doc,
        ["Table", "Important column", "Kaj"],
        [
            ["snapshots", "id captured_at label", "Kokhon snapshot neya hoyeche"],
            ["services", "snapshot_id name state start_type pid", "Oi time er service data"],
        ],
        [1.3, 2.6, 2.7],
    )
    doc.add_heading("Diff kibhabe hoy", level=2)
    doc.add_paragraph(
        "Before ebong after snapshot e service name diye match kora hoy. State, startup type, ba PID change "
        "hole CHANGED dekhay. Notun service hole ADDED, ar ager service pore na thakle REMOVED dekhay."
    )
    doc.add_heading("PID grouping kibhabe hoy", level=2)
    doc.add_paragraph(
        "Prottek running service ke tar PID diye group kora hoy. PID 0 mane kono running process nei, tai "
        "eta grouping theke bad deya hoy. Je PID te beshi service ache seta age dekhano hoy."
    )

    doc.add_heading("Install ebong run korar niyom", level=1)
    doc.add_heading("Step 1 PowerShell open", level=2)
    doc.add_paragraph(
        "Project folder open kore address bar e powershell likhe Enter dao. Ba PowerShell theke cd command diye folder e jao."
    )
    add_code(doc, 'cd "E:\\Operating-system-project-3-team-techtonic"')
    doc.add_heading("Step 2 Python check", level=2)
    add_code(doc, "python --version")
    doc.add_paragraph("Python 3.10 ba tar cheye notun version hole thik ache.")
    doc.add_heading("Step 3 Test run", level=2)
    add_code(doc, "python -m unittest discover -s tests -v")
    doc.add_paragraph("Shob test er pashe ok ebong sheshe OK dekhale logic thik ache.")
    doc.add_heading("Step 4 Live list", level=2)
    add_code(doc, "python main.py live")
    doc.add_paragraph(
        "Eta shob service print korbe. Column gulo holo SERVICE, STATE, START TYPE, PID, ar DISPLAY NAME."
    )
    doc.add_heading("Step 5 Snapshot save", level=2)
    add_code(
        doc,
        "python main.py snapshot --label before\n"
        "python main.py snapshot --label after"
    )
    doc.add_heading("Step 6 Snapshot list ebong diff", level=2)
    add_code(
        doc,
        "python main.py snapshots\n"
        "python main.py diff\n"
        "python main.py diff 1 2"
    )
    doc.add_paragraph(
        "ID na dile latest duita snapshot compare hobe. ID dile oi duita selected snapshot compare hobe."
    )
    doc.add_heading("Step 7 Shared PID analysis", level=2)
    add_code(doc, "python main.py groups --limit 10")
    doc.add_heading("Step 8 Continuous watch", level=2)
    add_code(doc, "python main.py watch --interval 3")
    doc.add_paragraph("Bondho korte Ctrl+C press korbe.")

    doc.add_heading("Nijer hate demo korar steps", level=1)
    doc.add_paragraph(
        "Teacher ke real state change dekhate hole safe service select korte hobe. Je service somporke sure na, "
        "seta kokhono stop koro na. Lab PC te instructor je service approve korben shudhu seta use koro."
    )
    add_bullets(
        doc,
        [
            "python main.py snapshot --label before command dao.",
            "services.msc open koro.",
            "Instructor approved service start ba stop koro.",
            "State stable howa porjonto kichu second wait koro.",
            "python main.py snapshot --label after command dao.",
            "python main.py diff command dao.",
            "Demo sheshe service original state e ferot nao.",
        ],
        numbered=True,
    )
    doc.add_paragraph(
        "Alternative hisebe watch mode chalu rekhe instructor approved service change korle program change "
        "sathe sathe print korbe. Project code nij theke kono service modify kore na."
    )

    doc.add_heading("Output kibhabe bujhbe", level=1)
    doc.add_heading("Ei machine er verified result", level=2)
    add_table(
        doc,
        ["Result", "Value"],
        [
            ["Total service", "314"],
            ["Running", "139"],
            ["Stopped", "175"],
            ["Automatic", "92"],
            ["Automatic delayed", "14"],
            ["Manual", "197"],
            ["Disabled", "10"],
        ],
        [3.7, 2.9],
    )
    doc.add_heading("Shared host example", level=2)
    doc.add_paragraph(
        "PID 1556 er vitore 5 ta service chilo: BrokerInfrastructure, DcomLaunch, PlugPlay, Power, ebong "
        "SystemEventsBroker. Ek PID te 5 ta service dekha mane service ar process ek jinis na."
    )
    doc.add_heading("No changes message", level=2)
    doc.add_paragraph(
        "No service state, startup type, or PID changes were detected message mane duita snapshot er majhe "
        "monitor korar moto kono change hoyni. Eta error na."
    )
    doc.add_heading("ACCESS_DENIED mane ki", level=2)
    doc.add_paragraph(
        "Kono protected service runtime list e dekhate pare, kintu tar configuration read permission na-o dite "
        "pare. Tokhon program crash na kore START TYPE column e ACCESS_DENIED rakhe."
    )

    doc.add_heading("Common problem ebong solution", level=1)
    add_table(
        doc,
        ["Problem", "Possible karon", "Solution"],
        [
            ["python command not found", "Python PATH e nei", "Python install kore Add Python to PATH select koro"],
            ["scm.py can only run on Windows", "Linux ba macOS e run", "Windows 10 ba 11 use koro"],
            ["At least two snapshots are required", "Database e ekta ba zero snapshot", "Aro duita snapshot nao"],
            ["Snapshot ID does not exist", "Vul ID deya", "python main.py snapshots diye ID dekho"],
            ["ACCESS_DENIED", "Protected service", "Field ta note koro; program ke admin deya mandatory na"],
            ["Diff e kono change nei", "Duita time e state same", "Approved service change kore abar snapshot nao"],
            ["Output onek lomba", "Machine e onek service", "groups command use koro ba output scroll koro"],
        ],
        [1.65, 2.1, 2.85],
    )

    doc.add_heading("Viva preparation", level=1)
    qa = [
        ("SCM ki", "Service Control Manager Windows er service database ebong lifecycle manage kore."),
        ("Service ar process er difference ki", "Service logical managed component. Process running executable. Onek service ek process share korte pare."),
        ("Runtime state ki", "Ei muhurte service running stopped paused ba pending kina."),
        ("Startup type ki", "Boot ba demand er somoy Windows service kibhabe start korbe tar configuration."),
        ("Two call sizing pattern keno", "Prothom call required buffer size ney, second call actual data ney."),
        ("ctypes keno use kora hoyeche", "Third party binding chara Python theke raw Win32 DLL function call korar jonno."),
        ("PID 0 mane ki", "Service er ekhon kono running host process nei, usually service stopped."),
        ("SQLite keno", "Lightweight file database, Python standard library te ache, snapshot query easy."),
        ("Diff ki ki compare kore", "State, startup type, PID, added service, ebong removed service."),
        ("Handle close na korle ki hoy", "Native resource leak hote pare."),
        ("Program ki service stop kore", "Na. Program query-only; observation kore, modification kore na."),
        ("Delayed automatic kivabe dhora hoy", "QueryServiceConfig2W er delayed auto-start information diye."),
    ]
    for question, answer in qa:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run("Q  " + question)
        run.bold = True
        p = doc.add_paragraph("A  " + answer)
        p.paragraph_format.left_indent = Inches(0.22)
        p.paragraph_format.space_after = Pt(6)

    doc.add_heading("Short revision sheet", level=1)
    add_bullets(
        doc,
        [
            "OpenSCManagerW diye SCM handle nei.",
            "EnumServicesStatusExW diye name state PID pai.",
            "QueryServiceConfigW diye startup type pai.",
            "QueryServiceConfig2W diye delayed automatic check kori.",
            "ServiceInfo object e combined data rakhi.",
            "SQLite e timestamp soho snapshot save kori.",
            "Diff e before ebong after compare kori.",
            "PID group diye shared svchost concept prove kori.",
            "Program query-only tai nij theke service change kore na.",
        ],
    )
    doc.add_paragraph(
        "Ei noyta line bujhe bolte parle project er main architecture ebong operating-system concept duita-i explain korte parbe."
    )

    path = OUT / "Services_Monitor_Banglish_Beginner_Guide_Submission.docx"
    doc.save(path)
    return path


if __name__ == "__main__":
    for report in (build_english_report(), build_banglish_guide()):
        print(report)
