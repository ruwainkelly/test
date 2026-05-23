"""
Generate a professional C-suite CV for Ruwain Kelly as a Word document.
Layout follows a clean executive style:
- Bold name banner with title
- Contact line
- Section headers with coloured background bars
- Two-column skill grids
- Role headings with dates right-aligned
- Bullet-pointed responsibilities and achievements
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement


# ---------- Theme ----------
PRIMARY = RGBColor(0x0B, 0x2E, 0x4F)      # Deep navy
ACCENT = RGBColor(0xB8, 0x8A, 0x2E)       # Muted gold
TEXT = RGBColor(0x22, 0x22, 0x22)         # Near-black
MUTED = RGBColor(0x55, 0x55, 0x55)        # Grey
LINE = RGBColor(0xCF, 0xCF, 0xCF)         # Light grey divider

FONT_BODY = "Calibri"
FONT_HEAD = "Calibri"


# ---------- Helpers ----------
def set_cell_shading(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)


def remove_cell_borders(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement('w:tcBorders')
    for border in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{border}')
        b.set(qn('w:val'), 'nil')
        tc_borders.append(b)
    tc_pr.append(tc_borders)


def add_horizontal_line(paragraph, color=LINE):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'),
               f"{color[0]:02X}{color[1]:02X}{color[2]:02X}")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def section_header(doc, text):
    """Coloured bar with white section title."""
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.columns[0].width = Cm(17)
    cell = table.cell(0, 0)
    cell.width = Cm(17)
    set_cell_shading(cell, '0B2E4F')
    remove_cell_borders(cell)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text.upper())
    run.font.name = FONT_HEAD
    run.font.size = Pt(11.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # spacing after the bar
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(2)
    spacer.paragraph_format.space_before = Pt(0)


def add_paragraph(doc, text, size=10.5, bold=False, italic=False,
                  color=TEXT, align=None, space_after=4, space_before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if align:
        p.alignment = align
    run = p.add_run(text)
    run.font.name = FONT_BODY
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return p


def add_bullet(doc, text, size=10.5, color=TEXT):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.6)
    run = p.runs[0] if p.runs else p.add_run()
    run.text = text
    run.font.name = FONT_BODY
    run.font.size = Pt(size)
    run.font.color.rgb = color
    return p


def add_role_header(doc, company, role, location_or_dates, dates=None):
    """Company on left bold, dates right-aligned."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    tab_stops = p.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Cm(17), WD_TAB_ALIGNMENT.RIGHT)

    r1 = p.add_run(company)
    r1.font.name = FONT_HEAD
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.color.rgb = PRIMARY

    if dates:
        p.add_run("\t")
        r2 = p.add_run(dates)
        r2.font.name = FONT_BODY
        r2.font.size = Pt(10.5)
        r2.font.italic = True
        r2.font.color.rgb = MUTED

    # Role line
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(2)
    r = p2.add_run(role)
    r.font.name = FONT_BODY
    r.font.size = Pt(11)
    r.font.italic = True
    r.font.color.rgb = ACCENT

    if location_or_dates and not dates:
        # legacy single-line variant
        pass


def two_column_skills(doc, columns):
    """columns: list of (title, [items]) groups; renders 2 per row."""
    # Pair up
    rows_data = []
    for i in range(0, len(columns), 2):
        left = columns[i]
        right = columns[i + 1] if i + 1 < len(columns) else ("", [])
        rows_data.append((left, right))

    table = doc.add_table(rows=len(rows_data), cols=2)
    table.autofit = False
    for col in table.columns:
        col.width = Cm(8.5)

    for r_idx, (left, right) in enumerate(rows_data):
        for c_idx, (title, items) in enumerate([left, right]):
            cell = table.cell(r_idx, c_idx)
            cell.width = Cm(8.5)
            remove_cell_borders(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

            # Clear default paragraph
            cell.paragraphs[0].text = ""
            if not title:
                continue

            ph = cell.paragraphs[0]
            ph.paragraph_format.space_after = Pt(2)
            ph.paragraph_format.space_before = Pt(2)
            rh = ph.add_run(title)
            rh.font.name = FONT_HEAD
            rh.font.size = Pt(10.5)
            rh.font.bold = True
            rh.font.color.rgb = PRIMARY

            for item in items:
                p = cell.add_paragraph()
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.left_indent = Cm(0.3)
                rb = p.add_run(f"•  {item}")
                rb.font.name = FONT_BODY
                rb.font.size = Pt(10)
                rb.font.color.rgb = TEXT

    # spacer
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def two_column_check_list(doc, items):
    """Render impact areas as a two-column ticked list."""
    half = (len(items) + 1) // 2
    left = items[:half]
    right = items[half:]
    rows = max(len(left), len(right))

    table = doc.add_table(rows=rows, cols=2)
    table.autofit = False
    for col in table.columns:
        col.width = Cm(8.5)

    for r in range(rows):
        for c, src in enumerate([left, right]):
            cell = table.cell(r, c)
            cell.width = Cm(8.5)
            remove_cell_borders(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            cell.paragraphs[0].text = ""
            if r >= len(src):
                continue
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r1 = p.add_run("✓  ")
            r1.font.name = FONT_BODY
            r1.font.size = Pt(10.5)
            r1.font.bold = True
            r1.font.color.rgb = ACCENT
            r2 = p.add_run(src[r])
            r2.font.name = FONT_BODY
            r2.font.size = Pt(10.5)
            r2.font.color.rgb = TEXT

    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_cert_row(doc, table, name, year):
    row = table.add_row()
    row.cells[0].width = Cm(13)
    row.cells[1].width = Cm(4)
    for cell in row.cells:
        remove_cell_borders(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    p = row.cells[0].paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(name)
    r.font.name = FONT_BODY
    r.font.size = Pt(10.5)
    r.font.color.rgb = TEXT

    p2 = row.cells[1].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.paragraph_format.space_after = Pt(0)
    r2 = p2.add_run(year)
    r2.font.name = FONT_BODY
    r2.font.size = Pt(10.5)
    r2.font.italic = True
    r2.font.color.rgb = MUTED


def cert_subheading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.font.name = FONT_HEAD
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = PRIMARY


def cert_table(doc):
    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    for col in table.columns:
        col.width = Cm(13) if col is table.columns[0] else Cm(4)
    return table


# ---------- Build Document ----------
doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

# Default style
style = doc.styles['Normal']
style.font.name = FONT_BODY
style.font.size = Pt(10.5)

# === HEADER BANNER ===
header_table = doc.add_table(rows=1, cols=1)
header_table.autofit = False
header_table.columns[0].width = Cm(17)
hcell = header_table.cell(0, 0)
hcell.width = Cm(17)
set_cell_shading(hcell, '0B2E4F')
remove_cell_borders(hcell)

# Name
p = hcell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(8)
p.paragraph_format.space_after = Pt(0)
run = p.add_run("RUWAIN KELLY")
run.font.name = FONT_HEAD
run.font.size = Pt(28)
run.font.bold = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

# Title
p2 = hcell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_before = Pt(2)
p2.paragraph_format.space_after = Pt(2)
r2 = p2.add_run("National Supply Chain, Procurement & Operations Executive")
r2.font.name = FONT_BODY
r2.font.size = Pt(12)
r2.font.color.rgb = RGBColor(0xE6, 0xC9, 0x7A)  # gold

# Tagline
p3 = hcell.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.paragraph_format.space_before = Pt(0)
p3.paragraph_format.space_after = Pt(8)
r3 = p3.add_run(
    "Transforming supply chains into competitive advantage through "
    "operational excellence, commercial discipline and strategic leadership."
)
r3.font.name = FONT_BODY
r3.font.size = Pt(10)
r3.font.italic = True
r3.font.color.rgb = RGBColor(0xE0, 0xE6, 0xEE)

# Contact line below banner
contact = doc.add_paragraph()
contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
contact.paragraph_format.space_before = Pt(6)
contact.paragraph_format.space_after = Pt(2)
contact_text = ("Durbanville, Cape Town, South Africa   |   "
                "(+27) 74 760 7776   |   admin@ruwainkelly.co.za")
cr = contact.add_run(contact_text)
cr.font.name = FONT_BODY
cr.font.size = Pt(10.5)
cr.font.color.rgb = PRIMARY
cr.font.bold = True

# Divider
divider = doc.add_paragraph()
divider.paragraph_format.space_before = Pt(0)
divider.paragraph_format.space_after = Pt(4)
add_horizontal_line(divider, color=ACCENT)

# === EXECUTIVE VALUE PROPOSITION ===
section_header(doc, "Executive Value Proposition")

paragraphs = [
    ("Senior Supply Chain, Procurement and Operations Executive with more than "
     "20 years of leadership experience across FMCG, manufacturing, logistics, "
     "warehousing and distribution environments."),
    ("Recognised for transforming complex operational environments into scalable, "
     "high-performance business functions that improve profitability, strengthen "
     "service delivery and optimise working capital."),
    ("Proven track record leading enterprise-wide procurement, supply chain and "
     "operational strategies with accountability for multi-site operations, "
     "supplier ecosystems, inventory investment, logistics performance, workforce "
     "leadership and governance frameworks."),
    ("Combines strong commercial acumen with operational execution excellence to "
     "drive sustainable cost reduction, improve supply continuity, enhance customer "
     "service levels and support long-term organisational growth."),
    ("Trusted by executive leadership teams to lead transformation initiatives, "
     "manage strategic supplier relationships, improve operational resilience and "
     "deliver measurable business outcomes in highly competitive and regulated "
     "environments."),
]
for para in paragraphs:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(para)
    r.font.name = FONT_BODY
    r.font.size = Pt(10.5)
    r.font.color.rgb = TEXT

# Impact Areas subhead
add_paragraph(doc, "Executive Impact Areas", size=11, bold=True,
              color=PRIMARY, space_before=4, space_after=4)

impact_areas = [
    "Procurement Cost Optimisation",
    "Working Capital Improvement",
    "Inventory & Demand Planning Excellence",
    "Multi-Site Operational Leadership",
    "Supply Chain Transformation",
    "Strategic Sourcing & Contract Negotiation",
    "Logistics & Distribution Optimisation",
    "Enterprise Risk & Governance",
    "SAP-Driven Operational Excellence",
    "Executive Stakeholder Management",
]
two_column_check_list(doc, impact_areas)

# === EXECUTIVE LEADERSHIP PROFILE ===
section_header(doc, "Executive Leadership Profile")

leadership_columns = [
    ("Commercial Leadership", [
        "Procurement Strategy Development",
        "Strategic Sourcing & Category Management",
        "Supplier Relationship Management",
        "Contract Negotiation",
        "Cost-to-Serve Optimisation",
        "Margin Protection Initiatives",
    ]),
    ("Supply Chain Excellence", [
        "End-to-End Supply Chain Leadership",
        "Sales & Operations Planning (S&OP)",
        "Demand Planning & Forecasting",
        "Inventory Optimisation",
        "Warehouse & Distribution Management",
        "Transport & Fleet Optimisation",
        "Supply Chain Risk Management",
    ]),
    ("Operational Leadership", [
        "Multi-Site Operations Management",
        "P&L Accountability",
        "Continuous Improvement",
        "Lean Six Sigma Methodologies",
        "Performance Management Frameworks",
        "Operational Governance",
        "Health & Safety Compliance",
        "Industrial Relations Leadership",
    ]),
    ("Executive Governance", [
        "Executive Committee Engagement",
        "Strategic Planning",
        "Business Transformation",
        "Change Leadership",
        "Risk Management",
        "Audit Readiness",
        "Regulatory Compliance",
        "Cross-Functional Leadership",
    ]),
]
two_column_skills(doc, leadership_columns)

# === PROFESSIONAL EXPERIENCE ===
section_header(doc, "Professional Experience")

# --- Quantum Foods ---
add_role_header(doc, "Quantum Foods",
                "National Procurement & Supply Chain Manager",
                None, dates="January 2023 – Present")
add_paragraph(doc, "Executive Mandate", size=10.5, bold=True,
              color=PRIMARY, space_before=4, space_after=2)
add_paragraph(doc,
    "Entrusted with strengthening procurement governance, improving supply chain "
    "resilience, optimising working capital and enhancing operational performance "
    "across national procurement and supply chain functions.",
    size=10.5, space_after=4)

add_paragraph(doc, "Strategic Leadership", size=10.5, bold=True,
              color=PRIMARY, space_before=2, space_after=2)
for b in [
    "Lead national procurement strategy across direct and indirect spend categories.",
    "Drive supplier performance, strategic sourcing and commercial negotiations.",
    "Oversee inventory investment, replenishment strategies and demand planning disciplines.",
    "Align procurement, operations and logistics through structured S&OP governance.",
    "Lead warehousing, transportation and distribution performance management.",
    "Implement procurement governance frameworks ensuring compliance and risk mitigation.",
    "Support executive decision-making through operational performance reporting and analytics.",
]:
    add_bullet(doc, b)

add_paragraph(doc, "Executive Contributions", size=10.5, bold=True,
              color=PRIMARY, space_before=4, space_after=2)
for b in [
    "Delivered sustainable procurement savings through supplier consolidation and strategic sourcing initiatives.",
    "Improved supplier performance through structured KPI governance and performance management frameworks.",
    "Reduced inventory exposure while maintaining product availability and service performance.",
    "Enhanced procurement controls, strengthening compliance and audit readiness.",
    "Increased supply chain visibility through executive reporting dashboards and performance scorecards.",
    "Improved collaboration between procurement, planning and operational teams to support business objectives.",
]:
    add_bullet(doc, b)

# --- Ecolab ---
add_role_header(doc, "Ecolab",
                "Regional Operations & Supply Chain Manager",
                None, dates="January 2020 – December 2022")
add_paragraph(doc, "Executive Mandate", size=10.5, bold=True,
              color=PRIMARY, space_before=4, space_after=2)
add_paragraph(doc,
    "Appointed to strengthen regional operational performance, improve service delivery, "
    "optimise supply chain execution and ensure sustainable profitability across multiple "
    "operational sites.",
    size=10.5, space_after=4)

add_paragraph(doc, "Leadership Scope", size=10.5, bold=True,
              color=PRIMARY, space_before=2, space_after=2)
for b in [
    "Regional operational budget exceeding R15 million.",
    "Multi-site warehousing and distribution operations.",
    "Procurement and inventory management.",
    "Fleet and transport operations.",
    "Customer service and service delivery.",
    "Workforce leadership across three operational facilities.",
]:
    add_bullet(doc, b)

add_paragraph(doc, "Executive Contributions", size=10.5, bold=True,
              color=PRIMARY, space_before=4, space_after=2)
for b in [
    "Improved OTIF performance to above 95% through operational process optimisation.",
    "Delivered significant procurement savings through strategic sourcing and supplier negotiations.",
    "Reduced fulfilment lead times through workflow redesign and operational improvements.",
    "Strengthened inventory control disciplines improving stock accuracy and visibility.",
    "Re-engineered logistics and third-party provider performance frameworks.",
    "Enhanced health, safety and compliance standards across all facilities.",
    "Improved operational governance and performance accountability through KPI management systems.",
]:
    add_bullet(doc, b)

# --- Let Me Do I.T. ---
add_role_header(doc, "Let Me Do I.T.", "Operations Director",
                None, dates="September 2014 – December 2019")
add_paragraph(doc, "Executive Mandate", size=10.5, bold=True,
              color=PRIMARY, space_before=4, space_after=2)
add_paragraph(doc,
    "Transform operational performance, improve scalability and implement systems and "
    "governance structures capable of supporting sustainable business growth.",
    size=10.5, space_after=4)

add_paragraph(doc, "Executive Contributions", size=10.5, bold=True,
              color=PRIMARY, space_before=2, space_after=2)
for b in [
    "Led operational transformation initiatives improving efficiency and profitability.",
    "Implemented ERP-enabled procurement and inventory management frameworks.",
    "Negotiated and managed strategic supplier agreements and service contracts.",
    "Established performance management systems improving operational visibility and accountability.",
    "Reduced operational costs through process redesign and workflow optimisation.",
    "Developed management reporting frameworks supporting strategic decision-making.",
    "Built scalable operational structures supporting business expansion.",
]:
    add_bullet(doc, b)

# --- Unitrans ---
add_role_header(doc, "Unitrans", "Operations & Logistics Manager",
                "Cape Town, South Africa", dates="January 2010 – August 2014")
add_paragraph(doc, "Cape Town, South Africa", size=10, italic=True,
              color=MUTED, space_after=2)
add_paragraph(doc,
    "Managed large-scale multi-site warehouse and transport operations within a "
    "high-pressure logistics environment.",
    size=10.5, space_after=4)
for b in [
    "Led 120+ staff across warehouse and fleet operations.",
    "Governed SAP-supported dispatch and inventory controls using WMS integration, reducing errors by 30%.",
    "Renegotiated transport and vendor contracts achieving 10% logistics cost reduction.",
    "Developed structured KPI frameworks improving throughput by 18%.",
    "Enforced Occupational Health & Safety Act compliance with zero major audit findings.",
    "Chaired disciplinary processes and supported labour compliance alignment.",
    "Balanced stock across regional nodes to maintain JIT service performance.",
    "Managed distributed operations across multiple sites using Agile-aligned sprint cadence for continuous operational improvement cycles.",
]:
    add_bullet(doc, b)

# --- Premier Foods ---
add_role_header(doc, "Premier Foods", "Warehouse & Logistics Manager",
                "Cape Town, South Africa", dates="January 2005 – December 2009")
add_paragraph(doc, "Cape Town, South Africa", size=10, italic=True,
              color=MUTED, space_after=2)
add_paragraph(doc,
    "Led high-volume national warehousing and distribution operations within a "
    "fast-paced FMCG environment.",
    size=10.5, space_after=4)
for b in [
    "Managed high-volume national distribution operations supporting production and retail supply chains.",
    "Directed 24/7 warehouse operations including receiving, storage, dispatch, and transport coordination.",
    "Implemented ERP-aligned inventory governance and WMS controls reducing stock loss by 15%.",
    "Redesigned route planning frameworks reducing national logistics costs by 10%.",
    "Led structured daily operational cadence meetings reviewing OTIF, dispatch variances, fleet utilisation, and exception management.",
    "Enforced OHSA compliance including safety audits, risk assessments, and incident reporting protocols.",
    "Managed union engagement and workforce governance, handling grievances and maintaining labour compliance alignment.",
]:
    add_bullet(doc, b)

# --- Wayne's Constructions ---
add_role_header(doc, "Wayne's Constructions",
                "HR & Operations Systems Generalist",
                "George, South Africa", dates="February 2001 – December 2004")
add_paragraph(doc, "George, South Africa", size=10, italic=True,
              color=MUTED, space_after=2)
add_paragraph(doc,
    "Provided HR, workforce governance, and operational systems support within a "
    "labour-intensive environment.",
    size=10.5, space_after=4)
for b in [
    "Managed end-to-end HR operations including recruitment, workforce planning, and labour compliance.",
    "Implemented early-stage HRIS and Workday-aligned HR system processes; digitised employee records to improve administrative control and reporting accuracy.",
    "Reduced hiring lead times by 30% through structured recruitment and onboarding processes.",
    "Improved employee retention by 15% through structured engagement and performance alignment initiatives.",
    "Delivered health, safety, and cross-training programmes aligned to regulatory and operational standards.",
    "Supported cross-functional operational systems governance across payroll, procurement, and project controls.",
    "Enforced Occupational Health & Safety Act compliance through formalised risk assessments, safety audits, and incident investigation frameworks.",
]:
    add_bullet(doc, b)

add_paragraph(doc, "Industrial Relations & Workforce Governance",
              size=10.5, bold=True, color=PRIMARY,
              space_before=4, space_after=2)
for b in [
    "Led union negotiations and grievance handling processes.",
    "Chaired disciplinary forums aligned with Labour Relations Act.",
    "Directed workforce planning and productivity modelling across distribution nodes.",
]:
    add_bullet(doc, b)

add_paragraph(doc, "OHSA Compliance", size=10.5, bold=True, color=PRIMARY,
              space_before=4, space_after=2)
add_paragraph(doc, "Enforced Occupational Health & Safety Act adherence through:",
              size=10.5, space_after=2)
for b in [
    "Formal risk assessments",
    "Safety audits",
    "Incident investigations",
    "Executive compliance reporting",
]:
    add_bullet(doc, b)
add_paragraph(doc, "Achieved zero major audit findings.",
              size=10.5, italic=True, color=MUTED, space_after=4)

# === ENTERPRISE ACHIEVEMENTS ===
section_header(doc, "Enterprise Achievements")

achievements_columns = [
    ("Financial Performance", [
        "Managed operational budgets exceeding R15 million.",
        "Delivered recurring procurement savings through strategic sourcing and supplier negotiations.",
        "Reduced operational expenditure through process optimisation and improved supplier management.",
        "Improved inventory efficiency and working capital utilisation.",
    ]),
    ("Operational Excellence", [
        "Achieved OTIF performance exceeding 95%.",
        "Improved inventory accuracy through strengthened operational controls.",
        "Reduced logistics and fulfilment lead times.",
        "Optimised warehouse and distribution operations across multiple sites.",
    ]),
    ("Leadership & Governance", [
        "Led teams exceeding 120 employees.",
        "Managed unionised workforces and industrial relations processes.",
        "Implemented governance frameworks supporting audit readiness and compliance.",
        "Established KPI-driven performance cultures focused on accountability and continuous improvement.",
    ]),
    ("Strategic Transformation", [
        "Led enterprise operational improvement programmes.",
        "Implemented ERP-enabled process improvements supporting procurement and supply chain functions.",
        "Improved operational visibility through executive performance reporting.",
        "Enhanced organisational resilience through strengthened supplier and operational governance.",
    ]),
]
two_column_skills(doc, achievements_columns)

# === CERTIFICATIONS & EDUCATION ===
section_header(doc, "Certifications & Education")

cert_subheading(doc, "Technical & Engineering Certifications")
t = cert_table(doc)
for name, year in [
    ("Computer Science Certificate — Harvard University", "2023"),
    ("Ethical Hacker — Cisco Network Academy", "2025"),
    ("AI for Everyone + Deep Learning Specialisation — Coursera (Andrew Ng)", "2025"),
    ("App Development Programme — FNB App Academy", "2025"),
    ("Full Stack Development — FreeCodeCamp", "2025"),
    ("Back End Development and APIs — FreeCodeCamp", "2025"),
    ("Machine Learning with Python — FreeCodeCamp", "2025"),
    ("Web Design — FreeCodeCamp", "2025"),
]:
    add_cert_row(doc, t, name, year)

cert_subheading(doc, "Supply Chain & Operations Certifications")
t = cert_table(doc)
for name, year in [
    ("Certified Supply Chain Professional (CSCP) — APICS", "2025"),
    ("Lean Six Sigma (Black Belt) — SSGI", "2025"),
    ("Diploma in Supply Chain Management — Coursera", "2023"),
]:
    add_cert_row(doc, t, name, year)

cert_subheading(doc, "Business, Project Management & Compliance")
t = cert_table(doc)
for name, year in [
    ("Diploma in Project Management — Coursera", "2023"),
    ("Diploma in Human Resources Management — Coursera", "2019"),
    ("Diploma in Health & Safety — Coursera", "2023"),
]:
    add_cert_row(doc, t, name, year)

# === LANGUAGES & WORK RIGHTS ===
section_header(doc, "Languages & Work Rights")
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
r = p.add_run("Languages:  ")
r.font.bold = True
r.font.name = FONT_BODY
r.font.size = Pt(10.5)
r.font.color.rgb = PRIMARY
r2 = p.add_run("English (Native)  ·  Afrikaans (Native)  ·  German (Conversational - dual SA/DE residency)")
r2.font.name = FONT_BODY
r2.font.size = Pt(10.5)
r2.font.color.rgb = TEXT

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
r = p.add_run("Work Rights:  ")
r.font.bold = True
r.font.name = FONT_BODY
r.font.size = Pt(10.5)
r.font.color.rgb = PRIMARY
r2 = p.add_run(
    "Cape Town: 10 Anysberg Crescent, Durbanville, ZA  |  "
    "Germany: Uerdinger Straße 44, Moers, DE  |  Remote, Hybrid or On-Site"
)
r2.font.name = FONT_BODY
r2.font.size = Pt(10.5)
r2.font.color.rgb = TEXT

# === LEADERSHIP & COMMUNITY ENGAGEMENT ===
section_header(doc, "Leadership & Community Engagement")
for b in [
    "Tech-for-Good Mentor – AI & SCM upskilling for underserved communities.",
    "Open-Source Contributor – Logistics automations, dashboards, backend tools.",
    "Cross-Cultural Advocate – EU/SA team bridging, stakeholder communication.",
    "Discipline-Driven Leader – Calisthenics & high-discipline routines.",
]:
    add_bullet(doc, b)

# === AVAILABILITY & LOGISTICS ===
section_header(doc, "Availability & Logistics")
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
r = p.add_run("Availability:  ")
r.font.bold = True
r.font.name = FONT_BODY
r.font.size = Pt(10.5)
r.font.color.rgb = PRIMARY
r2 = p.add_run("Immediate")
r2.font.name = FONT_BODY
r2.font.size = Pt(10.5)
r2.font.color.rgb = TEXT

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
r = p.add_run("Licences:  ")
r.font.bold = True
r.font.name = FONT_BODY
r.font.size = Pt(10.5)
r.font.color.rgb = PRIMARY
r2 = p.add_run("C1 Driver's Licence  |  Own vehicle")
r2.font.name = FONT_BODY
r2.font.size = Pt(10.5)
r2.font.color.rgb = TEXT

# === SAVE ===
output_path = "/projects/sandbox/test/Ruwain_Kelly_Executive_CV.docx"
doc.save(output_path)
print(f"Saved: {output_path}")
