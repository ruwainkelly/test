"""Institutional-grade C-suite CV for Ruwain Kelly.

Visual identity
---------------
- Full-width navy hero banner with white name, gold title, grey tagline
  and white contact lines.
- Stats strip just under the banner (key numbers in cream tint with gold rule).
- Section headers rendered as full-width navy bars with white uppercase text.
- Role headers with company in uppercase navy + gold rule under the role line.
- Bold lead-in labels on quantified bullets.

ATS safety
----------
- Zero layout tables. Every visual block is a single paragraph with shading
  and/or borders applied via OOXML. ATS parsers strip these and read clean
  linear text.
- Standard heading names recognised by all major ATS systems.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---- Theme ---------------------------------------------------------------
NAVY = RGBColor(0x0F, 0x25, 0x44)
NAVY_HEX = "0F2544"
GOLD = RGBColor(0xC0, 0x97, 0x3C)
GOLD_HEX = "C0973C"
TEXT = RGBColor(0x1B, 0x1B, 0x1B)
MUTED = RGBColor(0x55, 0x5C, 0x66)
SOFT = RGBColor(0xD8, 0xDC, 0xE2)
CREAM = RGBColor(0xF6, 0xF1, 0xE5)
CREAM_HEX = "F6F1E5"
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD_LIGHT = RGBColor(0xE6, 0xC8, 0x7A)

FONT = "Calibri"
PAGE_W_CM = 17.0  # A4 minus 2cm margins



# ---- Low-level OOXML helpers ---------------------------------------------
def _hex(c: RGBColor) -> str:
    return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def add_run(p, text, *, bold=False, italic=False, size=11,
            color=TEXT, font=FONT, caps=False, spacing=None,
            underline=False):
    r = p.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.underline = underline
    rPr = r._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    rFonts.set(qn('w:cs'), font)
    rFonts.set(qn('w:eastAsia'), font)
    if caps:
        cap = OxmlElement('w:caps')
        cap.set(qn('w:val'), '1')
        rPr.append(cap)
    if spacing is not None:
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:val'), str(spacing))
        rPr.append(sp)
    return r


def shade_paragraph(paragraph, fill_hex):
    """Fill a paragraph background with the given colour (full text-frame width)."""
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    p_pr.append(shd)



def paragraph_border(paragraph, *, sides=('bottom',), color=GOLD,
                     size=8, space=2):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = paragraph._p.find(qn('w:pPr')).find(qn('w:pBdr'))
    if p_bdr is None:
        p_bdr = OxmlElement('w:pBdr')
        p_pr.append(p_bdr)
    for side in sides:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), str(size))
        b.set(qn('w:space'), str(space))
        b.set(qn('w:color'), _hex(color))
        p_bdr.append(b)


def paragraph_left_bar(paragraph, color=GOLD, size=24):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement('w:pBdr')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), str(size))
    left.set(qn('w:space'), '6')
    left.set(qn('w:color'), _hex(color))
    p_bdr.append(left)
    p_pr.append(p_bdr)


def new_p(doc, *, align=None, space_before=0, space_after=0,
          line_spacing=1.15, left_indent=None, right_indent=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if left_indent is not None:
        p.paragraph_format.left_indent = Cm(left_indent)
    if right_indent is not None:
        p.paragraph_format.right_indent = Cm(right_indent)
    return p



# ---- Hero banner ---------------------------------------------------------
def hero_banner(doc, name, title, tagline, contact1, contact2):
    """Full-width navy banner with name, title, tagline and contact lines.
    Implemented as consecutive shaded paragraphs (no tables)."""
    # Top padding strip
    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    shade_paragraph(p, NAVY_HEX)
    add_run(p, " ", size=4, color=NAVY)

    # Name
    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    shade_paragraph(p, NAVY_HEX)
    add_run(p, name, bold=True, size=32, color=WHITE,
            caps=True, spacing=140)

    # Tiny gap (still navy)
    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    shade_paragraph(p, NAVY_HEX)
    add_run(p, " ", size=4, color=NAVY)

    # Title in gold
    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    shade_paragraph(p, NAVY_HEX)
    add_run(p, title, italic=True, size=13, color=GOLD_LIGHT,
            caps=True, spacing=40)

    # Spacer
    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    shade_paragraph(p, NAVY_HEX)
    add_run(p, " ", size=2, color=NAVY)

    # Tagline
    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.15,
              left_indent=1.5, right_indent=1.5)
    shade_paragraph(p, NAVY_HEX)
    add_run(p, tagline, italic=True, size=10.5,
            color=RGBColor(0xE0, 0xE6, 0xEE))

    # Inner gold rule (full-width within the banner)
    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0,
              space_before=4)
    shade_paragraph(p, NAVY_HEX)
    paragraph_border(p, sides=('bottom',), color=GOLD, size=8)
    add_run(p, " ", size=2, color=NAVY)

    # Contact line 1 (white)
    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.2,
              space_before=2)
    shade_paragraph(p, NAVY_HEX)
    add_run(p, contact1, size=10.5, color=WHITE)

    # Contact line 2 (white)
    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.2)
    shade_paragraph(p, NAVY_HEX)
    add_run(p, contact2, size=10.5, color=WHITE)

    # Bottom padding
    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    shade_paragraph(p, NAVY_HEX)
    add_run(p, " ", size=4, color=NAVY)



# ---- Stats strip ---------------------------------------------------------
def stats_strip(doc, items):
    """Cream-tinted strip with bold navy stats separated by gold pipes."""
    # Top spacer
    new_p(doc, space_before=2, space_after=0, line_spacing=1.0)

    p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER,
              left_indent=0.2, right_indent=0.2,
              space_before=2, space_after=2, line_spacing=1.3)
    shade_paragraph(p, CREAM_HEX)
    paragraph_border(p, sides=('top', 'bottom'), color=GOLD, size=6)
    for i, item in enumerate(items):
        if i > 0:
            add_run(p, "  |  ", bold=True, size=11, color=GOLD)
        add_run(p, item, bold=True, size=11, color=NAVY,
                caps=True, spacing=20)


# ---- Section header (full-width navy bar) --------------------------------
def section_header(doc, text):
    p = new_p(doc, space_before=12, space_after=4, line_spacing=1.15,
              left_indent=0.25)
    shade_paragraph(p, NAVY_HEX)
    add_run(p, text, bold=True, size=12.5, color=WHITE,
            caps=True, spacing=80)


# ---- Body / bullets / sub-headings ---------------------------------------
def body(doc, text, *, justify=True, size=11, color=TEXT,
         italic=False, bold=False, space_after=4, space_before=0):
    p = new_p(doc, space_before=space_before, space_after=space_after,
              line_spacing=1.2)
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(p, text, size=size, color=color, italic=italic, bold=bold)
    return p


def bullet(doc, text, *, size=11, bold_lead=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.line_spacing = 1.18
    if p.runs:
        p.runs[0].text = ""
    if bold_lead:
        add_run(p, bold_lead, size=size, color=NAVY, bold=True)
        add_run(p, text, size=size, color=TEXT)
    else:
        add_run(p, text, size=size, color=TEXT)
    return p


def sub_heading(doc, text, *, size=11, space_before=4, space_after=1):
    p = new_p(doc, space_before=space_before, space_after=space_after,
              line_spacing=1.0)
    add_run(p, text, bold=True, size=size, color=NAVY,
            caps=True, spacing=30)



# ---- Role headers --------------------------------------------------------
def role_header(doc, company, role, location, dates):
    """Two-line role header with gold accent rule, no tables."""
    # Line 1: COMPANY (uppercase navy, larger) ........ DATES (italic muted)
    p = new_p(doc, space_before=10, space_after=0, line_spacing=1.0)
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(PAGE_W_CM), WD_TAB_ALIGNMENT.RIGHT)
    add_run(p, company, bold=True, size=13, color=NAVY,
            caps=True, spacing=60)
    p.add_run("\t")
    add_run(p, dates, italic=True, size=10.5, color=MUTED)

    # Line 2: Role (italic gold bold) ........ Location (muted)
    p = new_p(doc, space_before=0, space_after=2, line_spacing=1.0)
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(PAGE_W_CM), WD_TAB_ALIGNMENT.RIGHT)
    add_run(p, role, italic=True, bold=True, size=11.5, color=GOLD)
    if location:
        p.add_run("\t")
        add_run(p, location, italic=True, size=10, color=MUTED)
    paragraph_border(p, sides=('bottom',), color=GOLD, size=4)


# ---- Specialised lines ---------------------------------------------------
def keyword_line(doc, label, items, *, separator="  |  "):
    p = new_p(doc, space_before=1, space_after=3, line_spacing=1.2)
    add_run(p, f"{label}:  ", bold=True, size=11, color=NAVY)
    add_run(p, separator.join(items), size=11, color=TEXT)


def labeled_line(doc, label, value, *, size=11, space_after=2):
    p = new_p(doc, space_before=0, space_after=space_after, line_spacing=1.2)
    add_run(p, f"{label}  ", bold=True, size=size, color=NAVY)
    add_run(p, value, size=size, color=TEXT)


def category_line(doc, label, value):
    p = new_p(doc, space_before=2, space_after=3, line_spacing=1.25)
    add_run(p, f"{label} — ", bold=True, size=11, color=NAVY,
            caps=True, spacing=30)
    add_run(p, value, size=11, color=TEXT)


def callout_line(doc, label, value):
    """Gold left-bar callout used for the 'At a Glance' detail block."""
    p = new_p(doc, space_before=1, space_after=1, line_spacing=1.2,
              left_indent=0.3)
    shade_paragraph(p, CREAM_HEX)
    paragraph_left_bar(p, color=GOLD, size=24)
    add_run(p, f"{label}  ", bold=True, size=11, color=NAVY,
            caps=True, spacing=30)
    add_run(p, value, size=11, color=TEXT)



# ==========================================================================
# Build document
# ==========================================================================
doc = Document()

for section in doc.sections:
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

normal = doc.styles['Normal']
normal.font.name = FONT
normal.font.size = Pt(11)
normal.font.color.rgb = TEXT
normal.paragraph_format.line_spacing = 1.18

try:
    lb = doc.styles['List Bullet']
    lb.font.name = FONT
    lb.font.size = Pt(11)
except KeyError:
    pass


# === HERO BANNER =========================================================
hero_banner(
    doc,
    name="Ruwain Kelly",
    title="National Supply Chain, Procurement & Operations Executive",
    tagline=("Transforming supply chains into competitive advantage through "
             "operational excellence, commercial discipline and strategic leadership."),
    contact1="Durbanville, Cape Town, South Africa  |  +27 74 760 7776  |  info@ruwainkelly.com",
    contact2="linkedin.com/in/ruwainkelly",
)


# === STATS STRIP =========================================================
stats_strip(doc, [
    "15+ Years Executive",
    "R25M+ Budgets",
    "120+ Team Led",
    ">95% OTIF",
    "Zero Audit Findings",
])


# === EXECUTIVE SUMMARY ===================================================
section_header(doc, "Executive Summary")

body(doc,
     "Senior Supply Chain, Procurement and Operations Executive with 15+ years of "
     "leadership across FMCG, manufacturing, logistics, warehousing and distribution. "
     "I turn fragmented, high-cost operations into scalable, high-performance "
     "functions that improve EBITDA, protect working capital and strengthen service "
     "delivery.",
     space_after=6)

body(doc,
     "Trusted by executive teams to lead enterprise-wide transformation, manage "
     "multi-site P&Ls, re-engineer supplier ecosystems and embed governance that "
     "survives audit scrutiny. My approach combines commercial rigour with "
     "operational execution — delivering sustainable cost reduction, supply "
     "continuity and measurable business outcomes in competitive and regulated "
     "environments.",
     space_after=4)



# === IMPACT AT A GLANCE ==================================================
section_header(doc, "Impact at a Glance")
callout_line(doc, "Experience",
             "15+ years executive leadership across FMCG, manufacturing, logistics & distribution.")
callout_line(doc, "Financial Scale",
             "Managed budgets exceeding R25M; recurring procurement savings; OPEX reduction via process redesign.")
callout_line(doc, "Service Delivery",
             "OTIF performance >95%; reduced fulfilment lead times; multi-site distribution at national scale.")
callout_line(doc, "Team Leadership",
             "Led teams of 120+ across warehouse, fleet, procurement and operations.")
callout_line(doc, "Transformation",
             "Enterprise improvement programmes, ERP / SAP / WMS deployment, S&OP governance, KPI-driven cultures.")
callout_line(doc, "Compliance",
             "Zero major audit findings; OHSA, Labour Relations Act and regulatory compliance leadership.")
callout_line(doc, "Geography",
             "Dual SA / DE residency  |  Remote, Hybrid or On-Site  |  Available immediately.")


# === CORE COMPETENCIES ===================================================
section_header(doc, "Core Competencies")
keyword_line(doc, "Strategic Focus", [
    "Procurement Cost Optimisation", "Working Capital Improvement",
    "Inventory Optimisation", "Supply Chain Transformation",
    "SAP-Driven Operational Excellence",
])
keyword_line(doc, "Commercial Leadership", [
    "Strategic Sourcing & Category Management", "Contract Negotiation",
    "Supplier Relationship Management", "Cost-to-Serve Optimisation",
    "Margin Protection",
])
keyword_line(doc, "Supply Chain Excellence", [
    "End-to-End Supply Chain", "S&OP / Demand Planning",
    "Inventory & Demand Planning", "Warehouse & Distribution",
    "Transport & Fleet Optimisation",
])
keyword_line(doc, "Operational Leadership", [
    "Multi-Site P&L Accountability", "Continuous Improvement / Lean Six Sigma",
    "Performance & KPI Management", "Health & Safety Compliance",
    "Industrial & Union Relations",
])
keyword_line(doc, "Executive Governance", [
    "Executive Committee Engagement", "Strategic Planning & Transformation",
    "Change Leadership", "Risk & Audit Readiness",
    "Cross-Functional Leadership",
])
keyword_line(doc, "Systems & Tools", [
    "SAP", "ERP (SYSPRO, Odoo, Sage)", "WMS", "Power BI",
    "Advanced Excel", "MS Project", "BI Dashboards",
])



# === PROFESSIONAL EXPERIENCE =============================================
section_header(doc, "Professional Experience")

# --- Quantum Foods ---
role_header(doc, "Quantum Foods",
            "National Procurement & Supply Chain Manager",
            "South Africa", "Jan 2023 – Present")
sub_heading(doc, "Executive Mandate")
body(doc,
     "Strengthen procurement governance, improve supply chain resilience and "
     "optimise working capital across national operations.")
sub_heading(doc, "Strategic Leadership")
for b in [
    "Lead national procurement strategy across direct and indirect spend, supplier performance and commercial negotiations.",
    "Govern inventory investment, replenishment strategies and demand planning disciplines.",
    "Align procurement, operations and logistics through structured S&OP governance.",
    "Drive warehousing, transportation and distribution performance management.",
    "Provide executive decision support via operational performance analytics and dashboards.",
]:
    bullet(doc, b)
sub_heading(doc, "Key Contributions")
for lead, b in [
    ("Cost: ", "delivered sustainable procurement savings through supplier consolidation and strategic sourcing."),
    ("Supplier Performance: ", "improved OTIF (On-Time, In-Full) via structured KPI governance."),
    ("Working Capital: ", "reduced inventory exposure while maintaining product availability."),
    ("Governance: ", "enhanced procurement controls, strengthening audit readiness and compliance."),
    ("Visibility: ", "launched executive dashboards that improved cross-functional collaboration."),
]:
    bullet(doc, b, bold_lead=lead)

# --- Ecolab ---
role_header(doc, "Ecolab",
            "Regional Operations & Supply Chain Manager",
            "South Africa", "Jan 2020 – Dec 2022")
sub_heading(doc, "Executive Mandate")
body(doc,
     "Turn around regional operational performance, service delivery and "
     "profitability across multiple sites.")
sub_heading(doc, "Leadership Scope")
for lead, b in [
    ("P&L: ", "R15M+ operational budget."),
    ("Footprint: ", "multi-site warehousing and distribution operations."),
    ("Functions: ", "procurement, inventory, fleet and customer service."),
    ("Workforce: ", "75 employees across three operational facilities."),
]:
    bullet(doc, b, bold_lead=lead)
sub_heading(doc, "Key Contributions")
for lead, b in [
    ("Service: ", "improved OTIF to >95% through process optimisation."),
    ("Cost: ", "delivered double-digit procurement savings via strategic sourcing."),
    ("Lead Time: ", "reduced fulfilment lead times through workflow redesign."),
    ("Inventory: ", "strengthened stock accuracy and visibility."),
    ("3PL Management: ", "re-engineered logistics provider performance frameworks."),
    ("Compliance: ", "enhanced health, safety and regulatory standards across all sites."),
    ("Governance: ", "embedded KPI management systems to drive accountability."),
]:
    bullet(doc, b, bold_lead=lead)



# --- Let Me Do I.T. ---
role_header(doc, "Let Me Do I.T.", "Chief of Operations",
            "South Africa", "Sep 2014 – Dec 2019")
sub_heading(doc, "Executive Mandate")
body(doc,
     "Transform operational performance, scalability and governance to support "
     "business growth.")
sub_heading(doc, "Key Contributions")
for b in [
    "Led operational transformation initiatives, improving efficiency and profitability.",
    "Implemented ERP-enabled procurement and inventory management frameworks.",
    "Negotiated and managed strategic supplier agreements and service contracts.",
    "Established performance management systems improving operational visibility and accountability.",
    "Reduced operational costs through process redesign and workflow optimisation.",
    "Developed management reporting frameworks supporting strategic decision-making.",
    "Built scalable operational structures that supported business expansion.",
]:
    bullet(doc, b)

# --- Unitrans ---
role_header(doc, "Unitrans", "Operations & Logistics Manager",
            "Cape Town, South Africa", "Jan 2010 – Aug 2014")
sub_heading(doc, "Executive Mandate")
body(doc,
     "Manage large-scale, multi-site warehouse and transport operations in a "
     "high-pressure logistics environment.")
sub_heading(doc, "Key Contributions")
for lead, b in [
    ("Team: ", "led 120+ staff across warehouse and fleet operations."),
    ("Technology: ", "governed SAP-supported WMS integration — reduced dispatch errors by 30%."),
    ("Cost: ", "renegotiated transport and vendor contracts — achieved 10% logistics cost reduction."),
    ("Productivity: ", "developed structured KPI frameworks — improved throughput by 18%."),
    ("Compliance: ", "enforced OHSA compliance — zero major audit findings."),
    ("Labour: ", "chaired disciplinary processes and supported labour alignment."),
    ("Service: ", "balanced stock across regional nodes to maintain JIT performance."),
    ("Agile Operations: ", "applied Agile sprint cadence for continuous improvement cycles across multiple sites."),
]:
    bullet(doc, b, bold_lead=lead)

# --- Premier Foods ---
role_header(doc, "Premier Foods", "Warehouse & Logistics Manager",
            "Cape Town, South Africa", "Jan 2005 – Dec 2009")
sub_heading(doc, "Executive Mandate")
body(doc,
     "Lead high-volume national warehousing and distribution operations in a "
     "fast-paced FMCG environment.")
sub_heading(doc, "Key Contributions")
for lead, b in [
    ("Scope: ", "managed 24/7 national distribution — receiving, storage, dispatch and transport coordination."),
    ("Inventory: ", "implemented ERP-aligned WMS controls — reduced stock loss by 15%."),
    ("Cost: ", "redesigned route planning — cut national logistics costs by 10%."),
    ("Governance: ", "led daily operational cadence meetings — tracked OTIF, dispatch variances, fleet utilisation and exceptions."),
    ("Safety: ", "enforced OHSA compliance — safety audits, risk assessments and incident reporting."),
    ("Labour: ", "managed union engagement and workforce governance — handled grievances and maintained labour compliance."),
]:
    bullet(doc, b, bold_lead=lead)



# --- Wayne's Constructions ---
role_header(doc, "Wayne's Constructions",
            "HR & Operations Systems Generalist",
            "George, South Africa", "Feb 2001 – Dec 2004")
sub_heading(doc, "Executive Mandate")
body(doc,
     "Provide HR, workforce governance and operational systems support in a "
     "labour-intensive environment.")
sub_heading(doc, "Key Contributions")
for lead, b in [
    ("HR Operations: ", "end-to-end recruitment, workforce planning and labour compliance."),
    ("Digitisation: ", "implemented early HRIS / Workday-aligned processes — digitised employee records, improving control and reporting accuracy."),
    ("Efficiency: ", "reduced hiring lead times by 30% via structured recruitment and onboarding."),
    ("Retention: ", "improved employee retention by 15% through structured engagement and performance alignment."),
    ("Safety & Training: ", "delivered OHSA compliance — risk assessments, safety audits and incident investigation frameworks."),
    ("Cross-Functional Governance: ", "supported payroll, procurement and project controls."),
]:
    bullet(doc, b, bold_lead=lead)

# --- Cross-role expertise ---
sub_heading(doc, "Industrial Relations & Workforce Governance — Cross-Role Expertise",
            space_before=10)
for b in [
    "Led union negotiations and grievance handling processes across multiple operational environments.",
    "Chaired disciplinary forums aligned with the Labour Relations Act.",
    "Directed workforce planning and productivity modelling across distribution nodes.",
    "Enforced OHSA compliance through formal risk assessments, safety audits and incident investigations — zero major audit findings across roles.",
]:
    bullet(doc, b)


# === ENTERPRISE ACHIEVEMENTS =============================================
section_header(doc, "Enterprise Achievements — Roll-up of Metrics")
category_line(doc, "Financial Performance",
              "Managed budgets exceeding R25M; delivered recurring procurement savings; "
              "reduced OPEX via process optimisation; improved inventory efficiency and working capital.")
category_line(doc, "Operational Excellence",
              "Achieved >95% OTIF; improved inventory accuracy; reduced logistics lead times; "
              "optimised multi-site warehouse and distribution operations.")
category_line(doc, "Leadership & Governance",
              "Led teams of 120+; managed unionised workforces; implemented audit-ready "
              "governance frameworks; built KPI-driven accountability cultures.")
category_line(doc, "Strategic Transformation",
              "Led enterprise improvement programmes; deployed ERP-enabled processes; "
              "improved executive visibility via performance reporting; strengthened "
              "supplier and operational resilience.")



# === EDUCATION & CERTIFICATIONS ==========================================
section_header(doc, "Education & Certifications")

sub_heading(doc, "Supply Chain & Operations")
body(doc,
     "APICS CSCP (Certified Supply Chain Professional)  ·  Lean Six Sigma Black Belt  ·  "
     "Diploma in Supply Chain Management (Coursera).",
     justify=False, space_after=4)

sub_heading(doc, "Business & Compliance")
body(doc,
     "Diploma in Project Management  ·  Diploma in Health & Safety  ·  "
     "Diploma in Human Resources Management  (all Coursera).",
     justify=False, space_after=4)

sub_heading(doc, "Technology & Engineering")
body(doc,
     "Computer Science Certificate (Harvard CS50)  ·  Ethical Hacker (Cisco Networking Academy)  ·  "
     "AI for Everyone + Deep Learning Specialisation (Andrew Ng / DeepLearning.AI)  ·  "
     "FNB App Development Programme  ·  Full Stack Development, Back End & APIs, "
     "Machine Learning with Python, Web Design (freeCodeCamp).",
     justify=False, space_after=4)


# === LANGUAGES & WORK RIGHTS ============================================
section_header(doc, "Languages & Work Rights")
labeled_line(doc, "Languages:",
             "English (Native)  ·  Afrikaans (Native)  ·  German (Conversational)")
labeled_line(doc, "Residency:", "Dual South African / German citizenship.")
labeled_line(doc, "Cape Town:", "10 Anysberg Crescent, Durbanville, ZA")
labeled_line(doc, "Germany:", "Uerdinger Straße 44, Moers, DE")
labeled_line(doc, "Work Mode:", "Remote, Hybrid or On-Site")
labeled_line(doc, "Notice Period:", "Immediate")
labeled_line(doc, "Licences:", "Code C1 Driver's Licence  ·  Own Vehicle")


# === LEADERSHIP & COMMUNITY ENGAGEMENT ===================================
section_header(doc, "Leadership & Community Engagement")
for lead, b in [
    ("Tech-for-Good Mentor — ", "AI and supply chain upskilling for underserved communities."),
    ("Open-Source Contributor — ", "logistics automations, dashboards and backend tools."),
    ("Cross-Cultural Advocate — ", "bridging EU and SA teams and stakeholder communication."),
    ("Discipline-Driven Leader — ", "calisthenics and high-discipline daily routines, demonstrating resilience and consistency."),
]:
    bullet(doc, b, bold_lead=lead)


# === CLOSING =============================================================
p = new_p(doc, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=14, space_after=0, line_spacing=1.0)
paragraph_border(p, sides=('bottom',), color=GOLD, size=8)
add_run(p, "References available on request",
        italic=True, size=10, color=MUTED)


# ---- Save ---------------------------------------------------------------
output_path = "/projects/sandbox/test/Ruwain_Kelly_Executive_CV.docx"
doc.save(output_path)
print(f"Saved: {output_path}")
