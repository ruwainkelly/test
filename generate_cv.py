"""
Generate a corporate, ATS-friendly executive CV for Ruwain Kelly.

Design principles
-----------------
- Single-column linear flow (parses cleanly in Workday, Greenhouse, Lever,
  Taleo, iCIMS, SmartRecruiters, BambooHR, etc.).
- Zero layout tables. Right-aligned dates use tab stops, not table cells.
- Standard section headings ("Executive Summary", "Core Competencies",
  "Professional Experience", "Education & Certifications", "Languages")
  that ATS keyword extractors recognise.
- Calibri throughout - universally rendered, ATS-friendly, modern.
- Visual identity carried through colour and typography only:
    * Deep navy primary
    * Warm gold accent (rules, role titles)
    * Charcoal body text, refined grey for muted metadata
- Inline pipe-separated keyword lists for skills - dense ATS keyword
  coverage in a single readable line.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------
NAVY = RGBColor(0x1A, 0x36, 0x5D)       # primary corporate navy
GOLD = RGBColor(0xB7, 0x90, 0x2C)       # warm gold accent
TEXT = RGBColor(0x20, 0x20, 0x20)       # near-black body
MUTED = RGBColor(0x55, 0x5C, 0x66)      # refined grey for metadata
SOFT = RGBColor(0xD8, 0xDC, 0xE2)       # soft separator grey

FONT = "Calibri"
PAGE_WIDTH_CM = 17.0   # 21cm A4 minus 2cm margins each side


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------
def _hex(c: RGBColor) -> str:
    return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def add_run(p, text, *, bold=False, italic=False, size=10.5,
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
    # Force eastAsia font too so Word doesn't substitute on some readers
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    rFonts.set(qn('w:cs'), font)
    if caps:
        cap = OxmlElement('w:caps')
        cap.set(qn('w:val'), '1')
        rPr.append(cap)
    if spacing is not None:
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:val'), str(spacing))
        rPr.append(sp)
    return r


def paragraph_border_bottom(paragraph, color=GOLD, size=8):
    """Add a thin coloured rule under a paragraph (used for section dividers)."""
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), '2')
    bottom.set(qn('w:color'), _hex(color))
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def horizontal_rule(doc, color=SOFT, size=6, space_before=2, space_after=2):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    paragraph_border_bottom(p, color=color, size=size)


# ---------------------------------------------------------------------------
# Composition primitives
# ---------------------------------------------------------------------------
def name_block(doc, name, title, contact, tagline):
    # Name - large, uppercase, letter-spaced, navy
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    add_run(p, name, bold=True, size=26, color=NAVY,
            caps=True, spacing=80)

    # Gold accent rule under name
    rule = doc.add_paragraph()
    rule.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rule.paragraph_format.space_before = Pt(0)
    rule.paragraph_format.space_after = Pt(4)
    paragraph_border_bottom(rule, color=GOLD, size=12)

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    add_run(p, title, italic=True, size=12.5, color=NAVY)

    # Tagline
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    add_run(p, tagline, italic=True, size=10, color=MUTED)

    # Contact
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    add_run(p, contact, size=10.5, color=TEXT)

    # Soft divider after header block
    horizontal_rule(doc, color=SOFT, size=6, space_before=4, space_after=2)


def section_header(doc, text):
    """Bold, uppercase, navy, letter-spaced section title with gold rule under it.
    Standard heading text - ATS-friendly."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(1)
    add_run(p, text, bold=True, size=12, color=NAVY,
            caps=True, spacing=60)
    paragraph_border_bottom(p, color=GOLD, size=8)
    # micro spacer
    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(2)


def body(doc, text, *, justify=True, size=10.5, color=TEXT,
         italic=False, bold=False, space_after=4, space_before=0):
    p = doc.add_paragraph()
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    add_run(p, text, size=size, color=color, italic=italic, bold=bold)
    return p


def bullet(doc, text, *, size=10.5):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(0.6)
    if p.runs:
        p.runs[0].text = ""
    add_run(p, text, size=size, color=TEXT)
    return p


def role_header(doc, company, role, location, dates):
    """Two-line role header, no tables. Uses tab stops for right-alignment."""
    # Line 1: Company (bold navy) ... dates (italic muted)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(PAGE_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT)
    add_run(p, company, bold=True, size=12, color=NAVY)
    p.add_run("\t")
    add_run(p, dates, italic=True, size=10.5, color=MUTED)

    # Line 2: Role (italic gold) ... location (muted)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(PAGE_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT)
    add_run(p, role, italic=True, bold=True, size=11, color=GOLD)
    if location:
        p.add_run("\t")
        add_run(p, location, size=10, color=MUTED)


def sub_heading(doc, text, *, color=NAVY, size=10.5,
                space_before=4, space_after=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    add_run(p, text, bold=True, size=size, color=color)


def keyword_line(doc, label, items, *, separator="  |  "):
    """ATS-optimised inline keyword list."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(3)
    add_run(p, f"{label}: ", bold=True, size=10.5, color=NAVY)
    add_run(p, separator.join(items), size=10.5, color=TEXT)


def cert_line(doc, name, year):
    """Single-line certification entry with year right-aligned via tab stop."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(PAGE_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT)
    add_run(p, name, size=10.5, color=TEXT)
    p.add_run("\t")
    add_run(p, year, size=10.5, italic=True, color=MUTED)


def labeled_line(doc, label, value):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    add_run(p, f"{label}  ", bold=True, size=10.5, color=NAVY)
    add_run(p, value, size=10.5, color=TEXT)


# ---------------------------------------------------------------------------
# Build document
# ---------------------------------------------------------------------------
doc = Document()

for section in doc.sections:
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

normal = doc.styles['Normal']
normal.font.name = FONT
normal.font.size = Pt(10.5)
normal.font.color.rgb = TEXT

# Tighten the default List Bullet style
try:
    lb = doc.styles['List Bullet']
    lb.font.name = FONT
    lb.font.size = Pt(10.5)
except KeyError:
    pass


# === HEADER ============================================================
name_block(
    doc,
    name="Ruwain Kelly",
    title="National Supply Chain, Procurement & Operations Executive",
    tagline=("Transforming supply chains into competitive advantage through "
             "operational excellence, commercial discipline and strategic leadership."),
    contact=("Durbanville, Cape Town, South Africa  |  +27 74 760 7776  |  "
             "admin@ruwainkelly.co.za"),
)

# === EXECUTIVE SUMMARY =================================================
section_header(doc, "Executive Summary")

for para in [
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
]:
    body(doc, para, justify=True, space_after=4)


# === CORE COMPETENCIES (ATS keyword block) =============================
section_header(doc, "Core Competencies")

keyword_line(doc, "Strategic Focus", [
    "Procurement Cost Optimisation",
    "Working Capital Improvement",
    "Inventory & Demand Planning",
    "Multi-Site Operational Leadership",
    "Supply Chain Transformation",
    "Strategic Sourcing & Contract Negotiation",
    "Logistics & Distribution Optimisation",
    "Enterprise Risk & Governance",
    "SAP-Driven Operational Excellence",
    "Executive Stakeholder Management",
])

keyword_line(doc, "Commercial Leadership", [
    "Procurement Strategy Development",
    "Strategic Sourcing & Category Management",
    "Supplier Relationship Management",
    "Contract Negotiation",
    "Cost-to-Serve Optimisation",
    "Margin Protection",
])

keyword_line(doc, "Supply Chain Excellence", [
    "End-to-End Supply Chain Leadership",
    "Sales & Operations Planning (S&OP)",
    "Demand Planning & Forecasting",
    "Inventory Optimisation",
    "Warehouse & Distribution Management",
    "Transport & Fleet Optimisation",
    "Supply Chain Risk Management",
])

keyword_line(doc, "Operational Leadership", [
    "Multi-Site Operations Management",
    "P&L Accountability",
    "Continuous Improvement",
    "Lean Six Sigma",
    "Performance Management",
    "Operational Governance",
    "Health & Safety Compliance",
    "Industrial Relations",
])

keyword_line(doc, "Executive Governance", [
    "Executive Committee Engagement",
    "Strategic Planning",
    "Business Transformation",
    "Change Leadership",
    "Risk Management",
    "Audit Readiness",
    "Regulatory Compliance",
    "Cross-Functional Leadership",
])

keyword_line(doc, "Systems & Tools", [
    "SAP", "ERP", "WMS", "S&OP Tools", "BI Dashboards",
    "Power BI", "Excel (Advanced)", "MS Project",
])


# === PROFESSIONAL EXPERIENCE ===========================================
section_header(doc, "Professional Experience")

# --- Quantum Foods ---
role_header(doc,
            company="Quantum Foods",
            role="National Procurement & Supply Chain Manager",
            location="South Africa",
            dates="January 2023 – Present")

sub_heading(doc, "Executive Mandate")
body(doc,
     "Entrusted with strengthening procurement governance, improving supply chain "
     "resilience, optimising working capital and enhancing operational performance "
     "across national procurement and supply chain functions.")

sub_heading(doc, "Strategic Leadership")
for b in [
    "Lead national procurement strategy across direct and indirect spend categories.",
    "Drive supplier performance, strategic sourcing and commercial negotiations.",
    "Oversee inventory investment, replenishment strategies and demand planning disciplines.",
    "Align procurement, operations and logistics through structured S&OP governance.",
    "Lead warehousing, transportation and distribution performance management.",
    "Implement procurement governance frameworks ensuring compliance and risk mitigation.",
    "Support executive decision-making through operational performance reporting and analytics.",
]:
    bullet(doc, b)

sub_heading(doc, "Key Contributions")
for b in [
    "Delivered sustainable procurement savings through supplier consolidation and strategic sourcing initiatives.",
    "Improved supplier performance through structured KPI governance and performance management frameworks.",
    "Reduced inventory exposure while maintaining product availability and service performance.",
    "Enhanced procurement controls, strengthening compliance and audit readiness.",
    "Increased supply chain visibility through executive reporting dashboards and performance scorecards.",
    "Improved collaboration between procurement, planning and operational teams to support business objectives.",
]:
    bullet(doc, b)

# --- Ecolab ---
role_header(doc,
            company="Ecolab",
            role="Regional Operations & Supply Chain Manager",
            location="South Africa",
            dates="January 2020 – December 2022")

sub_heading(doc, "Executive Mandate")
body(doc,
     "Appointed to strengthen regional operational performance, improve service delivery, "
     "optimise supply chain execution and ensure sustainable profitability across multiple "
     "operational sites.")

sub_heading(doc, "Leadership Scope")
for b in [
    "Regional operational budget exceeding R15 million.",
    "Multi-site warehousing and distribution operations.",
    "Procurement and inventory management.",
    "Fleet and transport operations.",
    "Customer service and service delivery.",
    "Workforce leadership across three operational facilities.",
]:
    bullet(doc, b)

sub_heading(doc, "Key Contributions")
for b in [
    "Improved OTIF performance to above 95% through operational process optimisation.",
    "Delivered significant procurement savings through strategic sourcing and supplier negotiations.",
    "Reduced fulfilment lead times through workflow redesign and operational improvements.",
    "Strengthened inventory control disciplines improving stock accuracy and visibility.",
    "Re-engineered logistics and third-party provider performance frameworks.",
    "Enhanced health, safety and compliance standards across all facilities.",
    "Improved operational governance and performance accountability through KPI management systems.",
]:
    bullet(doc, b)

# --- Let Me Do I.T. ---
role_header(doc,
            company="Let Me Do I.T.",
            role="Operations Director",
            location="South Africa",
            dates="September 2014 – December 2019")

sub_heading(doc, "Executive Mandate")
body(doc,
     "Transform operational performance, improve scalability and implement systems and "
     "governance structures capable of supporting sustainable business growth.")

sub_heading(doc, "Key Contributions")
for b in [
    "Led operational transformation initiatives improving efficiency and profitability.",
    "Implemented ERP-enabled procurement and inventory management frameworks.",
    "Negotiated and managed strategic supplier agreements and service contracts.",
    "Established performance management systems improving operational visibility and accountability.",
    "Reduced operational costs through process redesign and workflow optimisation.",
    "Developed management reporting frameworks supporting strategic decision-making.",
    "Built scalable operational structures supporting business expansion.",
]:
    bullet(doc, b)

# --- Unitrans ---
role_header(doc,
            company="Unitrans",
            role="Operations & Logistics Manager",
            location="Cape Town, South Africa",
            dates="January 2010 – August 2014")

body(doc,
     "Managed large-scale multi-site warehouse and transport operations within a "
     "high-pressure logistics environment.",
     space_before=2)

sub_heading(doc, "Key Contributions")
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
    bullet(doc, b)

# --- Premier Foods ---
role_header(doc,
            company="Premier Foods",
            role="Warehouse & Logistics Manager",
            location="Cape Town, South Africa",
            dates="January 2005 – December 2009")

body(doc,
     "Led high-volume national warehousing and distribution operations within a "
     "fast-paced FMCG environment.",
     space_before=2)

sub_heading(doc, "Key Contributions")
for b in [
    "Managed high-volume national distribution operations supporting production and retail supply chains.",
    "Directed 24/7 warehouse operations including receiving, storage, dispatch and transport coordination.",
    "Implemented ERP-aligned inventory governance and WMS controls reducing stock loss by 15%.",
    "Redesigned route planning frameworks reducing national logistics costs by 10%.",
    "Led structured daily operational cadence meetings reviewing OTIF, dispatch variances, fleet utilisation and exception management.",
    "Enforced OHSA compliance including safety audits, risk assessments and incident reporting protocols.",
    "Managed union engagement and workforce governance, handling grievances and maintaining labour compliance alignment.",
]:
    bullet(doc, b)

# --- Wayne's Constructions ---
role_header(doc,
            company="Wayne's Constructions",
            role="HR & Operations Systems Generalist",
            location="George, South Africa",
            dates="February 2001 – December 2004")

body(doc,
     "Provided HR, workforce governance and operational systems support within a "
     "labour-intensive environment.",
     space_before=2)

sub_heading(doc, "Key Contributions")
for b in [
    "Managed end-to-end HR operations including recruitment, workforce planning and labour compliance.",
    "Implemented early-stage HRIS and Workday-aligned HR system processes; digitised employee records to improve administrative control and reporting accuracy.",
    "Reduced hiring lead times by 30% through structured recruitment and onboarding processes.",
    "Improved employee retention by 15% through structured engagement and performance alignment initiatives.",
    "Delivered health, safety and cross-training programmes aligned to regulatory and operational standards.",
    "Supported cross-functional operational systems governance across payroll, procurement and project controls.",
    "Enforced Occupational Health & Safety Act compliance through formalised risk assessments, safety audits and incident investigation frameworks.",
]:
    bullet(doc, b)

sub_heading(doc, "Industrial Relations & Workforce Governance")
for b in [
    "Led union negotiations and grievance handling processes.",
    "Chaired disciplinary forums aligned with the Labour Relations Act.",
    "Directed workforce planning and productivity modelling across distribution nodes.",
]:
    bullet(doc, b)

sub_heading(doc, "OHSA Compliance")
for b in [
    "Enforced Occupational Health & Safety Act adherence through formal risk assessments, safety audits, incident investigations and executive compliance reporting.",
    "Achieved zero major audit findings.",
]:
    bullet(doc, b)


# === ENTERPRISE ACHIEVEMENTS ===========================================
section_header(doc, "Enterprise Achievements")

sub_heading(doc, "Financial Performance")
for b in [
    "Managed operational budgets exceeding R15 million.",
    "Delivered recurring procurement savings through strategic sourcing and supplier negotiations.",
    "Reduced operational expenditure through process optimisation and improved supplier management.",
    "Improved inventory efficiency and working capital utilisation.",
]:
    bullet(doc, b)

sub_heading(doc, "Operational Excellence")
for b in [
    "Achieved OTIF performance exceeding 95%.",
    "Improved inventory accuracy through strengthened operational controls.",
    "Reduced logistics and fulfilment lead times.",
    "Optimised warehouse and distribution operations across multiple sites.",
]:
    bullet(doc, b)

sub_heading(doc, "Leadership & Governance")
for b in [
    "Led teams exceeding 120 employees.",
    "Managed unionised workforces and industrial relations processes.",
    "Implemented governance frameworks supporting audit readiness and compliance.",
    "Established KPI-driven performance cultures focused on accountability and continuous improvement.",
]:
    bullet(doc, b)

sub_heading(doc, "Strategic Transformation")
for b in [
    "Led enterprise operational improvement programmes.",
    "Implemented ERP-enabled process improvements supporting procurement and supply chain functions.",
    "Improved operational visibility through executive performance reporting.",
    "Enhanced organisational resilience through strengthened supplier and operational governance.",
]:
    bullet(doc, b)


# === EDUCATION & CERTIFICATIONS ========================================
section_header(doc, "Education & Certifications")

sub_heading(doc, "Supply Chain & Operations")
for name, year in [
    ("Certified Supply Chain Professional (CSCP) - APICS", "2025"),
    ("Lean Six Sigma Black Belt - SSGI", "2025"),
    ("Diploma in Supply Chain Management - Coursera", "2023"),
]:
    cert_line(doc, name, year)

sub_heading(doc, "Business, Project Management & Compliance")
for name, year in [
    ("Diploma in Project Management - Coursera", "2023"),
    ("Diploma in Health & Safety - Coursera", "2023"),
    ("Diploma in Human Resources Management - Coursera", "2019"),
]:
    cert_line(doc, name, year)

sub_heading(doc, "Technology & Engineering")
for name, year in [
    ("Computer Science Certificate - Harvard University (CS50)", "2023"),
    ("Ethical Hacker - Cisco Networking Academy", "2025"),
    ("AI for Everyone + Deep Learning Specialisation - Coursera (Andrew Ng)", "2025"),
    ("App Development Programme - FNB App Academy", "2025"),
    ("Full Stack Development - freeCodeCamp", "2025"),
    ("Back End Development and APIs - freeCodeCamp", "2025"),
    ("Machine Learning with Python - freeCodeCamp", "2025"),
    ("Web Design - freeCodeCamp", "2025"),
]:
    cert_line(doc, name, year)


# === LANGUAGES & WORK RIGHTS ==========================================
section_header(doc, "Languages & Work Rights")
labeled_line(doc, "Languages:",
             "English (Native)  ·  Afrikaans (Native)  ·  German (Conversational)")
labeled_line(doc, "Residency:",
             "Dual SA / DE - Cape Town: 10 Anysberg Crescent, Durbanville, ZA  |  "
             "Germany: Uerdinger Straße 44, Moers, DE")
labeled_line(doc, "Work Mode:",
             "Remote, Hybrid or On-Site")


# === LEADERSHIP & COMMUNITY ENGAGEMENT ================================
section_header(doc, "Leadership & Community Engagement")
for b in [
    "Tech-for-Good Mentor - AI and supply chain upskilling for underserved communities.",
    "Open-Source Contributor - logistics automations, dashboards and backend tools.",
    "Cross-Cultural Advocate - bridging EU and SA teams and stakeholder communication.",
    "Discipline-Driven Leader - calisthenics and high-discipline daily routines.",
]:
    bullet(doc, b)


# === AVAILABILITY =====================================================
section_header(doc, "Availability")
labeled_line(doc, "Notice Period:", "Immediate")
labeled_line(doc, "Licences:", "Code C1 Driver's Licence  |  Own Vehicle")


# ---------------------------------------------------------------------------
# Save
# ---------------------------------------------------------------------------
output_path = "/projects/sandbox/test/Ruwain_Kelly_Executive_CV.docx"
doc.save(output_path)
print(f"Saved: {output_path}")
