"""
Ruwain Kelly – Institutional-Grade Executive CV
Simple, elegant, ATS-friendly. No blocks, no shading overload.
Clean single-column layout with refined typography and minimal accent lines.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─── Theme ────────────────────────────────────────────────────────────────────
NAVY = RGBColor(0x1B, 0x2A, 0x4A)
GOLD = RGBColor(0xA6, 0x83, 0x2E)
CHARCOAL = RGBColor(0x2C, 0x2C, 0x2C)
GREY = RGBColor(0x6B, 0x6B, 0x6B)
RULE_GREY = RGBColor(0xCC, 0xCC, 0xCC)
FONT = "Calibri"
PW = 17.0


# ─── Helpers ──────────────────────────────────────────────────────────────────
def _hex(c):
    return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def run(p, text, *, bold=False, italic=False, size=11, color=CHARCOAL,
        caps=False, spacing=None):
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    rPr = r._element.get_or_add_rPr()
    rF = rPr.find(qn('w:rFonts'))
    if rF is None:
        rF = OxmlElement('w:rFonts')
        rPr.insert(0, rF)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rF.set(qn(attr), FONT)
    if caps:
        c = OxmlElement('w:caps')
        c.set(qn('w:val'), '1')
        rPr.append(c)
    if spacing:
        s = OxmlElement('w:spacing')
        s.set(qn('w:val'), str(spacing))
        rPr.append(s)
    return r


def border_bottom(p, color=GOLD, size=6, space=1):
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single')
    b.set(qn('w:sz'), str(size))
    b.set(qn('w:space'), str(space))
    b.set(qn('w:color'), _hex(color))
    bdr.append(b)
    pPr.append(bdr)


def thin_rule(doc, color=RULE_GREY, size=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    border_bottom(p, color=color, size=size)


# ─── Components ───────────────────────────────────────────────────────────────
def header(doc):
    # Name
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    run(p, "RUWAIN KELLY", bold=True, size=28, color=NAVY, spacing=120)

    # Gold rule
    r = doc.add_paragraph()
    r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r.paragraph_format.space_before = Pt(0)
    r.paragraph_format.space_after = Pt(4)
    border_bottom(r, color=GOLD, size=10)

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    run(p, "National Supply Chain, Procurement & Operations Executive",
        size=13, color=NAVY, italic=True)

    # Tagline
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    run(p, "Transforming supply chains into competitive advantage through "
        "operational excellence, commercial discipline and strategic leadership.",
        size=10, color=GREY, italic=True)

    # Contact
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    run(p, "Durbanville, Cape Town, South Africa  |  +27 74 760 7776  |  "
        "info@ruwainkelly.com", size=10.5, color=CHARCOAL)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    run(p, "linkedin.com/in/ruwainkelly", size=10.5, color=NAVY, italic=True)

    thin_rule(doc, color=RULE_GREY)


def section(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(3)
    run(p, title, bold=True, size=12, color=NAVY, caps=True, spacing=60)
    border_bottom(p, color=GOLD, size=6)


def body_text(doc, text, *, justify=True, space_after=5):
    p = doc.add_paragraph()
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.2
    run(p, text)
    return p


def role(doc, company, title, location, dates):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(PW), WD_TAB_ALIGNMENT.RIGHT)
    run(p, company, bold=True, size=12, color=NAVY)
    p.add_run("\t")
    run(p, dates, italic=True, size=10.5, color=GREY)

    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(3)
    p2.paragraph_format.tab_stops.add_tab_stop(Cm(PW), WD_TAB_ALIGNMENT.RIGHT)
    run(p2, title, italic=True, size=11, color=GOLD, bold=True)
    if location:
        p2.add_run("\t")
        run(p2, location, italic=True, size=10, color=GREY)


def sub(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    run(p, text, bold=True, size=10.5, color=NAVY)


def bullet(doc, text, bold_lead=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.line_spacing = 1.15
    if p.runs:
        p.runs[0].text = ""
    if bold_lead:
        run(p, bold_lead, bold=True, size=10.5, color=NAVY)
        run(p, text, size=10.5)
    else:
        run(p, text, size=10.5)


def kw_line(doc, label, items):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.2
    run(p, f"{label}:  ", bold=True, size=10.5, color=NAVY)
    run(p, "  |  ".join(items), size=10.5, color=CHARCOAL)


def info_line(doc, label, value):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    run(p, f"{label}  ", bold=True, size=10.5, color=NAVY)
    run(p, value, size=10.5)


def achievement(doc, label, value):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.2
    run(p, f"{label}:  ", bold=True, size=10.5, color=NAVY)
    run(p, value, size=10.5)


# ─── Build ────────────────────────────────────────────────────────────────────
doc = Document()
for s in doc.sections:
    s.top_margin = Cm(1.8)
    s.bottom_margin = Cm(1.8)
    s.left_margin = Cm(2.2)
    s.right_margin = Cm(2.2)

normal = doc.styles['Normal']
normal.font.name = FONT
normal.font.size = Pt(10.5)
normal.font.color.rgb = CHARCOAL

try:
    doc.styles['List Bullet'].font.name = FONT
    doc.styles['List Bullet'].font.size = Pt(10.5)
except:
    pass

# === HEADER ===
header(doc)

# === EXECUTIVE SUMMARY ===
section(doc, "Executive Summary")
body_text(doc,
    "Senior Supply Chain, Procurement and Operations Executive with 15+ years of "
    "leadership across FMCG, manufacturing, logistics, warehousing and distribution. "
    "I turn fragmented, high-cost operations into scalable, high-performance functions "
    "that improve EBITDA, protect working capital and strengthen service delivery.")
body_text(doc,
    "Trusted by executive teams to lead enterprise-wide transformation, manage "
    "multi-site P&Ls, re-engineer supplier ecosystems and embed governance that "
    "survives audit scrutiny. My approach combines commercial rigour with operational "
    "execution — delivering sustainable cost reduction, supply continuity and "
    "measurable business outcomes in competitive and regulated environments.")

# === CAREER HIGHLIGHTS ===
section(doc, "Career Highlights")
highlights = [
    ("Experience", "15+ years executive leadership across FMCG, manufacturing, logistics and distribution."),
    ("Financial Scale", "Managed budgets exceeding R25M; recurring procurement savings; OPEX reduction."),
    ("Service Delivery", "OTIF performance >95%; reduced fulfilment lead times at national scale."),
    ("Team Leadership", "Led teams of 120+ across warehouse, fleet, procurement and operations."),
    ("Transformation", "Enterprise programmes: ERP / SAP / WMS deployment, S&OP governance, KPI cultures."),
    ("Compliance", "Zero major audit findings; OHSA and Labour Relations Act leadership."),
    ("Geography", "Dual SA / DE residency  |  Remote, Hybrid or On-Site  |  Available immediately."),
]
for label, value in highlights:
    achievement(doc, label, value)

# === CORE COMPETENCIES ===
section(doc, "Core Competencies")
kw_line(doc, "Strategic", [
    "Procurement Cost Optimisation", "Working Capital Improvement",
    "Inventory Optimisation", "Supply Chain Transformation",
    "SAP-Driven Operational Excellence"])
kw_line(doc, "Commercial", [
    "Strategic Sourcing & Category Management", "Contract Negotiation",
    "Supplier Relationship Management", "Cost-to-Serve Optimisation",
    "Margin Protection"])
kw_line(doc, "Supply Chain", [
    "End-to-End Supply Chain", "S&OP / Demand Planning",
    "Warehouse & Distribution", "Transport & Fleet Optimisation",
    "Supply Chain Risk Management"])
kw_line(doc, "Operations", [
    "Multi-Site P&L Accountability", "Lean Six Sigma",
    "Performance & KPI Management", "Health & Safety Compliance",
    "Industrial & Union Relations"])
kw_line(doc, "Governance", [
    "Executive Committee Engagement", "Strategic Planning",
    "Change Leadership", "Risk & Audit Readiness",
    "Cross-Functional Leadership"])
kw_line(doc, "Systems", [
    "SAP", "ERP (SYSPRO, Odoo, Sage)", "WMS", "Power BI",
    "Advanced Excel", "MS Project"])

# === PROFESSIONAL EXPERIENCE ===
section(doc, "Professional Experience")

# Quantum Foods
role(doc, "Quantum Foods", "National Procurement & Supply Chain Manager",
     "South Africa", "Jan 2023 – Present")
sub(doc, "Executive Mandate")
body_text(doc, "Strengthen procurement governance, improve supply chain resilience "
    "and optimise working capital across national operations.", space_after=3)
sub(doc, "Strategic Leadership")
for b in [
    "Lead national procurement strategy across direct and indirect spend, supplier performance and commercial negotiations.",
    "Govern inventory investment, replenishment strategies and demand planning disciplines.",
    "Align procurement, operations and logistics through structured S&OP governance.",
    "Drive warehousing, transportation and distribution performance management.",
    "Provide executive decision support via operational performance analytics and dashboards.",
]: bullet(doc, b)
sub(doc, "Key Contributions")
for l, b in [
    ("Cost: ", "delivered sustainable procurement savings through supplier consolidation and strategic sourcing."),
    ("Supplier Performance: ", "improved OTIF (On-Time, In-Full) via structured KPI governance."),
    ("Working Capital: ", "reduced inventory exposure while maintaining product availability."),
    ("Governance: ", "enhanced procurement controls, strengthening audit readiness and compliance."),
    ("Visibility: ", "launched executive dashboards that improved cross-functional collaboration."),
]: bullet(doc, b, bold_lead=l)

# Ecolab
role(doc, "Ecolab", "Regional Operations & Supply Chain Manager",
     "South Africa", "Jan 2020 – Dec 2022")
sub(doc, "Executive Mandate")
body_text(doc, "Turn around regional operational performance, service delivery and "
    "profitability across multiple sites.", space_after=3)
sub(doc, "Leadership Scope")
for l, b in [
    ("P&L: ", "R15M+ operational budget."),
    ("Footprint: ", "multi-site warehousing and distribution operations."),
    ("Functions: ", "procurement, inventory, fleet and customer service."),
    ("Workforce: ", "75 employees across three operational facilities."),
]: bullet(doc, b, bold_lead=l)
sub(doc, "Key Contributions")
for l, b in [
    ("Service: ", "improved OTIF to >95% through process optimisation."),
    ("Cost: ", "delivered double-digit procurement savings via strategic sourcing."),
    ("Lead Time: ", "reduced fulfilment lead times through workflow redesign."),
    ("Inventory: ", "strengthened stock accuracy and visibility."),
    ("3PL: ", "re-engineered logistics provider performance frameworks."),
    ("Compliance: ", "enhanced health, safety and regulatory standards across all sites."),
    ("Governance: ", "embedded KPI management systems to drive accountability."),
]: bullet(doc, b, bold_lead=l)

# Let Me Do I.T.
role(doc, "Let Me Do I.T.", "Chief of Operations",
     "South Africa", "Sep 2014 – Dec 2019")
sub(doc, "Executive Mandate")
body_text(doc, "Transform operational performance, scalability and governance to "
    "support business growth.", space_after=3)
sub(doc, "Key Contributions")
for b in [
    "Led operational transformation initiatives, improving efficiency and profitability.",
    "Implemented ERP-enabled procurement and inventory management frameworks.",
    "Negotiated and managed strategic supplier agreements and service contracts.",
    "Established performance management systems improving operational visibility.",
    "Reduced operational costs through process redesign and workflow optimisation.",
    "Developed management reporting frameworks supporting strategic decision-making.",
    "Built scalable operational structures that supported business expansion.",
]: bullet(doc, b)

# Unitrans
role(doc, "Unitrans", "Operations & Logistics Manager",
     "Cape Town, South Africa", "Jan 2010 – Aug 2014")
sub(doc, "Executive Mandate")
body_text(doc, "Manage large-scale, multi-site warehouse and transport operations "
    "in a high-pressure logistics environment.", space_after=3)
sub(doc, "Key Contributions")
for l, b in [
    ("Team: ", "led 120+ staff across warehouse and fleet operations."),
    ("Technology: ", "governed SAP-supported WMS integration — reduced dispatch errors by 30%."),
    ("Cost: ", "renegotiated transport and vendor contracts — achieved 10% logistics cost reduction."),
    ("Productivity: ", "developed structured KPI frameworks — improved throughput by 18%."),
    ("Compliance: ", "enforced OHSA compliance — zero major audit findings."),
    ("Labour: ", "chaired disciplinary processes and supported labour alignment."),
    ("Service: ", "balanced stock across regional nodes to maintain JIT performance."),
    ("Agile Ops: ", "applied Agile sprint cadence for continuous improvement across multiple sites."),
]: bullet(doc, b, bold_lead=l)

# Premier Foods
role(doc, "Premier Foods", "Warehouse & Logistics Manager",
     "Cape Town, South Africa", "Jan 2005 – Dec 2009")
sub(doc, "Executive Mandate")
body_text(doc, "Lead high-volume national warehousing and distribution operations "
    "in a fast-paced FMCG environment.", space_after=3)
sub(doc, "Key Contributions")
for l, b in [
    ("Scope: ", "managed 24/7 national distribution — receiving, storage, dispatch and transport."),
    ("Inventory: ", "implemented ERP-aligned WMS controls — reduced stock loss by 15%."),
    ("Cost: ", "redesigned route planning — cut national logistics costs by 10%."),
    ("Governance: ", "led daily cadence meetings — tracked OTIF, dispatch variances, fleet utilisation."),
    ("Safety: ", "enforced OHSA compliance — safety audits, risk assessments and incident reporting."),
    ("Labour: ", "managed union engagement and workforce governance — maintained compliance."),
]: bullet(doc, b, bold_lead=l)

# Wayne's Constructions
role(doc, "Wayne's Constructions", "HR & Operations Systems Generalist",
     "George, South Africa", "Feb 2001 – Dec 2004")
sub(doc, "Executive Mandate")
body_text(doc, "Provide HR, workforce governance and operational systems support "
    "in a labour-intensive environment.", space_after=3)
sub(doc, "Key Contributions")
for l, b in [
    ("HR Operations: ", "end-to-end recruitment, workforce planning and labour compliance."),
    ("Digitisation: ", "implemented early HRIS processes — digitised records, improving reporting accuracy."),
    ("Efficiency: ", "reduced hiring lead times by 30% via structured recruitment and onboarding."),
    ("Retention: ", "improved employee retention by 15% through engagement and performance alignment."),
    ("Safety: ", "delivered OHSA compliance — risk assessments, safety audits, incident frameworks."),
    ("Governance: ", "supported payroll, procurement and project controls."),
]: bullet(doc, b, bold_lead=l)

# Cross-role
sub(doc, "Industrial Relations & Workforce Governance (Cross-Role)")
for b in [
    "Led union negotiations and grievance handling across multiple environments.",
    "Chaired disciplinary forums aligned with the Labour Relations Act.",
    "Directed workforce planning and productivity modelling across distribution nodes.",
    "Zero major audit findings — OHSA compliance via formal risk assessments and safety audits.",
]: bullet(doc, b)

# === ENTERPRISE ACHIEVEMENTS ===
section(doc, "Enterprise Achievements")
achievement(doc, "Financial", "Managed budgets exceeding R25M; recurring procurement savings; reduced OPEX; improved working capital.")
achievement(doc, "Operational", "Achieved >95% OTIF; improved inventory accuracy; reduced lead times; optimised multi-site operations.")
achievement(doc, "Leadership", "Led 120+ teams; managed unionised workforces; audit-ready governance; KPI-driven cultures.")
achievement(doc, "Transformation", "Enterprise improvement programmes; ERP deployment; executive reporting; supplier resilience.")

# === EDUCATION & CERTIFICATIONS ===
section(doc, "Education & Certifications")
sub(doc, "Supply Chain & Operations")
body_text(doc, "APICS CSCP (Certified Supply Chain Professional)  ·  Lean Six Sigma Black Belt  ·  "
    "Diploma in Supply Chain Management (Coursera).", justify=False, space_after=4)
sub(doc, "Business & Compliance")
body_text(doc, "Diploma in Project Management  ·  Diploma in Health & Safety  ·  "
    "Diploma in Human Resources Management (all Coursera).", justify=False, space_after=4)
sub(doc, "Technology & Engineering")
body_text(doc, "Computer Science (Harvard CS50)  ·  Ethical Hacker (Cisco)  ·  "
    "AI + Deep Learning (Andrew Ng)  ·  FNB App Development  ·  "
    "Full Stack, Back End, ML, Web Design (freeCodeCamp).", justify=False, space_after=4)

# === LANGUAGES & WORK RIGHTS ===
section(doc, "Languages & Work Rights")
info_line(doc, "Languages:", "English (Native)  ·  Afrikaans (Native)  ·  German (Conversational)")
info_line(doc, "Residency:", "Dual South African / German")
info_line(doc, "Addresses:", "Cape Town: 10 Anysberg Crescent, Durbanville  |  Germany: Uerdinger Str. 44, Moers")
info_line(doc, "Work Mode:", "Remote, Hybrid or On-Site")
info_line(doc, "Notice:", "Immediate")
info_line(doc, "Licences:", "Code C1 Driver's Licence  ·  Own Vehicle")

# === LEADERSHIP & COMMUNITY ===
section(doc, "Leadership & Community Engagement")
for l, b in [
    ("Tech-for-Good Mentor — ", "AI and supply chain upskilling for underserved communities."),
    ("Open-Source Contributor — ", "logistics automations, dashboards and backend tools."),
    ("Cross-Cultural Advocate — ", "bridging EU and SA teams and stakeholder communication."),
    ("Discipline-Driven Leader — ", "calisthenics and high-discipline routines demonstrating resilience."),
]: bullet(doc, b, bold_lead=l)

# === CLOSING ===
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(16)
p.paragraph_format.space_after = Pt(0)
border_bottom(p, color=GOLD, size=6)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(4)
run(p, "References available on request", italic=True, size=10, color=GREY)

# ─── Save ─────────────────────────────────────────────────────────────────────
doc.save("/projects/sandbox/test/Ruwain_Kelly_Executive_CV.docx")
print("Saved: Ruwain_Kelly_Executive_CV.docx")
