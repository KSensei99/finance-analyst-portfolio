import os
import re
import sys
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
)
from reportlab.pdfgen import canvas

# ─── Two-Pass Canvas for Page Numbering ───────────────────────
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()

    def draw_page_number(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#6B7280")) # Cool Grey
        
        # Header (on all pages except page 1)
        if self._pageNumber > 1:
            self.drawString(54, 750, "CASE STUDY: FINANCIAL PROCESS ENGINEERING & AUTOMATION")
            self.setStrokeColor(colors.HexColor("#E5E7EB"))
            self.setLineWidth(0.5)
            self.line(54, 742, 558, 742)
            
        # Footer (on all pages)
        self.drawString(54, 36, "CONFIDENTIAL — NEXORIA COMMERCE INC.")
        page_str = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(558, 36, page_str)
        self.restoreState()


from reportlab.graphics.shapes import Drawing, Rect, String, Line

def create_architecture_drawing():
    # 504 x 195 drawing area
    d = Drawing(504, 195)
    
    # 1. Background box for the database (Midnight Navy)
    db_bg = Rect(12, 160, 480, 30, rx=4, ry=4)
    db_bg.fillColor = colors.HexColor("#0F172A")
    db_bg.strokeColor = colors.HexColor("#1E293B")
    db_bg.strokeWidth = 1
    d.add(db_bg)
    
    t1 = String(252, 177, "Relational Database Warehouse Layer (SQLite — nexoria.db)", textAnchor='middle')
    t1.fontName = 'Helvetica-Bold'
    t1.fontSize = 8.5
    t1.fillColor = colors.white
    d.add(t1)
    
    t2 = String(252, 167, "orders (32K+ rows) · customers · products · state_tax_rates · shipments · gl_journal_entries · budget · actuals", textAnchor='middle')
    t2.fontName = 'Helvetica'
    t2.fontSize = 7.0
    t2.fillColor = colors.HexColor("#94A3B8")
    d.add(t2)
    
    # 2. 9 Modules (3 Columns x 3 Rows)
    # Box dimensions: width 150, height 26
    # Column X: Col 1 = 12, Col 2 = 177, Col 3 = 342
    # Row Y: Row 1 = 122, Row 2 = 88, Row 3 = 54
    
    modules = [
        # Col 1: Operations (Teal Stripe)
        (12, 122, "01. Invoice Automation", "VBA & Python PDF Engine", "#0D9488"),
        (12, 88, "02. Sales Tax Engine", "PQ & Multi-State Exemption", "#0D9488"),
        (12, 54, "03. Shipping Allocator", "Pandas Cost Allocation", "#0D9488"),
        # Col 2: FP&A & Analytics (Blue Stripe)
        (177, 122, "04. BI Dashboard", "Power BI / DAX Star Schema", "#2563EB"),
        (177, 88, "05. P&L Automation", "openpyxl & COM month slicer", "#2563EB"),
        (177, 54, "06. AR Aging Report", "Pivots & Color-Coded KPIs", "#2563EB"),
        # Col 3: Strategy & Cloud (Purple Stripe)
        (342, 122, "07. Revenue Forecasting", "statsmodels SARIMA Engine", "#7C3AED"),
        (342, 88, "08. Quote Pipeline", "GAS Google Sheet to Gmail", "#7C3AED"),
        (342, 54, "09. Portfolio Delivery", "Case Study & Upwork Pack", "#7C3AED")
    ]
    
    for x, y, title, desc, stripe_color in modules:
        # Main Box
        box = Rect(x, y, 150, 26, rx=2, ry=2)
        box.fillColor = colors.HexColor("#F8FAFC")
        box.strokeColor = colors.HexColor("#E2E8F0")
        box.strokeWidth = 0.5
        d.add(box)
        
        # Left Accent Stripe
        stripe = Rect(x, y, 3.5, 26, rx=0.5, ry=0.5)
        stripe.fillColor = colors.HexColor(stripe_color)
        stripe.strokeColor = colors.transparent
        d.add(stripe)
        
        # Title text
        m_title = String(x + 77, y + 15, title, textAnchor='middle')
        m_title.fontName = 'Helvetica-Bold'
        m_title.fontSize = 7.5
        m_title.fillColor = colors.HexColor("#1E293B")
        d.add(m_title)
        
        # Desc text
        m_desc = String(x + 77, y + 6, desc, textAnchor='middle')
        m_desc.fontName = 'Helvetica'
        m_desc.fontSize = 6.5
        m_desc.fillColor = colors.HexColor("#64748B")
        d.add(m_desc)

    # 3. Deliverables (Bottom Row - Y = 10, height 20)
    # Output 1 (Blue/Teal theme)
    deliv_bg1 = Rect(12, 10, 150, 20, rx=2, ry=2)
    deliv_bg1.fillColor = colors.HexColor("#EFF6FF")
    deliv_bg1.strokeColor = colors.HexColor("#3B82F6")
    deliv_bg1.strokeWidth = 0.5
    d.add(deliv_bg1)
    
    o1 = String(87, 16, "Branded PDF Invoices & Run Logs", textAnchor='middle')
    o1.fontName = 'Helvetica-Bold'
    o1.fontSize = 7.0
    o1.fillColor = colors.HexColor("#1D4ED8")
    d.add(o1)
    
    # Output 2 (Green theme)
    deliv_bg2 = Rect(177, 10, 150, 20, rx=2, ry=2)
    deliv_bg2.fillColor = colors.HexColor("#ECFDF5")
    deliv_bg2.strokeColor = colors.HexColor("#10B981")
    deliv_bg2.strokeWidth = 0.5
    d.add(deliv_bg2)
    
    o2 = String(252, 16, "Automated P&L, AR & Tax Sheets", textAnchor='middle')
    o2.fontName = 'Helvetica-Bold'
    o2.fontSize = 7.0
    o2.fillColor = colors.HexColor("#047857")
    d.add(o2)
    
    # Output 3 (Purple theme)
    deliv_bg3 = Rect(342, 10, 150, 20, rx=2, ry=2)
    deliv_bg3.fillColor = colors.HexColor("#F5F3FF")
    deliv_bg3.strokeColor = colors.HexColor("#8B5CF6")
    deliv_bg3.strokeWidth = 0.5
    d.add(deliv_bg3)
    
    o3 = String(417, 16, "Power BI Dashboards & Forecasts", textAnchor='middle')
    o3.fontName = 'Helvetica-Bold'
    o3.fontSize = 7.0
    o3.fillColor = colors.HexColor("#6D28D9")
    d.add(o3)

    # 4. Connecting Lines
    # Data Layer to top row modules
    d.add(Line(87, 160, 87, 148, strokeColor=colors.HexColor("#94A3B8"), strokeWidth=0.5))
    d.add(Line(252, 160, 252, 148, strokeColor=colors.HexColor("#94A3B8"), strokeWidth=0.5))
    d.add(Line(417, 160, 417, 148, strokeColor=colors.HexColor("#94A3B8"), strokeWidth=0.5))
    
    # Vertical flows inside columns
    d.add(Line(87, 122, 87, 114, strokeColor=colors.HexColor("#CBD5E1"), strokeWidth=0.5))
    d.add(Line(87, 88, 87, 80, strokeColor=colors.HexColor("#CBD5E1"), strokeWidth=0.5))
    d.add(Line(87, 54, 87, 30, strokeColor=colors.HexColor("#3B82F6"), strokeWidth=0.5))
    
    d.add(Line(252, 122, 252, 114, strokeColor=colors.HexColor("#CBD5E1"), strokeWidth=0.5))
    d.add(Line(252, 88, 252, 80, strokeColor=colors.HexColor("#CBD5E1"), strokeWidth=0.5))
    d.add(Line(252, 54, 252, 30, strokeColor=colors.HexColor("#10B981"), strokeWidth=0.5))
    
    d.add(Line(417, 122, 417, 114, strokeColor=colors.HexColor("#CBD5E1"), strokeWidth=0.5))
    d.add(Line(417, 88, 417, 80, strokeColor=colors.HexColor("#CBD5E1"), strokeWidth=0.5))
    d.add(Line(417, 54, 417, 30, strokeColor=colors.HexColor("#8B5CF6"), strokeWidth=0.5))

    return d


# ─── Markdown Parser and PDF Generator ────────────────────────
def convert_md_to_pdf(md_path, pdf_path):
    # Initialize Document
    # Page height: 792. Left/right margin 45 (0.625in). Top margin 54 (0.75in), bottom margin 45.
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=45,
        rightMargin=45,
        topMargin=54,
        bottomMargin=45
    )

    styles = getSampleStyleSheet()
    
    # Custom Palette
    CHARCOAL = colors.HexColor("#1F2937")
    NAVY = colors.HexColor("#0F172A")
    TINT_BG = colors.HexColor("#F3F4F6")
    ACCENT_GOLD = colors.HexColor("#C69B3C")
    
    # Define custom typography styles
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12.5,
        leading=15,
        textColor=NAVY,
        spaceAfter=4
    )
    
    h2_style = ParagraphStyle(
        'Heading2',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12.5,
        textColor=colors.HexColor("#1E3A8A"), # Dark Blue
        spaceBefore=6,
        spaceAfter=2,
        keepWithNext=True
    )
    
    h3_style = ParagraphStyle(
        'Heading3',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=CHARCOAL,
        spaceBefore=4,
        spaceAfter=1,
        keepWithNext=True
    )

    body_style = ParagraphStyle(
        'BodyText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        textColor=CHARCOAL,
        spaceAfter=2
    )

    bullet_style = ParagraphStyle(
        'BulletItem',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        textColor=CHARCOAL,
        leftIndent=12,
        firstLineIndent=-8,
        spaceAfter=1
    )
    
    code_style = ParagraphStyle(
        'CodeText',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=6,
        leading=8,
        textColor=colors.HexColor("#111827"),
    )

    meta_key_style = ParagraphStyle(
        'MetaKey',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#4B5563")
    )
    
    meta_val_style = ParagraphStyle(
        'MetaVal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        textColor=CHARCOAL
    )

    story = []
    
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # Pre-parse metadata card
    metadata = {}
    md_content_lines = []
    
    for line in lines:
        stripped = line.strip()
        # Parse **Key**: Value metadata
        meta_match = re.match(r'^\*\*(Client|Specialization|Author)\*\*:\s*(.*)', stripped)
        if meta_match:
            metadata[meta_match.group(1)] = meta_match.group(2)
        elif stripped == '---' or stripped == '':
            continue
        elif stripped.startswith('# '):
            # Capture title
            doc_title = stripped[2:]
        else:
            md_content_lines.append(line)

    # 1. Document Title
    story.append(Spacer(1, 10))
    story.append(Paragraph(doc_title, title_style))
    
    # Simple line decoration under title
    title_line = Table([['']], colWidths=[522], rowHeights=[2.0])
    title_line.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), ACCENT_GOLD),
        ('PADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(title_line)
    story.append(Spacer(1, 10))

    # 2. Project Brief / Metadata Card
    if metadata:
        meta_data = []
        for k, v in metadata.items():
            meta_data.append([Paragraph(k, meta_key_style), Paragraph(v, meta_val_style)])
        
        meta_table = Table(meta_data, colWidths=[110, 412])
        meta_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F9FAFB")), # Off-white gray
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#E5E7EB")), # Light Border
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(meta_table)
        story.append(Spacer(1, 12))

    # 3. Main Content Parser
    in_code_block = False
    code_block_lines = []
    
    p_accumulator = []

    def flush_accumulator():
        if p_accumulator:
            text = " ".join(p_accumulator)
            # Replace bold markdown with ReportLab tags
            text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', text)
            story.append(Paragraph(text, body_style))
            p_accumulator.clear()

    for line in md_content_lines:
        stripped = line.strip()
        
        # Code block toggle
        if stripped.startswith('```'):
            if in_code_block:
                # End of code block
                # Check if this code block represents the system architecture diagram
                is_architecture = any("Relational Data Layer" in l or "Central SQLite Database" in l or "nexoria" in l or "01. Invoicing" in l for l in code_block_lines)
                if is_architecture:
                    from reportlab.platypus import Image as RLImage
                    img_path = os.path.join(base_dir, "..", "docs", "architecture_diagram.png")
                    img = RLImage(img_path, width=240, height=240)
                    img_table = Table([[img]], colWidths=[522])
                    img_table.setStyle(TableStyle([
                        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                        ('TOPPADDING', (0,0), (-1,-1), 6),
                    ]))
                    story.append(img_table)
                else:
                    # Draw a code block table
                    code_table_data = [[Paragraph("<br/>".join(code_block_lines), code_style)]]
                    code_table = Table(code_table_data, colWidths=[522])
                    code_table.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,-1), TINT_BG),
                        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
                        ('VALIGN', (0,0), (-1,-1), 'TOP'),
                        ('TOPPADDING', (0,0), (-1,-1), 8),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                        ('LEFTPADDING', (0,0), (-1,-1), 10),
                        ('RIGHTPADDING', (0,0), (-1,-1), 10),
                    ]))
                    story.append(code_table)
                story.append(Spacer(1, 10))
                code_block_lines.clear()
                in_code_block = False
            else:
                # Start of code block
                flush_accumulator()
                in_code_block = True
            continue

        if in_code_block:
            # Escape HTML characters for code block to prevent ReportLab rendering errors
            escaped = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').rstrip('\n')
            # Retain leading spaces for formatting
            leading_spaces = len(line) - len(line.lstrip(' '))
            escaped = '&nbsp;' * leading_spaces + escaped.lstrip()
            code_block_lines.append(escaped)
            continue

        # Headings
        if stripped.startswith('## '):
            flush_accumulator()
            header_text = re.sub(r'\*\*(.*?)\*\*', r'\1', stripped[3:])
            story.append(Paragraph(header_text, h2_style))
            continue
        elif stripped.startswith('### '):
            flush_accumulator()
            header_text = re.sub(r'\*\*(.*?)\*\*', r'\1', stripped[4:])
            story.append(Paragraph(header_text, h3_style))
            continue

        # Lists
        bullet_match = re.match(r'^[*+-]\s+(.*)', stripped)
        num_match = re.match(r'^(\d+)\.\s+(.*)', stripped)
        
        if bullet_match:
            flush_accumulator()
            item_text = bullet_match.group(1)
            item_text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', item_text)
            story.append(Paragraph(f"&bull; {item_text}", bullet_style))
            continue
        elif num_match:
            flush_accumulator()
            num_val = num_match.group(1)
            item_text = num_match.group(2)
            item_text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', item_text)
            story.append(Paragraph(f"{num_val}. {item_text}", bullet_style))
            continue

        # Empty lines
        if not stripped:
            flush_accumulator()
            continue

        # Accumulate paragraph text
        p_accumulator.append(stripped)

    # Flush any trailing paragraph
    flush_accumulator()

    # Build Document
    doc.build(story, canvasmaker=NumberedCanvas)


if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    md_file = os.path.join(base_dir, "Nexoria_Finance_Case_Study.md")
    pdf_file = os.path.join(base_dir, "Nexoria_Finance_Case_Study.pdf")
    
    if len(sys.argv) > 2:
        md_file = sys.argv[1]
        pdf_file = sys.argv[2]
        
    print(f"Reading markdown from: {md_file}")
    print(f"Generating PDF at: {pdf_file}")
    
    convert_md_to_pdf(md_file, pdf_file)
    print("PDF generation complete!")
