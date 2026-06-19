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


# ─── Markdown Parser and PDF Generator ────────────────────────
def convert_md_to_pdf(md_path, pdf_path):
    # Initialize Document
    # Page height: 792. Left/right margin 54 (0.75in). Top/bottom margin 72 (1in) to clear header/footer.
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=54,
        rightMargin=54,
        topMargin=72,
        bottomMargin=72
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
        fontSize=24,
        leading=30,
        textColor=NAVY,
        spaceAfter=15
    )
    
    h2_style = ParagraphStyle(
        'Heading2',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=15,
        leading=20,
        textColor=colors.HexColor("#1E3A8A"), # Dark Blue
        spaceBefore=18,
        spaceAfter=8,
        keepWithNext=True
    )
    
    h3_style = ParagraphStyle(
        'Heading3',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=16,
        textColor=CHARCOAL,
        spaceBefore=14,
        spaceAfter=6,
        keepWithNext=True
    )

    body_style = ParagraphStyle(
        'BodyText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14.5,
        textColor=CHARCOAL,
        spaceAfter=10
    )

    bullet_style = ParagraphStyle(
        'BulletItem',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14.5,
        textColor=CHARCOAL,
        leftIndent=15,
        firstLineIndent=-10,
        spaceAfter=6
    )
    
    code_style = ParagraphStyle(
        'CodeText',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#111827"),
    )

    meta_key_style = ParagraphStyle(
        'MetaKey',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor("#4B5563")
    )
    
    meta_val_style = ParagraphStyle(
        'MetaVal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13,
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
    story.append(Spacer(1, 15))
    story.append(Paragraph(doc_title, title_style))
    
    # Simple line decoration under title
    title_line = Table([['']], colWidths=[504], rowHeights=[2.5])
    title_line.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), ACCENT_GOLD),
        ('PADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(title_line)
    story.append(Spacer(1, 15))

    # 2. Project Brief / Metadata Card
    if metadata:
        meta_data = []
        for k, v in metadata.items():
            meta_data.append([Paragraph(k, meta_key_style), Paragraph(v, meta_val_style)])
        
        meta_table = Table(meta_data, colWidths=[110, 374])
        meta_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F9FAFB")), # Off-white gray
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#E5E7EB")), # Light Border
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(meta_table)
        story.append(Spacer(1, 20))

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
                # Draw a code block table
                code_table_data = [[Paragraph("<br/>".join(code_block_lines), code_style)]]
                code_table = Table(code_table_data, colWidths=[504])
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
