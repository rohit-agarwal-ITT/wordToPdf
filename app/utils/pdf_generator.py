from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
import os
from reportlab.lib.utils import ImageReader
import io
from reportlab.graphics import renderPDF
from svglib.svglib import svg2rlg

class PDFGenerator:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
    
    def _setup_custom_styles(self):
        """Setup custom paragraph styles for better formatting"""
        # Custom styles for different paragraph types
        self.styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=18,
            spaceAfter=12,
            alignment=TA_CENTER
        ))
        
        self.styles.add(ParagraphStyle(
            name='CustomHeading',
            parent=self.styles['Heading2'],
            fontSize=14,
            spaceAfter=8,
            spaceBefore=12
        ))
        
        self.styles.add(ParagraphStyle(
            name='CustomNormal',
            parent=self.styles['Normal'],
            fontSize=11,
            spaceAfter=6,
            alignment=TA_JUSTIFY
        ))
    
    def _draw_header(self, canvas, doc):
        """Draws the company header on every page, matching the new sample exactly."""
        canvas.saveState()
        width, height = doc.pagesize
        y_top = height - 40
        # Draw SVG logo centered
        logo_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../samples/itt logo.svg'))
        if os.path.exists(logo_path):
            drawing = svg2rlg(logo_path)
            logo_width = 220
            logo_height = 48
            renderPDF.draw(drawing, canvas, width/2 - logo_width/2, y_top - logo_height)
            y_offset = y_top - logo_height - 5
        else:
            y_offset = y_top
        # CIN (top right)
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(width - 40, height - 48, "CIN: U72200RJ2009PTC028316")
        # Company Name (bold)
        canvas.setFont("Helvetica-Bold", 12)
        canvas.drawCentredString(width / 2, y_offset, "InTimeTec Visionsoft Pvt. Ltd.")
        # Registered Office
        canvas.setFont("Helvetica", 9)
        canvas.drawCentredString(width / 2, y_offset - 14, "Registered Office : Plot No. 1 / 2, Kanakpura Industrial  Area, Kanakpura, Sirsi Road, Jaipur-302034")
        # Branch Office
        canvas.drawCentredString(width / 2, y_offset - 26, "Branch Office : Site No. 271, Sri Ganesha Complex, Hosur Main Road, Madiwala, BTM 1st Stage, Bangalore- 560068")
        # Email, Website, Office No.
        canvas.drawCentredString(width / 2, y_offset - 38, "E-mail : info@intimetec.com, Website : www.intimetec.com, Office No. : 77378 53360")
        # Horizontal line
        canvas.setLineWidth(1)
        canvas.line(40, y_offset - 48, width - 40, y_offset - 48)
        canvas.restoreState()

    def create_pdf(self, content, output_path=None, in_memory=False):
        """
        Create a PDF document from extracted Word content.
        If in_memory is True, return a BytesIO object instead of saving to disk.
        """
        try:
            buffer = io.BytesIO() if in_memory else None
            doc = SimpleDocTemplate(buffer if in_memory else output_path, pagesize=A4)
            story = []
            
            # Add spacer to push content below header
            story.append(Spacer(1, 120))
            
            # Add title if available
            if content.get('metadata', {}).get('title'):
                title = Paragraph(content['metadata']['title'], self.styles['CustomTitle'])
                story.append(title)
                story.append(Spacer(1, 12))
            
            # Add author if available
            if content.get('metadata', {}).get('author'):
                author = Paragraph(f"Author: {content['metadata']['author']}", self.styles['CustomNormal'])
                story.append(author)
                story.append(Spacer(1, 6))
            
            # Process paragraphs
            for para_data in content.get('paragraphs', []):
                text = para_data['text']
                style_name = para_data.get('style', 'Normal')
                
                # Map Word styles to PDF styles
                if 'heading' in style_name.lower() or 'title' in style_name.lower():
                    pdf_style = self.styles['CustomHeading']
                else:
                    pdf_style = self.styles['CustomNormal']
                
                # Apply formatting from runs if available
                if para_data.get('runs'):
                    formatted_text = self._apply_run_formatting(para_data['runs'])
                    paragraph = Paragraph(formatted_text, pdf_style)
                else:
                    paragraph = Paragraph(text, pdf_style)
                
                story.append(paragraph)
                story.append(Spacer(1, 6))
            
            # Process tables
            for table_data in content.get('tables', []):
                if table_data.get('rows'):
                    table_story = self._create_table(table_data)
                    story.extend(table_story)
                    story.append(Spacer(1, 12))
            
            # Build PDF with header on every page
            doc.build(story, onFirstPage=self._draw_header, onLaterPages=self._draw_header)
            
            if in_memory:
                buffer.seek(0)
                return buffer
        except Exception as e:
            raise Exception(f"Error creating PDF: {str(e)}")
    
    def _apply_run_formatting(self, runs):
        """
        Apply formatting from Word runs to PDF text
        """
        formatted_text = ""
        
        for run in runs:
            text = run.get('text', '')
            
            # Apply bold formatting
            if run.get('bold'):
                text = f"<b>{text}</b>"
            
            # Apply italic formatting
            if run.get('italic'):
                text = f"<i>{text}</i>"
            
            # Apply underline formatting
            if run.get('underline'):
                text = f"<u>{text}</u>"
            
            formatted_text += text
        
        return formatted_text
    
    def _create_table(self, table_data):
        """
        Create a PDF table from table data with improved styling
        """
        story = []
        table_rows = []
        for row_data in table_data.get('rows', []):
            row = []
            for cell_data in row_data.get('cells', []):
                cell_text = ""
                for para in cell_data.get('paragraphs', []):
                    cell_text += para.get('text', '') + " "
                row.append(cell_text.strip())
            table_rows.append(row)
        if table_rows:
            table = Table(table_rows)
            # Improved styling: yellow header, blue highlights, bold totals, grid, etc.
            style = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#ffe599')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f3f6fa')),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]
            # Bold and color for total rows (example: last row)
            if len(table_rows) > 1:
                style.append(('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'))
                style.append(('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#fff2cc')))
            table.setStyle(TableStyle(style))
            story.append(table)
        return story
    
    def create_simple_pdf(self, text_content, output_path):
        """
        Create a simple PDF with just text content
        """
        try:
            doc = SimpleDocTemplate(output_path, pagesize=A4)
            story = []
            
            # Split text into paragraphs
            paragraphs = text_content.split('\n\n')
            
            for para in paragraphs:
                if para.strip():
                    paragraph = Paragraph(para.strip(), self.styles['CustomNormal'])
                    story.append(paragraph)
                    story.append(Spacer(1, 6))
            
            doc.build(story)
            
        except Exception as e:
            raise Exception(f"Error creating simple PDF: {str(e)}")

    def extract_candidate_name(self, content):
        """
        Try to extract the candidate's name from the content (e.g., from the first paragraph after 'Dear').
        """
        for para in content.get('paragraphs', []):
            if para['text'].strip().lower().startswith('dear'):
                # e.g., 'Dear Rohit Agarwal,'
                parts = para['text'].split(' ', 1)
                if len(parts) > 1:
                    name_part = parts[1].replace(',', '').strip()
                    return name_part
        return None 